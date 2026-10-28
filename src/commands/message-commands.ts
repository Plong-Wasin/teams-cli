import { graphRequest, type Context } from './auth-commands.js';
import { resolveChannel, resolveTeam } from './team-commands.js';
import { usage } from '../utils/errors.js';

interface Message {
  id?: string;
  createdDateTime?: string;
  from?: { user?: { displayName?: string; id?: string } };
  body?: { content?: string; contentType?: string };
  replies?: Message[];
  [key: string]: unknown;
}
interface Collection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

const DEFAULT_TOP = 20;
const PAGE_CAP = 100;
const BODY_WIDTH = 80;
const REPLY_LIMIT = 5;

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}

function option(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  if (index < 0) return undefined;
  const value = args[index + 1];
  if (!value || value.startsWith('-')) usage(`Option '${name}' requires a value.`);
  return value;
}

function plainText(content: string | undefined): string {
  return (content ?? '')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<\/(p|div|li|tr)>/gi, ' ')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function truncate(value: string, max = BODY_WIDTH): string {
  return value.length > max ? `${value.slice(0, max - 1)}…` : value;
}

function validateContinuation(link: string, ctx: Context): void {
  const base = ctx.env.TEAMS_GRAPH_URL || 'https://graph.microsoft.com/v1.0';
  let url: URL;
  try { url = new URL(link, base); }
  catch { usage('Graph returned an invalid continuation URL.'); }
  const configured = new URL(base);
  const basePath = configured.pathname.replace(/\/$/, '');
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== configured.origin || !inBasePath) usage('Graph returned a continuation URL outside the configured Graph URL.');
}

function cursorUrl(token: string, ctx: Context): string {
  let link: string;
  try { link = Buffer.from(token, 'base64url').toString('utf8'); }
  catch { usage('Invalid cursor. Copy the cursor from the previous command output.'); }
  validateContinuation(link, ctx);
  return link;
}

function graphPage(ctx: Context, url: string): Promise<unknown> {
  return /^https?:\/\//i.test(url)
    ? graphRequest(ctx, 'GET', url, { rawUrl: url })
    : graphRequest(ctx, 'GET', url);
}

function printContinuation(link: string | undefined): void {
  if (link) process.stderr.write(`more available: --cursor ${Buffer.from(link).toString('base64url')} (or --all)\n`);
}

function parseTop(args: string[]): number {
  const value = option(args, '--top');
  if (value === undefined) return DEFAULT_TOP;
  const top = Number(value);
  if (!Number.isSafeInteger(top) || top < 1 || top > 1000) usage("Option '--top' must be an integer from 1 to 1000.");
  return top;
}

export async function runMessages(ctx: Context, args: string[]): Promise<void> {
  const valueOptions = new Set(['--top', '--cursor']);
  const flags = new Set(['--all', '--full', '--json', '--replies']);
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (valueOptions.has(arg)) { option(args, arg); index += 1; }
    else if (!flags.has(arg)) usage(`Unknown messages option or argument '${arg}'.`);
  }

  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
  if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
  if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
  const team = await resolveTeam(ctx, teamValue);
  const channel = await resolveChannel(ctx, team.id!, channelValue);
  const path = `/teams/${encodeURIComponent(team.id!)}/channels/${encodeURIComponent(channel.id!)}/messages`;
  const top = parseTop(args);
  const cursor = option(args, '--cursor');
  let url = cursor ? cursorUrl(cursor, ctx) : `${path}?$top=${top}`;
  const all = args.includes('--all');
  const pages: Collection<Message>[] = [];
  let nextLink: string | undefined;
  for (let page = 0; page < (all ? PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url) as Collection<Message>;
    pages.push(result);
    nextLink = result['@odata.nextLink'];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const messages = pages.flatMap((page) => page.value ?? []);
  const response: Collection<Message> = all ? { ...first, value: messages, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } : first;
  if (all && !nextLink) delete response['@odata.nextLink'];

  const replyPreviews = new Map<string, string>();
  if (args.includes('--replies') && !args.includes('--json')) {
    for (const message of messages) {
      if (!message.id) continue;
      const replyData = await graphRequest(ctx, 'GET', `${path}/${encodeURIComponent(message.id)}/replies`, { query: new URLSearchParams({ '$top': String(REPLY_LIMIT) }) }) as Collection<Message>;
      const replies = (replyData.value ?? []).slice().sort((a, b) => (a.createdDateTime ?? '').localeCompare(b.createdDateTime ?? '')).slice(-REPLY_LIMIT);
      const preview = replies.map((reply) => `${reply.from?.user?.displayName ?? 'Unknown'}: ${plainText(reply.body?.content)}`).filter((value) => !value.endsWith(': ')).join(' | ');
      if (preview) replyPreviews.set(message.id, truncate(preview));
    }
  }

  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }
  const full = args.includes('--full');
  const rows = messages.map((message) => {
    const row = [message.id ?? '', message.createdDateTime ?? '', message.from?.user?.displayName ?? 'Unknown', full ? plainText(message.body?.content) : truncate(plainText(message.body?.content))];
    if (args.includes('--replies')) row.push(replyPreviews.get(message.id ?? '') ?? '');
    return row;
  });
  process.stdout.write(table(['ID', 'TIME', 'SENDER', 'BODY', ...(args.includes('--replies') ? ['REPLIES'] : [])], rows));
  printContinuation(nextLink);
}
