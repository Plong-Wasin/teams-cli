import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { ID, parseLink } from '../utils/links.js';
import { resolveChannel, resolveTeam } from './team-commands.js';

interface Message {
  id?: string;
  replyToId?: string | null;
  createdDateTime?: string;
  from?: { user?: { displayName?: string } };
  body?: { content?: string };
}
interface Collection<T> { value?: T[]; '@odata.nextLink'?: string }

const DEFAULT_TOP = 20;

function validId(value: string): boolean { return ID.test(value); }

function messageId(value: string): string {
  const link = parseLink(value);
  if (link) return link.messageId;
  if (/^https?:\/\//i.test(value) || !validId(value)) usage('Expected a valid Channel Message id or a Teams message link.');
  return value;
}

function option(args: string[], name: string): string | undefined {
  const index = args.indexOf(name);
  if (index < 0) return undefined;
  const value = args[index + 1];
  if (!value || value.startsWith('-')) usage(`Option '${name}' requires a value.`);
  return value;
}

function validateContinuation(link: string, ctx: Context): void {
  const base = ctx.env.TEAMS_GRAPH_URL || 'https://graph.microsoft.com/v1.0';
  let url: URL;
  try { url = new URL(link, base); } catch { usage('Graph returned an invalid continuation URL.'); }
  const configured = new URL(base);
  const basePath = configured.pathname.replace(/\/$/, '');
  if (url.origin !== configured.origin || (basePath && url.pathname !== basePath && !url.pathname.startsWith(`${basePath}/`))) {
    usage('Graph returned a continuation URL outside the configured Graph URL.');
  }
}

function printContinuation(link: string | undefined): void {
  if (link) process.stderr.write(`more available: --cursor ${Buffer.from(link).toString('base64url')}\n`);
}

function plainText(content: string | undefined): string {
  return (content ?? '').replace(/<br\s*\/?\s*>/gi, ' ').replace(/<\/[^>]+>/g, ' ').replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'").replace(/\s+/g, ' ').trim();
}

function table(rows: string[][]): string {
  const headers = ['ID', 'TIME', 'SENDER', 'BODY'];
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}\n`;
}

export async function runThread(ctx: Context, args: string[]): Promise<void> {
  const values = new Set(['--top', '--cursor']);
  const flags = new Set(['--full', '--json']);
  const positional: string[] = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (values.has(arg)) { option(args, arg); index += 1; }
    else if (flags.has(arg)) continue;
    else if (arg.startsWith('-')) usage(`Unknown thread option '${arg}'.`);
    else positional.push(arg);
  }
  if (positional.length !== 1) usage("Use 'teams thread <id|link> [--top N] [--cursor token] [--full] [--json]'.");

  const supplied = positional[0]!;
  const link = parseLink(supplied);
  if (/^https?:\/\//i.test(supplied) && !link) usage('Expected a valid Teams message link.');
  const id = link?.messageId ?? messageId(supplied);
  let teamValue = link?.teamId ?? ''; let channelValue = link?.channelId ?? '';
  if (!link) {
    const teamName = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    const channelName = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!teamName) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!channelName) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    const team = await resolveTeam(ctx, teamName);
    teamValue = team.id ?? '';
    channelValue = (await resolveChannel(ctx, teamValue, channelName)).id ?? '';
  }

  const topValue = option(args, '--top');
  const top = topValue === undefined ? DEFAULT_TOP : Number(topValue);
  if (!Number.isSafeInteger(top) || top < 1 || top > 50) usage("Option '--top' must be an integer from 1 to 50.");
  const cursor = option(args, '--cursor');
  const base = `/teams/${encodeURIComponent(teamValue)}/channels/${encodeURIComponent(channelValue)}/messages`;
  // Graph cannot fetch a Reply by its own id, so the Thread is read from its Root Message id: the id given,
  // or the parentMessageId carried by a pasted Reply link.
  const rootId = link?.parentId ?? id;
  let root: Message;
  try { root = await graphRequest(ctx, 'GET', `${base}/${encodeURIComponent(rootId)}`) as Message; }
  catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) {
      usage('Not found: the id is not a Root Message in this Channel, the Team or Channel is wrong, or the message was deleted. A Reply id cannot be read on its own: pass its Root Message id, or paste a Reply link that carries parentMessageId.');
    }
    throw error;
  }

  let repliesUrl: string;
  if (cursor) {
    if (!/^[A-Za-z0-9_-]+$/.test(cursor)) usage('Invalid cursor. Copy the cursor from the previous command output.');
    const decoded = Buffer.from(cursor, 'base64url').toString('utf8');
    validateContinuation(decoded, ctx);
    repliesUrl = decoded!;
  } else {
    repliesUrl = `${base}/${encodeURIComponent(rootId)}/replies?$top=${top}`;
  }
  const pages: Collection<Message>[] = [];
  let nextLink: string | undefined;
  const result = await graphRequest(ctx, 'GET', repliesUrl, /^https?:\/\//i.test(repliesUrl) ? { rawUrl: repliesUrl } : {}) as Collection<Message>;
  pages.push(result);
  nextLink = result['@odata.nextLink'];
  const pageReplies = (pages[0]?.value ?? []).slice().sort((a, b) => (a.createdDateTime ?? '').localeCompare(b.createdDateTime ?? ''));
  const response = { root, replies: pageReplies, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) };
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }
  const all = [root, ...pageReplies];
  const rows = all.map((message) => [message.id ?? '', message.createdDateTime ?? '', message.from?.user?.displayName ?? 'Unknown', plainText(message.body?.content)]);
  process.stdout.write(table(rows));
  printContinuation(nextLink);
}
