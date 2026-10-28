import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';

interface Message {
  id?: string;
  replyToId?: string | null;
  createdDateTime?: string;
  from?: { user?: { displayName?: string } };
  body?: { content?: string };
}
interface Collection<T> { value?: T[]; '@odata.nextLink'?: string }
interface ThreadLink { teamId: string; channelId: string; messageId: string; parentId?: string }

const DEFAULT_TOP = 20;
const ID = /^[A-Za-z0-9._~=-]{1,512}$/;

function validId(value: string): boolean { return ID.test(value); }

function parseLink(value: string): ThreadLink | undefined {
  let url: URL;
  try { url = new URL(value); } catch { return undefined; }
  if (url.protocol !== 'https:' || url.hostname !== 'teams.microsoft.com' || url.port || url.username || url.password || url.hash) return undefined;
  const match = /^\/l\/message\/([^/]+)\/([^/]+)$/.exec(url.pathname);
  if (!match) return undefined;
  let channelId: string; let messageId: string;
  try { channelId = decodeURIComponent(match[1]!); messageId = decodeURIComponent(match[2]!); } catch { return undefined; }
  const groupIds = url.searchParams.getAll('groupId');
  const teamId = groupIds.length === 1 ? groupIds[0]! : '';
  if (!validId(teamId) || !validId(channelId) || !validId(messageId)) return undefined;
  const parents = url.searchParams.getAll('parentMessageId');
  if (parents.length > 1 || (parents[0] !== undefined && !validId(parents[0]))) return undefined;
  return { teamId, channelId, messageId, ...(parents[0] && parents[0] !== messageId ? { parentId: parents[0] } : {}) };
}

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

function validateContinuation(link: string, ctx: Context, expectedPath: string): void {
  const base = ctx.env.TEAMS_GRAPH_URL || 'https://graph.microsoft.com/v1.0';
  let url: URL;
  try { url = new URL(link, base); } catch { usage('Graph returned an invalid continuation URL.'); }
  const configured = new URL(base);
  const basePath = configured.pathname.replace(/\/$/, '');
  const expectedUrl = new URL(`${base.replace(/\/$/, '')}${expectedPath}`, base);
  if (url.origin !== configured.origin || (basePath && url.pathname !== basePath && !url.pathname.startsWith(`${basePath}/`)) || url.pathname !== expectedUrl.pathname) {
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
  const teamValue = link?.teamId || ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  const channelValue = link?.channelId || ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
  if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
  if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");

  const topValue = option(args, '--top');
  const top = topValue === undefined ? DEFAULT_TOP : Number(topValue);
  if (!Number.isSafeInteger(top) || top < 1 || top > 1000) usage("Option '--top' must be an integer from 1 to 1000.");
  const cursor = option(args, '--cursor');
  const base = `/teams/${encodeURIComponent(teamValue)}/channels/${encodeURIComponent(channelValue)}/messages`;
  // Graph cannot fetch a Reply by its own id, so the Thread is read from its Root Message id: the id given,
  // or the parentMessageId carried by a pasted Reply link.
  const rootId = link?.parentId ?? id;
  let root: Message;
  try { root = await graphRequest(ctx, 'GET', `${base}/${encodeURIComponent(rootId)}`) as Message; }
  catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) {
      usage('The id is not a Root Message in this Channel. A Reply id cannot be read on its own: pass its Root Message id, or paste a Reply link that carries parentMessageId.');
    }
    throw error;
  }

  let repliesUrl: string;
  if (cursor) {
    if (!/^[A-Za-z0-9_-]+$/.test(cursor)) usage('Invalid cursor. Copy the cursor from the previous command output.');
    const decoded = Buffer.from(cursor, 'base64url').toString('utf8');
    validateContinuation(decoded, ctx, `${base}/${encodeURIComponent(rootId)}/replies`);
    repliesUrl = decoded!;
  } else {
    repliesUrl = `${base}/${encodeURIComponent(rootId)}/replies?$top=${top}&$orderby=createdDateTime%20desc`;
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
