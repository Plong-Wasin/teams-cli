import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { resolveChannel, resolveTeam } from './team-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { ID, parseLink } from '../utils/links.js';

interface User { id?: string; displayName?: string }
interface Message { id?: string; subject?: string | null }
interface Mention {
  id: number;
  mentionText: string;
  mentioned: { user: { id: string; displayName: string; userIdentityType: 'aadUser' } };
}
type Payload = { body: { contentType: 'html'; content: string }; mentions?: Mention[] };


function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\r\n|\r|\n/g, '<br>');
}

function parseId(value: string): string {
  const link = parseLink(value);
  if (link) return link.messageId;
  if (/^https?:\/\//i.test(value) || !ID.test(value)) usage('Expected a valid Channel Message id or a Teams message link.');
  return value;
}

async function stdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}

async function payload(ctx: Context, args: string[], allowHtml: boolean): Promise<Payload> {
  const positional: string[] = []; const mentions: string[] = [];
  let bodyOption: string | undefined; let html = false; let dryRun = false;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i]!;
    if (arg === '--html' && allowHtml) { html = true; continue; }
    if (arg === '--dry-run') { dryRun = true; continue; }
    if (arg === '--body' || arg.startsWith('--body=')) {
      if (bodyOption !== undefined) usage("Option '--body' may only be provided once.");
      bodyOption = arg === '--body' ? args[++i] : arg.slice(7);
      if (bodyOption === undefined) usage("Option '--body' requires text, @file, or '-'.");
      continue;
    }
    if (arg === '--mention' || arg.startsWith('--mention=')) {
      const value = arg === '--mention' ? args[++i] : arg.slice(10);
      if (!value) usage("Option '--mention' requires a user principal name or email.");
      mentions.push(value);
      continue;
    }
    if (arg.startsWith('-') && arg !== '-') usage(`Unknown message option: ${arg}. To post text that starts with '-', pass it with --body.`);
    positional.push(arg);
  }
  if (bodyOption !== undefined && positional.length) usage("Provide the message body as an argument or with '--body', not both.");
  if (bodyOption === undefined && positional.length === 0) usage("Provide message text as an argument or with '--body @file|-'.");
  const source = bodyOption ?? positional.join(' ');
  const raw = source === '-' ? await stdin() : source.startsWith('@') ? await readFile(source.slice(1), 'utf8') : source;
  const result: Payload = { body: { contentType: 'html', content: html ? raw : escapeHtml(raw) } };
  if (mentions.length) {
    const resolved: Mention[] = []; const tags: string[] = [];
    for (const upn of mentions) {
      const user = await graphRequest(ctx, 'GET', `/users/${encodeURIComponent(upn)}`) as User;
      const id = resolved.length; const userId = user.id ?? ''; const displayName = user.displayName ?? upn;
      resolved.push({ id, mentionText: displayName, mentioned: { user: { id: userId, displayName, userIdentityType: 'aadUser' } } });
      tags.push(`<at id="${id}">${escapeHtml(displayName)}</at>`);
    }
    result.body.content = `${tags.join(' ')}${raw ? ` ${result.body.content}` : ''}`;
    result.mentions = resolved;
  }
  if (dryRun) process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  return result;
}

interface Resolved { path: string; rootId: string; selectedId: string; isReply: boolean; subject?: string }

// Graph cannot fetch a Reply by its own id (GET .../messages/{replyId} returns 404), so a Reply is only
// addressable together with its Root Message id: `--root`, or `parentMessageId` in a pasted Teams link.
async function resolveTarget(ctx: Context, supplied: string, rootOption: string | undefined): Promise<Resolved> {
  const link = parseLink(supplied);
  if (/^https?:\/\//i.test(supplied) && !link) usage('Expected a valid Teams Channel Message link; Chat links cannot be written.');
  const selectedId = link?.messageId ?? parseId(supplied);
  if (rootOption !== undefined && !ID.test(rootOption)) usage("Option '--root' requires a valid Root Message id.");
  const rootHint = rootOption ?? link?.parentId;
  let teamId: string; let channelId: string;
  if (link) {
    teamId = link.teamId;
    channelId = link.channelId;
  } else {
    const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    const team = await resolveTeam(ctx, teamValue);
    const channel = await resolveChannel(ctx, team.id ?? '', channelValue);
    teamId = team.id ?? '';
    channelId = channel.id ?? '';
  }
  const path = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
  const isReply = rootHint !== undefined && rootHint !== selectedId;
  const rootId = isReply ? rootHint! : selectedId;
  let root: Message;
  try { root = await graphRequest(ctx, 'GET', `${path}/${encodeURIComponent(rootId)}`) as Message; }
  catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) {
      usage('Not found: the id is not a Root Message in this Channel (Chat messages cannot be written), the Team or Channel is wrong, or the message was deleted. To target a Reply, also pass its Root Message id with --root.');
    }
    throw error;
  }
  return { path, rootId, selectedId, isReply, ...(root.subject ? { subject: root.subject } : {}) };
}

function takeRoot(args: string[]): { rest: string[]; root: string | undefined } {
  const rest: string[] = []; let root: string | undefined;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i]!;
    if (arg === '--root' || arg.startsWith('--root=')) {
      if (root !== undefined) usage("Option '--root' may only be provided once.");
      root = arg === '--root' ? args[++i] : arg.slice(7);
      if (!root) usage("Option '--root' requires a Root Message id.");
    } else rest.push(arg);
  }
  return { rest, root };
}

export async function runReply(ctx: Context, args: string[]): Promise<void> {
  const { rest, root } = takeRoot(args);
  if (rest.length < 2) usage("Use 'teams reply <id|link> <text> [--root <root id>] [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, rest[0]!, root);
  const body = await payload(ctx, rest.slice(1), true);
  if (rest.includes('--dry-run')) return;
  const result = await graphRequest(ctx, 'POST', `${target.path}/${encodeURIComponent(target.rootId)}/replies`, { body: JSON.stringify(body) });
  process.stdout.write(`${JSON.stringify(result)}\n`);
}

// Graph clears a Root Message's subject on any PATCH that omits it, so the current subject is re-sent.
export async function runEdit(ctx: Context, args: string[]): Promise<void> {
  const { rest, root } = takeRoot(args);
  if (rest.length < 2) usage("Use 'teams edit <id|link> <text> [--root <root id>] [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, rest[0]!, root);
  const body = await payload(ctx, rest.slice(1), true);
  if (rest.includes('--dry-run')) return;
  const endpoint = target.isReply
    ? `${target.path}/${encodeURIComponent(target.rootId)}/replies/${encodeURIComponent(target.selectedId)}`
    : `${target.path}/${encodeURIComponent(target.selectedId)}`;
  const patch = target.isReply || !target.subject ? body : { subject: target.subject, ...body };
  const result = await graphRequest(ctx, 'PATCH', endpoint, { body: JSON.stringify(patch) });
  process.stdout.write(result === undefined ? 'Edited.\n' : `${JSON.stringify(result)}\n`);
}
