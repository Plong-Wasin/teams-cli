import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { resolveChannel, resolveTeam } from './team-commands.js';
import { CliError, usage } from '../utils/errors.js';

interface User { id?: string; displayName?: string }
interface Message { id?: string; replyToId?: string | null }
interface Mention {
  id: number;
  mentionText: string;
  mentioned: { user: { id: string; displayName: string; userIdentityType: 'aadUser' } };
}
interface Target { teamId: string; channelId: string; messageId: string }
type Payload = { body: { contentType: 'html'; content: string }; mentions?: Mention[] };

const ID = /^[A-Za-z0-9._~=-]{1,512}$/;

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\r\n|\r|\n/g, '<br>');
}

function parseLink(value: string): Target | undefined {
  let url: URL;
  try { url = new URL(value); } catch { return undefined; }
  if (url.protocol !== 'https:' || url.hostname !== 'teams.microsoft.com' || url.port || url.username || url.password || url.hash) return undefined;
  const match = /^\/l\/message\/([^/]+)\/([^/]+)$/.exec(url.pathname);
  if (!match) return undefined;
  let channelId: string; let messageId: string;
  try { channelId = decodeURIComponent(match[1]!); messageId = decodeURIComponent(match[2]!); } catch { return undefined; }
  const groupIds = url.searchParams.getAll('groupId');
  const teamId = groupIds.length === 1 ? groupIds[0]! : '';
  if (!ID.test(teamId) || !ID.test(channelId) || !ID.test(messageId)) return undefined;
  return { teamId, channelId, messageId };
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
    if (arg.startsWith('-') && arg !== '-') usage(`Unknown message option: ${arg}.`);
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

async function resolveTarget(ctx: Context, supplied: string): Promise<{ path: string; rootId: string; selectedId: string }> {
  const link = parseLink(supplied);
  if (/^https?:\/\//i.test(supplied) && !link) usage('Expected a valid Teams Channel Message link; Chat links cannot be written.');
  const selectedId = link?.messageId ?? parseId(supplied);
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
  let selected: Message;
  try { selected = await graphRequest(ctx, 'GET', `${path}/${encodeURIComponent(selectedId)}`) as Message; }
  catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) usage('The id is not a Channel Message (Chat messages cannot be written).');
    throw error;
  }
  return { path, rootId: selected.replyToId || selected.id || selectedId, selectedId };
}

export async function runReply(ctx: Context, args: string[]): Promise<void> {
  if (args.length < 2) usage("Use 'teams reply <id|link> <text> [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, args[0]!);
  const body = await payload(ctx, args.slice(1), true);
  if (args.includes('--dry-run')) return;
  const result = await graphRequest(ctx, 'POST', `${target.path}/${encodeURIComponent(target.rootId)}/replies`, { body: JSON.stringify(body) });
  process.stdout.write(`${JSON.stringify(result)}\n`);
}

export async function runEdit(ctx: Context, args: string[]): Promise<void> {
  if (args.length < 2) usage("Use 'teams edit <id|link> <text> [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, args[0]!);
  const body = await payload(ctx, args.slice(1), true);
  if (args.includes('--dry-run')) return;
  const result = await graphRequest(ctx, 'PATCH', `${target.path}/${encodeURIComponent(target.selectedId)}`, { body: JSON.stringify(body) });
  process.stdout.write(`${JSON.stringify(result)}\n`);
}
