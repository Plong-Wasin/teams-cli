import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { resolveChannel, resolveTeam } from './team-commands.js';
import { usage } from '../utils/errors.js';

interface User { id?: string; displayName?: string }
interface Mention {
  id: number;
  mentionText: string;
  mentioned: { user: { id: string; displayName: string; userIdentityType: 'aadUser' } };
}

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\r\n|\r|\n/g, '<br>');
}

async function stdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}

export async function runSend(ctx: Context, args: string[]): Promise<void> {
  const positional: string[] = []; const mentions: string[] = [];
  let bodyOption: string | undefined; let subject: string | undefined; let html = false; let dryRun = false;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i]!;
    if (arg === '--html') { html = true; continue; }
    if (arg === '--dry-run') { dryRun = true; continue; }
    if (arg === '--body' || arg.startsWith('--body=')) {
      if (bodyOption !== undefined) usage("Option '--body' may only be provided once.");
      bodyOption = arg === '--body' ? args[++i] : arg.slice(7);
      if (bodyOption === undefined) usage("Option '--body' requires text, @file, or '-'.");
      continue;
    }
    if (arg === '--subject' || arg.startsWith('--subject=')) {
      if (subject !== undefined) usage("Option '--subject' may only be provided once.");
      subject = arg === '--subject' ? args[++i] : arg.slice(10);
      if (subject === undefined) usage("Option '--subject' requires a value.");
      continue;
    }
    if (arg === '--mention' || arg.startsWith('--mention=')) {
      const value = arg === '--mention' ? args[++i] : arg.slice(10);
      if (!value) usage("Option '--mention' requires a user principal name or email.");
      mentions.push(value);
      continue;
    }
    if (arg.startsWith('-') && arg !== '-') usage(`Unknown send option: ${arg}. To post text that starts with '-', pass it with --body.`);
    positional.push(arg);
  }
  if (bodyOption !== undefined && positional.length) usage("Provide the message body as an argument or with '--body', not both.");
  if (bodyOption === undefined && positional.length === 0) usage("Use 'teams send <text> [--body @file|-] [--subject text] [--mention upn] [--html] [--dry-run]'.");
  const source = bodyOption ?? positional.join(' ');
  const raw = source === '-' ? await stdin() : source.startsWith('@') ? await readFile(source.slice(1), 'utf8') : source;
  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
  if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
  if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
  const team = await resolveTeam(ctx, teamValue);
  const teamId = team.id ?? '';
  const channel = await resolveChannel(ctx, teamId, channelValue);
  const payload: { subject?: string; body: { contentType: 'html'; content: string }; mentions?: Mention[] } = {
    ...(subject === undefined ? {} : { subject }),
    body: { contentType: 'html', content: html ? raw : escapeHtml(raw) },
  };
  if (mentions.length) {
    const resolved: Mention[] = [];
    const tags: string[] = [];
    for (const upn of mentions) {
      const user = await graphRequest(ctx, 'GET', `/users/${encodeURIComponent(upn)}`) as User;
      const userId = user.id ?? '';
      const displayName = user.displayName ?? upn;
      const id = resolved.length;
      resolved.push({ id, mentionText: displayName, mentioned: { user: { id: userId, displayName, userIdentityType: 'aadUser' } } });
      tags.push(`<at id="${id}">${escapeHtml(displayName)}</at>`);
    }
    payload.body.content = `${tags.join(' ')}${raw ? ` ${payload.body.content}` : ''}`;
    payload.mentions = resolved;
  }
  if (dryRun) {
    process.stdout.write(`${JSON.stringify(payload, null, 2)}\n`);
    return;
  }
  const result = await graphRequest(ctx, 'POST', `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channel.id ?? '')}/messages`, { body: JSON.stringify(payload) });
  process.stdout.write(`${JSON.stringify(result)}\n`);
}
