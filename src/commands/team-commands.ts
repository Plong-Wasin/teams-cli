import type { Context } from './auth-commands.js';
import { graphRequest } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';

interface Named { id?: string; displayName?: string; name?: string; [key: string]: unknown }
interface Collection<T> { value?: T[]; [key: string]: unknown }
function label(item: Named): string { return item.displayName ?? item.name ?? ''; }
function id(item: Named): string { return item.id ?? ''; }
function resolve(items: Named[], value: string, kind: string): Named {
  const byId = items.find((item) => id(item) === value);
  if (byId) return byId;
  const exact = items.filter((item) => label(item) === value);
  const matches = exact.length ? exact : items.filter((item) => label(item).toLocaleLowerCase() === value.toLocaleLowerCase());
  if (matches.length === 1) return matches[0]!;
  const candidates = (matches.length ? matches : items).map((item) => `${label(item)} (${id(item)})`).join(', ');
  throw new CliError(`${matches.length ? 'Ambiguous' : 'No matching'} ${kind} '${value}'. Candidates: ${candidates || '(none)'}.`);
}
function table(items: Named[], columns: Array<[string, (item: Named) => string]>): string {
  const widths = columns.map(([title, value]) => Math.max(title.length, ...items.map((item) => value(item).length)));
  const row = (values: string[]) => values.map((value, i) => value.padEnd(widths[i]!)).join('  ').trimEnd();
  return `${row(columns.map(([title]) => title))}\n${items.map((item) => row(columns.map(([, value]) => value(item)))).join('\n')}${items.length ? '\n' : ''}`;
}
function output(data: unknown, json: boolean, columns: Array<[string, (item: Named) => string]>): void {
  if (json) process.stdout.write(`${JSON.stringify(data, null, 2)}\n`);
  else process.stdout.write(table((data as Collection<Named>).value ?? [], columns));
}
async function teams(ctx: Context, json: boolean): Promise<void> {
  const data = await graphRequest(ctx, 'GET', '/me/joinedTeams') as Collection<Named>;
  output(data, json, [['ID', id], ['NAME', label]]);
}
export async function resolveTeam(ctx: Context, value: string): Promise<Named> {
  const data = await graphRequest(ctx, 'GET', '/me/joinedTeams') as Collection<Named>;
  return resolve(data.value ?? [], value, 'Team');
}
export async function resolveChannel(ctx: Context, teamId: string, value: string): Promise<Named> {
  const data = await graphRequest(ctx, 'GET', `/teams/${encodeURIComponent(teamId)}/channels`) as Collection<Named>;
  return resolve(data.value ?? [], value, 'Channel');
}
async function channels(ctx: Context, teamValue: string, json: boolean): Promise<void> {
  const team = await resolveTeam(ctx, teamValue);
  const data = await graphRequest(ctx, 'GET', `/teams/${encodeURIComponent(id(team))}/channels`) as Collection<Named>;
  output(data, json, [['ID', id], ['NAME', label], ['TYPE', (item) => String(item.membershipType ?? '')]]);
}
async function channel(ctx: Context, teamValue: string, channelValue: string, json: boolean): Promise<void> {
  const team = await resolveTeam(ctx, teamValue);
  const chan = await resolveChannel(ctx, id(team), channelValue);
  const data = await graphRequest(ctx, 'GET', `/teams/${encodeURIComponent(id(team))}/channels/${encodeURIComponent(id(chan))}`) as Named;
  if (json) process.stdout.write(`${JSON.stringify(data, null, 2)}\n`);
  else process.stdout.write(`ID  NAME  TYPE  LAYOUT\n${id(data)}  ${label(data)}  ${String(data.membershipType ?? '')}  ${String(data.layout ?? '')}\n`);
}
export async function runTeamsCommand(ctx: Context, command: string, args: string[]): Promise<void> {
  const json = args.includes('--json');
  const positional = args.filter((arg) => arg !== '--json');
  if (command === 'teams' && positional.length === 0) return teams(ctx, json);
  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  if (command === 'channels') {
    if (positional.length > 1) usage("Use 'teams channels [team] [--json]'.");
    const selected = positional[0] || teamValue;
    if (!selected) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    return channels(ctx, selected, json);
  }
  if (command === 'channel') {
    if (positional.length > 2) usage("Use 'teams channel [channel] [--team team] [--json]'.");
    const selectedTeam = positional.length === 2 ? positional[0]! : teamValue;
    const selectedChannel = positional.length === 2 ? positional[1]! : positional[0] || ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!selectedTeam) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!selectedChannel) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    return channel(ctx, selectedTeam, selectedChannel, json);
  }
  throw new CliError(`Invalid ${command} arguments.` , 2);
}
