import { maskSecrets, saveConfig } from '../utils/config.js';
import type { Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { resolveChannel, resolveTeam } from './team-commands.js';
export async function runConfig(ctx: Context, args: string[]): Promise<void> {
  if (args.length === 1 && args[0] === 'show') { process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}\n`); return; }
  if (args.length === 3 && args[0] === 'set' && args[1] === 'team') {
    const team = await resolveTeam(ctx, args[2]!);
    ctx.profile.defaultTeam = team.id;
    ctx.config.profiles[ctx.profileName] = ctx.profile;
    ctx.config.defaultProfile ??= ctx.profileName;
    saveConfig(ctx.path, ctx.config);
    process.stdout.write(`Default Team set to ${team.displayName ?? team.id}.\n`);
    return;
  }
  if (args.length === 3 && args[0] === 'set' && args[1] === 'channel') {
    const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    if (!teamValue) usage("Choose a Team with --team, TEAMS_TEAM, or 'teams config set team' first.");
    const team = await resolveTeam(ctx, teamValue);
    const channel = await resolveChannel(ctx, team.id!, args[2]!);
    ctx.profile.defaultChannel = channel.id;
    ctx.config.profiles[ctx.profileName] = ctx.profile;
    ctx.config.defaultProfile ??= ctx.profileName;
    saveConfig(ctx.path, ctx.config);
    process.stdout.write(`Default Channel set to ${channel.displayName ?? channel.id}.\n`);
    return;
  }
  usage("Use 'teams config show', 'teams config set team <name|id>', or 'teams config set channel <name|id>'.");
}
