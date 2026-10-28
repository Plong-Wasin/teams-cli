import { HELP } from './help.js';
import { VERSION } from './utils/version.js';
import { runAuth, selectedContext, type Options } from './commands/auth-commands.js';
import { runConfig } from './commands/config-commands.js';
import { runApi } from './commands/api-commands.js';
import { runSearch } from './commands/search-commands.js';
import { graphRequest } from './commands/auth-commands.js';
import { CliError, usage } from './utils/errors.js';
import { runTeamsCommand } from './commands/team-commands.js';
import { runMessages } from './commands/message-commands.js';
import { runSend } from './commands/send-command.js';
import { runThread } from './commands/thread-command.js';
import { runEdit, runReply } from './commands/reply-edit-command.js';

function globals(argv: string[]): { options: Options; args: string[] } {
  const options: Options = {}; const args: string[] = [];
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]!;
    if (arg === '-p' || arg === '--profile' || arg === '--client-id' || arg === '--team' || arg === '--channel') {
      const value = argv[i + 1]; if (!value || value.startsWith('-')) usage(`Option '${arg}' requires a value.`);
      if (arg === '--client-id') options.clientId = value;
      else if (arg === '--team') options.team = value;
      else if (arg === '--channel') options.channel = value;
      else options.profile = value;
      i += 1;
    } else if (arg.startsWith('--profile=')) options.profile = arg.slice(10);
    else if (arg.startsWith('--client-id=')) options.clientId = arg.slice(12);
    else if (arg.startsWith('--team=')) options.team = arg.slice(7);
    else if (arg.startsWith('--channel=')) options.channel = arg.slice(10);
    else args.push(arg);
  }
  return { options, args };
}
async function main(argv: string[]): Promise<void> {
  if (!argv.length || argv.includes('--help') || argv.includes('-h')) { process.stdout.write(`${HELP}\n`); return; }
  if (argv.length === 1 && (argv[0] === '--version' || argv[0] === '-v')) { process.stdout.write(`${VERSION}\n`); return; }
  const { options, args } = globals(argv); const [command, sub, ...rest] = args;
  if (!command) usage("Run 'teams --help' for usage.");
  const ctx = selectedContext(process.env, options);
  if (command === 'auth') return runAuth(ctx, sub);
  if (command === 'config') return runConfig(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'api') return runApi(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'teams' || command === 'channels' || command === 'channel') return runTeamsCommand(ctx, command, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'search') return runSearch(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'messages') return runMessages(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'send') return runSend(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'reply') return runReply(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'edit') return runEdit(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'thread') return runThread(ctx, [sub, ...rest].filter((x): x is string => x !== undefined));
  if (command === 'me') {
    if (sub && sub !== '--json') usage("Use 'teams me [--json]'.");
    const identity = await graphRequest(ctx, 'GET', '/me') as { displayName?: string; userPrincipalName?: string; mail?: string; id?: string };
    if (sub === '--json') process.stdout.write(`${JSON.stringify(identity, null, 2)}\n`);
    else process.stdout.write(`NAME  EMAIL\n${identity.displayName ?? 'Current user'}  ${identity.mail ?? identity.userPrincipalName ?? ''}\n`);
    return;
  }
  throw new CliError(`Unknown command: ${command}. Run 'teams --help' for usage.`, 2);
}
main(process.argv.slice(2)).catch((error: unknown) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  process.stderr.write(`teams: ${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = code;
});
