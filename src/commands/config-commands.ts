import { maskSecrets } from '../utils/config.js';
import type { Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
export function runConfig(ctx: Context, args: string[]): void {
  if (args.length === 1 && args[0] === 'show') { process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}\n`); return; }
  usage("Use 'teams config show'.");
}
