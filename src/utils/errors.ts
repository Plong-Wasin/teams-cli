export class CliError extends Error {
  constructor(message: string, readonly exitCode = 1) { super(message); }
}
export function usage(message: string): never { throw new CliError(message, 2); }
