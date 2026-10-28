import { spawn } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const BIN = new URL('../../dist/teams.js', import.meta.url).pathname;
export interface Result { code: number | null; stdout: string; stderr: string }
export interface Cli {
  home: string;
  run(args: string[], env?: Record<string, string>, stdin?: string): Promise<Result>;
  cleanup(): void;
}
export function createCli(env: Record<string, string> = {}): Cli {
  const home = mkdtempSync(join(tmpdir(), 'teams-cli-test-'));
  return {
    home,
    run: (args, extra = {}, stdin) => new Promise((resolve, reject) => {
      const child = spawn(process.execPath, [BIN, ...args], {
        env: { PATH: process.env.PATH ?? '', HOME: home, TEAMS_CONFIG: join(home, 'config.json'), ...env, ...extra },
        stdio: [stdin === undefined ? 'ignore' : 'pipe', 'pipe', 'pipe'],
      });
      let stdout = ''; let stderr = '';
      child.stdout?.setEncoding('utf8').on('data', (s: string) => { stdout += s; });
      child.stderr?.setEncoding('utf8').on('data', (s: string) => { stderr += s; });
      child.on('error', reject);
      child.on('close', (code) => resolve({ code, stdout, stderr }));
      if (stdin !== undefined) { child.stdin?.write(stdin); child.stdin?.end(); }
    }),
    cleanup: () => rmSync(home, { recursive: true, force: true }),
  };
}
