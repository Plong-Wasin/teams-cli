import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

export async function runApi(ctx: Context, args: string[]): Promise<void> {
  const method = args[0]?.toUpperCase(); const path = args[1];
  if (!method || !path) usage("Use 'teams api <METHOD> <path> [key=value ...]'.");
  if (!['GET', 'POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);
  const query = new URLSearchParams(); const headers: Record<string, string> = {}; let body: string | undefined; let all = false;
  for (let i = 2; i < args.length; i += 1) {
    const arg = args[i]!;
    if (arg === '--all') { all = true; continue; }
    if (arg === '--header' || arg.startsWith('--header=')) {
      const value = arg === '--header' ? args[++i] : arg.slice(9); const colon = value?.indexOf(':') ?? -1;
      if (!value || colon < 1) usage("Option '--header' requires 'Name: value'.");
      headers[value.slice(0, colon).trim()] = value.slice(colon + 1).trim(); continue;
    }
    if (arg === '--data' || arg.startsWith('--data=')) {
      if (body !== undefined) usage("Option '--data' may only be provided once.");
      const value = arg === '--data' ? args[++i] : arg.slice(7); if (value === undefined) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = value === '-' ? await readStdin() : value.startsWith('@') ? await readFile(value.slice(1), 'utf8') : value;
      try { body = JSON.stringify(JSON.parse(text)); } catch { usage("Option '--data' must contain valid JSON."); }
      continue;
    }
    const equal = arg.indexOf('='); if (equal < 1) usage(`Expected a query pair in key=value form: ${arg}.`);
    query.append(arg.slice(0, equal), arg.slice(equal + 1));
  }
  if (all && method !== 'GET') usage("Option '--all' is only supported with GET requests.");
  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, data);
  if (data !== undefined) process.stdout.write(`${JSON.stringify(data)}\n`);
}
async function collectPages(ctx: Context, initial: unknown): Promise<unknown> {
  if (!isCollection(initial)) return initial;
  const result: Record<string, unknown> = { ...initial, value: [...initial.value] }; let next = initial['@odata.nextLink'];
  for (let count = 1; typeof next === 'string' && count < 50; count += 1) {
    const url = validateLink(ctx, next); const page: unknown = await graphRequest(ctx, 'GET', url.href, { rawUrl: url.href });
    if (!isCollection(page)) break;
    result.value = [...result.value as unknown[], ...page.value]; next = page['@odata.nextLink'];
  }
  delete result['@odata.nextLink']; return result;
}
function isCollection(value: unknown): value is Record<string, unknown> & { value: unknown[]; '@odata.nextLink'?: string } {
  return typeof value === 'object' && value !== null && Array.isArray((value as { value?: unknown }).value);
}
function validateLink(ctx: Context, link: string): URL {
  const base = new URL(ctx.env.TEAMS_GRAPH_URL || 'https://graph.microsoft.com/v1.0'); let url: URL;
  try { url = new URL(link, base); } catch { usage('Graph returned an invalid continuation URL.'); }
  const prefix = base.pathname.replace(/\/$/, '');
  if (url.origin !== base.origin || (prefix && url.pathname !== prefix && !url.pathname.startsWith(`${prefix}/`))) usage('Graph returned a continuation URL outside the configured Graph URL.');
  return url;
}
async function readStdin(): Promise<string> { const chunks: Buffer[] = []; for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk)); return Buffer.concat(chunks).toString('utf8'); }
