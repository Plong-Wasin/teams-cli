import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

interface SearchHit {
  summary?: string;
  resource?: {
    id?: string;
    replyToId?: string | null;
    channelIdentity?: { teamId?: string; channelId?: string };
    from?: { user?: { displayName?: string } };
    createdDateTime?: string;
    body?: { content?: string };
  };
}

export async function runSearch(ctx: Context, args: string[]): Promise<void> {
  let size = 25;
  let json = false;
  const terms: string[] = [];
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i]!;
    if (arg === '--json') { json = true; continue; }
    if (arg === '--size' || arg.startsWith('--size=')) {
      const value = arg === '--size' ? args[++i] : arg.slice(7);
      if (!value || !/^\d+$/.test(value) || Number(value) < 1) usage("Option '--size' requires a positive integer.");
      size = Number(value);
      continue;
    }
    if (arg.startsWith('-')) usage(`Unknown search option: ${arg}.`);
    terms.push(arg);
  }
  const query = terms.join(' ').trim();
  if (!query) usage("Use 'teams search <query> [--size N] [--json]'.");

  const result = await graphRequest(ctx, 'POST', '/search/query', { body: JSON.stringify({
    requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: 0, size, enableTopResults: true }],
  }) });
  if (json) {
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    return;
  }
  const hits = searchHits(result);
  const rows = hits.map((hit) => {
    const resource = hit.resource ?? {};
    const teamId = resource.channelIdentity?.teamId;
    const channelId = resource.channelIdentity?.channelId;
    const kind = teamId && channelId ? 'Channel' : 'Chat';
    return [kind, teamId ?? '-', channelId ?? '-', resource.id ?? '-', resource.replyToId ?? '-', resource.from?.user?.displayName ?? '-', resource.createdDateTime ?? '-', plainText(resource.body?.content ?? hit.summary ?? '')];
  });
  process.stdout.write(`TYPE  TEAM  CHANNEL  MESSAGE  ROOT  SENDER  TIME  BODY\n`);
  for (const row of rows) process.stdout.write(`${row.map((value) => value.replaceAll('\n', ' ').slice(0, 80)).join('  ')}\n`);
}

function searchHits(value: unknown): SearchHit[] {
  if (!value || typeof value !== 'object') return [];
  const containers = (value as { value?: Array<{ hitsContainers?: Array<{ hits?: SearchHit[] }> }> }).value;
  return containers?.[0]?.hitsContainers?.[0]?.hits ?? [];
}

function plainText(value: string): string {
  return value.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
}
