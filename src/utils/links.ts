export interface MessageLink { teamId: string; channelId: string; messageId: string; parentId?: string }

// Message ids are numeric strings and team ids GUIDs; channel ids look like 19:<hash>@thread.tacv2, so they may carry ':' and '@'.
const TEAMS_HOSTS = new Set(['teams.microsoft.com', 'teams.cloud.microsoft']);
export const ID = /^[A-Za-z0-9._~=-]{1,512}$/;
export const CHANNEL_ID = /^[A-Za-z0-9._~=:@-]{1,512}$/;

export function parseLink(value: string): MessageLink | undefined {
  let url: URL;
  try { url = new URL(value); } catch { return undefined; }
  if (url.protocol !== 'https:' || !TEAMS_HOSTS.has(url.hostname) || url.port || url.username || url.password || url.hash) return undefined;
  const match = /^\/l\/message\/([^/]+)\/([^/]+)$/.exec(url.pathname);
  if (!match) return undefined;
  let channelId: string; let messageId: string;
  try { channelId = decodeURIComponent(match[1]!); messageId = decodeURIComponent(match[2]!); } catch { return undefined; }
  const groupIds = url.searchParams.getAll('groupId');
  const teamId = groupIds.length === 1 ? groupIds[0]! : '';
  if (!ID.test(teamId) || !CHANNEL_ID.test(channelId) || !ID.test(messageId)) return undefined;
  const parents = url.searchParams.getAll('parentMessageId');
  if (parents.length > 1 || (parents[0] !== undefined && !ID.test(parents[0]))) return undefined;
  return { teamId, channelId, messageId, ...(parents[0] && parents[0] !== messageId ? { parentId: parents[0] } : {}) };
}
