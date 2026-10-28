import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers;
let cli: Cli;
const rootPath = '/teams/team-1/channels/chan-1/messages';
const rootMessage = { id: 'message-1', createdDateTime: '2026-10-28T10:00:00Z', from: { user: { displayName: 'Ada' } }, body: { content: '<p>Root text</p>' } };
const replies = [
  { id: 'reply-old', createdDateTime: '2026-10-28T10:01:00Z', from: { user: { displayName: 'Lin' } }, body: { content: '<p>Older reply</p>' } },
  { id: 'reply-new', createdDateTime: '2026-10-28T10:02:00Z', from: { user: { displayName: 'Sam' } }, body: { content: '<p>Newer reply</p>' } },
];

before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
  servers.graph.setReply((request, response) => {
    const url = new URL(request.path, 'http://graph.test');
    response.setHeader('Content-Type', 'application/json');
    if (url.pathname === '/me/joinedTeams') { response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] })); return; }
    if (url.pathname === '/teams/team-1/channels') { response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] })); return; }
    if (url.pathname === `${rootPath}/message-1`) { response.end(JSON.stringify(rootMessage)); return; }
    if (url.pathname === `${rootPath}/message-1/replies`) {
      const older = url.searchParams.has('$skiptoken');
      const top = Number(url.searchParams.get('$top') ?? 20);
      response.end(JSON.stringify(older ? { value: [replies[0]] } : { value: top === 1 ? [replies[1]] : [replies[1], replies[0]], '@odata.nextLink': `${servers.graph.url}${rootPath}/message-1/replies?$skiptoken=older` }));
      return;
    }
    response.writeHead(404).end(JSON.stringify({ error: { message: 'Not found' } }));
  });
});
function messagePaths(): string[] { return servers.graph.requests.map((request) => request.path).filter((path) => path.startsWith(rootPath)); }
after(async () => { await servers.close(); cli.cleanup(); });

it('reads a Root Message from a Teams link without Team or Channel defaults', async () => {
  servers.graph.reset();
  const link = 'https://teams.microsoft.com/l/message/chan-1/message-1?tenantId=tenant-1&groupId=team-1';
  const result = await cli.run(['thread', link]);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /message-1[\s\S]*Root text[\s\S]*reply-old[\s\S]*Older reply[\s\S]*reply-new[\s\S]*Newer reply/);
  assert.deepEqual(messagePaths(), [
    `${rootPath}/message-1`,
    `${rootPath}/message-1/replies?$top=20`,
  ]);
  assert.match(result.stderr, /more available: --cursor/);
});

it('pages back to older Replies from a Root Message id', async () => {
  servers.graph.reset();
  const first = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'thread', 'message-1', '--top', '1']);
  assert.equal(first.code, 0, first.stderr);
  assert.match(first.stdout, /Root text[\s\S]*reply-new/);
  assert.deepEqual(messagePaths(), [
    `${rootPath}/message-1`,
    `${rootPath}/message-1/replies?$top=1`,
  ]);
  const cursor = first.stderr.match(/--cursor ([A-Za-z0-9_-]+)/)?.[1];
  assert.ok(cursor, first.stderr);

  servers.graph.reset();
  const older = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'thread', 'message-1', '--cursor', cursor]);
  assert.equal(older.code, 0, older.stderr);
  assert.match(older.stdout, /Root text[\s\S]*reply-old[\s\S]*Older reply/);
  assert.doesNotMatch(older.stdout, /reply-new/);
  assert.equal(servers.graph.requests.at(-1)?.path, `${rootPath}/message-1/replies?$skiptoken=older`);
  assert.equal(older.stderr, '');
});

it('reads the Thread of a Reply link through its parentMessageId', async () => {
  servers.graph.reset();
  const link = 'https://teams.microsoft.com/l/message/chan-1/reply-new?tenantId=tenant-1&groupId=team-1&parentMessageId=message-1';
  const result = await cli.run(['thread', link]);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(messagePaths(), [
    `${rootPath}/message-1`,
    `${rootPath}/message-1/replies?$top=20`,
  ]);
});

it('explains that a bare Reply id cannot be read on its own', async () => {
  servers.graph.reset();
  const result = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'thread', 'reply-new']);
  assert.equal(result.code, 2);
  assert.match(result.stderr, /Root Message id/);
  assert.deepEqual(messagePaths(), [`${rootPath}/reply-new`]);
});

it('resolves Team and Channel names for a Root Message id', async () => {
  servers.graph.reset();
  const result = await cli.run(['--team', 'engineering', '--channel', 'general', 'thread', 'message-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(messagePaths()[0], `${rootPath}/message-1`);
});

it('reads a Thread from a link with a real-shaped channel id', async () => {
  servers.graph.reset();
  const link = 'https://teams.microsoft.com/l/message/19%3Aabc123%40thread.tacv2/message-1?tenantId=tenant-1&groupId=team-1';
  await cli.run(['thread', link]);
  assert.equal(servers.graph.requests[0]?.path, '/teams/team-1/channels/19%3Aabc123%40thread.tacv2/messages/message-1');
});

it('accepts links on the teams.cloud.microsoft host', async () => {
  servers.graph.reset();
  const result = await cli.run(['thread', 'https://teams.cloud.microsoft/l/message/chan-1/message-1?tenantId=tenant-1&groupId=team-1']);
  assert.equal(result.code, 0, result.stderr);
});

it('rejects malformed ids and look-alike Teams links without a Graph request', async () => {
  for (const value of ['bad/id', 'https://teams.microsoft.com.evil.test/l/message/chan-1/message-1?groupId=team-1', 'https://teams.microsoft.com/l/message/chan-1/message-1/extra?groupId=team-1']) {
    servers.graph.reset();
    const result = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'thread', value]);
    assert.equal(result.code, 2, `${value}: ${result.stderr}`);
    assert.equal(servers.graph.requests.length, 0, value);
  }
});
