import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers;
let cli: Cli;
const rootPath = '/teams/team-1/channels/chan-1/messages';

before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
  servers.graph.setReply((request, response) => {
    const url = new URL(request.path, 'http://graph.test');
    response.setHeader('Content-Type', 'application/json');
    if (url.pathname === '/me/joinedTeams') response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] }));
    else if (url.pathname === '/teams/team-1/channels') response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] }));
    else if (url.pathname.startsWith('/users/')) response.end(JSON.stringify({ id: 'user-1', displayName: 'Sam & Co' }));
    else if (url.pathname === `${rootPath}/reply-1`) response.end(JSON.stringify({ id: 'reply-1', replyToId: 'root-1' }));
    else if (url.pathname === `${rootPath}/root-1`) response.end(JSON.stringify({ id: 'root-1' }));
    else if (url.pathname === `${rootPath}/root-edit`) response.end(JSON.stringify({ id: 'root-edit' }));
    else if (url.pathname === `${rootPath}/root-1/replies`) response.end(JSON.stringify({ id: 'reply-2' }));
    else response.writeHead(404).end(JSON.stringify({ error: { message: 'Not found' } }));
  });
});

it('edits a Reply body without changing its parent', async () => {
  servers.graph.reset();
  const result = await cli.run(['edit', 'reply-1', 'Corrected text', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.slice(-2).map(({ method, path, body }) => ({ method, path, body })), [
    { method: 'GET', path: `${rootPath}/reply-1`, body: '' },
    { method: 'PATCH', path: `${rootPath}/reply-1`, body: JSON.stringify({ body: { contentType: 'html', content: 'Corrected text' } }) },
  ]);
});

it('uses file bodies, mentions, HTML and dry-run for write commands', async () => {
  const file = `${cli.home}/message.txt`;
  writeFileSync(file, '<b>draft</b>');
  servers.graph.reset();
  const result = await cli.run(['reply', 'root-1', '--body', `@${file}`, '--mention', 'sam@example.com', '--html', '--dry-run', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), {
    body: { contentType: 'html', content: '<at id="0">Sam &amp; Co</at> <b>draft</b>' },
    mentions: [{ id: 0, mentionText: 'Sam & Co', mentioned: { user: { id: 'user-1', displayName: 'Sam & Co', userIdentityType: 'aadUser' } } }],
  });
  assert.equal(servers.graph.requests.some((request) => request.method === 'POST'), false);
  assert.deepEqual(servers.graph.requests.slice(-3).map(({ method, path }) => ({ method, path })), [
    { method: 'GET', path: '/teams/team-1/channels' },
    { method: 'GET', path: '/teams/team-1/channels/chan-1/messages/root-1' },
    { method: 'GET', path: '/users/sam%40example.com' },
  ]);
});

it('reads an edit body from stdin and sends plain text as escaped HTML', async () => {
  servers.graph.reset();
  const result = await cli.run(['edit', 'root-edit', '--body', '-', '--team', 'team-1', '--channel', 'chan-1'], {}, 'Correction & clarification');
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.slice(-1).map(({ method, path, body }) => ({ method, path, body })), [
    { method: 'PATCH', path: `${rootPath}/root-edit`, body: JSON.stringify({ body: { contentType: 'html', content: 'Correction &amp; clarification' } }) },
  ]);
});

it('accepts a Channel Message link for editing a Root Message', async () => {
  servers.graph.reset();
  const result = await cli.run(['edit', 'https://teams.microsoft.com/l/message/chan-1/root-edit?tenantId=tenant-1&groupId=team-1', 'Updated',]);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => ({ method, path })), [
    { method: 'GET', path: `${rootPath}/root-edit` },
    { method: 'PATCH', path: `${rootPath}/root-edit` },
  ]);
});

it('refuses Chat links and ids with usage exit code 2', async () => {
  servers.graph.reset();
  const chatLink = 'https://teams.microsoft.com/l/message/chat-1/message-1?tenantId=tenant-1';
  const invalidLink = await cli.run(['reply', chatLink, 'Hello']);
  assert.equal(invalidLink.code, 2);
  assert.match(invalidLink.stderr, /Chat links cannot be written/);
  assert.equal(servers.graph.requests.length, 0);

  const chatId = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'edit', 'chat-message-1', 'Hello']);
  assert.equal(chatId.code, 2);
  assert.match(chatId.stderr, /Chat messages cannot be written/);
  assert.equal(servers.graph.requests.at(-1)?.path, `${rootPath}/chat-message-1`);
});

after(async () => { await servers.close(); cli.cleanup(); });

it('posts a Reply under the Root Message when given a Reply id', async () => {
  servers.graph.reset();
  const result = await cli.run(['reply', 'reply-1', 'Thanks <team>', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.slice(-2).map(({ method, path, body }) => ({ method, path, body })), [
    { method: 'GET', path: `${rootPath}/reply-1`, body: '' },
    { method: 'POST', path: `${rootPath}/root-1/replies`, body: JSON.stringify({ body: { contentType: 'html', content: 'Thanks &lt;team&gt;' } }) },
  ]);
});
