import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers; let cli: Cli;
before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/joinedTeams') response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] }));
    else if (request.path === '/teams/team-1/channels') response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] }));
    else if (request.path.startsWith('/users/')) {
      const upn = decodeURIComponent(request.path.slice('/users/'.length));
      response.end(JSON.stringify({ id: upn === 'jane@example.com' ? 'user-1' : 'user-2', displayName: upn === 'jane@example.com' ? 'Jane Doe' : 'Sam & Co' }));
    } else if (request.path === '/teams/team-1/channels/chan-1/messages') response.end(JSON.stringify({ id: 'message-1' }));
    else response.writeHead(404).end(JSON.stringify({ error: { message: 'Not found' } }));
  });
});
after(async () => { await servers.close(); cli.cleanup(); });
function config(value: unknown): void { writeFileSync(`${cli.home}/config.json`, JSON.stringify(value)); }

it('sends a Root Message with plain text escaped and newlines preserved', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', 'Hello <team> & friends\nNext line'], { TEAMS_TEAM: 'team-1', TEAMS_CHANNEL: 'chan-1' });
  assert.equal(result.code, 0, result.stderr);
  const post = servers.graph.requests.at(-1)!;
  assert.equal(post.method, 'POST');
  assert.equal(post.path, '/teams/team-1/channels/chan-1/messages');
  assert.deepEqual(JSON.parse(post.body), { body: { contentType: 'html', content: 'Hello &lt;team&gt; &amp; friends<br>Next line' } });
});

it('reads the message body from a file or stdin and sends raw HTML with a subject', async () => {
  const file = `${cli.home}/message.txt`; writeFileSync(file, 'From file & safe');
  servers.graph.reset();
  let result = await cli.run(['send', '--body', `@${file}`, '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(JSON.parse(servers.graph.requests.at(-1)!.body), { body: { contentType: 'html', content: 'From file &amp; safe' } });
  servers.graph.reset();
  result = await cli.run(['send', '-', '--html', '--subject', 'Release notes', '--team', 'team-1', '--channel', 'chan-1'], {}, '<b>from stdin</b>');
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(JSON.parse(servers.graph.requests.at(-1)!.body), { subject: 'Release notes', body: { contentType: 'html', content: '<b>from stdin</b>' } });
});

it('resolves every repeated mention and adds escaped tags and mention objects', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', '--body', 'please review', '--mention', 'jane@example.com', '--mention', 'sam@example.com', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.filter((request) => request.path.startsWith('/users/')).map((request) => request.path), ['/users/jane%40example.com', '/users/sam%40example.com']);
  assert.deepEqual(JSON.parse(servers.graph.requests.at(-1)!.body), {
    body: { contentType: 'html', content: '<at id="0">Jane Doe</at> <at id="1">Sam &amp; Co</at> please review' },
    mentions: [
      { id: 0, mentionText: 'Jane Doe', mentioned: { user: { id: 'user-1', displayName: 'Jane Doe', userIdentityType: 'aadUser' } } },
      { id: 1, mentionText: 'Sam & Co', mentioned: { user: { id: 'user-2', displayName: 'Sam & Co', userIdentityType: 'aadUser' } } },
    ],
  });
});

it('prints the request body for dry run without posting', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', 'Preview <this>', '--subject', 'Draft', '--dry-run', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), { subject: 'Draft', body: { contentType: 'html', content: 'Preview &lt;this&gt;' } });
  assert.equal(servers.graph.requests.some((request) => request.method === 'POST'), false);
});
