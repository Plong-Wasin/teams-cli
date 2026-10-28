import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers;
let cli: Cli;
const root = '/teams/team-1/channels/chan-1/messages';

before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
  servers.graph.setReply((request, response) => {
    const url = new URL(request.path, 'http://graph.test');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (url.pathname === '/me/joinedTeams') {
      response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] }));
      return;
    }
    if (url.pathname === '/teams/team-1/channels') {
      response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] }));
      return;
    }
    if (url.pathname === root) {
      response.end(JSON.stringify({ value: [{ id: 'message-1', createdDateTime: '2026-10-28T10:00:00Z', from: { user: { displayName: 'Ada' } }, body: { contentType: 'html', content: `<p>${'Hello Teams '.repeat(12)}</p>` } }], '@odata.nextLink': `${servers.graph.url}${root}?$skiptoken=older` }));
      return;
    }
    if (url.pathname === `${root}/message-1/replies`) {
      response.end(JSON.stringify({ value: [{ id: 'reply-1', createdDateTime: '2026-10-28T10:01:00Z', from: { user: { displayName: 'Lin' } }, body: { content: '<p>Thanks!</p>' } }] }));
      return;
    }
    response.writeHead(404).end(JSON.stringify({ error: { message: 'Not found' } }));
  });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('lists recent Root Messages in a compact table and uses the default page size', async () => {
  servers.graph.reset();
  const result = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ID\s+TIME\s+SENDER\s+BODY/);
  assert.match(result.stdout, /message-1\s+2026-10-28T10:00:00Z\s+Ada\s+Hello Teams/);
  assert.doesNotMatch(result.stdout, /Hello Teams (?:Hello Teams ){10}/);
  assert.match(result.stderr, /more available: --cursor/);
  assert.deepEqual(servers.graph.requests.slice(-1).map((request) => request.path), [`${root}?$top=20`]);
});

it('passes --top, starts at --cursor, prints full text, and emits raw JSON', async () => {
  servers.graph.reset();
  const top = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1', '--top', '7', '--full']);
  assert.equal(top.code, 0, top.stderr);
  assert.match(top.stdout, /Hello Teams (?:Hello Teams ){10}/);
  assert.deepEqual(servers.graph.requests.at(-1)!.path, `${root}?$top=7`);

  const cursor = Buffer.from(`${servers.graph.url}/teams/team-1/channels/chan-1/messages?$skiptoken=older`).toString('base64url');
  const page = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1', '--cursor', cursor, '--json']);
  assert.equal(page.code, 0, page.stderr);
  assert.deepEqual(JSON.parse(page.stdout).value[0].id, 'message-1');
  assert.deepEqual(servers.graph.requests.at(-1)!.path, '/teams/team-1/channels/chan-1/messages?$skiptoken=older');
});

it('fetches a Reply preview only when --replies is set', async () => {
  servers.graph.reset();
  const without = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1']);
  assert.equal(without.code, 0, without.stderr);
  assert.equal(servers.graph.requests.filter((request) => request.path.endsWith('/messages') || request.path.includes('/messages?')).length, 1);

  const withReplies = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1', '--replies']);
  assert.equal(withReplies.code, 0, withReplies.stderr);
  assert.deepEqual(servers.graph.requests.slice(-2).map((request) => request.path), [root + '?$top=20', `${root}/message-1/replies?%24top=5`]);
  assert.match(withReplies.stdout, /Lin: Thanks!/);
});

it('follows @odata.nextLink with --all and retains a cursor after the page cap', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    const url = new URL(request.path, 'http://graph.test');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (url.pathname === '/me/joinedTeams') { response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] })); return; }
    if (url.pathname === '/teams/team-1/channels') { response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] })); return; }
    const page = Number(url.searchParams.get('$skiptoken') ?? 0);
    response.end(JSON.stringify({ value: [{ id: `message-${page}`, body: { content: `page ${page}` } }], ...(page < 1 ? { '@odata.nextLink': `${servers.graph.url}${root}?$skiptoken=${page + 1}` } : {}) }));
  });
  const result = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1', '--all']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /message-0/);
  assert.match(result.stdout, /message-1/);
  assert.equal(servers.graph.requests.filter((request) => new URL(request.path, 'http://graph.test').pathname === root).length, 2);
  assert.equal(result.stderr, '');

  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    const url = new URL(request.path, 'http://graph.test');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (url.pathname === '/me/joinedTeams') { response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }] })); return; }
    if (url.pathname === '/teams/team-1/channels') { response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General' }] })); return; }
    const page = Number(url.searchParams.get('$skiptoken') ?? 0);
    response.end(JSON.stringify({ value: [{ id: `message-${page}` }], '@odata.nextLink': `${servers.graph.url}${root}?$skiptoken=${page + 1}` }));
  });
  const capped = await cli.run(['messages', '--team', 'team-1', '--channel', 'chan-1', '--all']);
  assert.equal(capped.code, 0, capped.stderr);
  assert.equal(servers.graph.requests.filter((request) => new URL(request.path, 'http://graph.test').pathname === root).length, 100);
  assert.match(capped.stderr, /more available: --cursor/);
});

it('rejects --top above the Graph page limit of 50', async () => {
  servers.graph.reset();
  const result = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'messages', '--top', '51']);
  assert.equal(result.code, 2);
  assert.match(result.stderr, /1 to 50/);
  assert.equal(servers.graph.requests.some((request) => request.path.includes('/messages')), false);
});

it('accepts a continuation link in OData key form on the configured Graph host', async () => {
  servers.graph.reset();
  const cursor = Buffer.from(`${servers.graph.url}/teams('team-1')/channels('chan-1')/messages?$skiptoken=1`).toString('base64url');
  const result = await cli.run(['--team', 'team-1', '--channel', 'chan-1', 'messages', '--cursor', cursor]);
  assert.notEqual(result.code, 2, result.stderr);
  assert.equal(servers.graph.requests.at(-1)?.path, "/teams('team-1')/channels('chan-1')/messages?$skiptoken=1");
});
