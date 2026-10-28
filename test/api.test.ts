import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers; let cli: Cli;
before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('sends method, path, query, JSON body, and custom header for raw API', async () => {
  servers.graph.setReply((request, response) => { response.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({ ok: true, body: request.body })); });
  const result = await cli.run(['api', 'POST', '/teams/team-1/channels', 'filter=one', '--header', 'Prefer: outlook.body-content-type="text"', '--data', '{"name":"General"}']);
  assert.equal(result.code, 0, result.stderr);
  const request = servers.graph.requests[0]!;
  assert.equal(request.method, 'POST');
  assert.equal(request.path, '/teams/team-1/channels?filter=one');
  assert.equal(request.headers.authorization, 'Bearer test-token');
  assert.equal(request.headers.prefer, 'outlook.body-content-type="text"');
  assert.deepEqual(JSON.parse(request.body), { name: 'General' });
});

it('accepts file data, collects continuation pages with GET --all, and accepts an empty 204', async () => {
  servers.graph.reset();
  const file = join(cli.home, 'body.json'); writeFileSync(file, '{"displayName":"New channel"}');
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path === '/me' ? { value: [{ id: 'one' }], '@odata.nextLink': `${servers.graph.url}/me?$skiptoken=next` } : { value: [{ id: 'two' }] }));
  });
  const all = await cli.run(['api', 'GET', '/me', '--all']);
  assert.equal(all.code, 0, all.stderr);
  assert.deepEqual(JSON.parse(all.stdout).value, [{ id: 'one' }, { id: 'two' }]);
  assert.equal(servers.graph.requests.length, 2);
  servers.graph.reset();
  servers.graph.setReply((_request, response) => { response.writeHead(204).end(); });
  const empty = await cli.run(['api', 'PATCH', '/teams/team-1/channels/chan-1', '--data', `@${file}`]);
  assert.equal(empty.code, 0, empty.stderr);
  assert.equal(empty.stdout, '');
  assert.deepEqual(JSON.parse(servers.graph.requests[0]!.body), { displayName: 'New channel' });
});

it('returns documented error and usage exit codes and lists commands in help', async () => {
  for (const [status, path, hint] of [[401, '/me', 'auth check'], [403, '/me', 'scopes'], [404, '/users/missing', 'Check'], [429, '/me', 'Wait']] as const) {
    servers.graph.reset();
    servers.graph.setReply((_request, response) => { response.writeHead(status, { 'Content-Type': 'application/json' }).end(JSON.stringify({ error: { message: 'Missing' } })); });
    const failed = await cli.run(['api', 'GET', path]);
    assert.equal(failed.code, 1);
    assert.match(failed.stderr, new RegExp(`teams:.*GET ${path.replaceAll('/', '\\/')}.*${hint}`, 'i'));
  }
  const bad = await cli.run(['api', 'TRACE', '/me']);
  assert.equal(bad.code, 2);
  assert.match(bad.stderr, /^teams: /);
  const help = await cli.run(['--help']);
  assert.equal(help.code, 0);
  for (const command of ['auth', 'config', 'me', 'api']) assert.match(help.stdout, new RegExp(command));
});
