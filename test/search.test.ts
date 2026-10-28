import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers; let cli: Cli;
before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_TOKEN: 'test-token' });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('posts the KQL query intact with the requested size and renders Channel and Chat hits', async () => {
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({
      value: [{ hitsContainers: [{ hits: [
        { resource: { id: 'message-1', channelIdentity: { teamId: 'team-1', channelId: 'chan-1' }, from: { user: { displayName: 'Ari' } }, createdDateTime: '2026-10-28T10:00:00Z', body: { content: 'deploy update' } } },
        { resource: { id: 'message-2', from: { user: { displayName: 'Bea' } }, createdDateTime: '2026-10-28T11:00:00Z', body: { content: 'chat update' } } },
      ] }] }], extra: 'preserved',
    }));
  });
  const query = 'deploy from:ari sent>2026-10-01';
  const result = await cli.run(['search', query, '--size', '7']);
  assert.equal(result.code, 0, result.stderr);
  const request = servers.graph.requests[0]!;
  assert.equal(request.method, 'POST');
  assert.equal(request.path, '/search/query');
  assert.deepEqual(JSON.parse(request.body), { requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: 0, size: 7, enableTopResults: true }] });
  assert.match(result.stdout, /team-1/);
  assert.match(result.stdout, /chan-1/);
  assert.match(result.stdout, /message-1/);
  assert.match(result.stdout, /Chat/);
  assert.match(result.stdout, /message-2/);

  servers.graph.reset();
  const raw = { value: [{ hitsContainers: [{ hits: [{ resource: { id: 'message-3', channelIdentity: { teamId: 'team-1', channelId: 'chan-1' } } }] }] }], rawField: true };
  servers.graph.setReply((_request, response) => { response.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify(raw)); });
  const json = await cli.run(['search', query, '--json']);
  assert.equal(json.code, 0, json.stderr);
  assert.deepEqual(JSON.parse(json.stdout), raw);
});
