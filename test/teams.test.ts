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
    if (request.path === '/me/joinedTeams') response.end(JSON.stringify({ value: [{ id: 'team-1', displayName: 'Engineering' }, { id: 'team-2', displayName: 'Support' }, { id: 'team-3', displayName: 'engineering' }] }));
    else if (request.path === '/teams/team-1/channels') response.end(JSON.stringify({ value: [{ id: 'chan-1', displayName: 'General', membershipType: 'standard' }, { id: 'chan-2', displayName: 'Private', membershipType: 'private' }, { id: 'chan-4', displayName: 'general', membershipType: 'standard' }] }));
    else if (request.path === '/teams/team-1/channels/chan-1') response.end(JSON.stringify({ id: 'chan-1', displayName: 'General', membershipType: 'standard', layout: 'threaded' }));
    else if (request.path === '/teams/team-2/channels') response.end(JSON.stringify({ value: [{ id: 'chan-3', displayName: 'General', membershipType: 'standard' }] }));
    else if (request.path === '/teams/team-2/channels/chan-3') response.end(JSON.stringify({ id: 'chan-3', displayName: 'General', membershipType: 'standard', layout: 'threaded' }));
    else response.writeHead(404).end(JSON.stringify({ error: { message: 'Not found' } }));
  });
});
after(async () => { await servers.close(); cli.cleanup(); });
function config(value: unknown): void { writeFileSync(`${cli.home}/config.json`, JSON.stringify(value), { mode: 0o600 }); }
function saved(): any { return JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')); }

it('lists joined Teams as a compact table and as JSON', async () => {
  servers.graph.reset();
  const table = await cli.run(['teams']);
  assert.equal(table.code, 0, table.stderr);
  assert.match(table.stdout, /ID\s+NAME/);
  assert.match(table.stdout, /team-1\s+Engineering/);
  assert.deepEqual(servers.graph.requests[0]!.path, '/me/joinedTeams');
  const json = await cli.run(['teams', '--json']);
  assert.equal(json.code, 0, json.stderr);
  assert.deepEqual(JSON.parse(json.stdout).value[0], { id: 'team-1', displayName: 'Engineering' });
  assert.equal(servers.graph.requests.at(-1)!.method, 'GET');
});

it('lists Channels for a Team name and shows a Channel using its single-item layout', async () => {
  servers.graph.reset();
  const result = await cli.run(['channels', 'Engineering']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /chan-1\s+General\s+standard/);
  assert.deepEqual(servers.graph.requests.map((request) => [request.method, request.path]), [
    ['GET', '/me/joinedTeams'], ['GET', '/teams/team-1/channels'],
  ]);
  servers.graph.reset();
  const detail = await cli.run(['channel', 'Engineering', 'General', '--json']);
  assert.equal(detail.code, 0, detail.stderr);
  assert.equal(JSON.parse(detail.stdout).layout, 'threaded');
  assert.deepEqual(servers.graph.requests.map((request) => request.path), [
    '/me/joinedTeams', '/teams/team-1/channels', '/teams/team-1/channels/chan-1',
  ]);
});

it('stores resolved Team and Channel ids as Profile defaults', async () => {
  servers.graph.reset();
  const team = await cli.run(['config', 'set', 'team', 'Engineering']);
  assert.equal(team.code, 0, team.stderr);
  assert.equal(saved().profiles.default.defaultTeam, 'team-1');
  const channel = await cli.run(['config', 'set', 'channel', 'General']);
  assert.equal(channel.code, 0, channel.stderr);
  assert.equal(saved().profiles.default.defaultChannel, 'chan-1');
  assert.deepEqual(servers.graph.requests.map((request) => request.path), [
    '/me/joinedTeams', '/me/joinedTeams', '/teams/team-1/channels',
  ]);
});

it('resolves exact names before case-insensitive names and lists candidates on ambiguity or absence', async () => {
  config({ profiles: { default: { defaultTeam: 'team-1' } } });
  const exact = await cli.run(['channels', 'Engineering']);
  assert.equal(exact.code, 0, exact.stderr);
  assert.equal(servers.graph.requests.at(-1)!.path, '/teams/team-1/channels');
  const ambiguousTeam = await cli.run(['channels', 'ENGINEERING']);
  assert.equal(ambiguousTeam.code, 1);
  assert.match(ambiguousTeam.stderr, /Ambiguous Team.*Engineering \(team-1\).*engineering \(team-3\)/);
  const ambiguousChannel = await cli.run(['channel', 'Engineering', 'GENERAL']);
  assert.equal(ambiguousChannel.code, 1);
  assert.match(ambiguousChannel.stderr, /Ambiguous Channel.*General \(chan-1\).*general \(chan-4\)/);
  const missing = await cli.run(['channels', 'Unknown']);
  assert.equal(missing.code, 1);
  assert.match(missing.stderr, /No matching Team 'Unknown'.*Engineering \(team-1\)/);
  const missingChannel = await cli.run(['channel', 'Engineering', 'Unknown']);
  assert.equal(missingChannel.code, 1);
  assert.match(missingChannel.stderr, /No matching Channel 'Unknown'.*General \(chan-1\).*general \(chan-4\)/);
});

it('uses Team and Channel flags before environment and Profile defaults', async () => {
  config({ profiles: { default: { defaultTeam: 'team-1', defaultChannel: 'chan-1' } } });
  const result = await cli.run(['channel', '--team', 'team-2', '--channel', 'chan-3', '--json'], { TEAMS_TEAM: 'team-1', TEAMS_CHANNEL: 'chan-1' });
  assert.equal(result.code, 0, result.stderr);
  assert.equal(JSON.parse(result.stdout).id, 'chan-3');
  assert.equal(servers.graph.requests.at(-1)!.path, '/teams/team-2/channels/chan-3');
  const fromEnv = await cli.run(['channel', '--json'], { TEAMS_TEAM: 'team-2', TEAMS_CHANNEL: 'chan-3' });
  assert.equal(fromEnv.code, 0, fromEnv.stderr);
  assert.equal(JSON.parse(fromEnv.stdout).id, 'chan-3');
});
