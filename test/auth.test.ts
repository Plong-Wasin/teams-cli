import assert from 'node:assert/strict';
import { readFileSync, statSync, writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startServers, type Servers } from './helpers/fake-server.ts';

let servers: Servers; let cli: Cli;
before(async () => {
  servers = await startServers();
  cli = createCli({ TEAMS_GRAPH_URL: servers.graph.url, TEAMS_LOGIN_URL: servers.login.url });
  servers.login.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path.endsWith('/devicecode')) response.end(JSON.stringify({ device_code: 'device-secret', user_code: 'ABCD-EFGH', verification_uri: 'https://login.example.test/device', interval: 0, expires_in: 30 }));
    else response.end(JSON.stringify({ access_token: 'access-secret', refresh_token: 'refresh-secret', expires_in: 3600 }));
  });
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path === '/me' ? { displayName: 'Example User', userPrincipalName: 'user@example.com' } : { value: [] }));
  });
});
after(async () => { await servers.close(); cli.cleanup(); });
function config(value: unknown): void { writeFileSync(`${cli.home}/config.json`, JSON.stringify(value), { mode: 0o600 }); }

it('logs in with device code, User.Read baseline scope, and saves mode 600 credentials', async () => {
  const result = await cli.run(['auth', 'login', '--client-id', 'test-client'], { TEAMS_CLIENT_ID: 'ignored-client' });
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /ABCD-EFGH/);
  const device = servers.login.requests.find((request) => request.path.endsWith('/devicecode'))!;
  const deviceBody = new URLSearchParams(device.body);
  assert.equal(device.method, 'POST');
  assert.equal(deviceBody.get('client_id'), 'test-client');
  assert.match(deviceBody.get('scope') ?? '', /User\.Read/);
  assert.match(deviceBody.get('scope') ?? '', /offline_access/);
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8'));
  assert.equal(saved.profiles.default.refreshToken, 'refresh-secret');
  assert.equal(statSync(`${cli.home}/config.json`).mode & 0o777, 0o600);
  servers.login.reset();
  const fromEnvironment = await cli.run(['auth', 'login'], { TEAMS_CLIENT_ID: 'environment-client' });
  assert.equal(fromEnvironment.code, 0, fromEnvironment.stderr);
  const environmentRequest = servers.login.requests.find((request) => request.path.endsWith('/devicecode'))!;
  assert.equal(new URLSearchParams(environmentRequest.body).get('client_id'), 'environment-client');
});

it('refreshes expired credentials and check sends the refreshed bearer token', async () => {
  config({ defaultProfile: 'work', profiles: { work: { accessToken: 'old', refreshToken: 'old-refresh', expiresAt: 1 } } });
  servers.login.reset(); servers.graph.reset();
  const result = await cli.run(['auth', 'check']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Example User/);
  assert.equal(new URLSearchParams(servers.login.requests[0]!.body).get('grant_type'), 'refresh_token');
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer access-secret');
});

it('selects Profiles by flag, supports headless TEAMS_TOKEN, masks config, and logs out', async () => {
  config({ defaultProfile: 'one', profiles: { one: { accessToken: 'secret-one' }, two: { accessToken: 'secret-two', refreshToken: 'refresh-two' } } });
  servers.graph.reset();
  const selected = await cli.run(['auth', 'check', '-p', 'two'], { TEAMS_PROFILE: 'one' });
  assert.equal(selected.code, 0, selected.stderr);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer secret-two');
  const headless = await cli.run(['me'], { TEAMS_TOKEN: 'Bearer headless-secret' });
  assert.equal(headless.code, 0, headless.stderr);
  assert.equal(servers.graph.requests.at(-1)!.headers.authorization, 'Bearer headless-secret');
  const shown = await cli.run(['config', 'show']);
  assert.equal(shown.code, 0);
  assert.doesNotMatch(shown.stdout, /secret-one|secret-two|refresh-two/);
  const loggedOut = await cli.run(['auth', 'logout', '-p', 'two']);
  assert.equal(loggedOut.code, 0, loggedOut.stderr);
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8'));
  assert.equal(saved.profiles.two.accessToken, undefined);
  assert.equal(saved.profiles.one.accessToken, 'secret-one');
});

it('prints the signed-in identity and gives useful HTTP errors', async () => {
  config({ profiles: { default: { accessToken: 'token' } } });
  servers.graph.reset();
  const me = await cli.run(['me']);
  assert.equal(me.code, 0, me.stderr);
  assert.match(me.stdout, /Example User/);
  servers.graph.setReply((_request, response) => { response.writeHead(403, { 'Content-Type': 'application/json' }).end(JSON.stringify({ error: { message: 'Denied' } })); });
  const denied = await cli.run(['me']);
  assert.equal(denied.code, 1);
  assert.match(denied.stderr, /teams:.*GET \/me.*scope/i);
});
