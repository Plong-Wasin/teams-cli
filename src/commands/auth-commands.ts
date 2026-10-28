import { CliError, usage } from '../utils/errors.js';
import { configPath, loadConfig, saveConfig, type Config, type Profile } from '../utils/config.js';

const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const SCOPES = 'User.Read User.ReadBasic.All Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All ChannelMessage.Send ChannelMessage.ReadWrite Chat.Read offline_access';
export interface Options { profile?: string; clientId?: string; team?: string; channel?: string }
export interface Context { env: NodeJS.ProcessEnv; options: Options; config: Config; path: string; profileName: string; profile: Profile }
export function selectedContext(env: NodeJS.ProcessEnv, options: Options): Context {
  const path = configPath(env); const config = loadConfig(path);
  const profileName = options.profile || env.TEAMS_PROFILE || config.defaultProfile || 'default';
  config.profiles[profileName] ??= {};
  return { env, options, config, path, profileName, profile: config.profiles[profileName]! };
}
function clientId(ctx: Context): string { return ctx.options.clientId || ctx.env.TEAMS_CLIENT_ID || ctx.profile.clientId || DEFAULT_CLIENT_ID; }
function loginBase(ctx: Context): string { return (ctx.env.TEAMS_LOGIN_URL || 'https://login.microsoftonline.com').replace(/\/$/, ''); }
function tokenEndpoint(ctx: Context): string { return `${loginBase(ctx)}/common/oauth2/v2.0/token`; }
function persist(ctx: Context): void { ctx.config.profiles[ctx.profileName] = ctx.profile; ctx.config.defaultProfile ??= ctx.profileName; saveConfig(ctx.path, ctx.config); }

export async function getAccessToken(ctx: Context): Promise<string> {
  const envToken = ctx.env.TEAMS_TOKEN?.replace(/^Bearer\s+/i, '');
  if (envToken) return envToken;
  if (!ctx.profile.accessToken) throw new CliError(`Profile '${ctx.profileName}' is not signed in. Run 'teams auth login'.`);
  if (ctx.profile.expiresAt && ctx.profile.expiresAt <= Date.now() + 60_000) await refresh(ctx);
  return ctx.profile.accessToken!.replace(/^Bearer\s+/i, '');
}
interface Token { access_token?: string; refresh_token?: string; expires_in?: number; error?: string }
async function refresh(ctx: Context): Promise<void> {
  if (!ctx.profile.refreshToken) throw new CliError("Sign-in expired. Run 'teams auth login' to sign in again.");
  let response: Response;
  try {
    response = await fetch(tokenEndpoint(ctx), { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: clientId(ctx), grant_type: 'refresh_token', refresh_token: ctx.profile.refreshToken, scope: SCOPES }) });
  } catch { throw new CliError("Could not refresh sign-in. Run 'teams auth login' to sign in again."); }
  const data = await response.json() as Token;
  if (!response.ok || !data.access_token) throw new CliError("Sign-in expired. Run 'teams auth login' to sign in again.");
  ctx.profile.accessToken = data.access_token; ctx.profile.refreshToken = data.refresh_token || ctx.profile.refreshToken;
  ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1000; persist(ctx);
}

export async function graphRequest(ctx: Context, method: string, path: string, options: { query?: URLSearchParams; body?: string; rawUrl?: string; headers?: Record<string, string> } = {}): Promise<unknown> {
  const base = (ctx.env.TEAMS_GRAPH_URL || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
  const url = new URL(options.rawUrl || `${base}${path.startsWith('/') ? path : `/${path}`}`);
  options.query?.forEach((value, key) => url.searchParams.append(key, value));
  const token = await getAccessToken(ctx);
  let response: Response;
  try { response = await fetch(url, { method, headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', ...options.headers, ...(options.body === undefined ? {} : { 'Content-Type': 'application/json' }) }, ...(options.body === undefined ? {} : { body: options.body }) }); }
  catch { throw new CliError(`Cannot reach ${url} for ${method} ${path}. Check the network and Graph URL.`); }
  if (!response.ok) {
    let detail = '';
    try { const data = await response.json() as { error?: string | { code?: string; message?: string }; message?: string }; detail = typeof data.error === 'string' ? data.error : [data.error?.code, data.error?.message ?? data.message].filter(Boolean).join(': '); } catch { /* empty response */ }
    const hint: Record<number, string> = {
      401: `Authentication failed for ${method} ${path}. Run 'teams auth check' or 'teams auth login'.`,
      403: `Access denied for ${method} ${path}. Check that your token has the required Microsoft Graph scopes.`,
      404: `Not found: ${method} ${path}. Check the id and path.`,
      429: `Rate limited for ${method} ${path}. Wait, then try again.`,
    };
    throw new CliError(`${hint[response.status] || `HTTP ${response.status} for ${method} ${path}.`}${detail ? ` Graph said: ${detail}` : ''}`);
  }
  const text = await response.text(); return text ? JSON.parse(text) as unknown : undefined;
}

export async function runAuth(ctx: Context, subcommand: string | undefined): Promise<void> {
  if (subcommand === 'login') return login(ctx);
  if (subcommand === 'check') {
    const me = await graphRequest(ctx, 'GET', '/me') as { displayName?: string; userPrincipalName?: string; mail?: string };
    process.stdout.write(`Signed in as ${me.displayName || me.userPrincipalName || me.mail || 'current user'}\n`); return;
  }
  if (subcommand === 'logout') {
    delete ctx.profile.accessToken; delete ctx.profile.refreshToken; delete ctx.profile.expiresAt; persist(ctx);
    process.stdout.write(`Signed out of Profile '${ctx.profileName}'.\n`); return;
  }
  usage("Use 'teams auth login', 'teams auth check', or 'teams auth logout'.");
}
async function login(ctx: Context): Promise<void> {
  const id = clientId(ctx); const deviceUrl = `${loginBase(ctx)}/common/oauth2/v2.0/devicecode`;
  let deviceResponse: Response;
  try { deviceResponse = await fetch(deviceUrl, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: id, scope: SCOPES }) }); }
  catch { throw new CliError(`Cannot reach ${deviceUrl}. Check the login URL and network.`); }
  const device = await deviceResponse.json() as { device_code?: string; user_code?: string; verification_uri?: string; interval?: number; expires_in?: number };
  if (!deviceResponse.ok || !device.device_code) throw new CliError('Could not start device-code sign-in. Check the client id and try again.');
  process.stdout.write(`To sign in, visit ${device.verification_uri} and enter code ${device.user_code}.\n`);
  const interval = Math.max(1, device.interval || 5) * 1000; const expires = Date.now() + (device.expires_in || 900) * 1000;
  while (Date.now() < expires) {
    const response = await fetch(tokenEndpoint(ctx), { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams({ client_id: id, grant_type: 'urn:ietf:params:oauth:grant-type:device_code', device_code: device.device_code }) });
    const data = await response.json() as Token;
    if (response.ok && data.access_token) {
      ctx.profile.accessToken = data.access_token; ctx.profile.refreshToken = data.refresh_token; ctx.profile.clientId = id;
      ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1000; persist(ctx);
      process.stdout.write(`Signed in to Profile '${ctx.profileName}'.\n`); return;
    }
    if (data.error === 'authorization_pending') { await new Promise((resolve) => setTimeout(resolve, interval)); continue; }
    if (data.error === 'slow_down') { await new Promise((resolve) => setTimeout(resolve, interval + 5000)); continue; }
    throw new CliError("Device-code sign-in failed. Run 'teams auth login' to try again.");
  }
  throw new CliError("Device code expired. Run 'teams auth login' to try again.");
}
