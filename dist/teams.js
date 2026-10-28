#!/usr/bin/env node

// src/help.ts
var HELP = `teams \u2014 Microsoft Teams Channel CLI

Usage: teams [options] <command>

Options:
  -p, --profile <name>    Select a Profile
  --client-id <id>        Microsoft public client id for sign-in
  --team <name|id>        Select a Team (overrides environment and Profile defaults)
  --channel <name|id>     Select a Channel (overrides environment and Profile defaults)
  -h, --help              Show help
  -v, --version           Show version

Commands:
  auth login | check | logout
  config show
  config set team <name|id>
  config set channel <name|id>
  me
  teams [--json]
  channels [team] [--json]
  channel [channel] [--team team] [--json]
  messages [--team team] [--channel channel] [--top N] [--cursor token] [--all] [--full] [--replies] [--json]
  thread <id|link> [--top N] [--cursor token] [--full] [--json]
  search <query> [--size N] [--json]
  send <text> [--body @file|-] [--subject text] [--mention upn] [--html] [--dry-run]
  reply <id|link> <text> [--root id] [--body @file|-] [--mention upn] [--html] [--dry-run]
  edit <id|link> <text> [--root id] [--body @file|-] [--mention upn] [--html] [--dry-run]
  api <GET|POST|PUT|PATCH|DELETE> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]

Environment:
  TEAMS_PROFILE, TEAMS_TOKEN, TEAMS_CLIENT_ID, TEAMS_CONFIG
  TEAMS_TEAM, TEAMS_CHANNEL
  TEAMS_GRAPH_URL, TEAMS_LOGIN_URL

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;

// package.json
var package_default = {
  name: "teams-cli",
  version: "0.1.0",
  description: "Command-line client for Microsoft Teams",
  type: "module",
  bin: { teams: "./dist/teams.js" },
  private: true,
  scripts: {
    typecheck: "tsc --noEmit && tsc --noEmit -p tsconfig.test.json",
    build: 'pnpm typecheck && esbuild src/index.ts --bundle --platform=node --target=node20 --format=esm --banner:js="#!/usr/bin/env node" --outfile=dist/teams.js && chmod +x dist/teams.js',
    test: 'pnpm build && node --test "test/*.test.ts"'
  },
  engines: { node: ">=20" },
  devDependencies: { "@types/node": "^20.19.0", esbuild: "^0.28.2", typescript: "^7.0.2" },
  license: "MIT",
  packageManager: "pnpm@12.3.4"
};

// src/utils/version.ts
var VERSION = package_default.version;

// src/utils/errors.ts
var CliError = class extends Error {
  constructor(message, exitCode = 1) {
    super(message);
    this.exitCode = exitCode;
  }
  exitCode;
};
function usage(message) {
  throw new CliError(message, 2);
}

// src/utils/config.ts
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { homedir } from "node:os";
function configPath(env) {
  return env.TEAMS_CONFIG || `${homedir()}/.config/teams/config.json`;
}
function loadConfig(path) {
  try {
    const value = JSON.parse(readFileSync(path, "utf8"));
    return value.profiles && typeof value.profiles === "object" ? { defaultProfile: value.defaultProfile, profiles: value.profiles } : { profiles: {} };
  } catch (error) {
    if (error.code === "ENOENT") return { profiles: {} };
    throw new CliError(`Cannot read config file ${path}. Check that it is valid JSON, or remove it and sign in again.`);
  }
}
function saveConfig(path, config) {
  mkdirSync(dirname(path), { recursive: true, mode: 448 });
  writeFileSync(path, `${JSON.stringify(config, null, 2)}
`, { mode: 384 });
  chmodSync(path, 384);
}
function maskSecrets(config) {
  return { ...config, profiles: Object.fromEntries(Object.entries(config.profiles).map(([name, profile]) => [name, {
    ...profile,
    ...profile.accessToken ? { accessToken: "********" } : {},
    ...profile.refreshToken ? { refreshToken: "********" } : {}
  }])) };
}

// src/commands/auth-commands.ts
var DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
var SCOPES = "User.Read User.ReadBasic.All Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All ChannelMessage.Send ChannelMessage.ReadWrite Chat.Read offline_access";
function selectedContext(env, options) {
  const path = configPath(env);
  const config = loadConfig(path);
  const profileName = options.profile || env.TEAMS_PROFILE || config.defaultProfile || "default";
  config.profiles[profileName] ??= {};
  return { env, options, config, path, profileName, profile: config.profiles[profileName] };
}
function clientId(ctx) {
  return ctx.options.clientId || ctx.env.TEAMS_CLIENT_ID || ctx.profile.clientId || DEFAULT_CLIENT_ID;
}
function loginBase(ctx) {
  return (ctx.env.TEAMS_LOGIN_URL || "https://login.microsoftonline.com").replace(/\/$/, "");
}
function tokenEndpoint(ctx) {
  return `${loginBase(ctx)}/common/oauth2/v2.0/token`;
}
function persist(ctx) {
  ctx.config.profiles[ctx.profileName] = ctx.profile;
  ctx.config.defaultProfile ??= ctx.profileName;
  saveConfig(ctx.path, ctx.config);
}
async function getAccessToken(ctx) {
  const envToken = ctx.env.TEAMS_TOKEN?.replace(/^Bearer\s+/i, "");
  if (envToken) return envToken;
  if (!ctx.profile.accessToken) throw new CliError(`Profile '${ctx.profileName}' is not signed in. Run 'teams auth login'.`);
  if (ctx.profile.expiresAt && ctx.profile.expiresAt <= Date.now() + 6e4) await refresh(ctx);
  return ctx.profile.accessToken.replace(/^Bearer\s+/i, "");
}
async function refresh(ctx) {
  if (!ctx.profile.refreshToken) throw new CliError("Sign-in expired. Run 'teams auth login' to sign in again.");
  let response;
  try {
    response = await fetch(tokenEndpoint(ctx), { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: clientId(ctx), grant_type: "refresh_token", refresh_token: ctx.profile.refreshToken, scope: SCOPES }) });
  } catch {
    throw new CliError("Could not refresh sign-in. Run 'teams auth login' to sign in again.");
  }
  const data = await response.json();
  if (!response.ok || !data.access_token) throw new CliError("Sign-in expired. Run 'teams auth login' to sign in again.");
  ctx.profile.accessToken = data.access_token;
  ctx.profile.refreshToken = data.refresh_token || ctx.profile.refreshToken;
  ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
  persist(ctx);
}
async function graphRequest(ctx, method, path, options = {}) {
  const base = (ctx.env.TEAMS_GRAPH_URL || "https://graph.microsoft.com/v1.0").replace(/\/$/, "");
  const url = new URL(options.rawUrl || `${base}${path.startsWith("/") ? path : `/${path}`}`);
  options.query?.forEach((value, key) => url.searchParams.append(key, value));
  const token = await getAccessToken(ctx);
  let response;
  try {
    response = await fetch(url, { method, headers: { Authorization: `Bearer ${token}`, Accept: "application/json", ...options.headers, ...options.body === void 0 ? {} : { "Content-Type": "application/json" } }, ...options.body === void 0 ? {} : { body: options.body } });
  } catch {
    throw new CliError(`Cannot reach ${url} for ${method} ${path}. Check the network and Graph URL.`);
  }
  if (!response.ok) {
    let detail = "";
    try {
      const data = await response.json();
      detail = typeof data.error === "string" ? data.error : [data.error?.code, data.error?.message ?? data.message].filter(Boolean).join(": ");
    } catch {
    }
    const hint = {
      401: `Authentication failed for ${method} ${path}. Run 'teams auth check' or 'teams auth login'.`,
      403: `Access denied for ${method} ${path}. Check that your token has the required Microsoft Graph scopes.`,
      404: `Not found: ${method} ${path}. Check the id and path.`,
      429: `Rate limited for ${method} ${path}. Wait, then try again.`
    };
    throw new CliError(`${hint[response.status] || `HTTP ${response.status} for ${method} ${path}.`}${detail ? ` Graph said: ${detail}` : ""}`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : void 0;
}
async function runAuth(ctx, subcommand) {
  if (subcommand === "login") return login(ctx);
  if (subcommand === "check") {
    const me = await graphRequest(ctx, "GET", "/me");
    process.stdout.write(`Signed in as ${me.displayName || me.userPrincipalName || me.mail || "current user"}
`);
    return;
  }
  if (subcommand === "logout") {
    delete ctx.profile.accessToken;
    delete ctx.profile.refreshToken;
    delete ctx.profile.expiresAt;
    persist(ctx);
    process.stdout.write(`Signed out of Profile '${ctx.profileName}'.
`);
    return;
  }
  usage("Use 'teams auth login', 'teams auth check', or 'teams auth logout'.");
}
async function login(ctx) {
  const id2 = clientId(ctx);
  const deviceUrl = `${loginBase(ctx)}/common/oauth2/v2.0/devicecode`;
  let deviceResponse;
  try {
    deviceResponse = await fetch(deviceUrl, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: id2, scope: SCOPES }) });
  } catch {
    throw new CliError(`Cannot reach ${deviceUrl}. Check the login URL and network.`);
  }
  const device = await deviceResponse.json();
  if (!deviceResponse.ok || !device.device_code) throw new CliError("Could not start device-code sign-in. Check the client id and try again.");
  process.stdout.write(`To sign in, visit ${device.verification_uri} and enter code ${device.user_code}.
`);
  const interval = Math.max(1, device.interval || 5) * 1e3;
  const expires = Date.now() + (device.expires_in || 900) * 1e3;
  while (Date.now() < expires) {
    const response = await fetch(tokenEndpoint(ctx), { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams({ client_id: id2, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: device.device_code }) });
    const data = await response.json();
    if (response.ok && data.access_token) {
      ctx.profile.accessToken = data.access_token;
      ctx.profile.refreshToken = data.refresh_token;
      ctx.profile.clientId = id2;
      ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
      persist(ctx);
      process.stdout.write(`Signed in to Profile '${ctx.profileName}'.
`);
      return;
    }
    if (data.error === "authorization_pending") {
      await new Promise((resolve2) => setTimeout(resolve2, interval));
      continue;
    }
    if (data.error === "slow_down") {
      await new Promise((resolve2) => setTimeout(resolve2, interval + 5e3));
      continue;
    }
    throw new CliError("Device-code sign-in failed. Run 'teams auth login' to try again.");
  }
  throw new CliError("Device code expired. Run 'teams auth login' to try again.");
}

// src/commands/team-commands.ts
function label(item) {
  return item.displayName ?? item.name ?? "";
}
function id(item) {
  return item.id ?? "";
}
function resolve(items, value, kind) {
  const byId = items.find((item) => id(item) === value);
  if (byId) return byId;
  const exact = items.filter((item) => label(item) === value);
  const matches = exact.length ? exact : items.filter((item) => label(item).toLocaleLowerCase() === value.toLocaleLowerCase());
  if (matches.length === 1) return matches[0];
  const candidates = (matches.length ? matches : items).map((item) => `${label(item)} (${id(item)})`).join(", ");
  throw new CliError(`${matches.length ? "Ambiguous" : "No matching"} ${kind} '${value}'. Candidates: ${candidates || "(none)"}.`);
}
function table(items, columns) {
  const widths = columns.map(([title, value]) => Math.max(title.length, ...items.map((item) => value(item).length)));
  const row = (values) => values.map((value, i) => value.padEnd(widths[i])).join("  ").trimEnd();
  return `${row(columns.map(([title]) => title))}
${items.map((item) => row(columns.map(([, value]) => value(item)))).join("\n")}${items.length ? "\n" : ""}`;
}
function output(data, json, columns) {
  if (json) process.stdout.write(`${JSON.stringify(data, null, 2)}
`);
  else process.stdout.write(table(data.value ?? [], columns));
}
async function teams(ctx, json) {
  const data = await graphRequest(ctx, "GET", "/me/joinedTeams");
  output(data, json, [["ID", id], ["NAME", label]]);
}
async function resolveTeam(ctx, value) {
  const data = await graphRequest(ctx, "GET", "/me/joinedTeams");
  return resolve(data.value ?? [], value, "Team");
}
async function resolveChannel(ctx, teamId, value) {
  const data = await graphRequest(ctx, "GET", `/teams/${encodeURIComponent(teamId)}/channels`);
  return resolve(data.value ?? [], value, "Channel");
}
async function channels(ctx, teamValue, json) {
  const team = await resolveTeam(ctx, teamValue);
  const data = await graphRequest(ctx, "GET", `/teams/${encodeURIComponent(id(team))}/channels`);
  output(data, json, [["ID", id], ["NAME", label], ["TYPE", (item) => String(item.membershipType ?? "")]]);
}
async function channel(ctx, teamValue, channelValue, json) {
  const team = await resolveTeam(ctx, teamValue);
  const chan = await resolveChannel(ctx, id(team), channelValue);
  const data = await graphRequest(ctx, "GET", `/teams/${encodeURIComponent(id(team))}/channels/${encodeURIComponent(id(chan))}`);
  if (json) process.stdout.write(`${JSON.stringify(data, null, 2)}
`);
  else process.stdout.write(`ID  NAME  TYPE  LAYOUT
${id(data)}  ${label(data)}  ${String(data.membershipType ?? "")}  ${String(data.layout ?? "")}
`);
}
async function runTeamsCommand(ctx, command, args) {
  const json = args.includes("--json");
  const positional = args.filter((arg) => arg !== "--json");
  if (command === "teams" && positional.length === 0) return teams(ctx, json);
  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  if (command === "channels") {
    if (positional.length > 1) usage("Use 'teams channels [team] [--json]'.");
    const selected = positional[0] || teamValue;
    if (!selected) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    return channels(ctx, selected, json);
  }
  if (command === "channel") {
    if (positional.length > 2) usage("Use 'teams channel [channel] [--team team] [--json]'.");
    const selectedTeam = positional.length === 2 ? positional[0] : teamValue;
    const selectedChannel = positional.length === 2 ? positional[1] : positional[0] || ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!selectedTeam) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!selectedChannel) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    return channel(ctx, selectedTeam, selectedChannel, json);
  }
  throw new CliError(`Invalid ${command} arguments.`, 2);
}

// src/commands/config-commands.ts
async function runConfig(ctx, args) {
  if (args.length === 1 && args[0] === "show") {
    process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}
`);
    return;
  }
  if (args.length === 3 && args[0] === "set" && args[1] === "team") {
    const team = await resolveTeam(ctx, args[2]);
    ctx.profile.defaultTeam = team.id;
    ctx.config.profiles[ctx.profileName] = ctx.profile;
    ctx.config.defaultProfile ??= ctx.profileName;
    saveConfig(ctx.path, ctx.config);
    process.stdout.write(`Default Team set to ${team.displayName ?? team.id}.
`);
    return;
  }
  if (args.length === 3 && args[0] === "set" && args[1] === "channel") {
    const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    if (!teamValue) usage("Choose a Team with --team, TEAMS_TEAM, or 'teams config set team' first.");
    const team = await resolveTeam(ctx, teamValue);
    const channel2 = await resolveChannel(ctx, team.id, args[2]);
    ctx.profile.defaultChannel = channel2.id;
    ctx.config.profiles[ctx.profileName] = ctx.profile;
    ctx.config.defaultProfile ??= ctx.profileName;
    saveConfig(ctx.path, ctx.config);
    process.stdout.write(`Default Channel set to ${channel2.displayName ?? channel2.id}.
`);
    return;
  }
  usage("Use 'teams config show', 'teams config set team <name|id>', or 'teams config set channel <name|id>'.");
}

// src/commands/api-commands.ts
import { readFile } from "node:fs/promises";
async function runApi(ctx, args) {
  const method = args[0]?.toUpperCase();
  const path = args[1];
  if (!method || !path) usage("Use 'teams api <METHOD> <path> [key=value ...]'.");
  if (!["GET", "POST", "PUT", "PATCH", "DELETE"].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);
  const query = new URLSearchParams();
  const headers = {};
  let body;
  let all = false;
  for (let i = 2; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--all") {
      all = true;
      continue;
    }
    if (arg === "--header" || arg.startsWith("--header=")) {
      const value = arg === "--header" ? args[++i] : arg.slice(9);
      const colon = value?.indexOf(":") ?? -1;
      if (!value || colon < 1) usage("Option '--header' requires 'Name: value'.");
      headers[value.slice(0, colon).trim()] = value.slice(colon + 1).trim();
      continue;
    }
    if (arg === "--data" || arg.startsWith("--data=")) {
      if (body !== void 0) usage("Option '--data' may only be provided once.");
      const value = arg === "--data" ? args[++i] : arg.slice(7);
      if (value === void 0) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = value === "-" ? await readStdin() : value.startsWith("@") ? await readFile(value.slice(1), "utf8") : value;
      try {
        body = JSON.stringify(JSON.parse(text));
      } catch {
        usage("Option '--data' must contain valid JSON.");
      }
      continue;
    }
    const equal = arg.indexOf("=");
    if (equal < 1) usage(`Expected a query pair in key=value form: ${arg}.`);
    query.append(arg.slice(0, equal), arg.slice(equal + 1));
  }
  if (all && method !== "GET") usage("Option '--all' is only supported with GET requests.");
  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, data);
  if (data !== void 0) process.stdout.write(`${JSON.stringify(data)}
`);
}
async function collectPages(ctx, initial) {
  if (!isCollection(initial)) return initial;
  const result = { ...initial, value: [...initial.value] };
  let next = initial["@odata.nextLink"];
  for (let count = 1; typeof next === "string" && count < 50; count += 1) {
    const url = validateLink(ctx, next);
    const page = await graphRequest(ctx, "GET", url.href, { rawUrl: url.href });
    if (!isCollection(page)) break;
    result.value = [...result.value, ...page.value];
    next = page["@odata.nextLink"];
  }
  delete result["@odata.nextLink"];
  return result;
}
function isCollection(value) {
  return typeof value === "object" && value !== null && Array.isArray(value.value);
}
function validateLink(ctx, link) {
  const base = new URL(ctx.env.TEAMS_GRAPH_URL || "https://graph.microsoft.com/v1.0");
  let url;
  try {
    url = new URL(link, base);
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const prefix = base.pathname.replace(/\/$/, "");
  if (url.origin !== base.origin || prefix && url.pathname !== prefix && !url.pathname.startsWith(`${prefix}/`)) usage("Graph returned a continuation URL outside the configured Graph URL.");
  return url;
}
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}

// src/commands/search-commands.ts
async function runSearch(ctx, args) {
  let size = 25;
  let json = false;
  const terms = [];
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--size" || arg.startsWith("--size=")) {
      const value = arg === "--size" ? args[++i] : arg.slice(7);
      if (!value || !/^\d+$/.test(value) || Number(value) < 1) usage("Option '--size' requires a positive integer.");
      size = Number(value);
      continue;
    }
    if (arg.startsWith("-")) usage(`Unknown search option: ${arg}.`);
    terms.push(arg);
  }
  const query = terms.join(" ").trim();
  if (!query) usage("Use 'teams search <query> [--size N] [--json]'.");
  const result = await graphRequest(ctx, "POST", "/search/query", { body: JSON.stringify({
    requests: [{ entityTypes: ["chatMessage"], query: { queryString: query }, from: 0, size, enableTopResults: true }]
  }) });
  if (json) {
    process.stdout.write(`${JSON.stringify(result, null, 2)}
`);
    return;
  }
  const hits = searchHits(result);
  const rows = hits.map((hit) => {
    const resource = hit.resource ?? {};
    const teamId = resource.channelIdentity?.teamId;
    const channelId = resource.channelIdentity?.channelId;
    const kind = teamId && channelId ? "Channel" : "Chat";
    return [kind, teamId ?? "-", channelId ?? "-", resource.id ?? "-", resource.replyToId ?? "-", resource.from?.user?.displayName ?? "-", resource.createdDateTime ?? "-", plainText(resource.body?.content ?? hit.summary ?? "")];
  });
  process.stdout.write(`TYPE  TEAM  CHANNEL  MESSAGE  ROOT  SENDER  TIME  BODY
`);
  for (const row of rows) process.stdout.write(`${row.map((value) => value.replaceAll("\n", " ").slice(0, 80)).join("  ")}
`);
}
function searchHits(value) {
  if (!value || typeof value !== "object") return [];
  const containers = value.value;
  return containers?.[0]?.hitsContainers?.[0]?.hits ?? [];
}
function plainText(value) {
  return value.replace(/<[^>]*>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/\s+/g, " ").trim();
}

// src/commands/message-commands.ts
var DEFAULT_TOP = 20;
var PAGE_CAP = 100;
var BODY_WIDTH = 80;
var REPLY_LIMIT = 5;
function table2(headers, rows) {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values) => values.map((value, index) => value.padEnd(widths[index])).join("  ").trimEnd();
  return `${line(headers)}
${rows.map(line).join("\n")}${rows.length ? "\n" : ""}`;
}
function option(args, name) {
  const index = args.indexOf(name);
  if (index < 0) return void 0;
  const value = args[index + 1];
  if (!value || value.startsWith("-")) usage(`Option '${name}' requires a value.`);
  return value;
}
function plainText2(content) {
  return (content ?? "").replace(/<br\s*\/?>/gi, " ").replace(/<\/(p|div|li|tr)>/gi, " ").replace(/<[^>]*>/g, "").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'").replace(/\s+/g, " ").trim();
}
function truncate(value, max = BODY_WIDTH) {
  return value.length > max ? `${value.slice(0, max - 1)}\u2026` : value;
}
function validateContinuation(link, ctx) {
  const base = ctx.env.TEAMS_GRAPH_URL || "https://graph.microsoft.com/v1.0";
  let url;
  try {
    url = new URL(link, base);
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const configured = new URL(base);
  const basePath = configured.pathname.replace(/\/$/, "");
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== configured.origin || !inBasePath) usage("Graph returned a continuation URL outside the configured Graph URL.");
}
function cursorUrl(token, ctx) {
  let link;
  try {
    link = Buffer.from(token, "base64url").toString("utf8");
  } catch {
    usage("Invalid cursor. Copy the cursor from the previous command output.");
  }
  validateContinuation(link, ctx);
  return link;
}
function graphPage(ctx, url) {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, "GET", url, { rawUrl: url }) : graphRequest(ctx, "GET", url);
}
function printContinuation(link) {
  if (link) process.stderr.write(`more available: --cursor ${Buffer.from(link).toString("base64url")} (or --all)
`);
}
function parseTop(args) {
  const value = option(args, "--top");
  if (value === void 0) return DEFAULT_TOP;
  const top = Number(value);
  if (!Number.isSafeInteger(top) || top < 1 || top > 50) usage("Option '--top' must be an integer from 1 to 50.");
  return top;
}
async function runMessages(ctx, args) {
  const valueOptions = /* @__PURE__ */ new Set(["--top", "--cursor"]);
  const flags = /* @__PURE__ */ new Set(["--all", "--full", "--json", "--replies"]);
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (valueOptions.has(arg)) {
      option(args, arg);
      index += 1;
    } else if (!flags.has(arg)) usage(`Unknown messages option or argument '${arg}'.`);
  }
  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
  if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
  if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
  const team = await resolveTeam(ctx, teamValue);
  const channel2 = await resolveChannel(ctx, team.id, channelValue);
  const path = `/teams/${encodeURIComponent(team.id)}/channels/${encodeURIComponent(channel2.id)}/messages`;
  const top = parseTop(args);
  const cursor = option(args, "--cursor");
  let url = cursor ? cursorUrl(cursor, ctx) : `${path}?$top=${top}`;
  const all = args.includes("--all");
  const pages = [];
  let nextLink;
  for (let page = 0; page < (all ? PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url);
    pages.push(result);
    nextLink = result["@odata.nextLink"];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const messages = pages.flatMap((page) => page.value ?? []);
  const response = all ? { ...first, value: messages, ...nextLink ? { "@odata.nextLink": nextLink } : {} } : first;
  if (all && !nextLink) delete response["@odata.nextLink"];
  const replyPreviews = /* @__PURE__ */ new Map();
  if (args.includes("--replies") && !args.includes("--json")) {
    for (const message of messages) {
      if (!message.id) continue;
      const replyData = await graphRequest(ctx, "GET", `${path}/${encodeURIComponent(message.id)}/replies`, { query: new URLSearchParams({ "$top": String(REPLY_LIMIT) }) });
      const replies = (replyData.value ?? []).slice().sort((a, b) => (a.createdDateTime ?? "").localeCompare(b.createdDateTime ?? "")).slice(-REPLY_LIMIT);
      const preview = replies.map((reply) => `${reply.from?.user?.displayName ?? "Unknown"}: ${plainText2(reply.body?.content)}`).filter((value) => !value.endsWith(": ")).join(" | ");
      if (preview) replyPreviews.set(message.id, truncate(preview));
    }
  }
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    printContinuation(nextLink);
    return;
  }
  const full = args.includes("--full");
  const rows = messages.map((message) => {
    const row = [message.id ?? "", message.createdDateTime ?? "", message.from?.user?.displayName ?? "Unknown", full ? plainText2(message.body?.content) : truncate(plainText2(message.body?.content))];
    if (args.includes("--replies")) row.push(replyPreviews.get(message.id ?? "") ?? "");
    return row;
  });
  process.stdout.write(table2(["ID", "TIME", "SENDER", "BODY", ...args.includes("--replies") ? ["REPLIES"] : []], rows));
  printContinuation(nextLink);
}

// src/commands/send-command.ts
import { readFile as readFile2 } from "node:fs/promises";
function escapeHtml(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/\r\n|\r|\n/g, "<br>");
}
async function stdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}
async function runSend(ctx, args) {
  const positional = [];
  const mentions = [];
  let bodyOption;
  let subject;
  let html = false;
  let dryRun = false;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--html") {
      html = true;
      continue;
    }
    if (arg === "--dry-run") {
      dryRun = true;
      continue;
    }
    if (arg === "--body" || arg.startsWith("--body=")) {
      if (bodyOption !== void 0) usage("Option '--body' may only be provided once.");
      bodyOption = arg === "--body" ? args[++i] : arg.slice(7);
      if (bodyOption === void 0) usage("Option '--body' requires text, @file, or '-'.");
      continue;
    }
    if (arg === "--subject" || arg.startsWith("--subject=")) {
      if (subject !== void 0) usage("Option '--subject' may only be provided once.");
      subject = arg === "--subject" ? args[++i] : arg.slice(10);
      if (subject === void 0) usage("Option '--subject' requires a value.");
      continue;
    }
    if (arg === "--mention" || arg.startsWith("--mention=")) {
      const value = arg === "--mention" ? args[++i] : arg.slice(10);
      if (!value) usage("Option '--mention' requires a user principal name or email.");
      mentions.push(value);
      continue;
    }
    if (arg.startsWith("-") && arg !== "-") usage(`Unknown send option: ${arg}. To post text that starts with '-', pass it with --body.`);
    positional.push(arg);
  }
  if (bodyOption !== void 0 && positional.length) usage("Provide the message body as an argument or with '--body', not both.");
  if (bodyOption === void 0 && positional.length === 0) usage("Use 'teams send <text> [--body @file|-] [--subject text] [--mention upn] [--html] [--dry-run]'.");
  const source = bodyOption ?? positional.join(" ");
  const raw = source === "-" ? await stdin() : source.startsWith("@") ? await readFile2(source.slice(1), "utf8") : source;
  const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
  const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
  if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
  if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
  const team = await resolveTeam(ctx, teamValue);
  const teamId = team.id ?? "";
  const channel2 = await resolveChannel(ctx, teamId, channelValue);
  const payload2 = {
    ...subject === void 0 ? {} : { subject },
    body: { contentType: "html", content: html ? raw : escapeHtml(raw) }
  };
  if (mentions.length) {
    const resolved = [];
    const tags = [];
    for (const upn of mentions) {
      const user = await graphRequest(ctx, "GET", `/users/${encodeURIComponent(upn)}`);
      const userId = user.id ?? "";
      const displayName = user.displayName ?? upn;
      const id2 = resolved.length;
      resolved.push({ id: id2, mentionText: displayName, mentioned: { user: { id: userId, displayName, userIdentityType: "aadUser" } } });
      tags.push(`<at id="${id2}">${escapeHtml(displayName)}</at>`);
    }
    payload2.body.content = `${tags.join(" ")}${raw ? ` ${payload2.body.content}` : ""}`;
    payload2.mentions = resolved;
  }
  if (dryRun) {
    process.stdout.write(`${JSON.stringify(payload2, null, 2)}
`);
    return;
  }
  const result = await graphRequest(ctx, "POST", `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channel2.id ?? "")}/messages`, { body: JSON.stringify(payload2) });
  process.stdout.write(`${JSON.stringify(result)}
`);
}

// src/utils/links.ts
var TEAMS_HOSTS = /* @__PURE__ */ new Set(["teams.microsoft.com", "teams.cloud.microsoft"]);
var ID = /^[A-Za-z0-9._~=-]{1,512}$/;
var CHANNEL_ID = /^[A-Za-z0-9._~=:@-]{1,512}$/;
function parseLink(value) {
  let url;
  try {
    url = new URL(value);
  } catch {
    return void 0;
  }
  if (url.protocol !== "https:" || !TEAMS_HOSTS.has(url.hostname) || url.port || url.username || url.password || url.hash) return void 0;
  const match = /^\/l\/message\/([^/]+)\/([^/]+)$/.exec(url.pathname);
  if (!match) return void 0;
  let channelId;
  let messageId2;
  try {
    channelId = decodeURIComponent(match[1]);
    messageId2 = decodeURIComponent(match[2]);
  } catch {
    return void 0;
  }
  const groupIds = url.searchParams.getAll("groupId");
  const teamId = groupIds.length === 1 ? groupIds[0] : "";
  if (!ID.test(teamId) || !CHANNEL_ID.test(channelId) || !ID.test(messageId2)) return void 0;
  const parents = url.searchParams.getAll("parentMessageId");
  if (parents.length > 1 || parents[0] !== void 0 && !ID.test(parents[0])) return void 0;
  return { teamId, channelId, messageId: messageId2, ...parents[0] && parents[0] !== messageId2 ? { parentId: parents[0] } : {} };
}

// src/commands/thread-command.ts
var DEFAULT_TOP2 = 20;
function validId(value) {
  return ID.test(value);
}
function messageId(value) {
  const link = parseLink(value);
  if (link) return link.messageId;
  if (/^https?:\/\//i.test(value) || !validId(value)) usage("Expected a valid Channel Message id or a Teams message link.");
  return value;
}
function option2(args, name) {
  const index = args.indexOf(name);
  if (index < 0) return void 0;
  const value = args[index + 1];
  if (!value || value.startsWith("-")) usage(`Option '${name}' requires a value.`);
  return value;
}
function validateContinuation2(link, ctx) {
  const base = ctx.env.TEAMS_GRAPH_URL || "https://graph.microsoft.com/v1.0";
  let url;
  try {
    url = new URL(link, base);
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const configured = new URL(base);
  const basePath = configured.pathname.replace(/\/$/, "");
  if (url.origin !== configured.origin || basePath && url.pathname !== basePath && !url.pathname.startsWith(`${basePath}/`)) {
    usage("Graph returned a continuation URL outside the configured Graph URL.");
  }
}
function printContinuation2(link) {
  if (link) process.stderr.write(`more available: --cursor ${Buffer.from(link).toString("base64url")}
`);
}
function plainText3(content) {
  return (content ?? "").replace(/<br\s*\/?\s*>/gi, " ").replace(/<\/[^>]+>/g, " ").replace(/<[^>]*>/g, "").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'").replace(/\s+/g, " ").trim();
}
function table3(rows) {
  const headers = ["ID", "TIME", "SENDER", "BODY"];
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values) => values.map((value, index) => value.padEnd(widths[index])).join("  ").trimEnd();
  return `${line(headers)}
${rows.map(line).join("\n")}
`;
}
async function runThread(ctx, args) {
  const values = /* @__PURE__ */ new Set(["--top", "--cursor"]);
  const flags = /* @__PURE__ */ new Set(["--full", "--json"]);
  const positional = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (values.has(arg)) {
      option2(args, arg);
      index += 1;
    } else if (flags.has(arg)) continue;
    else if (arg.startsWith("-")) usage(`Unknown thread option '${arg}'.`);
    else positional.push(arg);
  }
  if (positional.length !== 1) usage("Use 'teams thread <id|link> [--top N] [--cursor token] [--full] [--json]'.");
  const supplied = positional[0];
  const link = parseLink(supplied);
  if (/^https?:\/\//i.test(supplied) && !link) usage("Expected a valid Teams message link.");
  const id2 = link?.messageId ?? messageId(supplied);
  let teamValue = link?.teamId ?? "";
  let channelValue = link?.channelId ?? "";
  if (!link) {
    const teamName = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    const channelName = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!teamName) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!channelName) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    const team = await resolveTeam(ctx, teamName);
    teamValue = team.id ?? "";
    channelValue = (await resolveChannel(ctx, teamValue, channelName)).id ?? "";
  }
  const topValue = option2(args, "--top");
  const top = topValue === void 0 ? DEFAULT_TOP2 : Number(topValue);
  if (!Number.isSafeInteger(top) || top < 1 || top > 50) usage("Option '--top' must be an integer from 1 to 50.");
  const cursor = option2(args, "--cursor");
  const base = `/teams/${encodeURIComponent(teamValue)}/channels/${encodeURIComponent(channelValue)}/messages`;
  const rootId = link?.parentId ?? id2;
  let root;
  try {
    root = await graphRequest(ctx, "GET", `${base}/${encodeURIComponent(rootId)}`);
  } catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) {
      usage("Not found: the id is not a Root Message in this Channel, the Team or Channel is wrong, or the message was deleted. A Reply id cannot be read on its own: pass its Root Message id, or paste a Reply link that carries parentMessageId.");
    }
    throw error;
  }
  let repliesUrl;
  if (cursor) {
    if (!/^[A-Za-z0-9_-]+$/.test(cursor)) usage("Invalid cursor. Copy the cursor from the previous command output.");
    const decoded = Buffer.from(cursor, "base64url").toString("utf8");
    validateContinuation2(decoded, ctx);
    repliesUrl = decoded;
  } else {
    repliesUrl = `${base}/${encodeURIComponent(rootId)}/replies?$top=${top}`;
  }
  const pages = [];
  let nextLink;
  const result = await graphRequest(ctx, "GET", repliesUrl, /^https?:\/\//i.test(repliesUrl) ? { rawUrl: repliesUrl } : {});
  pages.push(result);
  nextLink = result["@odata.nextLink"];
  const pageReplies = (pages[0]?.value ?? []).slice().sort((a, b) => (a.createdDateTime ?? "").localeCompare(b.createdDateTime ?? ""));
  const response = { root, replies: pageReplies, ...nextLink ? { "@odata.nextLink": nextLink } : {} };
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    printContinuation2(nextLink);
    return;
  }
  const all = [root, ...pageReplies];
  const rows = all.map((message) => [message.id ?? "", message.createdDateTime ?? "", message.from?.user?.displayName ?? "Unknown", plainText3(message.body?.content)]);
  process.stdout.write(table3(rows));
  printContinuation2(nextLink);
}

// src/commands/reply-edit-command.ts
import { readFile as readFile3 } from "node:fs/promises";
function escapeHtml2(text) {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;").replace(/\r\n|\r|\n/g, "<br>");
}
function parseId(value) {
  const link = parseLink(value);
  if (link) return link.messageId;
  if (/^https?:\/\//i.test(value) || !ID.test(value)) usage("Expected a valid Channel Message id or a Teams message link.");
  return value;
}
async function stdin2() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}
async function payload(ctx, args, allowHtml) {
  const positional = [];
  const mentions = [];
  let bodyOption;
  let html = false;
  let dryRun = false;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--html" && allowHtml) {
      html = true;
      continue;
    }
    if (arg === "--dry-run") {
      dryRun = true;
      continue;
    }
    if (arg === "--body" || arg.startsWith("--body=")) {
      if (bodyOption !== void 0) usage("Option '--body' may only be provided once.");
      bodyOption = arg === "--body" ? args[++i] : arg.slice(7);
      if (bodyOption === void 0) usage("Option '--body' requires text, @file, or '-'.");
      continue;
    }
    if (arg === "--mention" || arg.startsWith("--mention=")) {
      const value = arg === "--mention" ? args[++i] : arg.slice(10);
      if (!value) usage("Option '--mention' requires a user principal name or email.");
      mentions.push(value);
      continue;
    }
    if (arg.startsWith("-") && arg !== "-") usage(`Unknown message option: ${arg}. To post text that starts with '-', pass it with --body.`);
    positional.push(arg);
  }
  if (bodyOption !== void 0 && positional.length) usage("Provide the message body as an argument or with '--body', not both.");
  if (bodyOption === void 0 && positional.length === 0) usage("Provide message text as an argument or with '--body @file|-'.");
  const source = bodyOption ?? positional.join(" ");
  const raw = source === "-" ? await stdin2() : source.startsWith("@") ? await readFile3(source.slice(1), "utf8") : source;
  const result = { body: { contentType: "html", content: html ? raw : escapeHtml2(raw) } };
  if (mentions.length) {
    const resolved = [];
    const tags = [];
    for (const upn of mentions) {
      const user = await graphRequest(ctx, "GET", `/users/${encodeURIComponent(upn)}`);
      const id2 = resolved.length;
      const userId = user.id ?? "";
      const displayName = user.displayName ?? upn;
      resolved.push({ id: id2, mentionText: displayName, mentioned: { user: { id: userId, displayName, userIdentityType: "aadUser" } } });
      tags.push(`<at id="${id2}">${escapeHtml2(displayName)}</at>`);
    }
    result.body.content = `${tags.join(" ")}${raw ? ` ${result.body.content}` : ""}`;
    result.mentions = resolved;
  }
  if (dryRun) process.stdout.write(`${JSON.stringify(result, null, 2)}
`);
  return result;
}
async function resolveTarget(ctx, supplied, rootOption) {
  const link = parseLink(supplied);
  if (/^https?:\/\//i.test(supplied) && !link) usage("Expected a valid Teams Channel Message link; Chat links cannot be written.");
  const selectedId = link?.messageId ?? parseId(supplied);
  if (rootOption !== void 0 && !ID.test(rootOption)) usage("Option '--root' requires a valid Root Message id.");
  const rootHint = rootOption ?? link?.parentId;
  let teamId;
  let channelId;
  if (link) {
    teamId = link.teamId;
    channelId = link.channelId;
  } else {
    const teamValue = ctx.options.team || ctx.env.TEAMS_TEAM || ctx.profile.defaultTeam;
    const channelValue = ctx.options.channel || ctx.env.TEAMS_CHANNEL || ctx.profile.defaultChannel;
    if (!teamValue) usage("Set a Team with --team, TEAMS_TEAM, or 'teams config set team'.");
    if (!channelValue) usage("Set a Channel with --channel, TEAMS_CHANNEL, or 'teams config set channel'.");
    const team = await resolveTeam(ctx, teamValue);
    const channel2 = await resolveChannel(ctx, team.id ?? "", channelValue);
    teamId = team.id ?? "";
    channelId = channel2.id ?? "";
  }
  const path = `/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/messages`;
  const isReply = rootHint !== void 0 && rootHint !== selectedId;
  const rootId = isReply ? rootHint : selectedId;
  let root;
  try {
    root = await graphRequest(ctx, "GET", `${path}/${encodeURIComponent(rootId)}`);
  } catch (error) {
    if (error instanceof CliError && /Not found: GET /.test(error.message)) {
      usage("Not found: the id is not a Root Message in this Channel (Chat messages cannot be written), the Team or Channel is wrong, or the message was deleted. To target a Reply, also pass its Root Message id with --root.");
    }
    throw error;
  }
  return { path, rootId, selectedId, isReply, ...root.subject ? { subject: root.subject } : {} };
}
function takeRoot(args) {
  const rest = [];
  let root;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--root" || arg.startsWith("--root=")) {
      if (root !== void 0) usage("Option '--root' may only be provided once.");
      root = arg === "--root" ? args[++i] : arg.slice(7);
      if (!root) usage("Option '--root' requires a Root Message id.");
    } else rest.push(arg);
  }
  return { rest, root };
}
async function runReply(ctx, args) {
  const { rest, root } = takeRoot(args);
  if (rest.length < 2) usage("Use 'teams reply <id|link> <text> [--root <root id>] [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, rest[0], root);
  const body = await payload(ctx, rest.slice(1), true);
  if (rest.includes("--dry-run")) return;
  const result = await graphRequest(ctx, "POST", `${target.path}/${encodeURIComponent(target.rootId)}/replies`, { body: JSON.stringify(body) });
  process.stdout.write(`${JSON.stringify(result)}
`);
}
async function runEdit(ctx, args) {
  const { rest, root } = takeRoot(args);
  if (rest.length < 2) usage("Use 'teams edit <id|link> <text> [--root <root id>] [--body @file|-] [--mention upn] [--html] [--dry-run]'.");
  const target = await resolveTarget(ctx, rest[0], root);
  const body = await payload(ctx, rest.slice(1), true);
  if (rest.includes("--dry-run")) return;
  const endpoint = target.isReply ? `${target.path}/${encodeURIComponent(target.rootId)}/replies/${encodeURIComponent(target.selectedId)}` : `${target.path}/${encodeURIComponent(target.selectedId)}`;
  const patch = target.isReply || !target.subject ? body : { subject: target.subject, ...body };
  const result = await graphRequest(ctx, "PATCH", endpoint, { body: JSON.stringify(patch) });
  process.stdout.write(result === void 0 ? "Edited.\n" : `${JSON.stringify(result)}
`);
}

// src/index.ts
function globals(argv) {
  const options = {};
  const args = [];
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "-p" || arg === "--profile" || arg === "--client-id" || arg === "--team" || arg === "--channel") {
      const value = argv[i + 1];
      if (!value || value.startsWith("-")) usage(`Option '${arg}' requires a value.`);
      if (arg === "--client-id") options.clientId = value;
      else if (arg === "--team") options.team = value;
      else if (arg === "--channel") options.channel = value;
      else options.profile = value;
      i += 1;
    } else if (arg.startsWith("--profile=")) options.profile = arg.slice(10);
    else if (arg.startsWith("--client-id=")) options.clientId = arg.slice(12);
    else if (arg.startsWith("--team=")) options.team = arg.slice(7);
    else if (arg.startsWith("--channel=")) options.channel = arg.slice(10);
    else args.push(arg);
  }
  return { options, args };
}
async function main(argv) {
  if (!argv.length || argv.includes("--help") || argv.includes("-h")) {
    process.stdout.write(`${HELP}
`);
    return;
  }
  if (argv.length === 1 && (argv[0] === "--version" || argv[0] === "-v")) {
    process.stdout.write(`${VERSION}
`);
    return;
  }
  const { options, args } = globals(argv);
  const [command, sub, ...rest] = args;
  if (!command) usage("Run 'teams --help' for usage.");
  const ctx = selectedContext(process.env, options);
  if (command === "auth") return runAuth(ctx, sub);
  if (command === "config") return runConfig(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "api") return runApi(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "teams" || command === "channels" || command === "channel") return runTeamsCommand(ctx, command, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "search") return runSearch(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "messages") return runMessages(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "send") return runSend(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "reply") return runReply(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "edit") return runEdit(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "thread") return runThread(ctx, [sub, ...rest].filter((x) => x !== void 0));
  if (command === "me") {
    if (sub && sub !== "--json") usage("Use 'teams me [--json]'.");
    const identity = await graphRequest(ctx, "GET", "/me");
    if (sub === "--json") process.stdout.write(`${JSON.stringify(identity, null, 2)}
`);
    else process.stdout.write(`NAME  EMAIL
${identity.displayName ?? "Current user"}  ${identity.mail ?? identity.userPrincipalName ?? ""}
`);
    return;
  }
  throw new CliError(`Unknown command: ${command}. Run 'teams --help' for usage.`, 2);
}
main(process.argv.slice(2)).catch((error) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  process.stderr.write(`teams: ${error instanceof Error ? error.message : String(error)}
`);
  process.exitCode = code;
});
