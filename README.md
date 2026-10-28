# Teams CLI

`teams` is a command-line client for Microsoft Teams Channel Messages. It calls Microsoft Graph directly, so you can read and post from a shell, a script, or a headless job. The CLI covers Teams and Channels, Channel Messages and Threads, search, posting, replying, editing, and a raw Graph API command. Chat is not a command target; search can return Chat results, which are labeled.

## Install

Requires Node.js 20 or newer. Install the latest build from the `release` branch:

```sh
npm install -g --install-links git+https://github.com/<owner>/<repo>.git#release
teams --version
```

The first install is available after the repository's first successful release workflow run. Re-run the install command to update. The binary is not bundled with the Claude Code or Codex plugin, so install it separately.

To install from a checkout:

```sh
pnpm install
pnpm build
pnpm link --global
```

## Sign in

Start device-code sign-in and follow the printed URL and code:

```sh
teams auth login
teams auth check
```

Sign-in credentials are stored in `~/.config/teams/config.json` with file mode `0600`. `auth check` verifies the current Profile by calling Graph. Remove saved credentials with `teams auth logout`.

For a headless environment, provide `TEAMS_TOKEN`; it is used without creating a Profile. A token with a leading `Bearer ` prefix is accepted. To override the public client id used for sign-in, set `TEAMS_CLIENT_ID` or pass `--client-id`.

## Profiles and defaults

Profiles keep sign-in credentials and optional default Team and Channel selections separate:

```sh
teams auth login
teams -p work auth login
teams -p work config set team "Engineering"
teams -p work config set channel "General"
teams -p work config show
teams -p work messages
```

The first Profile is selected by default. Use `-p <name>` to select another. `config show` masks access and refresh tokens. Team and Channel defaults can be names or ids; the CLI resolves names and reports candidates when a name is missing or ambiguous.

For one-off selection, pass `--team` and `--channel`. Settings take precedence in this order: command-line flags, environment variables, then Profile defaults.

## Commands

Read commands are safe to use for inspection. `send`, `reply`, and `edit` write Channel Messages; use them only when the user explicitly instructs you to post, reply, or edit. Each supports `--dry-run` to print the payload without sending it. There is no delete command and Chat writing is not supported.

Microsoft Graph cannot fetch a Reply by its own id, so `thread`, `reply`, and `edit` take the Root Message id; a Reply is addressed with `--root` or a link with `parentMessageId`.

| Command | Purpose |
| --- | --- |
| `teams auth login\|check\|logout` | Sign in, verify credentials, or remove them |
| `teams config show` | Show Profile settings with secrets masked |
| `teams config set team <name\|id>` | Set the default Team |
| `teams config set channel <name\|id>` | Set the default Channel |
| `teams me` | Show the signed-in account |
| `teams teams` | List joined Teams |
| `teams channels [team]` | List a Team's Channels |
| `teams channel [channel]` | Show Channel details and layout |
| `teams messages` | List recent Root Messages; supports `--top`, `--cursor`, `--all`, `--full`, `--replies` |
| `teams thread <id\|link>` | Show a Root Message and its Replies oldest-first |
| `teams search <query>` | Search Channel Messages and Chat results |
| `teams send <text>` | Post a Root Message; supports `--subject`, `--body @file\|-`, `--mention`, `--html`, `--dry-run` |
| `teams reply <id\|link> <text>` | Reply in a Thread. `<id>` is the Root Message id. |
| `teams edit <id\|link> <text> [--root <id>]` | Edit a Root Message, keeping its subject. To edit a Reply, also pass its Root Message id with `--root`, or paste the Reply's Teams link, which carries `parentMessageId`. |
| `teams api <METHOD> <path>` | Send a raw Graph request with optional query pairs, `--data`, `--header`, and `--all` |

Add `--json` to list and detail commands for the Graph response. `messages` prints a `more available` cursor hint to stderr when another page exists; pass its cursor back with `--cursor` or use `--all`. Message ids can be supplied as ids or pasted Teams message links. Plain text in write commands is HTML-escaped and line breaks are preserved; use `--html` to send raw HTML. Message bodies can also come from a file (`--body @path`) or stdin (`--body -`). Repeat `--mention user@example.com` to mention people.

Examples:

```sh
teams teams
teams channels "Engineering"
teams messages --team "Engineering" --channel "General" --top 10 --full
teams thread '<message-link>'
teams search 'quarterly review'
teams api GET /me
teams send "Release notes are ready" --dry-run
```

## Environment variables

Flags override these variables, and environment values override Profile defaults:

| Variable | Purpose |
| --- | --- |
| `TEAMS_PROFILE` | Selected Profile |
| `TEAMS_TOKEN` | Headless bearer token, no Profile required |
| `TEAMS_CLIENT_ID` | OAuth public client id |
| `TEAMS_CONFIG` | Config file path |
| `TEAMS_TEAM` | Default Team name or id for this process |
| `TEAMS_CHANNEL` | Default Channel name or id for this process |
| `TEAMS_GRAPH_URL` | Microsoft Graph base URL |
| `TEAMS_LOGIN_URL` | OAuth login service base URL |

Exit codes are `0` for success, `1` for server or authentication errors, and `2` for usage errors.

## Agent plugins

The repository includes a `teams` skill for Claude Code and Codex. Replace `<owner>/<repo>` below with this repository's GitHub location:

```text
Claude Code:
/plugin marketplace add <owner>/<repo>
/plugin install teams-cli@teams-cli

Codex:
codex plugin marketplace add https://github.com/<owner>/<repo>.git
codex plugin add teams-cli@teams-cli
```

Install the CLI separately using the command above. The skill directs agents to check `teams --version` and `teams auth check` before use and to write only when the user explicitly instructs them to.

## Development

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm test
pnpm build
```

Tests invoke the bundled CLI against a local fake Graph server, with temporary configuration and environment state. They never call Microsoft Graph or use the real config. The fake server checks the requests made by the CLI, but it cannot prove that real Microsoft Graph responses have the shapes assumed by the implementation. The user should perform a real-Graph check before removing the `teams` MCP server from Claude Code or Codex configuration.
