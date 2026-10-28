export const HELP = `teams — Microsoft Teams Channel CLI

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
  search <query> [--size N] [--json]
  send <text> [--body @file|-] [--subject text] [--mention upn] [--html] [--dry-run]
  api <GET|POST|PUT|PATCH|DELETE> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]

Environment:
  TEAMS_PROFILE, TEAMS_TOKEN, TEAMS_CLIENT_ID, TEAMS_CONFIG
  TEAMS_TEAM, TEAMS_CHANNEL
  TEAMS_GRAPH_URL, TEAMS_LOGIN_URL

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
