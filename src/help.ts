export const HELP = `teams — Microsoft Teams Channel CLI

Usage: teams [options] <command>

Options:
  -p, --profile <name>    Select a Profile
  --client-id <id>        Microsoft public client id for sign-in
  -h, --help              Show help
  -v, --version           Show version

Commands:
  auth login | check | logout
  config show
  me
  search <query> [--size N] [--json]
  api <GET|POST|PUT|PATCH|DELETE> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]

Environment:
  TEAMS_PROFILE, TEAMS_TOKEN, TEAMS_CLIENT_ID, TEAMS_CONFIG
  TEAMS_GRAPH_URL, TEAMS_LOGIN_URL

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
