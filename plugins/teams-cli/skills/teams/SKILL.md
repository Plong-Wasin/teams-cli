---
name: teams
description: Use the teams CLI to inspect Microsoft Teams Channels, messages, and search results, or to post, reply, or edit when the user explicitly asks.
---

# Teams CLI

Use `teams` for Microsoft Teams Channel Messages from a shell. The CLI uses Microsoft Graph and stores credentials per Profile.

## Before use

1. Check that the binary is installed with `teams --version`. If it is missing, install it with `npm install -g --install-links git+https://github.com/<owner>/<repo>.git#release` and check the version again.
2. Check the selected Profile with `teams auth check`. If it is not signed in, follow the user's environment's sign-in process with `teams auth login`; never request or expose a password or token in a conversation.
3. Use `-p <name>` to select a Profile. `teams config show` prints its settings with secrets masked.

## Safety

Stay read-only by default. Do not run `send`, `reply`, or `edit` unless the user explicitly instructs you to post, reply, or edit a Channel Message. Before a write, confirm the target Team, Channel, or Thread from the user's request and resolved CLI output. Use `--dry-run` when checking the exact payload. Never write to Chat; the CLI has no Chat write command.

## Common commands

```sh
teams teams
teams channels [team]
teams channel [channel] [--team team]
teams messages [--team team] [--channel channel] [--top N] [--cursor token] [--all] [--full] [--replies]
teams thread <id|link> [--top N] [--cursor token] [--full]
teams search <query>
teams me
teams config set team <name|id>
teams config set channel <name|id>
```

Add `--json` to supported read commands for raw Graph output. Names resolve to ids; if the CLI reports multiple candidates, choose using the user's stated target rather than guessing. Message links can be used with `thread`, `reply`, and `edit`. Graph cannot fetch a Reply by its own id: pass the Root Message id to `thread`/`reply`, and to edit a Reply pass its Root Message id with `--root` (or use the Reply's link, which carries `parentMessageId`). Search may return Chat results; treat these as read-only search results.

Write commands are `send`, `reply`, and `edit`. They support `--dry-run`; `send` also accepts `--subject`, while the write commands accept `--body @file|-`, repeatable `--mention`, and `--html`.

## Raw Graph access

Use `teams api <METHOD> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]` when there is no shortcut. This command can perform writes and has no dry-run option, so follow the explicit-user-instruction rule and inspect the method and path carefully before running it.

## Exit codes

- `0`: success
- `1`: server or authentication error
- `2`: usage error
