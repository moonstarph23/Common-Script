# Agent Instructions

## GitHub Credentials

- Use GitHub CLI (`gh`) for authenticated GitHub operations.
- Confirm the active session before authenticated operations:

  ```bash
  gh auth status
  ```

- If authentication is required, use `gh auth login`. Do not manually read or export tokens from GitHub CLI's configuration.
- Prefer `gh` commands for GitHub resources such as repositories, pull requests, issues, checks, and releases.
- For authenticated Git commands, let GitHub CLI supply credentials through `gh auth git-credential`; never retrieve a token with `gh auth token` for use in a command.
- Do not place usernames, passwords, personal access tokens, or credential output in remote URLs, command arguments, patches, commits, logs, or responses.
- Redact credentials if an authentication error or remote URL unexpectedly includes them.
