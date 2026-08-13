# Agent Instructions

## GitHub Credentials

- Use `/workspace/work/myGH/git-credential-github` as the Git credential helper for authenticated GitHub operations.
- Pass it to Git per command so repository or global configuration is not changed:

  ```bash
  git -c credential.helper=/workspace/work/myGH/git-credential-github <command>
  ```

- Do not read, display, log, copy, modify, or commit `myEmail`, `myKeys`, `myUsername`, or the helper's output.
- Do not place usernames, passwords, personal access tokens, or helper output in remote URLs, command arguments, patches, commits, or responses.
- Do not run the credential helper directly for inspection. Let Git invoke it through the credential protocol.
- Redact credentials if an authentication error or remote URL unexpectedly includes them.
