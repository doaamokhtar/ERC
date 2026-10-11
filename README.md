# ERCJira — tooling for the ERC Jira Service Desk

Scripts and documentation for **`https://expo.sd.erc.sa/jira`** (Jira Service Management).

> ⚠️ **Private.** Internal URLs, custom field ids, group names and staff names end up in here.
> Bulk Jira data (issue extracts, CSV reports) is deliberately excluded — see `.gitignore`.

This is a separate project from `NLPJira`, which targets a different Jira
(`jira.neoleap.com.sa`). Nothing is shared between the two: different server, different
project keys, different field ids, different token.

---

## Layout

| Path | What |
|---|---|
| `tools/` | the scripts (PowerShell + Jira REST) |
| `tools/JiraClient.ps1` | the one place that holds the URL, the token and the REST call |
| `docs/` | process documentation, runbooks |
| `reports/` | generated output, design notes, browser snippets |

## Getting started

```powershell
# 1. store the token once (Windows User scope, survives reboots)
[Environment]::SetEnvironmentVariable("ERC_JIRA_TOKEN","<personal access token>","User")

# 2. open a NEW PowerShell window, then check it reaches the server
.\tools\test-connection.ps1
```

`test-connection.ps1` prints which account the token belongs to and lists the service desk
projects that account can see. If it fails, read [SETUP.md](SETUP.md) — it is almost always
the VPN or a token typed into the wrong window.

## Writing a new script

Dot-source the client and call `Invoke-JiraApi`. It handles the base URL, the context path
(`/jira`), the token, retries and paging, so a script only says what it wants:

```powershell
. "$PSScriptRoot\JiraClient.ps1"

$me = Invoke-JiraApi -Path "myself"
$issues = Get-JiraSearchResults -Jql "project = VS AND status != Closed" -Fields @("summary","status")
```

Two conventions carried over from the NeoLeap tooling, because both earned their place:

- **Every script takes `-DryRun`** and prints what it *would* write. A script that cannot be
  rehearsed does not get run against a live queue.
- **Every script is idempotent** — writes only when the value actually changed, so re-running
  is free and leaves no noise in the issue history.

## Current state

Empty scaffold. The Jira project itself (key, request types, workflow, queues) has not been
created yet, and no field ids are known — so nothing here is pointed at a real project.
