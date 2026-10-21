# ERCJira — tooling for the ERC Jira Service Desk

Scripts and documentation for ERC's **Jira Cloud** service desk: `https://expo.sd.erc.sa`, a custom
domain for `https://erc-sd.atlassian.net` (Jira Service Management).

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
| `tools/build-docs.ps1` | rebuilds the WB documents in `docs/` |
| `docs/` | process documentation, runbooks, the WB user guide and templates |
| `docs/user-guide-build/` | the scripts and images the WB documents are built from |
| `reports/` | generated output, design notes, browser snippets |

## Getting started

```powershell
# 1. store your email + API token once (Windows User scope, survives reboots)
[Environment]::SetEnvironmentVariable("ERC_JIRA_EMAIL","<your atlassian email>","User")
[Environment]::SetEnvironmentVariable("ERC_JIRA_TOKEN","<api token>","User")

# 2. open a NEW PowerShell window, then check it reaches the server
.\tools\test-connection.ps1
```

`test-connection.ps1` prints which account the token belongs to and lists the service desk
projects that account can see. If it fails, read [SETUP.md](SETUP.md) — it is almost always
a token from a different account than the email, or one set in a window that was already open.

## Building the documents

The WB user guide, SLA template and requirements workbook in `docs/` are generated. To change
one, edit its script in `docs/user-guide-build/` and rebuild — don't edit the .docx/.xlsx by hand,
the next build overwrites it.

Needs [Node.js](https://nodejs.org) 18+ and [Python](https://www.python.org) 3.10+ on PATH
(and Microsoft Word for `-Pdf`). The first run installs the rest (npm `docx`, Python `openpyxl`
and `Pillow`).

```powershell
.\tools\build-docs.ps1          # .docx + .xlsx
.\tools\build-docs.ps1 -Pdf     # also the PDFs, with the guide's contents page filled in
```

| Output in `docs/` | Built by |
|---|---|
| `WB-User-Guide-v1.2-DRAFT.docx` | `build.js` |
| `WB-SLA-Requirements-Template-v1.1.docx` | `sla-template.js` |
| `WB-Requirements-and-SLA-Template-v1.4.xlsx` | `requirements-template.py` |

`emails.js` holds the HTML of the WB automation emails, in the same Expo 2030 Riyadh theme.

If PowerShell refuses to run the script ("running scripts is disabled"), allow local scripts once:
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.

## Writing a new script

Dot-source the client and call `Invoke-JiraApi`. It handles the base URL, the email + token,
retries and paging, so a script only says what it wants:

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
