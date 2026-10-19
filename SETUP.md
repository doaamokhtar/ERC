# Setup

## 1. Which Jira this is

**Jira Cloud.** The browser shows `https://expo.sd.erc.sa/jira/...`, but that is a custom domain
in front of `https://erc-sd.atlassian.net`. The scripts talk to the `atlassian.net` address.
The `/jira` in the address bar is a page route in the web app. It is not part of the API path.

It is on the public internet, so you do not need the VPN.

## 2. An API token

Cloud has no personal access tokens. It uses an **API token paired with your Atlassian email**,
and you have to send both together. A token sent with the wrong email is rejected the same way
as a wrong token.

1. Sign in at `https://id.atlassian.com/manage-profile/security/api-tokens`
   as the same account you use on ERC's Jira
2. *Create API token*. Give it a name that says where it runs (e.g. `ERCJira scripts — laptop`)
   and the shortest expiry you can live with
3. Copy it. Atlassian only shows it once

Store both values in the Windows **User** scope, so they survive reboots and stay out of files:

```powershell
[Environment]::SetEnvironmentVariable("ERC_JIRA_EMAIL","<your atlassian email>","User")
[Environment]::SetEnvironmentVariable("ERC_JIRA_TOKEN","<paste the token here>","User")
```

Then **open a new PowerShell window**. A PowerShell window that was already open does not see a
variable set this way. This is the most common reason `test-connection.ps1` says nothing is set
right after you set it.

On a CI runner, the scripts also read `$env:ERC_JIRA_EMAIL` / `$env:ERC_JIRA_TOKEN`, where the
values come from secrets instead of the registry.

## 3. Check it

```powershell
.\tools\test-connection.ps1
```

You should see your display name, your accountId, and the list of service desk projects you can
see. If the list is empty, the token works but the account has no service desk access yet.

## Which account should the scripts run as?

The scripts run as whoever owns the token, so that is the name Jira records against every
comment, transition and field change they make. If automation runs under a person's token, the
issue history shows that person doing work no person did. It also stops working the day their
access changes. So anything scheduled or recurring should use its own service account, not yours.

`test-connection.ps1` prints the account on every run, so you can always see which one it is
without having to remember to check.
