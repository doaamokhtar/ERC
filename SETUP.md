# Setup

## 1. Network

The server is internal: `https://expo.sd.erc.sa/jira`. It is only reachable on the ERC
network or over the VPN. A failure here looks like a hang, then a timeout — not a login error.

```powershell
Test-NetConnection expo.sd.erc.sa -Port 443
```

## 2. A personal access token

A token can only be created by the account it belongs to — nobody can make one for you.

1. Sign in to `https://expo.sd.erc.sa/jira`
2. *Profile → Personal Access Tokens → Create token*
3. Give it a name that says where it runs (e.g. `ERCJira scripts — laptop`) and the shortest
   expiry you can live with

Store it in the Windows **User** scope so it survives reboots and is not sitting in a file:

```powershell
[Environment]::SetEnvironmentVariable("ERC_JIRA_TOKEN","<paste it here>","User")
```

Then **open a new PowerShell window**. A variable set this way does not reach a shell that
was already running, which is the single most common reason `test-connection.ps1` says the
token is not set right after you set it.

The scripts also accept `$env:ERC_JIRA_TOKEN` for a CI runner, where the token comes from a
secret rather than from the registry.

## 3. Check it

```powershell
.\tools\test-connection.ps1
```

Expected: your display name, your account name, and the list of service desk projects you can
see. If the list is empty, the token works but the account has no service desk access yet.

## Which account should the scripts run as?

Whatever the token belongs to, since that is the name Jira records against every comment,
transition and field change the scripts make. Running automation under a person's token
writes that person's name into the history of work no person did, and it stops the day their
access changes — so anything scheduled or recurring wants a service account of its own, not
yours.

`test-connection.ps1` prints the account name every run, so this stays visible rather than
being something you have to remember to check.
