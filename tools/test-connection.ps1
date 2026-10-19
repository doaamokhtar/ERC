<#
.SYNOPSIS
  Checks that the token reaches the ERC Jira, and prints what it can see.

.WHY
  The three things that go wrong when setting this up are indistinguishable from each other
  in any other script's error message: no VPN (a timeout), a token in the wrong window (not
  set), and a working token on an account with no service desk access (an empty list). This
  separates them and says which one it is.

.USAGE
  .\tools\test-connection.ps1
#>
[CmdletBinding()]
param()

. "$PSScriptRoot\JiraClient.ps1"

Write-Host "Server : $JiraUrl"

# ---- is the host even reachable? -------------------------------------------------------------
# Asked first, because a VPN problem otherwise surfaces as a 100-second hang inside the REST
# call and reads like the server is broken.
$hostName = ([Uri]$JiraUrl).Host
$reach = Test-NetConnection -ComputerName $hostName -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
if (-not $reach) {
  Write-Host "UNREACHABLE: $hostName`:443" -ForegroundColor Red
  Write-Host "Jira Cloud is on the public internet, so this is the connection or a proxy/firewall," -ForegroundColor Yellow
  Write-Host "not Jira. Check you are online and try again." -ForegroundColor Yellow
  exit 1
}
Write-Host "Network: $hostName`:443 reachable" -ForegroundColor Green

# ---- is there a token at all? ----------------------------------------------------------------
# Separated from the call below so "you never set it" and "Jira rejected it" do not print the
# same advice. They need opposite fixes.
try { [void](Get-JiraEmail); [void](Get-JiraToken) }
catch {
  Write-Host "NOT SET UP" -ForegroundColor Red
  Write-Host ("  " + $_.Exception.Message)
  Write-Host "`n  [Environment]::SetEnvironmentVariable(`"ERC_JIRA_EMAIL`",`"<your atlassian email>`",`"User`")" -ForegroundColor Yellow
  Write-Host "  [Environment]::SetEnvironmentVariable(`"ERC_JIRA_TOKEN`",`"<api token>`",`"User`")" -ForegroundColor Yellow
  Write-Host "  ...then open a NEW PowerShell window. See SETUP.md." -ForegroundColor Yellow
  exit 1
}
Write-Host "Email  : $(Get-JiraEmail) (the account the token must belong to)"

# ---- whose token is this? --------------------------------------------------------------------
try {
  $me = Get-JiraAccount
}
catch {
  Write-Host "TOKEN REJECTED" -ForegroundColor Red
  Write-Host ("  " + $_.Exception.Message)
  Write-Host "`nJira would not accept the email + token pair - usually a token from a different" -ForegroundColor Yellow
  Write-Host "Atlassian account than ERC_JIRA_EMAIL, or one that has expired or been revoked." -ForegroundColor Yellow
  Write-Host "Create a new one; see SETUP.md." -ForegroundColor Yellow
  exit 1
}

# Cloud has no username - accountId is the id Jira records against everything the scripts do.
Write-Host ("Account: {0} ({1})" -f $me.displayName, $me.accountId) -ForegroundColor Green

# ---- what can it see? ------------------------------------------------------------------------
Write-Host "`nService desks this account can see:"
try {
  $desks = Invoke-JiraApi -Api "servicedeskapi" -Path "servicedesk"
  if (-not $desks.values -or $desks.values.Count -eq 0) {
    Write-Host "  (none)" -ForegroundColor Yellow
    Write-Host "  The token works, but this account has no service desk access yet - ask a" -ForegroundColor Yellow
    Write-Host "  Jira admin to add it to the right project role." -ForegroundColor Yellow
  }
  else {
    foreach ($d in $desks.values) {
      Write-Host ("  {0,-8} {1}" -f $d.projectKey, $d.projectName)
    }
    Write-Host ("`n{0} service desk(s)." -f $desks.values.Count) -ForegroundColor Green
  }
}
catch {
  # Jira Software without JSM licensed, or the api disabled - worth saying, not worth failing on,
  # because the token and the network are already proven at this point.
  Write-Host "  could not list service desks:" -ForegroundColor Yellow
  Write-Host ("  " + $_.Exception.Message)
}
