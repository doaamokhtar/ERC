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
  Write-Host "This is an internal server - connect to the ERC network or the VPN and try again." -ForegroundColor Yellow
  exit 1
}
Write-Host "Network: $hostName`:443 reachable" -ForegroundColor Green

# ---- is there a token at all? ----------------------------------------------------------------
# Separated from the call below so "you never set it" and "Jira rejected it" do not print the
# same advice. They need opposite fixes.
try { [void](Get-JiraToken) }
catch {
  Write-Host "NO TOKEN" -ForegroundColor Red
  Write-Host ("  " + $_.Exception.Message)
  Write-Host "`n  [Environment]::SetEnvironmentVariable(`"ERC_JIRA_TOKEN`",`"<token>`",`"User`")" -ForegroundColor Yellow
  Write-Host "  ...then open a NEW PowerShell window. See SETUP.md." -ForegroundColor Yellow
  exit 1
}

# ---- whose token is this? --------------------------------------------------------------------
try {
  $me = Get-JiraAccount
}
catch {
  Write-Host "TOKEN REJECTED" -ForegroundColor Red
  Write-Host ("  " + $_.Exception.Message)
  Write-Host "`nThe token is set but Jira would not accept it - usually expired, revoked, or" -ForegroundColor Yellow
  Write-Host "created on a different Jira. Create a new one; see SETUP.md." -ForegroundColor Yellow
  exit 1
}

Write-Host ("Account: {0} ({1})" -f $me.displayName, $me.name) -ForegroundColor Green
if ($me.emailAddress) { Write-Host ("Email  : {0}" -f $me.emailAddress) }

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
