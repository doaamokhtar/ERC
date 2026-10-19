<#
.SYNOPSIS
  Makes sure the fixed "anonymous" customer exists and belongs to Anonymous TKT (AT), and prints
  its accountId for the Raise a Ticket web app (Script Property REPORTER_ID).

.WHY
  The web app raises every request on behalf of one customer. The REST api takes that customer
  as an accountId, not an email, and refuses one that is not a customer of the service desk -
  so the account has to exist, and be in AT, before the first submission.

.USAGE
  .\tools\ensure-anonymous-customer.ps1 -DryRun
  .\tools\ensure-anonymous-customer.ps1
#>
[CmdletBinding()]
param(
  [string] $Email = "Anonymous@Anonymous.com",
  [string] $DisplayName = "Anonymous",
  [string] $ServiceDeskId = "169",
  [switch] $DryRun
)

. "$PSScriptRoot\JiraClient.ps1"

function Find-DeskCustomer {
  $page = Invoke-JiraApi -Api "servicedeskapi" -Path "servicedesk/$ServiceDeskId/customer" -Query @{ query = $Email; limit = 50 }
  return $page.values | Where-Object { $_.emailAddress -and $_.emailAddress -ieq $Email } | Select-Object -First 1
}

$existing = Find-DeskCustomer
if ($existing) {
  Write-Host "Already a customer of service desk $ServiceDeskId - nothing to change." -ForegroundColor Green
  Write-Host ("REPORTER_ID = {0}" -f $existing.accountId) -ForegroundColor Cyan
  exit 0
}

if ($DryRun) {
  Write-Host "[DryRun] would create customer '$Email' ($DisplayName) and add it to service desk $ServiceDeskId" -ForegroundColor Yellow
  exit 0
}

# Created site-wide first, then added to the desk: one call cannot do both.
try {
  $created = Invoke-JiraApi -Api "servicedeskapi" -Path "customer" -Method POST -Body @{ email = $Email; displayName = $DisplayName }
  $accountId = $created.accountId
  Write-Host "Created customer $Email" -ForegroundColor Green
}
catch {
  # Exists on the site but not in AT. The customer api cannot look a portal-only account up by
  # email outside a desk, so say so rather than guess.
  Write-Host "Could not create $Email - it may already exist on the site under another desk." -ForegroundColor Red
  Write-Host ("  " + $_.Exception.Message)
  Write-Host "Add it by hand: AT -> Customers -> Add customers -> $Email, then run this again." -ForegroundColor Yellow
  exit 1
}

Invoke-JiraApi -Api "servicedeskapi" -Path "servicedesk/$ServiceDeskId/customer" -Method POST -Body @{ accountIds = @($accountId) } | Out-Null
Write-Host "Added to service desk $ServiceDeskId" -ForegroundColor Green
Write-Host ("REPORTER_ID = {0}" -f $accountId) -ForegroundColor Cyan
