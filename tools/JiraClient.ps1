<#
.SYNOPSIS
  The one place that knows the server, the token and how to make a REST call.

.WHY ONE FILE
  Every script needs the same four things - base URL, token, headers, paging - and when each
  one carries its own copy they drift: one gets a retry, one gets the context path wrong, one
  keeps working against the old hostname. Dot-source this instead:

    . "$PSScriptRoot\JiraClient.ps1"

.WHAT IT GIVES YOU
  $JiraUrl                 https://expo.sd.erc.sa/jira   (override with $env:ERC_JIRA_URL)
  Invoke-JiraApi           one REST call, any method, any api
  Get-JiraSearchResults    a JQL search, every page, as issue objects
  Get-JiraAccount          who the token belongs to
#>

# 1.0, not 2.0: strict mode 2.0 also throws on reading a property an object does not have, and
# half of what Jira returns is optional - myself has no emailAddress when it is hidden, a search
# page has no issues when it is empty. Those are normal answers, not errors.
Set-StrictMode -Version 1.0
$ErrorActionPreference = "Stop"

# Internal servers often refuse anything below TLS 1.2, and PowerShell 5.1 still defaults to
# TLS 1.0. Without this the failure is a bare "connection was closed", which reads like a
# network problem rather than a handshake one.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# The context path is part of the URL here - this Jira is served from /jira, not from the root
# of the host - so every REST path hangs off .../jira/rest/..., not .../rest/...
$script:JiraUrl = $env:ERC_JIRA_URL
if (-not $script:JiraUrl) { $script:JiraUrl = "https://expo.sd.erc.sa/jira" }
$script:JiraUrl = $script:JiraUrl.TrimEnd("/")
$JiraUrl = $script:JiraUrl

function Get-JiraToken {
  <# Environment variable first, so a CI runner can inject a secret; then the Windows User
     scope, which is where a laptop keeps it. #>
  $token = $env:ERC_JIRA_TOKEN
  if (-not $token) { $token = [Environment]::GetEnvironmentVariable("ERC_JIRA_TOKEN", "User") }
  if (-not $token) {
    throw "ERC_JIRA_TOKEN is not set. See SETUP.md. (If you just set it, open a new PowerShell window.)"
  }
  return $token
}

function Invoke-JiraApi {
  <#
  .SYNOPSIS
    One REST call against the ERC Jira.

  .PARAMETER Path
    The path AFTER the api root - "myself", "issue/VS-1", "search". A leading slash is fine.

  .PARAMETER Api
    Which api root. "api/2" is Jira itself; "servicedeskapi" is Jira Service Management
    (queues, request types, customers); "auditing/record" is the audit log.

  .PARAMETER Body
    A hashtable or PSObject - converted to JSON here, so callers never hand-build JSON.

  .EXAMPLE
    Invoke-JiraApi -Path "myself"
    Invoke-JiraApi -Path "issue/VS-1" -Method PUT -Body @{ fields = @{ summary = "new" } }
    Invoke-JiraApi -Api "servicedeskapi" -Path "servicedesk"
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string] $Path,
    [string] $Method = "GET",
    [string] $Api = "api/2",
    $Body,
    [hashtable] $Query,
    [int] $Retries = 2
  )

  $url = "$script:JiraUrl/rest/$Api/" + $Path.TrimStart("/")

  if ($Query -and $Query.Count -gt 0) {
    $pairs = foreach ($k in $Query.Keys) {
      "{0}={1}" -f [Uri]::EscapeDataString([string]$k), [Uri]::EscapeDataString([string]$Query[$k])
    }
    $url = $url + "?" + ($pairs -join "&")
  }

  $headers = @{
    Authorization = "Bearer " + (Get-JiraToken)
    # Jira rejects a REST write that looks like it came from a browser form unless this says
    # otherwise. Harmless on reads.
    "X-Atlassian-Token" = "no-check"
    # Several servicedeskapi endpoints are still flagged experimental and 400 without it.
    "X-ExperimentalApi" = "opt-in"
  }

  # not $args - that name is an automatic variable, and splatting it from inside an advanced
  # function is a confusing place to find a bug
  $req = @{
    Uri             = $url
    Method          = $Method
    Headers         = $headers
    ContentType     = "application/json; charset=utf-8"
    UseBasicParsing = $true
  }
  if ($null -ne $Body) {
    # Encoded here rather than passed as a string: Invoke-RestMethod in PowerShell 5.1 sends a
    # string body as ASCII regardless of ContentType, which turns Arabic text into question marks.
    $req["Body"] = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 20 -Compress))
  }

  $attempt = 0
  while ($true) {
    try {
      return Invoke-RestMethod @req
    }
    catch {
      $status = $null
      if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }

      # 401/403/404 will not get better by asking again - retrying only hides which one it was.
      $worthRetrying = ($null -eq $status) -or ($status -ge 500) -or ($status -eq 429)
      if (-not $worthRetrying -or $attempt -ge $Retries) {
        $detail = ""
        try {
          $reader = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())
          $detail = $reader.ReadToEnd()
        } catch { }
        throw ("{0} {1} -> {2} {3}`n{4}" -f $Method, $url, $status, $_.Exception.Message, $detail)
      }
      $attempt++
      Start-Sleep -Seconds ([Math]::Pow(2, $attempt))
    }
  }
}

function Get-JiraAccount {
  <# Who the token belongs to. A token is just a string - nothing about it says whose it is,
     and getting it wrong is silent: the scripts keep working and record a person's name for
     work no person did. #>
  return Invoke-JiraApi -Path "myself"
}

function Get-JiraSearchResults {
  <#
  .SYNOPSIS
    A JQL search, following paging to the end.

  .WHY
    Jira caps a search at maxResults and a caller that ignores the cap silently processes the
    first 50 issues and reports success.

  .EXAMPLE
    Get-JiraSearchResults -Jql "project = VS AND status != Closed" -Fields @("summary","status")
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string] $Jql,
    [string[]] $Fields = @("summary", "status", "issuetype", "created", "assignee"),
    [string[]] $Expand,
    [int] $PageSize = 100
  )

  $startAt = 0
  $issues = New-Object Collections.ArrayList
  while ($true) {
    $body = @{
      jql        = $Jql
      fields     = $Fields
      startAt    = $startAt
      maxResults = $PageSize
    }
    if ($Expand) { $body["expand"] = $Expand }

    $page = Invoke-JiraApi -Path "search" -Method POST -Body $body
    if ($page.issues) { [void]$issues.AddRange($page.issues) }

    $startAt += $PageSize
    if ($startAt -ge $page.total -or -not $page.issues -or $page.issues.Count -eq 0) { break }
  }
  return $issues.ToArray()
}
