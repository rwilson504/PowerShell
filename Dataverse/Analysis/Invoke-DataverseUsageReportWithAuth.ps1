<#
.SYNOPSIS
    Acquires an access token via device code flow and runs the full Dataverse usage-analysis
    suite, dropping every CSV into a single timestamped folder.

.DESCRIPTION
    Calls GetAccessTokenDeviceCode then invokes Invoke-DataverseUsageReport.ps1 to run
    all eight analysis scripts.

.PARAMETER TenantId
    The Azure AD tenant ID.
.PARAMETER ClientId
    The client ID of your registered Azure AD app.
.PARAMETER Environment
    "Public" / "GCC" / "GCCH" / "DoD". Default "Public".
.PARAMETER OrganizationUrl
    The Dataverse organization URL.
.PARAMETER Tables
    Optional list of table logical names to scope reports to.
.PARAMETER SolutionUniqueName
    Optional solution unique name to scope every report to.
.PARAMETER OutputRoot
    Root folder for the output subfolder. Default is the current directory.
.PARAMETER Skip
    One or more report names to skip: RecordCounts, Relationships, SolutionMembership,
    TableUsage, AttributeUsage, UIPresence, UserActivity, Audit.
.PARAMETER IncludeLastActivity
    Forwarded to GetRecordCountByTable.
.PARAMETER ActivityFallback
    Forwarded to GetRecordCountByTable.
.PARAMETER CustomEntitiesOnly
    Forwarded to GetRecordCountByTable when -SolutionUniqueName is NOT supplied.
.PARAMETER AuditDaysBack
    Forwarded to GetAttributeAuditHistory. 0 = AUTO (default; uses retention or env age).
.PARAMETER AutoDetectUserLookups
    Forwarded to GetUserActivityByTable. Auto-discovers every systemuser-targeted lookup.
.PARAMETER UserLookupAttributes
    Forwarded to GetUserActivityByTable. Custom user lookups beyond the standard 4.

.EXAMPLE
    .\Invoke-DataverseUsageReportWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -SolutionUniqueName "msf_Core" -IncludeLastActivity -ActivityFallback -AutoDetectUserLookups

    Runs the full suite scoped to the msf_Core solution and writes 8 CSVs into a timestamped
    folder under the current directory.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Public", "GCC", "GCCH", "DoD")]
    [string]$Environment = "Public",

    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $false)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [string]$SolutionUniqueName,

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [ValidateSet('RecordCounts','Relationships','SolutionMembership','TableUsage','AttributeUsage','UIPresence','UserActivity','Audit')]
    [string[]]$Skip,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLastActivity,

    [Parameter(Mandatory = $false)]
    [switch]$ActivityFallback,

    [Parameter(Mandatory = $false)]
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 3650)]
    [int]$AuditDaysBack = 0,

    [Parameter(Mandatory = $false)]
    [switch]$AutoDetectUserLookups,

    [Parameter(Mandatory = $false)]
    [string[]]$UserLookupAttributes
)

Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript  = Join-Path $scriptDir "..\..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

if (-not $accessToken) { Write-Error "Failed to acquire access token."; exit 1 }
Write-Host "Access token acquired successfully." -ForegroundColor Green

$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    OutputRoot      = $OutputRoot
}
if ($Tables -and $Tables.Count -gt 0)                             { $scriptParams.Tables                = $Tables }
if ($SolutionUniqueName)                                          { $scriptParams.SolutionUniqueName    = $SolutionUniqueName }
if ($Skip -and $Skip.Count -gt 0)                                 { $scriptParams.Skip                  = $Skip }
if ($IncludeLastActivity)                                         { $scriptParams.IncludeLastActivity   = $true }
if ($ActivityFallback)                                            { $scriptParams.ActivityFallback      = $true }
if ($CustomEntitiesOnly)                                          { $scriptParams.CustomEntitiesOnly    = $true }
if ($AuditDaysBack -gt 0)                                         { $scriptParams.AuditDaysBack         = $AuditDaysBack }
if ($AutoDetectUserLookups)                                       { $scriptParams.AutoDetectUserLookups = $true }
if ($UserLookupAttributes -and $UserLookupAttributes.Count -gt 0) { $scriptParams.UserLookupAttributes  = $UserLookupAttributes }

$mainScript = Join-Path $scriptDir "Invoke-DataverseUsageReport.ps1"
$result = & $mainScript @scriptParams
return $result
