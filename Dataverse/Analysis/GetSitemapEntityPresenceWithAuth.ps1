<#
.SYNOPSIS
    Acquires an access token via device code flow and reports which tables are surfaced in
    the sitemaps of model-driven apps (the strongest "user-facing?" signal in the suite).

.PARAMETER TenantId
    The Azure AD tenant ID.
.PARAMETER ClientId
    The client ID of your registered Azure AD app.
.PARAMETER Environment
    "Public" / "GCC" / "GCCH" / "DoD". Default "Public".
.PARAMETER OrganizationUrl
    The Dataverse organization URL.
.PARAMETER Tables
    Optional. Restrict output to SubAreas binding to one of these table logical names.
.PARAMETER SolutionUniqueName
    Optional. Restrict output to SubAreas binding to a table in this solution.
.PARAMETER AppUniqueNames
    Optional. Restrict the appmodules scanned to a specific list of unique names.
.PARAMETER IncludeUnpublished
    Switch. Include appmodules with componentstate other than 0.
.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".
.PARAMETER OutputPath
    Optional output file path.

.EXAMPLE
    .\GetSitemapEntityPresenceWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -SolutionUniqueName "msf_Core" -OutputFormat CSV
#>

param (
    [Parameter(Mandatory = $true)] [string]$TenantId,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [Parameter(Mandatory = $false)] [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public",
    [Parameter(Mandatory = $true)] [string]$OrganizationUrl,
    [Parameter(Mandatory = $false)] [string[]]$Tables,
    [Parameter(Mandatory = $false)] [string]$SolutionUniqueName,
    [Parameter(Mandatory = $false)] [string[]]$AppUniqueNames,
    [Parameter(Mandatory = $false)] [switch]$IncludeUnpublished,
    [Parameter(Mandatory = $false)] [ValidateSet("Table", "CSV", "JSON")] [string]$OutputFormat = "Table",
    [Parameter(Mandatory = $false)] [string]$OutputPath
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
    OutputFormat    = $OutputFormat
}
if ($Tables -and $Tables.Count -gt 0)             { $scriptParams.Tables             = $Tables }
if ($SolutionUniqueName)                          { $scriptParams.SolutionUniqueName = $SolutionUniqueName }
if ($AppUniqueNames -and $AppUniqueNames.Count -gt 0) { $scriptParams.AppUniqueNames = $AppUniqueNames }
if ($IncludeUnpublished)                          { $scriptParams.IncludeUnpublished = $true }
if ($OutputPath)                                  { $scriptParams.OutputPath         = $OutputPath }

$mainScript = Join-Path $scriptDir "GetSitemapEntityPresence.ps1"
$results = & $mainScript @scriptParams
return $results
