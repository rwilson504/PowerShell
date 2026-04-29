<#
.SYNOPSIS
    Acquires an access token via device code flow and reports per-attribute UI presence
    (forms, views, charts) for one or more Dataverse tables.

.PARAMETER TenantId
    The Azure AD tenant ID.
.PARAMETER ClientId
    The client ID of your registered Azure AD app.
.PARAMETER Environment
    "Public" / "GCC" / "GCCH" / "DoD". Default "Public".
.PARAMETER OrganizationUrl
    The Dataverse organization URL.
.PARAMETER Tables
    Required. One or more table logical names.
.PARAMETER IncludeUserQueries
    Also scan personal views (userquery).
.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".
.PARAMETER OutputPath
    Optional output file path.

.EXAMPLE
    .\GetFieldUIPresenceWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "msf_program" -OutputFormat CSV
#>

param (
    [Parameter(Mandatory = $true)] [string]$TenantId,
    [Parameter(Mandatory = $true)] [string]$ClientId,
    [Parameter(Mandatory = $false)] [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public",
    [Parameter(Mandatory = $true)] [string]$OrganizationUrl,
    [Parameter(Mandatory = $false)] [string[]]$Tables,
    [Parameter(Mandatory = $false)] [string]$SolutionUniqueName,
    [Parameter(Mandatory = $false)] [switch]$IncludeUserQueries,
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
if ($Tables -and $Tables.Count -gt 0) { $scriptParams.Tables             = $Tables }
if ($SolutionUniqueName)              { $scriptParams.SolutionUniqueName = $SolutionUniqueName }
if ($IncludeUserQueries) { $scriptParams.IncludeUserQueries = $true }
if ($OutputPath)         { $scriptParams.OutputPath         = $OutputPath }

$mainScript = Join-Path $scriptDir "GetFieldUIPresence.ps1"
$results = & $mainScript @scriptParams
return $results
