<#
.SYNOPSIS
    Acquires an access token via device code flow and reports per-user activity counts on
    Dataverse tables across createdby/modifiedby and any extra user-typed lookup columns.

.DESCRIPTION
    Calls GetAccessTokenDeviceCode and then invokes GetUserActivityByTable.ps1 to identify
    who is creating, modifying, and approving records.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    Azure environment. Valid values: "Public", "GCC", "GCCH", "DoD". Default "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER Tables
    Required. One or more table logical names to analyze.

.PARAMETER UserLookupAttributes
    Additional user lookups (beyond createdby/modifiedby/createdonbehalfby/modifiedonbehalfby)
    to include - e.g. "msf_approver","msf_reviewer".

.PARAMETER AutoDetectUserLookups
    Auto-discover and include EVERY systemuser-targeted lookup on each table.

.PARAMETER ExcludeStandardUserAttributes
    Skip the four standard audit lookups so you see custom approver/reviewer activity only.

.PARAMETER Filter
    Optional OData $filter expression that restricts which records are counted.

.PARAMETER TopUsersPerAttribute
    Limit output to top N users per (Table, Attribute) - 0 = no limit.

.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".

.PARAMETER OutputPath
    Optional file path for the export.

.EXAMPLE
    .\GetUserActivityByTableWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "msf_program" -AutoDetectUserLookups -OutputFormat CSV

    Produces useractivity_*.csv showing every user that has created / modified / been
    referenced via a user-lookup on every msf_program record.

.EXAMPLE
    .\GetUserActivityByTableWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "msf_program" -UserLookupAttributes "msf_approver","msf_reviewer" -ExcludeStandardUserAttributes

    Reports ONLY approval / review activity (skips audit columns).
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
    [string[]]$UserLookupAttributes,

    [Parameter(Mandatory = $false)]
    [switch]$AutoDetectUserLookups,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeStandardUserAttributes,

    [Parameter(Mandatory = $false)]
    [string]$Filter,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1000)]
    [int]$TopUsersPerAttribute = 0,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript  = Join-Path $scriptDir "..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    OutputFormat    = $OutputFormat
}
if ($Tables -and $Tables.Count -gt 0)                             { $scriptParams.Tables               = $Tables }
if ($SolutionUniqueName)                                          { $scriptParams.SolutionUniqueName   = $SolutionUniqueName }

if ($UserLookupAttributes -and $UserLookupAttributes.Count -gt 0) { $scriptParams.UserLookupAttributes        = $UserLookupAttributes }
if ($AutoDetectUserLookups)                                       { $scriptParams.AutoDetectUserLookups       = $true }
if ($ExcludeStandardUserAttributes)                               { $scriptParams.ExcludeStandardUserAttributes = $true }
if ($Filter)                                                      { $scriptParams.Filter                      = $Filter }
if ($TopUsersPerAttribute -gt 0)                                  { $scriptParams.TopUsersPerAttribute        = $TopUsersPerAttribute }
if ($OutputPath)                                                  { $scriptParams.OutputPath                  = $OutputPath }

$mainScript = Join-Path $scriptDir "GetUserActivityByTable.ps1"
$results = & $mainScript @scriptParams
return $results
