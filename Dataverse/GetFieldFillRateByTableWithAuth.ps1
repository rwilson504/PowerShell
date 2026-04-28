<#
.SYNOPSIS
    Acquires an access token via device code flow and reports field fill rates for
    one or more Dataverse tables.

.DESCRIPTION
    Calls the GetAccessTokenDeviceCode helper to obtain an access token and then
    invokes GetFieldFillRateByTable.ps1 to compute per-field populated record counts.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD".
    Default value is "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER Tables
    Required. One or more table logical names to analyze.

.PARAMETER Attributes
    Optional. Restrict analysis to the specified attribute logical names.

.PARAMETER IncludeSystemAttributes
    Include system-managed columns in the report.

.PARAMETER CustomAttributesOnly
    When set, only attributes where IsCustomAttribute is true are analyzed.

.PARAMETER StandardAttributesOnly
    When set, only attributes where IsCustomAttribute is false are analyzed.

.PARAMETER Filter
    Optional OData $filter expression that restricts which records are considered when
    computing fill rates. Example: "statecode eq 0".

.PARAMETER BatchRequestSize
    Number of per-attribute count requests bundled into a single OData $batch HTTP call.
    Default is 50. Set to 1 to disable batching.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetFieldFillRateByTableWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "msf_company"

    Reports field fill rates on msf_company.

.EXAMPLE
    .\GetFieldFillRateByTableWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "account","contact" -CustomAttributesOnly -OutputFormat CSV

    Reports custom-attribute fill rates on account and contact, exported to CSV.

.EXAMPLE
    .\GetFieldFillRateByTableWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "account" -Filter "statecode eq 0"

    Reports fill rates restricted to active accounts only.
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

    [Parameter(Mandatory = $true)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [string[]]$Attributes,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemAttributes,

    [Parameter(Mandatory = $false)]
    [switch]$CustomAttributesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$StandardAttributesOnly,

    [Parameter(Mandatory = $false)]
    [string]$Filter,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$BatchRequestSize = 100,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 60000)]
    [int]$RequestThrottleDelayMs = 0,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Get the access token using device code flow
Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript  = Join-Path $scriptDir "..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

# Build parameters for the main script
$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    Tables          = $Tables
    OutputFormat    = $OutputFormat
}

if ($Attributes -and $Attributes.Count -gt 0) {
    $scriptParams.Attributes = $Attributes
}

if ($IncludeSystemAttributes)  { $scriptParams.IncludeSystemAttributes  = $true }
if ($CustomAttributesOnly)     { $scriptParams.CustomAttributesOnly     = $true }
if ($StandardAttributesOnly)   { $scriptParams.StandardAttributesOnly   = $true }
if ($Filter)                   { $scriptParams.Filter                   = $Filter }
if ($BatchRequestSize -ne 100)         { $scriptParams.BatchRequestSize         = $BatchRequestSize }
if ($RequestThrottleDelayMs -gt 0)     { $scriptParams.RequestThrottleDelayMs   = $RequestThrottleDelayMs }
if ($OutputPath)                       { $scriptParams.OutputPath               = $OutputPath }

$mainScript = Join-Path $scriptDir "GetFieldFillRateByTable.ps1"
$results = & $mainScript @scriptParams
return $results
