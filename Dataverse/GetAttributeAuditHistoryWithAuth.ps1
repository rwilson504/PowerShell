<#
.SYNOPSIS
    Acquires an access token via device code flow and reports per-attribute audit history
    for one or more Dataverse tables.

.DESCRIPTION
    Calls the GetAccessTokenDeviceCode helper to obtain an access token and then invokes
    GetAttributeAuditHistory.ps1 to scan the audit log for per-attribute last-modified data.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER Tables
    Required. One or more table logical names to analyze.

.PARAMETER Attributes
    Optional. Restrict analysis to these attribute logical names.

.PARAMETER DaysBack
    Number of days of audit history to scan. Pass 1-3650 to override, or leave at default 0
    to use AUTO mode (uses configured retention; falls back to environment age when retention
    is unset/'never expire'). The base script reports what window was actually used.

.PARAMETER LookupAttributesOnly
    Only emit rows for Lookup attributes - useful for the "unused lookup" detection workflow.

.PARAMETER UnusedOnly
    Only emit rows where no audit entries were found in the window.

.PARAMETER IncludeAuditDisabledColumns
    By default, columns where IsAuditEnabled = false are omitted. Set this switch to include them.

.PARAMETER MaxAuditPageSize
    Maximum audit rows pulled per page (100-5000). Default 5000.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetAttributeAuditHistoryWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "account"

    Reports last-modified data for every audited attribute on account in the last 365 days.

.EXAMPLE
    .\GetAttributeAuditHistoryWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "account","contact" -LookupAttributesOnly -UnusedOnly -OutputFormat CSV

    Lists lookup fields on account and contact that were never modified in the last year -
    candidates for cleanup or relationship review.
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
    [string[]]$Attributes,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 3650)]
    [int]$DaysBack = 0,

    [Parameter(Mandatory = $false)]
    [switch]$LookupAttributesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$UnusedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeAuditDisabledColumns,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 5000)]
    [int]$MaxAuditPageSize = 5000,

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
    OutputFormat    = $OutputFormat
}
if ($Tables -and $Tables.Count -gt 0)         { $scriptParams.Tables                     = $Tables }
if ($SolutionUniqueName)                       { $scriptParams.SolutionUniqueName         = $SolutionUniqueName }

if ($Attributes -and $Attributes.Count -gt 0)  { $scriptParams.Attributes                  = $Attributes }
if ($DaysBack -gt 0)                            { $scriptParams.DaysBack                    = $DaysBack }
if ($LookupAttributesOnly)                      { $scriptParams.LookupAttributesOnly        = $true }
if ($UnusedOnly)                                { $scriptParams.UnusedOnly                  = $true }
if ($IncludeAuditDisabledColumns)               { $scriptParams.IncludeAuditDisabledColumns = $true }
if ($MaxAuditPageSize -ne 5000)                 { $scriptParams.MaxAuditPageSize            = $MaxAuditPageSize }
if ($OutputPath)                                { $scriptParams.OutputPath                  = $OutputPath }

$mainScript = Join-Path $scriptDir "GetAttributeAuditHistory.ps1"
$results = & $mainScript @scriptParams
return $results
