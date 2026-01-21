<#
.SYNOPSIS
    Acquires an access token and gets record counts for tables in Dataverse.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token 
    and then calls the GetRecordCountByTable script to retrieve record counts.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER Tables
    An optional array of table logical names to get counts for.
    If not provided, all readable tables will be queried.

.PARAMETER IncludeSystemTables
    When querying all tables, include system tables (those starting with 'sys').
    Default is $false.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables @("account", "contact")

    Gets record counts for account and contact tables using device code authentication.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -OutputFormat CSV -OutputPath "C:\temp\counts.csv"

    Gets record counts for all readable tables and exports to CSV.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Environment "GCCH"

    Gets record counts in a GCC High environment.
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
    [switch]$IncludeSystemTables = $false,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Get the access token using device code flow
Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

# Build parameters for the main script
$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken = $accessToken
    OutputFormat = $OutputFormat
}

if ($Tables -and $Tables.Count -gt 0) {
    $scriptParams.Tables = $Tables
}

if ($IncludeSystemTables) {
    $scriptParams.IncludeSystemTables = $true
}

if ($OutputPath) {
    $scriptParams.OutputPath = $OutputPath
}

# Get record counts
$mainScript = Join-Path $scriptDir "GetRecordCountByTable.ps1"
$results = & $mainScript @scriptParams

# Output the results
return $results
