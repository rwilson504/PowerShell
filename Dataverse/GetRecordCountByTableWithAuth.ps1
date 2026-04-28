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

.PARAMETER CustomEntitiesOnly
    When querying all tables (no -Tables parameter), restrict the result to custom entities only.
    Useful for auditing custom-built capabilities without out-of-the-box noise.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER BatchSize
    The number of tables to include per RetrieveTotalRecordCount API call.
    Default is 20. Reduce if you encounter URL length issues.

.PARAMETER OutputPath
    Optional file path to export the results.

.PARAMETER IncludeLastActivity
    When specified, retrieves the last CreatedOn, last ModifiedOn, and oldest CreatedOn timestamps
    for each table, plus computed DaysSinceLastCreated, DaysSinceLastModified, and a UsageBucket
    classification (Empty / Active (<=90d) / Dormant (91-365d) / Stale (>365d) / Unknown). Useful for identifying
    tables/capabilities that are no longer in active use. Adds three extra API calls per table
    with records, so it can significantly increase runtime.

.PARAMETER ActivityFallback
    Only meaningful with -IncludeLastActivity. When set, also runs the activity timestamp queries
    against tables whose RecordCount came back as 0 or N/A. Useful on test/sandbox environments
    where the count API can return stale 0 values. Adds API calls for every empty/unavailable
    table, so it can significantly increase runtime in environments with many empty tables.

.PARAMETER IncludeUnsupportedTypes
    By default, Virtual and Elastic tables are pre-skipped (RetrieveTotalRecordCount does not
    support them). Set this switch to attempt them anyway.

.NOTES
    Output always includes SchemaName, EntitySetName, and IsCustomEntity from table metadata.
    When a batch count call fails (one bad apple in the batch), the script automatically retries
    each table individually so the rest still get counts.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables @("account", "contact")

    Gets record counts for account and contact tables using device code authentication.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -OutputFormat CSV -OutputPath "C:\temp\counts.csv"

    Gets record counts for all readable tables and exports to CSV.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Environment "GCCH"

    Gets record counts in a GCC High environment.

.EXAMPLE
    .\GetRecordCountByTableWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -IncludeLastActivity -OutputFormat CSV -OutputPath "C:\temp\counts.csv"

    Gets record counts plus the last CreatedOn/ModifiedOn timestamps for each table
    to help identify unused tables.
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
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 200)]
    [int]$BatchSize = 20,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLastActivity,

    [Parameter(Mandatory = $false)]
    [switch]$ActivityFallback,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeUnsupportedTypes,

    [Parameter(Mandatory = $false)]
    [switch]$NoBatchActivityProbes,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 60000)]
    [int]$RequestThrottleDelayMs = 0
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

if ($CustomEntitiesOnly) {
    $scriptParams.CustomEntitiesOnly = $true
}

if ($BatchSize -ne 20) {
    $scriptParams.BatchSize = $BatchSize
}

if ($OutputPath) {
    $scriptParams.OutputPath = $OutputPath
}

if ($IncludeLastActivity) {
    $scriptParams.IncludeLastActivity = $true
}

if ($ActivityFallback) {
    $scriptParams.ActivityFallback = $true
}

if ($IncludeUnsupportedTypes) {
    $scriptParams.IncludeUnsupportedTypes = $true
}

if ($NoBatchActivityProbes) {
    $scriptParams.NoBatchActivityProbes = $true
}

if ($RequestThrottleDelayMs -gt 0) {
    $scriptParams.RequestThrottleDelayMs = $RequestThrottleDelayMs
}

# Get record counts
$mainScript = Join-Path $scriptDir "GetRecordCountByTable.ps1"
$results = & $mainScript @scriptParams

# Output the results
return $results
