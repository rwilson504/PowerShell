<#
.SYNOPSIS
    Acquires an access token via device code flow and lists Dataverse table relationships.

.DESCRIPTION
    Calls the GetAccessTokenDeviceCode helper to obtain an access token and then
    invokes GetTableRelationships.ps1 to enumerate 1:N, N:1, and M:M relationships.

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
    Optional array of table logical names. When supplied, only relationships in
    which at least one side is one of these tables are returned.

.PARAMETER RelationshipTypes
    Which relationship types to include. Valid values are "OneToMany" (1:N),
    "ManyToOne" (N:1), and "ManyToMany" (M:M). Default is all three.

.PARAMETER CustomEntitiesOnly
    When set, only return relationships where BOTH sides are custom entities.

.PARAMETER CustomRelationshipsOnly
    When set, only return relationships where IsCustomRelationship is true.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetTableRelationshipsWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com"

    Lists every relationship in the environment using device code authentication.

.EXAMPLE
    .\GetTableRelationshipsWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables @("account","contact") -OutputFormat CSV

    Lists relationships involving account or contact and exports to CSV.
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
    [ValidateSet("OneToMany", "ManyToOne", "ManyToMany")]
    [string[]]$RelationshipTypes = @("OneToMany", "ManyToOne", "ManyToMany"),

    [Parameter(Mandatory = $false)]
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$CustomRelationshipsOnly,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$SolutionUniqueName
)

# Get the access token using device code flow
Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript  = Join-Path $scriptDir "..\..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

# Build parameters for the main script
$scriptParams = @{
    OrganizationUrl   = $OrganizationUrl
    AccessToken       = $accessToken
    OutputFormat      = $OutputFormat
    RelationshipTypes = $RelationshipTypes
}

if ($Tables -and $Tables.Count -gt 0) {
    $scriptParams.Tables = $Tables
}

if ($CustomEntitiesOnly) {
    $scriptParams.CustomEntitiesOnly = $true
}

if ($CustomRelationshipsOnly) {
    $scriptParams.CustomRelationshipsOnly = $true
}

if ($OutputPath) {
    $scriptParams.OutputPath = $OutputPath
}

if ($SolutionUniqueName) {
    $scriptParams.SolutionUniqueName = $SolutionUniqueName
}

$mainScript = Join-Path $scriptDir "GetTableRelationships.ps1"
$results = & $mainScript @scriptParams
return $results
