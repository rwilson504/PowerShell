<#
.SYNOPSIS
    Acquires an access token and gets all first-party (Microsoft-owned) service principals.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token
    and then calls the GetFirstPartyServicePrincipals script to retrieve all Microsoft
    first-party service principals in the tenant.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.PARAMETER IncludeDisabled
    When specified, includes service principals that are disabled (accountEnabled = false).

.EXAMPLE
    .\GetFirstPartyServicePrincipalsWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -Environment "GCCH"

    Gets all first-party service principals in a GCC High tenant using device code authentication.

.EXAMPLE
    .\GetFirstPartyServicePrincipalsWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OutputFormat CSV -OutputPath "C:\temp\first-party-sps.csv"

    Gets all first-party service principals and exports to CSV.

.EXAMPLE
    .\GetFirstPartyServicePrincipalsWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -Environment "DoD" -IncludeDisabled

    Gets all first-party service principals including disabled ones in a DoD environment.

.AUTHOR
    Rick Wilson

.NOTES
    The app registration used must have Application.Read.All or Directory.Read.All
    permissions on Microsoft Graph.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Public", "GCC", "GCCH", "DoD")]
    [string]$Environment = "Public",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDisabled
)

# Determine the Microsoft Graph scope based on environment
switch ($Environment) {
    "Public" {
        $graphScope = "https://graph.microsoft.com/.default"
    }
    "GCC" {
        $graphScope = "https://graph.microsoft.com/.default"
    }
    "GCCH" {
        $graphScope = "https://graph.microsoft.us/.default"
    }
    "DoD" {
        $graphScope = "https://dod-graph.microsoft.us/.default"
    }
}

# Get the access token using device code flow
Write-Host "Acquiring access token for Microsoft Graph..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope $graphScope -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

# Build parameters for the main script
$scriptParams = @{
    AccessToken  = $accessToken
    Environment  = $Environment
    OutputFormat = $OutputFormat
}

if ($OutputPath) {
    $scriptParams.OutputPath = $OutputPath
}

if ($IncludeDisabled) {
    $scriptParams.IncludeDisabled = $true
}

# Get first-party service principals
$mainScript = Join-Path $scriptDir "GetFirstPartyServicePrincipals.ps1"
$results = & $mainScript @scriptParams

# Output the results
return $results
