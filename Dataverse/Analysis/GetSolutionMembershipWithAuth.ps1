<#
.SYNOPSIS
    Acquires an access token via device code flow and reports per-component solution
    membership for Dataverse tables, attributes, and relationships.

.PARAMETER TenantId
    The Azure AD tenant ID.
.PARAMETER ClientId
    The client ID of your registered Azure AD app.
.PARAMETER Environment
    "Public", "GCC", "GCCH", or "DoD". Default "Public".
.PARAMETER OrganizationUrl
    The Dataverse organization URL.
.PARAMETER Tables
    Optional list of table logical names to filter to.
.PARAMETER ComponentTypes
    "Entity", "Attribute", and/or "Relationship". Default all three.
.PARAMETER UnmanagedOnly
    Only emit unmanaged-solution memberships.
.PARAMETER ManagedOnly
    Only emit managed-solution memberships.
.PARAMETER ExcludeSystemSolutions
    Skip System / Active / Default / msdyn* / msft* / mscrm* solutions.
.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON".
.PARAMETER OutputPath
    Optional output file path.

.EXAMPLE
    .\GetSolutionMembershipWithAuth.ps1 -TenantId "..." -ClientId "..." -OrganizationUrl "https://your-org.crm.dynamics.com" -Tables "msf_program" -ComponentTypes Attribute,Relationship -UnmanagedOnly -OutputFormat CSV
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
    [ValidateSet("Entity","Attribute","Relationship")]
    [string[]]$ComponentTypes = @("Entity","Attribute","Relationship"),

    [Parameter(Mandatory = $false)]
    [switch]$UnmanagedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ManagedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeSystemSolutions,

    [Parameter(Mandatory = $false)]
    [string]$SolutionUniqueName,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
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
    ComponentTypes  = $ComponentTypes
}
if ($Tables -and $Tables.Count -gt 0) { $scriptParams.Tables                 = $Tables }
if ($UnmanagedOnly)                   { $scriptParams.UnmanagedOnly          = $true }
if ($ManagedOnly)                     { $scriptParams.ManagedOnly            = $true }
if ($ExcludeSystemSolutions)          { $scriptParams.ExcludeSystemSolutions = $true }
if ($SolutionUniqueName)              { $scriptParams.SolutionUniqueName     = $SolutionUniqueName }
if ($OutputPath)                      { $scriptParams.OutputPath             = $OutputPath }

$mainScript = Join-Path $scriptDir "GetSolutionMembership.ps1"
$results = & $mainScript @scriptParams
return $results
