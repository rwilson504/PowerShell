<#
.SYNOPSIS
    Acquires a Dataverse access token and downloads all note attachments directly related
    to a record.

.DESCRIPTION
    Uses the shared device-code authentication helper to acquire an access token, then
    invokes Save-RecordAttachments.ps1 for the supplied record GUID.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of the registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", and "DoD".
    Defaults to "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER RecordId
    The GUID of the record whose note attachments should be downloaded.

.PARAMETER OutputPath
    Directory where attachments are written. Defaults to a folder named
    DataverseAttachments-<RecordId> beneath the current directory.

.PARAMETER Overwrite
    Replace files that already exist at the resolved output path.

.EXAMPLE
    .\Save-RecordAttachmentsWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -RecordId "00000000-0000-0000-0000-000000000001"

    Authenticates interactively and downloads all note attachments for the record.

.EXAMPLE
    .\Save-RecordAttachmentsWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -RecordId "00000000-0000-0000-0000-000000000001" -OutputPath "C:\Exports\CaseAttachments" -Overwrite

    Downloads the attachments to a specific directory and replaces existing files.
#>

[CmdletBinding(SupportsShouldProcess)]
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
    [guid]$RecordId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path (Get-Location).Path "DataverseAttachments-$RecordId"),

    [Parameter(Mandatory = $false)]
    [switch]$Overwrite
)

Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$($OrganizationUrl.TrimEnd('/'))/user_impersonation" -Environment $Environment

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    RecordId        = $RecordId
    OutputPath      = $OutputPath
}

if ($Overwrite)       { $scriptParams.Overwrite = $true }
if ($WhatIfPreference) { $scriptParams.WhatIf = $true }

$mainScript = Join-Path $scriptDir "Save-RecordAttachments.ps1"
$results = & $mainScript @scriptParams
return $results
