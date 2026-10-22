<#
.SYNOPSIS
    Acquires a token and reads or sets one Dataverse OrgDBOrgSettings XML value.

.DESCRIPTION
    Uses the shared EntraID auth helper and delegates to SetOrgDbOrgSetting.ps1.
    Omit Value to read only. Writes affect the whole organization and require
    confirmation by default. Only use supported keys and values. WhatIf skips
    authentication and all API calls. Requires PowerShell 5.1 or later.

.PARAMETER OrganizationUrl
    HTTPS root URL of the Dataverse organization, without an API path.

.PARAMETER TenantId
    Microsoft Entra tenant ID.

.PARAMETER ClientId
    Application ID of the public-client registration used for authentication.

.PARAMETER Environment
    Azure cloud: Public, GCC, GCCH, or DoD. Defaults to Public.

.PARAMETER AuthenticationMode
    DeviceCode or Interactive browser sign-in with PKCE. Defaults to DeviceCode.

.PARAMETER Setting
    Exact, case-sensitive XML element name of a supported OrgDBOrgSettings key.
    Namespaces and XPath expressions are not accepted.

.PARAMETER Value
    Nonempty string to store as XML text. Omit to read only; no default is inferred.

.EXAMPLE
    .\SetOrgDbOrgSettingWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Setting AllowRoleAssignmentOnDisabledUsers -AuthenticationMode Interactive

    Signs in and reads the explicit setting without changing it.

.EXAMPLE
    .\SetOrgDbOrgSettingWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -Setting AllowRoleAssignmentOnDisabledUsers -Value "true" -WhatIf

    Previews an organization-wide change without sign-in or API calls.

.NOTES
    Author: Rick Wilson
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $uri = [uri]$_
        $uri.IsAbsoluteUri -and $uri.Scheme -eq 'https' -and
        $uri.AbsolutePath -eq '/' -and -not $uri.Query -and
        -not $uri.Fragment -and -not $uri.UserInfo
    })]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [ValidateSet('Public', 'GCC', 'GCCH', 'DoD')]
    [string]$Environment = 'Public',

    [ValidateSet('DeviceCode', 'Interactive')]
    [string]$AuthenticationMode = 'DeviceCode',

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z_][A-Za-z0-9_.-]*$')]
    [string]$Setting,

    [ValidateNotNullOrEmpty()]
    [string]$Value
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
$isWrite = $PSBoundParameters.ContainsKey('Value')
$action = "Authenticate and read OrgDBOrgSettings key '$Setting'"
if ($isWrite) { $action = "Authenticate and set OrgDBOrgSettings key '$Setting' to '$Value' (organization-wide)" }
if (($isWrite -or $WhatIfPreference) -and -not $PSCmdlet.ShouldProcess($OrganizationUrl, $action)) {
    return
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir '..\..\EntraID\GetAccessTokenDeviceCode.ps1'
Write-Host 'Acquiring access token...' -ForegroundColor Cyan
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment -AuthenticationMode $AuthenticationMode -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($accessToken)) {
    throw 'Failed to acquire access token. No organization settings were changed.'
}
$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    Setting         = $Setting
    Confirm         = $false
}
if ($isWrite) { $scriptParams.Value = $Value }
$mainScript = Join-Path $scriptDir 'SetOrgDbOrgSetting.ps1'
& $mainScript @scriptParams
