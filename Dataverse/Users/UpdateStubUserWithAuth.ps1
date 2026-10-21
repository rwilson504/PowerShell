<#
.SYNOPSIS
    Acquires an access token and updates selected Dataverse stub-user profile fields.

.DESCRIPTION
    Uses the shared EntraID authentication helper, then calls UpdateStubUser.ps1.
    Targets an existing stub by GUID and forwards only supplied profile fields.
    Requires at least one profile parameter. Does not enable users or change roles.
    Omitted fields are preserved; an empty Title explicitly clears the job title.
    EmailAddress and DomainName are independent. Use fictitious identities that
    will never belong to real Entra users. Profile updates may trigger existing
    automation. See UpdateStubUser.ps1 for stub criteria and concurrency checks.
    PowerShell 5.1 or later is required.

.PARAMETER TenantId
    Microsoft Entra tenant ID.

.PARAMETER ClientId
    Application ID of the public-client app registration used for authentication.

.PARAMETER Environment
    Azure cloud: Public, GCC, GCCH, or DoD. Defaults to Public.

.PARAMETER AuthenticationMode
    DeviceCode or Interactive (system browser with PKCE). Defaults to DeviceCode.

.PARAMETER OrganizationUrl
    HTTPS root URL of the Dataverse organization, without an API path.

.PARAMETER SystemUserId
    Required GUID of the existing stub user. Missing records are never created.

.PARAMETER FirstName
    Optional new first name, from 1 to 64 characters. Omitted by default.

.PARAMETER LastName
    Optional new last name, from 1 to 64 characters. Omitted by default.

.PARAMETER Title
    Optional new job title, up to 128 characters. Use -Title "" to clear it.
    Omitting this parameter preserves the current title.

.PARAMETER EmailAddress
    Optional new fictitious primary email address. Does not change DomainName.

.PARAMETER DomainName
    Optional new fictitious user name. Omitted by default.

.EXAMPLE
    .\UpdateStubUserWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -SystemUserId "11111111-1111-1111-1111-111111111111" -Title "Senior External Approver" -AuthenticationMode Interactive

    Signs in through the browser, updates only the title, and returns the verified record.

.EXAMPLE
    .\UpdateStubUserWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -SystemUserId "11111111-1111-1111-1111-111111111111" -Title "" -WhatIf

    Previews clearing the title without authenticating or making Web API requests.

.NOTES
    Author: Rick Wilson
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [ValidateSet("Public", "GCC", "GCCH", "DoD")]
    [string]$Environment = "Public",

    [ValidateSet("DeviceCode", "Interactive")]
    [string]$AuthenticationMode = "DeviceCode",

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        $uri = [uri]$_
        $uri.IsAbsoluteUri -and $uri.Scheme -eq 'https' -and
        $uri.AbsolutePath -eq '/' -and -not $uri.Query -and
        -not $uri.Fragment -and -not $uri.UserInfo
    })]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ $_ -ne [guid]::Empty })]
    [guid]$SystemUserId,

    [ValidatePattern('\S')]
    [ValidateLength(1, 64)]
    [string]$FirstName,

    [ValidatePattern('\S')]
    [ValidateLength(1, 64)]
    [string]$LastName,

    [AllowEmptyString()]
    [ValidateScript({ $_.Length -eq 0 -or $_ -match '\S' })]
    [ValidateLength(0, 128)]
    [string]$Title,

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [ValidateLength(1, 100)]
    [string]$EmailAddress,

    [ValidatePattern('^\S+$')]
    [ValidateLength(1, 1024)]
    [string]$DomainName
)

$profileParameters = @('FirstName', 'LastName', 'Title', 'EmailAddress', 'DomainName')
if (@($profileParameters | Where-Object { $PSBoundParameters.ContainsKey($_) }).Count -eq 0) {
    throw "Supply at least one profile parameter: FirstName, LastName, Title, EmailAddress, or DomainName."
}
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
if (-not $PSCmdlet.ShouldProcess("$SystemUserId in $OrganizationUrl", "Authenticate and update stub user profile")) {
    return
}

Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment -AuthenticationMode $AuthenticationMode -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($accessToken)) {
    throw "Failed to acquire access token. No user was updated."
}

$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    SystemUserId    = $SystemUserId
    Confirm         = $false
}
foreach ($name in $profileParameters) {
    if ($PSBoundParameters.ContainsKey($name)) {
        $scriptParams[$name] = $PSBoundParameters[$name]
    }
}

$mainScript = Join-Path $scriptDir "UpdateStubUser.ps1"
& $mainScript @scriptParams
