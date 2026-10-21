<#
.SYNOPSIS
    Acquires an access token and creates a disabled Dataverse stub user through the Web API.

.DESCRIPTION
    Uses the shared EntraID authentication helper, then calls NewStubUser.ps1.
    Creates a placeholder, not a sign-in account. Use fictitious email/user names
    that will never belong to real Entra users to avoid blocking later provisioning.
    No roles or organization settings are changed. The caller needs permission to
    read business units and read/create system users. See NewStubUser.ps1 for API
    references and limitations. PowerShell 5.1 or later is required.

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

.PARAMETER FirstName
    First name of the placeholder user.

.PARAMETER LastName
    Last name of the placeholder user. A suffix such as "(Stub)" is recommended.

.PARAMETER Title
    Optional job title, up to 128 characters. Omitted by default.

.PARAMETER EmailAddress
    Fictitious email address for internalemailaddress. Do not use a real user's email.

.PARAMETER DomainName
    Fictitious user name for domainname. Defaults to EmailAddress.

.PARAMETER BusinessUnitId
    Target business unit GUID. Defaults to the organization's root business unit.

.PARAMETER SystemUserId
    GUID for the new user. Defaults to a newly generated GUID in the base script.
    An existing ID is rejected, never updated.

.EXAMPLE
    .\NewStubUserWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -FirstName "Alex" -LastName "Former (Stub)" -EmailAddress "alex.former@stub.invalid" -Title "External Approver" -AuthenticationMode Interactive

    Signs in through the browser, creates the stub, and returns its verified record.

.EXAMPLE
    .\NewStubUserWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -FirstName "Alex" -LastName "Former (Stub)" -EmailAddress "alex.former@stub.invalid" -WhatIf

    Previews creation without authenticating or making Web API requests.

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
    [ValidatePattern('\S')]
    [ValidateLength(1, 64)]
    [string]$FirstName,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('\S')]
    [ValidateLength(1, 64)]
    [string]$LastName,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [ValidateLength(1, 100)]
    [string]$EmailAddress,

    [ValidatePattern('^\S+$')]
    [ValidateLength(1, 1024)]
    [string]$DomainName,

    [ValidateScript({ $_ -ne [guid]::Empty })]
    [guid]$BusinessUnitId,

    [ValidateScript({ $_ -ne [guid]::Empty })]
    [guid]$SystemUserId,

    [ValidatePattern('\S')]
    [ValidateLength(1, 128)]
    [string]$Title
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
if (-not $PSCmdlet.ShouldProcess("$EmailAddress in $OrganizationUrl", "Authenticate and create disabled stub user")) {
    return
}

Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment -AuthenticationMode $AuthenticationMode -ErrorAction Stop
if ([string]::IsNullOrWhiteSpace($accessToken)) {
    throw "Failed to acquire access token. No user was created."
}

$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    FirstName       = $FirstName
    LastName        = $LastName
    EmailAddress    = $EmailAddress
    Confirm         = $false
}
foreach ($name in @('Title', 'DomainName', 'BusinessUnitId', 'SystemUserId')) {
    if ($PSBoundParameters.ContainsKey($name)) {
        $scriptParams[$name] = $PSBoundParameters[$name]
    }
}

$mainScript = Join-Path $scriptDir "NewStubUser.ps1"
& $mainScript @scriptParams
