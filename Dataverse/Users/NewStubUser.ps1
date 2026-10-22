<#
.SYNOPSIS
    Creates a disabled Dataverse stub user through the Web API.

.DESCRIPTION
    Creates a placeholder systemuser for historical user references without creating
    a Microsoft Entra account. Sets isdisabled to true, islicensed to false, and
    issyncwithdirectory to false, and systemmanagedusertype to 0 (Entra User).
    The type label does not require an Entra identity for this disabled placeholder.
    Does not link an Entra object or application or request type 2 impersonation.
    Checks for an existing user by ID, email, or user name before creating anything,
    then retrieves the new record and verifies its stub-user flags.

    Use a fictitious email/user name that will never belong to a real Entra user:
    a stub can block later provisioning of a real user with the same identity.
    Stub users cannot sign in or be converted into regular synchronized users.
    This script does not assign roles, change organization settings, or guarantee
    that the user can own records; ownership depends on privileges and table rules.
    The caller needs permission to read business units and read/create system users.
    PowerShell 5.1 or later is required.

.PARAMETER OrganizationUrl
    HTTPS root URL of the Dataverse organization, without an API path.

.PARAMETER AccessToken
    Access token for the target Dataverse organization.

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
    GUID for the new user, for example to preserve a historical user ID during migration.
    Defaults to a newly generated GUID. An existing ID is rejected, never updated.

.EXAMPLE
    .\NewStubUser.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -FirstName "Alex" -LastName "Former (Stub)" -EmailAddress "alex.former@stub.invalid" -WhatIf

    Previews creation without making any Web API requests.

.EXAMPLE
    .\NewStubUser.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -FirstName "Alex" -LastName "Former (Stub)" -EmailAddress "alex.former@stub.invalid" -Title "External Approver"

    Creates a stub in the root business unit and returns the verified systemuser record.

.NOTES
    Author: Rick Wilson
    Microsoft's admin documentation describes stub creation through data import;
    the API approach and field choices are described in these practitioner references.
    Validate against a non-production environment before relying on this for migration.
    https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/systemuser
    https://learn.microsoft.com/en-us/power-platform/admin/create-users#how-stub-users-are-created
    https://theappsbakery.com/migrating-non-existing-users-to-dataverse/
    https://www.kingswaysoft.com/blog/2023/01/24/Manage-Unlicensed-Users-in-CRM-Using-KingswaySoft
#>

[CmdletBinding(SupportsShouldProcess)]
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
    [string]$AccessToken,

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
    [guid]$SystemUserId = [guid]::NewGuid(),

    [ValidatePattern('\S')]
    [ValidateLength(1, 128)]
    [string]$Title
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
if (-not $DomainName) {
    $DomainName = $EmailAddress
}

if (-not $PSCmdlet.ShouldProcess("$EmailAddress in $OrganizationUrl", "Create disabled stub user")) {
    return
}

$apiUrl = "$OrganizationUrl/api/data/v9.2"
$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "Content-Type"     = "application/json"
    "Accept"           = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
}

# Check both identity columns against both inputs to catch overlapping user identities.
$emailLiteral = $EmailAddress.Replace("'", "''")
$domainLiteral = $DomainName.Replace("'", "''")
$filter = "systemuserid eq $SystemUserId or internalemailaddress eq '$emailLiteral' or domainname eq '$emailLiteral' or internalemailaddress eq '$domainLiteral' or domainname eq '$domainLiteral'"
$encodedFilter = [uri]::EscapeDataString($filter)
$existing = Invoke-RestMethod -Method Get -Uri "$apiUrl/systemusers?`$select=systemuserid&`$filter=$encodedFilter&`$top=1" -Headers $headers -ErrorAction Stop
if (@($existing.value).Count -gt 0) {
    throw "A user already exists with the requested ID, email, or user name (systemuserid: $($existing.value[0].systemuserid)). No user was created or updated."
}

if (-not $PSBoundParameters.ContainsKey('BusinessUnitId')) {
    $rootUnits = Invoke-RestMethod -Method Get -Uri "$apiUrl/businessunits?`$select=businessunitid&`$filter=_parentbusinessunitid_value eq null&`$top=2" -Headers $headers -ErrorAction Stop
    if (@($rootUnits.value).Count -ne 1) {
        throw "Expected exactly one root business unit. Supply -BusinessUnitId explicitly. No user was created."
    }
    $BusinessUnitId = [guid]$rootUnits.value[0].businessunitid
    if ($BusinessUnitId -eq [guid]::Empty) {
        throw "The root business unit response did not contain a valid ID. No user was created."
    }
}

$payload = @{
    systemuserid                 = $SystemUserId.ToString()
    firstname                    = $FirstName
    lastname                     = $LastName
    internalemailaddress         = $EmailAddress
    domainname                   = $DomainName
    'businessunitid@odata.bind'   = "/businessunits($BusinessUnitId)"
    isdisabled                   = $true
    islicensed                   = $false
    issyncwithdirectory           = $false
    systemmanagedusertype        = 0
}
if ($PSBoundParameters.ContainsKey('Title')) {
    $payload.title = $Title
}

Write-Host "Creating disabled stub user in $OrganizationUrl (systemuserid: $SystemUserId)..." -ForegroundColor Cyan
$jsonPayload = $payload | ConvertTo-Json -Depth 10
# The ID is assigned client-side so a 204 response still permits an independent read-back.
$null = Invoke-RestMethod -Method Post -Uri "$apiUrl/systemusers" -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonPayload)) -ErrorAction Stop

$select = 'systemuserid,fullname,firstname,lastname,title,internalemailaddress,domainname,isdisabled,islicensed,issyncwithdirectory,systemmanagedusertype,azureactivedirectoryobjectid,applicationid,_businessunitid_value'
try {
    $user = Invoke-RestMethod -Method Get -Uri "$apiUrl/systemusers($SystemUserId)?`$select=$select" -Headers $headers -ErrorAction Stop
}
catch {
    throw "Create completed for systemuserid $SystemUserId, but read-back failed. Inspect that record before retrying. $($_.Exception.Message)"
}

if ($user.systemuserid -ne $SystemUserId.ToString() -or
    $user.isdisabled -ne $true -or $user.islicensed -ne $false -or
    $user.issyncwithdirectory -ne $false -or
    $user.systemmanagedusertype -ne 0 -or
    ($user.azureactivedirectoryobjectid -and $user.azureactivedirectoryobjectid -ne [guid]::Empty.ToString()) -or
    ($user.applicationid -and $user.applicationid -ne [guid]::Empty.ToString()) -or
    $user._businessunitid_value -ne $BusinessUnitId.ToString()) {
    throw "User $SystemUserId was created, but read-back did not verify a disabled, unlicensed, unsynchronized placeholder with systemmanagedusertype 0 (Entra User) in the requested business unit with no Entra/application identity. Inspect the record before retrying; it was not automatically deleted."
}
if ($PSBoundParameters.ContainsKey('Title') -and $user.title -cne $Title) {
    throw "User $SystemUserId was created, but read-back did not verify the requested title. Inspect the record before retrying; it was not automatically deleted."
}

Write-Host "Stub user created and verified (systemuserid: $SystemUserId)." -ForegroundColor Green
Write-Output $user
