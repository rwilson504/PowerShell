<#
.SYNOPSIS
    Updates selected profile fields on an existing disabled Dataverse stub user.

.DESCRIPTION
    Targets an existing systemuser by GUID and changes only supplied FirstName,
    LastName, Title, EmailAddress, or DomainName values. At least one is required.
    Omitted fields are preserved. An empty Title explicitly clears the job title.
    Changing EmailAddress does not silently change DomainName; supply both to
    change both identity columns.

    Before writing, verifies the built-in Stub Users view criteria: disabled,
    unlicensed, no Entra/application ID, name other than SYSTEM or INTEGRATION,
    and systemmanagedusertype 0 or 2. It does not infer stub status solely from
    being disabled or require issyncwithdirectory to be false.

    Rejects conflicting email/user names on other users. Uses the retrieved ETag
    in If-Match to prevent upserts and reject concurrent changes; the target must
    support optimistic concurrency. Does not retry a failed update automatically.
    Retrieves the record again to verify all profile fields and unchanged status,
    directory, access-mode and business-unit values. No state, roles, business unit
    or organization settings are changed by this script.

    Use fictitious email/user names that will never belong to real Entra users.
    Profile updates may trigger environment-specific automation. The caller needs
    read/write system-user privileges. PowerShell 5.1 or later is required.

.PARAMETER OrganizationUrl
    HTTPS root URL of the Dataverse organization, without an API path.

.PARAMETER AccessToken
    Access token for the target Dataverse organization.

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
    .\UpdateStubUser.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SystemUserId "11111111-1111-1111-1111-111111111111" -Title "Senior External Approver" -WhatIf

    Previews the update without making any Web API requests.

.EXAMPLE
    .\UpdateStubUser.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SystemUserId "11111111-1111-1111-1111-111111111111" -FirstName "Alex" -LastName "Updated (Stub)" -EmailAddress "alex.updated@stub.invalid" -DomainName "alex.updated@stub.invalid"

    Changes the specified name and identity fields, preserving the title and GUID.

.EXAMPLE
    .\UpdateStubUser.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SystemUserId "11111111-1111-1111-1111-111111111111" -Title ""

    Clears only the job title on the existing stub.

.NOTES
    Author: Rick Wilson
    Stub criteria were read from the built-in Stub Users view on 2026-10-21.
    https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/systemuser
    https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/perform-conditional-operations-using-web-api
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

$fieldMap = @{
    FirstName    = 'firstname'
    LastName     = 'lastname'
    Title        = 'title'
    EmailAddress = 'internalemailaddress'
    DomainName   = 'domainname'
}
$requested = @{}
foreach ($name in $fieldMap.Keys) {
    if ($PSBoundParameters.ContainsKey($name)) {
        $value = $PSBoundParameters[$name]
        if ($name -eq 'Title' -and $value -eq '') {
            $value = $null
        }
        $requested[$fieldMap[$name]] = $value
    }
}
if ($requested.Count -eq 0) {
    throw "Supply at least one profile parameter: FirstName, LastName, Title, EmailAddress, or DomainName."
}

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
if (-not $PSCmdlet.ShouldProcess("$SystemUserId in $OrganizationUrl", "Update stub user profile ($($requested.Keys -join ', '))")) {
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
$profileFields = @('firstname', 'lastname', 'title', 'internalemailaddress', 'domainname')
$protectedFields = @('isdisabled', 'islicensed', 'issyncwithdirectory', 'azureactivedirectoryobjectid', 'applicationid', 'systemmanagedusertype', 'accessmode', '_businessunitid_value')
$selectedFields = @('systemuserid', 'fullname') + $profileFields + $protectedFields
$recordUrl = "$apiUrl/systemusers($SystemUserId)?`$select=$($selectedFields -join ',')"

function Test-StubUser {
    param ([psobject]$User)

    foreach ($name in $selectedFields) {
        if ($User.PSObject.Properties.Name -notcontains $name) {
            return $false
        }
    }
    return ($User.systemuserid -eq $SystemUserId.ToString() -and
        $User.isdisabled -eq $true -and $User.islicensed -eq $false -and
        $null -eq $User.azureactivedirectoryobjectid -and $null -eq $User.applicationid -and
        $User.fullname -notin @('SYSTEM', 'INTEGRATION') -and
        $User.systemmanagedusertype -in @(0, 2))
}

$current = Invoke-RestMethod -Method Get -Uri $recordUrl -Headers $headers -ErrorAction Stop
if (-not (Test-StubUser -User $current)) {
    throw "User $SystemUserId did not verify against the built-in Stub Users criteria or the response was incomplete. No user was updated."
}

$payload = @{}
foreach ($field in $requested.Keys) {
    if ([string]$requested[$field] -cne [string]$current.$field) {
        $payload[$field] = $requested[$field]
    }
}
if ($payload.Count -eq 0) {
    Write-Host "Stub user $SystemUserId already has the requested values; no update was sent." -ForegroundColor Yellow
    Write-Output $current
    return
}

if ($payload.ContainsKey('internalemailaddress') -or $payload.ContainsKey('domainname')) {
    $identities = @($current.internalemailaddress, $current.domainname)
    if ($payload.ContainsKey('internalemailaddress')) { $identities[0] = $payload.internalemailaddress }
    if ($payload.ContainsKey('domainname')) { $identities[1] = $payload.domainname }
    $conditions = foreach ($identity in ($identities | Where-Object { $_ } | Select-Object -Unique)) {
        $literal = $identity.Replace("'", "''")
        "internalemailaddress eq '$literal' or domainname eq '$literal'"
    }
    $filter = "systemuserid ne $SystemUserId and ($($conditions -join ' or '))"
    $encodedFilter = [uri]::EscapeDataString($filter)
    $conflicts = Invoke-RestMethod -Method Get -Uri "$apiUrl/systemusers?`$select=systemuserid&`$filter=$encodedFilter&`$top=1" -Headers $headers -ErrorAction Stop
    if (@($conflicts.value).Count -gt 0) {
        throw "Another user already has a requested email or user name (systemuserid: $($conflicts.value[0].systemuserid)). No user was updated."
    }
}

$etag = $current.'@odata.etag'
if ([string]::IsNullOrWhiteSpace($etag) -or $etag -eq '*') {
    throw "User $SystemUserId did not return a usable ETag. No user was updated."
}
$updateHeaders = $headers.Clone()
$updateHeaders['If-Match'] = $etag
$jsonPayload = $payload | ConvertTo-Json -Depth 10
Write-Host "Updating stub user $SystemUserId in $OrganizationUrl..." -ForegroundColor Cyan
$null = Invoke-RestMethod -Method Patch -Uri "$apiUrl/systemusers($SystemUserId)" -Headers $updateHeaders -ContentType 'application/json; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonPayload)) -ErrorAction Stop

try {
    $updated = Invoke-RestMethod -Method Get -Uri $recordUrl -Headers $headers -ErrorAction Stop
}
catch {
    throw "Update completed for systemuserid $SystemUserId, but read-back failed. Inspect that record before retrying. $($_.Exception.Message)"
}
if (-not (Test-StubUser -User $updated)) {
    throw "User $SystemUserId was updated, but read-back did not verify a complete stub-user record. Inspect it before retrying; no automatic rollback was attempted."
}
foreach ($field in ($profileFields + $protectedFields)) {
    $expected = $current.$field
    if ($payload.ContainsKey($field)) { $expected = $payload[$field] }
    if ([string]$updated.$field -cne [string]$expected) {
        throw "User $SystemUserId was updated, but read-back did not verify field '$field'. Inspect it before retrying; no automatic rollback was attempted."
    }
}

Write-Host "Stub user updated and verified (systemuserid: $SystemUserId)." -ForegroundColor Green
Write-Output $updated
