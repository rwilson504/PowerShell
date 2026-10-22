<#
.SYNOPSIS
    Reads or sets one Dataverse OrgDBOrgSettings XML value.

.DESCRIPTION
    Reads one case-sensitive setting from the organization's OrgDBOrgSettings XML.
    Supply Value to add or replace that setting; omit Value for a read-only query.
    Preserves other XML settings, requires an ETag for writes, and independently
    reads back the saved value. A concurrent change rejects the write; no automatic
    retry or rollback is attempted. Requires organization read/write privileges.
    Only use documented, supported keys and values for the target environment.
    This script does not discover feature gates or edit direct organization columns.
    A successful XML update does not prove a feature is available or operational.
    WhatIf makes no API calls. Requires PowerShell 5.1 or later.

.PARAMETER OrganizationUrl
    HTTPS root URL of the Dataverse organization, without an API path.

.PARAMETER AccessToken
    Access token for the target Dataverse organization.

.PARAMETER Setting
    Exact, case-sensitive XML element name of a supported OrgDBOrgSettings key.
    Namespaces and XPath expressions are not accepted.

.PARAMETER Value
    Nonempty string to store as XML text, for example "true". Omit to read only.
    Valid values and defaults depend on the setting; no default value is inferred.

.EXAMPLE
    .\SetOrgDbOrgSetting.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Setting AllowRoleAssignmentOnDisabledUsers

    Reads the explicit setting. IsExplicit is false when the platform default applies.

.EXAMPLE
    .\SetOrgDbOrgSetting.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Setting AllowRoleAssignmentOnDisabledUsers -Value "true" -WhatIf

    Previews allowing role assignment to disabled users without making API calls.
    This is not the Stub User Impersonation feature.

.NOTES
    Author: Rick Wilson
    https://support.microsoft.com/en-us/topic/orgdborgsettings-tool-for-microsoft-dynamics-crm-20a10f46-2a24-a156-7144-365d49b842ba
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
    [string]$AccessToken,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[A-Za-z_][A-Za-z0-9_.-]*$')]
    [string]$Setting,

    [ValidateNotNullOrEmpty()]
    [string]$Value
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
$isWrite = $PSBoundParameters.ContainsKey('Value')
$action = "Read OrgDBOrgSettings key '$Setting'"
if ($isWrite) { $action = "Set OrgDBOrgSettings key '$Setting' to '$Value' (organization-wide)" }
if (($isWrite -or $WhatIfPreference) -and -not $PSCmdlet.ShouldProcess($OrganizationUrl, $action)) {
    return
}

function Read-OrgSettingsXml {
    param ([string]$Text)

    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $true
    $document.XmlResolver = $null
    if ([string]::IsNullOrWhiteSpace($Text)) { $Text = '<OrgSettings />' }
    $readerSettings = New-Object System.Xml.XmlReaderSettings
    $readerSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $readerSettings.XmlResolver = $null
    $stringReader = New-Object System.IO.StringReader($Text)
    $reader = $null
    try {
        $reader = [System.Xml.XmlReader]::Create($stringReader, $readerSettings)
        $document.Load($reader)
    }
    finally {
        if ($reader) { $reader.Dispose() }
        $stringReader.Dispose()
    }
    if ($document.DocumentElement.Name -cne 'OrgSettings' -or $document.DocumentElement.NamespaceURI) {
        throw 'Expected an unnamespaced OrgSettings XML root. No settings were written.'
    }
    return ,$document
}

function Get-SettingNode {
    param ([System.Xml.XmlDocument]$Document)

    $nodes = @($Document.DocumentElement.SelectNodes($Setting))
    if ($nodes.Count -gt 1) {
        throw "Duplicate '$Setting' elements found. Inspect the organization's settings before retrying."
    }
    if ($nodes.Count -eq 1) {
        if (@($nodes[0].SelectNodes('*')).Count -gt 0) {
            throw "Setting '$Setting' contains nested elements, not a scalar value."
        }
        return $nodes[0]
    }
}

$apiUrl = "$OrganizationUrl/api/data/v9.2"
$headers = @{
    Authorization      = "Bearer $AccessToken"
    'Content-Type'     = 'application/json'
    Accept             = 'application/json'
    'OData-MaxVersion' = '4.0'
    'OData-Version'    = '4.0'
}
$response = Invoke-RestMethod -Method Get -Uri "$apiUrl/organizations?`$select=organizationid,name,orgdborgsettings&`$top=2" -Headers $headers -ErrorAction Stop
if (@($response.value).Count -ne 1) {
    throw 'Expected exactly one organization row. No settings were written.'
}
$organization = $response.value[0]
$organizationId = [guid]$organization.organizationid
if ($organizationId -eq [guid]::Empty -or $organization.PSObject.Properties.Name -notcontains 'orgdborgsettings') {
    throw 'Organization response is incomplete. No settings were written.'
}
$document = Read-OrgSettingsXml -Text $organization.orgdborgsettings
$node = Get-SettingNode -Document $document
$currentValue = $null
if ($null -ne $node) { $currentValue = $node.InnerText }

if ($isWrite -and $currentValue -cne $Value) {
    $etag = $organization.'@odata.etag'
    if ([string]::IsNullOrWhiteSpace($etag) -or $etag -eq '*') {
        throw 'Organization response did not contain a usable ETag. No settings were written.'
    }
    if ($null -eq $node) {
        $node = $document.CreateElement($Setting)
        [void]$document.DocumentElement.AppendChild($node)
    }
    $node.InnerText = $Value
    $payload = @{ orgdborgsettings = $document.DocumentElement.OuterXml } | ConvertTo-Json -Depth 10
    $writeHeaders = $headers.Clone()
    $writeHeaders['If-Match'] = $etag
    Write-Host "Updating OrgDBOrgSettings key '$Setting' in $OrganizationUrl..." -ForegroundColor Cyan
    $null = Invoke-RestMethod -Method Patch -Uri "$apiUrl/organizations($organizationId)" -Headers $writeHeaders -ContentType 'application/json; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes($payload)) -ErrorAction Stop
    try {
        $saved = Invoke-RestMethod -Method Get -Uri "$apiUrl/organizations($organizationId)?`$select=organizationid,orgdborgsettings" -Headers $headers -ErrorAction Stop
        if ($saved.organizationid -ne $organizationId.ToString() -or $saved.PSObject.Properties.Name -notcontains 'orgdborgsettings') {
            throw 'Read-back returned an incomplete or mismatched organization.'
        }
        $savedDocument = Read-OrgSettingsXml -Text $saved.orgdborgsettings
        $node = Get-SettingNode -Document $savedDocument
        if ($null -eq $node -or $node.InnerText -cne $Value) {
            throw "Read-back did not match the requested value for '$Setting'."
        }
    }
    catch {
        throw "PATCH completed for organization $organizationId, but verification failed. Inspect '$Setting' before retrying; no rollback was attempted. $($_.Exception.Message)"
    }
    Write-Host "OrgDBOrgSettings key '$Setting' updated and verified." -ForegroundColor Green
}
elseif ($isWrite) {
    Write-Host "OrgDBOrgSettings key '$Setting' already has the requested value; no update was sent." -ForegroundColor Yellow
}

$resultValue = $null
if ($null -ne $node) { $resultValue = $node.InnerText }
[pscustomobject]@{
    OrganizationUrl = $OrganizationUrl
    OrganizationId  = $organizationId
    OrganizationName = $organization.name
    Setting         = $Setting
    IsExplicit      = ($null -ne $node)
    Value           = $resultValue
}
