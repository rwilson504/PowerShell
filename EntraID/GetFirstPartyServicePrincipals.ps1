<#
.SYNOPSIS
    Gets all first-party (Microsoft-owned) service principals in the tenant.

.DESCRIPTION
    This script queries Microsoft Graph to retrieve all service principals owned by Microsoft
    (first-party apps) in the current tenant. It identifies first-party apps by their
    appOwnerOrganizationId matching known Microsoft tenant IDs.

    The script handles pagination automatically using @odata.nextLink and supports
    Public, GCC, GCCH, and DoD environments.

.PARAMETER AccessToken
    The access token for authenticating with Microsoft Graph.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".
    This determines which Microsoft Graph endpoint to use.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results. If not provided, results are written to the console.

.PARAMETER IncludeDisabled
    When specified, includes service principals that are disabled (accountEnabled = false).
    By default, only enabled service principals are returned.

.EXAMPLE
    .\GetFirstPartyServicePrincipals.ps1 -AccessToken $token -Environment "GCCH"

    Gets all first-party service principals in a GCC High tenant and displays them in table format.

.EXAMPLE
    .\GetFirstPartyServicePrincipals.ps1 -AccessToken $token -OutputFormat CSV -OutputPath "C:\temp\first-party-sps.csv"

    Gets all first-party service principals and exports to CSV.

.EXAMPLE
    .\GetFirstPartyServicePrincipals.ps1 -AccessToken $token -OutputFormat JSON

    Gets all first-party service principals and outputs as JSON.

.EXAMPLE
    .\GetFirstPartyServicePrincipals.ps1 -AccessToken $token -IncludeDisabled

    Gets all first-party service principals including disabled ones.

.AUTHOR
    Rick Wilson

.NOTES
    Requires an access token with at least Application.Read.All or Directory.Read.All permissions
    on Microsoft Graph.

    You can use the built-in "Microsoft Graph PowerShell" app registration
    (AppId: 14d82eec-204b-4c2f-b7e8-296a70dab67e) with the device code flow to acquire a token
    with the necessary permissions.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

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

# Determine the Microsoft Graph endpoint based on environment
switch ($Environment) {
    "Public" {
        $GraphEndpoint = "https://graph.microsoft.com"
    }
    "GCC" {
        $GraphEndpoint = "https://graph.microsoft.com"
    }
    "GCCH" {
        $GraphEndpoint = "https://graph.microsoft.us"
    }
    "DoD" {
        $GraphEndpoint = "https://dod-graph.microsoft.us"
    }
}

# Known Microsoft tenant IDs that own first-party applications
$MicrosoftTenantIds = @(
    "f8cdef31-a31e-4b4a-93e4-5f571e91255a"  # Microsoft Services
    "72f988bf-86f1-41af-91ab-2d7cd011db47"  # Microsoft Corp
)

# Set up headers for Graph API calls
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json"
    "ConsistencyLevel" = "eventual"
}

# Build the filter for Microsoft-owned service principals
$filterParts = $MicrosoftTenantIds | ForEach-Object { "appOwnerOrganizationId eq $_" }
$filter = $filterParts -join " or "

# Select only the properties we need
$select = "id,appId,displayName,description,appOwnerOrganizationId,accountEnabled,servicePrincipalType,signInAudience,tags,appDisplayName,homepage,loginUrl,logoutUrl,replyUrls,servicePrincipalNames,publishedPermissionScopes,appRoles"

# Build the initial request URL with proper encoding
$baseUrl = "$GraphEndpoint/beta/servicePrincipals"
$encodedFilter = [System.Uri]::EscapeDataString($filter)
$encodedSelect = [System.Uri]::EscapeDataString($select)
$requestUrl = $baseUrl + '?$filter=' + $encodedFilter + '&$select=' + $encodedSelect + '&$top=100&$count=true'

Write-Host "Fetching first-party service principals from Microsoft Graph ($Environment)..." -ForegroundColor Cyan

$allServicePrincipals = [System.Collections.Generic.List[PSObject]]::new()
$pageCount = 0

# Paginate through all results
do {
    $pageCount++
    Write-Host "  Fetching page $pageCount..." -ForegroundColor Cyan

    try {
        $response = Invoke-RestMethod -Method Get -Uri $requestUrl -Headers $headers
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 429) {
            $retryAfter = 5
            if ($_.Exception.Response.Headers["Retry-After"]) {
                $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
            }
            Write-Warning "Rate limited (429). Waiting $retryAfter seconds before retrying..."
            Start-Sleep -Seconds $retryAfter
            continue
        }
        Write-Error "Failed to query Microsoft Graph: $($_.Exception.Message)"
        if ($_.ErrorDetails.Message) {
            Write-Error "Details: $($_.ErrorDetails.Message)"
        }
        exit 1
    }

    if ($response.value) {
        foreach ($sp in $response.value) {
            $allServicePrincipals.Add($sp)
        }
    }

    # Check for next page
    $requestUrl = $response.'@odata.nextLink'

} while ($requestUrl)

Write-Host "Retrieved $($allServicePrincipals.Count) first-party service principals." -ForegroundColor Green

# Filter out disabled if not requested
if (-not $IncludeDisabled) {
    $allServicePrincipals = $allServicePrincipals | Where-Object { $_.accountEnabled -eq $true }
    Write-Host "After filtering disabled: $($allServicePrincipals.Count) enabled service principals." -ForegroundColor Green
}

# Sort by display name
$allServicePrincipals = $allServicePrincipals | Sort-Object -Property displayName

# Build the output objects
$results = $allServicePrincipals | ForEach-Object {
    $ownerName = switch ($_.appOwnerOrganizationId) {
        "f8cdef31-a31e-4b4a-93e4-5f571e91255a" { "Microsoft Services" }
        "72f988bf-86f1-41af-91ab-2d7cd011db47" { "Microsoft Corp" }
        default { $_.appOwnerOrganizationId }
    }

    # Extract URL identifiers from servicePrincipalNames (exclude GUIDs)
    $identifierUris = @()
    if ($_.servicePrincipalNames) {
        $identifierUris = $_.servicePrincipalNames | Where-Object { $_ -match '^https?://' }
    }

    # Format delegated permissions (publishedPermissionScopes)
    $delegatedPermissions = ""
    if ($_.publishedPermissionScopes) {
        $delegatedPermissions = ($_.publishedPermissionScopes | ForEach-Object {
            $_.value
        }) -join '; '
    }

    # Format application roles (appRoles)
    $applicationRoles = ""
    if ($_.appRoles) {
        $applicationRoles = ($_.appRoles | ForEach-Object {
            $_.value
        }) -join '; '
    }

    [PSCustomObject]@{
        DisplayName              = $_.displayName
        Description              = $_.description
        AppId                    = $_.appId
        ObjectId                 = $_.id
        Enabled                  = $_.accountEnabled
        Homepage                 = $_.homepage
        LoginUrl                 = $_.loginUrl
        LogoutUrl                = $_.logoutUrl
        ReplyUrls                = ($_.replyUrls -join '; ')
        IdentifierUris           = ($identifierUris -join '; ')
        DelegatedPermissions     = $delegatedPermissions
        ApplicationRoles         = $applicationRoles
        ServicePrincipalType     = $_.servicePrincipalType
        SignInAudience           = $_.signInAudience
        AppOwner                 = $ownerName
        AppOwnerOrganizationId   = $_.appOwnerOrganizationId
    }
}

# Output results
switch ($OutputFormat) {
    "Table" {
        if ($OutputPath) {
            $results | Format-Table -AutoSize | Out-String | Set-Content -Path $OutputPath -Encoding UTF8
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        else {
            $results | Format-Table -AutoSize
        }
    }
    "CSV" {
        if ($OutputPath) {
            $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        else {
            $results | ConvertTo-Csv -NoTypeInformation
        }
    }
    "JSON" {
        $jsonOutput = $results | ConvertTo-Json -Depth 5
        if ($OutputPath) {
            $jsonOutput | Set-Content -Path $OutputPath -Encoding UTF8
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        else {
            $jsonOutput
        }
    }
}

return $results
