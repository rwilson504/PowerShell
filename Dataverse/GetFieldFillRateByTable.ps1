<#
.SYNOPSIS
    Reports per-field fill rates (number of records that contain data) for one or more
    Dataverse tables.

.DESCRIPTION
    For each requested table the script:
      1. Looks up the EntitySetName, primary key, and the list of attributes (columns)
         from the EntityDefinitions / Attributes metadata endpoints.
      2. Gets the total record count for the table (filtered by -Filter when supplied).
      3. For each in-scope column issues an OData query with
         ?$filter=<col> ne null&$count=true&$top=1 to count populated records.
         Multiple field queries are bundled into OData $batch HTTP calls so a typical
         200-attribute scan is ~4 round trips instead of 200.
      4. Computes a fill rate (PopulatedCount / TotalCount).

    Lookup attributes are queried using their _<name>_value alias so the null-check works.

    System-managed and bookkeeping columns are excluded by default (createdon, modifiedon,
    versionnumber, primary key, statecode, statuscode, owner-related lookups, etc.). Use
    -IncludeSystemAttributes to include them. Composite columns (e.g. fullname,
    address1_composite) and lookup-projection sub-attributes (e.g. _name / _yominame
    aliases) are always skipped because they duplicate values from other columns.

    Customer / Owner / activityparty polymorphic lookups are skipped automatically because
    they cannot be filtered with the simple "ne null" syntax.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Required. One or more table logical names to analyze (e.g., "account", "msf_program").

.PARAMETER Attributes
    Optional. Restrict analysis to the specified attribute logical names. When omitted
    every eligible attribute on the table is queried.

.PARAMETER IncludeSystemAttributes
    Include system-managed columns (createdon, modifiedon, versionnumber, statecode,
    statuscode, primary key, owner / createdby / modifiedby lookups, etc.). Default is
    to skip them because they are virtually always populated and add noise.

.PARAMETER CustomAttributesOnly
    When set, only attributes where IsCustomAttribute is true are analyzed.

.PARAMETER StandardAttributesOnly
    When set, only attributes where IsCustomAttribute is false are analyzed.
    Mutually exclusive with -CustomAttributesOnly.

.PARAMETER Filter
    Optional OData $filter expression that restricts which records are considered when
    computing fill rates. Useful for analyzing a subset (e.g., active records, records
    created in the last year). Example: "statecode eq 0", or
    "statecode eq 0 and createdon ge 2025-01-01T00:00:00Z".
    The filter is applied to BOTH the total count and every per-attribute count, so
    the resulting percentages reflect "of records matching <filter>, how many have
    this column populated".

.PARAMETER BatchRequestSize
    Number of per-attribute count requests bundled into a single OData $batch HTTP call.
    Default is 50. Dataverse allows up to 1000 sub-requests per batch but smaller batches
    keep the URL/body well under platform limits and provide more progress granularity.
    Set to 1 to disable batching (issue every request individually).

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetFieldFillRateByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_company"

    Reports the fill rate of every business attribute on the msf_company table.

.EXAMPLE
    .\GetFieldFillRateByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account","contact" -CustomAttributesOnly -OutputFormat CSV -OutputPath ".\fillrates.csv"

    Reports fill rates for custom attributes only on account and contact, exported to CSV.

.EXAMPLE
    .\GetFieldFillRateByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program" -Attributes "msf_status","msf_priority","msf_owner"

    Reports fill rates for only the specified attributes.

.EXAMPLE
    .\GetFieldFillRateByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account" -Filter "statecode eq 0"

    Reports fill rates restricted to active accounts only.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $true)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [string[]]$Attributes,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemAttributes,

    [Parameter(Mandatory = $false)]
    [switch]$CustomAttributesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$StandardAttributesOnly,

    [Parameter(Mandatory = $false)]
    [string]$Filter,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$BatchRequestSize = 50,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

if ($CustomAttributesOnly -and $StandardAttributesOnly) {
    Write-Error "-CustomAttributesOnly and -StandardAttributesOnly are mutually exclusive."
    exit 1
}

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Content-Type"     = "application/json; charset=utf-8"
    "Prefer"           = "odata.include-annotations=*"
}

# System / managed columns to skip by default. These are almost always populated and
# rarely interesting for fill-rate analysis.
$DefaultSkipAttributes = @(
    'createdon','createdby','createdonbehalfby',
    'modifiedon','modifiedby','modifiedonbehalfby',
    'overriddencreatedon','importsequencenumber',
    'versionnumber','timezoneruleversionnumber','utcconversiontimezonecode',
    'statecode','statuscode',
    'ownerid','owningbusinessunit','owninguser','owningteam',
    'organizationid'
)

# Attribute types that cannot be filtered with a simple "ne null" expression
$UnsupportedAttributeTypes = @('Customer','Owner','PartyList','Virtual','EntityName','CalendarRules','ManagedProperty','Uniqueidentifier_Sub')

function Get-EntityMetadata {
    <#
    .SYNOPSIS
        Returns metadata for a single table, or $null if it can't be resolved.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$LogicalName
    )

    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')?" +
        "`$select=LogicalName,SchemaName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,DisplayName"

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $displayName = if ($resp.DisplayName.UserLocalizedLabel) {
            $resp.DisplayName.UserLocalizedLabel.Label
        } else {
            $LogicalName
        }
        return [PSCustomObject]@{
            LogicalName          = $resp.LogicalName
            SchemaName           = $resp.SchemaName
            EntitySetName        = $resp.EntitySetName
            PrimaryIdAttribute   = $resp.PrimaryIdAttribute
            PrimaryNameAttribute = $resp.PrimaryNameAttribute
            DisplayName          = $displayName
        }
    }
    catch {
        Write-Warning "Failed to load metadata for table '$LogicalName': $_"
        return $null
    }
}

function Get-TableAttributes {
    <#
    .SYNOPSIS
        Returns the list of attribute metadata records for a table. Each item exposes
        LogicalName, SchemaName, DisplayName, AttributeType, IsCustomAttribute,
        IsValidForRead, IsLogical, AttributeOf.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$LogicalName
    )

    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes?" +
        "`$select=LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute,IsValidForRead,IsLogical,AttributeOf"

    $all = @()
    do {
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        foreach ($attr in $resp.value) {
            $displayName = if ($attr.DisplayName.UserLocalizedLabel) {
                $attr.DisplayName.UserLocalizedLabel.Label
            } else {
                $attr.LogicalName
            }
            $all += [PSCustomObject]@{
                LogicalName       = $attr.LogicalName
                SchemaName        = $attr.SchemaName
                DisplayName       = $displayName
                AttributeType     = $attr.AttributeType
                IsCustomAttribute = [bool]$attr.IsCustomAttribute
                IsValidForRead    = [bool]$attr.IsValidForRead
                IsLogical         = [bool]$attr.IsLogical
                AttributeOf       = $attr.AttributeOf
            }
        }
        $url = $resp.'@odata.nextLink'
    } while ($url)

    return $all
}

function Get-TotalRecordCount {
    <#
    .SYNOPSIS
        Returns the total record count for a table. When -Filter is supplied uses the
        $count=true endpoint with that filter; otherwise uses the lightweight
        /<set>/$count endpoint.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$EntitySetName,
        [string]$Filter
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Filter)) {
            $url = "$OrgUrl/api/data/v9.2/$EntitySetName/`$count"
            $resp = Invoke-WebRequest -Uri $url -Headers $Headers -Method Get -UseBasicParsing
            $text = if ($resp.Content -is [byte[]]) {
                [System.Text.Encoding]::UTF8.GetString($resp.Content)
            }
            else {
                [string]$resp.Content
            }
            $text = $text.Trim().TrimStart([char]0xFEFF)
            return [long]$text
        }
        else {
            $encoded = [System.Uri]::EscapeDataString($Filter)
            $url = "$OrgUrl/api/data/v9.2/$EntitySetName" +
                   "?`$filter=$encoded&`$count=true&`$top=1"
            $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
            return [long]$resp.'@odata.count'
        }
    }
    catch {
        Write-Warning "Failed to get total count for '$EntitySetName' (filter='$Filter'): $_"
        return $null
    }
}

function Build-PopulatedFilter {
    <#
    .SYNOPSIS
        Combines the user's -Filter with a "<col> ne null" predicate.
    #>
    param (
        [string]$ColumnFilterExpr,
        [string]$BaseFilter
    )
    if ([string]::IsNullOrWhiteSpace($BaseFilter)) {
        return "$ColumnFilterExpr ne null"
    }
    return "($BaseFilter) and ($ColumnFilterExpr ne null)"
}

function Get-FieldPopulatedCount {
    <#
    .SYNOPSIS
        Single-request fallback for when batching is disabled. Returns count or $null on error.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$EntitySetName,
        [string]$ColumnFilterExpr,
        [string]$BaseFilter
    )
    try {
        $combined = Build-PopulatedFilter -ColumnFilterExpr $ColumnFilterExpr -BaseFilter $BaseFilter
        $encoded  = [System.Uri]::EscapeDataString($combined)
        $url = "$OrgUrl/api/data/v9.2/$EntitySetName" +
               "?`$filter=$encoded&`$count=true&`$top=1&`$select=$ColumnFilterExpr"
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        return [long]$resp.'@odata.count'
    }
    catch {
        return $null
    }
}

function Invoke-ODataBatch {
    <#
    .SYNOPSIS
        Posts a multipart/mixed OData $batch containing GET sub-requests and returns
        the parsed JSON body of each response, in request order. Failed sub-requests
        return $null in their slot.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string[]]$RelativeRequests   # e.g. "accounts?$filter=...&$count=true&$top=1"
    )

    $boundary = "batch_$([guid]::NewGuid().ToString('N'))"
    $sb = [System.Text.StringBuilder]::new()
    foreach ($req in $RelativeRequests) {
        # CRLF line endings are required by RFC 2046 / OData batch spec
        [void]$sb.Append("--$boundary`r`n")
        [void]$sb.Append("Content-Type: application/http`r`n")
        [void]$sb.Append("Content-Transfer-Encoding: binary`r`n")
        [void]$sb.Append("`r`n")
        [void]$sb.Append("GET /api/data/v9.2/$req HTTP/1.1`r`n")
        [void]$sb.Append("Accept: application/json`r`n")
        [void]$sb.Append("OData-Version: 4.0`r`n")
        [void]$sb.Append("OData-MaxVersion: 4.0`r`n")
        [void]$sb.Append("`r`n")
    }
    [void]$sb.Append("--$boundary--`r`n")
    $body = $sb.ToString()

    # Build a request-specific header set: keep the auth header but swap Content-Type
    $batchHeaders = @{}
    foreach ($k in $Headers.Keys) {
        if ($k -ne 'Content-Type') { $batchHeaders[$k] = $Headers[$k] }
    }
    $batchHeaders['Content-Type'] = "multipart/mixed; boundary=$boundary"

    try {
        $resp = Invoke-WebRequest -Uri "$OrgUrl/api/data/v9.2/`$batch" `
            -Method Post -Headers $batchHeaders -Body $body -UseBasicParsing
    }
    catch {
        Write-Warning "OData batch request failed: $_"
        # Return one $null per request so the caller can fall back / mark as error
        return @($null) * $RelativeRequests.Count
    }

    # PowerShell 7's Invoke-WebRequest returns Content as a byte[]; Windows PowerShell
    # returns a string. Decode UTF-8 explicitly when needed.
    $respText = if ($resp.Content -is [byte[]]) {
        [System.Text.Encoding]::UTF8.GetString($resp.Content)
    }
    else {
        [string]$resp.Content
    }

    # Discover the response boundary from Content-Type
    $respCT = $resp.Headers['Content-Type']
    if ($respCT -is [array]) { $respCT = $respCT[0] }
    $respBoundary = $null
    if ($respCT -match 'boundary=([^;]+)') {
        $respBoundary = $matches[1].Trim('"')
    }
    if (-not $respBoundary) {
        Write-Warning "Could not parse response boundary from batch response."
        return @($null) * $RelativeRequests.Count
    }

    # Split into parts. The leading boundary marker plus trailing terminator are non-content.
    $parts = $respText -split [regex]::Escape("--$respBoundary")

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '' -or $trimmed -eq '--') { continue }

        # Each part body looks like:
        #   Content-Type: application/http
        #   Content-Transfer-Encoding: binary
        #
        #   HTTP/1.1 200 OK
        #   Header: value
        #   ...
        #
        #   {"@odata.context":"...","@odata.count":150,"value":[...]}
        #
        # Extract the JSON between the first '{' and the last '}' of the part.
        $startIdx = $part.IndexOf('{')
        $endIdx   = $part.LastIndexOf('}')
        if ($startIdx -lt 0 -or $endIdx -lt $startIdx) {
            $results.Add($null)
            continue
        }
        $json = $part.Substring($startIdx, $endIdx - $startIdx + 1)
        try {
            $obj = $json | ConvertFrom-Json -ErrorAction Stop
            $results.Add($obj)
        }
        catch {
            $results.Add($null)
        }
    }

    # Sanity check: the parsed result count should match the request count
    if ($results.Count -ne $RelativeRequests.Count) {
        Write-Warning "Batch response part count ($($results.Count)) does not match request count ($($RelativeRequests.Count))."
    }

    return $results.ToArray()
}

# Main script execution
try {
    $allResults = New-Object System.Collections.Generic.List[object]

    if ($Filter) {
        Write-Host "Applying -Filter to all counts: $Filter" -ForegroundColor Cyan
    }

    foreach ($logicalName in $Tables) {
        Write-Host "`n=== Processing table: $logicalName ===" -ForegroundColor Cyan

        $meta = Get-EntityMetadata -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        if (-not $meta) {
            Write-Warning "Skipping '$logicalName' (metadata lookup failed)."
            continue
        }
        Write-Host "  EntitySetName: $($meta.EntitySetName)" -ForegroundColor Gray

        $totalCount = Get-TotalRecordCount -OrgUrl $OrganizationUrl -Headers $headers `
            -EntitySetName $meta.EntitySetName -Filter $Filter
        if ($null -eq $totalCount) {
            Write-Warning "Could not determine total record count for '$logicalName'; skipping."
            continue
        }
        $totalLabel = if ($Filter) { "matching filter" } else { "total" }
        Write-Host "  Records ${totalLabel}: $totalCount" -ForegroundColor Gray

        if ($totalCount -eq 0) {
            Write-Host "  No matching records; emitting one row per attribute with PopulatedCount = 0." -ForegroundColor Yellow
        }

        # Load attributes
        $attrs = Get-TableAttributes -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        Write-Host "  Total attributes returned by metadata: $($attrs.Count)" -ForegroundColor Gray

        # Build filter set if user specified -Attributes
        $attrFilterSet = $null
        if ($Attributes -and $Attributes.Count -gt 0) {
            $attrFilterSet = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]@($Attributes | ForEach-Object { $_.ToLowerInvariant() }),
                [System.StringComparer]::OrdinalIgnoreCase
            )
        }

        # Decide which attributes to query
        $eligible = @()
        $skippedReasons = @{}

        foreach ($attr in $attrs) {
            $name = $attr.LogicalName

            if ($attrFilterSet -and -not $attrFilterSet.Contains($name)) {
                continue
            }

            if (-not $attr.IsValidForRead) {
                $skippedReasons[$name] = "NotReadable"
                continue
            }
            if ($attr.IsLogical) {
                $skippedReasons[$name] = "Logical"
                continue
            }
            # Skip lookup-projection sub-attributes (e.g. _name, _yominame versions of a lookup).
            # MscrmTools.AttributeUsageInspector applies the same rule.
            if ($attr.AttributeOf) {
                $skippedReasons[$name] = "SubAttributeOf($($attr.AttributeOf))"
                continue
            }
            # Skip composite columns (e.g. fullname, address1_composite). They are computed
            # concatenations and their fill rate just mirrors the underlying parts.
            if ($name -like '*composite*') {
                $skippedReasons[$name] = "Composite"
                continue
            }
            if ($CustomAttributesOnly -and -not $attr.IsCustomAttribute) {
                $skippedReasons[$name] = "NotCustom"
                continue
            }
            if ($StandardAttributesOnly -and $attr.IsCustomAttribute) {
                $skippedReasons[$name] = "NotStandard"
                continue
            }
            if ($attr.AttributeType -in $UnsupportedAttributeTypes) {
                $skippedReasons[$name] = "UnsupportedType ($($attr.AttributeType))"
                continue
            }
            if (-not $IncludeSystemAttributes) {
                if ($name -eq $meta.PrimaryIdAttribute) {
                    $skippedReasons[$name] = "PrimaryKey"
                    continue
                }
                if ($DefaultSkipAttributes -contains $name) {
                    $skippedReasons[$name] = "SystemColumn"
                    continue
                }
            }

            $eligible += $attr
        }

        Write-Host "  Eligible attributes to query: $($eligible.Count)" -ForegroundColor Green
        if ($skippedReasons.Count -gt 0 -and -not $attrFilterSet) {
            $reasonGroups = $skippedReasons.Values |
                ForEach-Object { ($_ -split ' ')[0] } |  # collapse "UnsupportedType (X)" -> "UnsupportedType"
                Group-Object | Sort-Object Count -Descending
            $summary = ($reasonGroups | ForEach-Object { "$($_.Count) $($_.Name)" }) -join ", "
            Write-Host "  Skipped attributes by reason: $summary" -ForegroundColor Gray
        }

        # Process eligible attributes in batches via OData $batch
        if ($totalCount -eq 0 -or $eligible.Count -eq 0) {
            # Skip API calls entirely; emit zero-count rows
            foreach ($attr in $eligible) {
                $allResults.Add([PSCustomObject][ordered]@{
                    TableLogicalName     = $logicalName
                    TableDisplayName     = $meta.DisplayName
                    TableSchemaName      = $meta.SchemaName
                    AttributeLogicalName = $attr.LogicalName
                    AttributeSchemaName  = $attr.SchemaName
                    AttributeDisplayName = $attr.DisplayName
                    AttributeType        = $attr.AttributeType
                    IsCustomAttribute    = $attr.IsCustomAttribute
                    TotalRecords         = $totalCount
                    PopulatedCount       = 0
                    EmptyCount           = 0
                    FillRatePercent      = 0.0
                    Status               = "Success"
                })
            }
        }
        else {
            $totalEligible = $eligible.Count
            $batchCount    = [math]::Ceiling($totalEligible / $BatchRequestSize)
            $processed     = 0

            for ($b = 0; $b -lt $batchCount; $b++) {
                $start = $b * $BatchRequestSize
                $end   = [math]::Min($start + $BatchRequestSize - 1, $totalEligible - 1)
                $chunk = $eligible[$start..$end]

                # Build relative request URLs for this batch
                $relRequests = foreach ($attr in $chunk) {
                    $name       = $attr.LogicalName
                    $filterExpr = if ($attr.AttributeType -eq 'Lookup') { "_${name}_value" } else { $name }
                    $combined   = Build-PopulatedFilter -ColumnFilterExpr $filterExpr -BaseFilter $Filter
                    $encoded    = [System.Uri]::EscapeDataString($combined)
                    "$($meta.EntitySetName)?`$filter=$encoded&`$count=true&`$top=1&`$select=$filterExpr"
                }

                Write-Progress -Activity "Field fill rate: $logicalName" `
                    -Status "Batch $($b + 1) of $batchCount - $($chunk.Count) attribute(s)" `
                    -PercentComplete (($b + 1) / $batchCount * 100)

                if ($BatchRequestSize -eq 1) {
                    # Caller asked us NOT to batch: issue one normal request
                    $attr       = $chunk[0]
                    $name       = $attr.LogicalName
                    $filterExpr = if ($attr.AttributeType -eq 'Lookup') { "_${name}_value" } else { $name }
                    $populated  = Get-FieldPopulatedCount -OrgUrl $OrganizationUrl -Headers $headers `
                        -EntitySetName $meta.EntitySetName -ColumnFilterExpr $filterExpr -BaseFilter $Filter
                    $batchResults = @( if ($null -ne $populated) { @{ '@odata.count' = $populated } } else { $null } )
                }
                else {
                    $batchResults = Invoke-ODataBatch -OrgUrl $OrganizationUrl -Headers $headers -RelativeRequests $relRequests
                }

                for ($k = 0; $k -lt $chunk.Count; $k++) {
                    $attr      = $chunk[$k]
                    $resp      = $batchResults[$k]
                    $populated = $null
                    $errored   = $true

                    if ($null -ne $resp) {
                        # Both real OData responses and our wrapped fallback hashtable have @odata.count
                        $countVal = $null
                        if ($resp.PSObject.Properties['@odata.count']) {
                            $countVal = $resp.'@odata.count'
                        } elseif ($resp -is [hashtable] -and $resp.ContainsKey('@odata.count')) {
                            $countVal = $resp['@odata.count']
                        }
                        if ($null -ne $countVal) {
                            $populated = [long]$countVal
                            $errored   = $false
                        }
                    }

                    $emptyCount = if ($errored) { $null } else { $totalCount - $populated }
                    $fillRate   = if ($errored) { $null } else { [math]::Round(($populated / $totalCount) * 100, 2) }

                    $allResults.Add([PSCustomObject][ordered]@{
                        TableLogicalName     = $logicalName
                        TableDisplayName     = $meta.DisplayName
                        TableSchemaName      = $meta.SchemaName
                        AttributeLogicalName = $attr.LogicalName
                        AttributeSchemaName  = $attr.SchemaName
                        AttributeDisplayName = $attr.DisplayName
                        AttributeType        = $attr.AttributeType
                        IsCustomAttribute    = $attr.IsCustomAttribute
                        TotalRecords         = $totalCount
                        PopulatedCount       = if ($errored) { "N/A" } else { $populated }
                        EmptyCount           = if ($null -eq $emptyCount) { "N/A" } else { $emptyCount }
                        FillRatePercent      = if ($null -eq $fillRate)   { "N/A" } else { $fillRate }
                        Status               = if ($errored) { "Error" } else { "Success" }
                    })
                    $processed++
                }
            }

            Write-Progress -Activity "Field fill rate: $logicalName" -Completed
        }
    }

    # Sort: by table, then by fill rate descending (errors last)
    $sorted = $allResults | Sort-Object `
        TableLogicalName,
        @{ Expression = { if ($_.FillRatePercent -eq "N/A") { -1 } else { [double]$_.FillRatePercent } }; Descending = $true },
        AttributeLogicalName

    # Summary
    $totalRows  = $sorted.Count
    $errorCount = ($sorted | Where-Object Status -eq 'Error').Count
    $emptyAttrs = ($sorted | Where-Object { $_.Status -eq 'Success' -and [double]$_.FillRatePercent -eq 0 }).Count
    $fullAttrs  = ($sorted | Where-Object { $_.Status -eq 'Success' -and [double]$_.FillRatePercent -eq 100 }).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total attribute rows: $totalRows"
    Write-Host "  Always populated (100%): $fullAttrs" -ForegroundColor Green
    Write-Host "  Never populated (0%):    $emptyAttrs" -ForegroundColor Yellow
    Write-Host "  Errors:                  $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
    Write-Host ""

    # Output
    switch ($OutputFormat) {
        "Table" {
            if ($OutputPath) {
                $sorted | Format-Table -AutoSize | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $sorted | Format-Table -AutoSize
            }
        }
        "CSV" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "attributeusage_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "attributeusage_$timestamp.json"
            }
            ($sorted | ConvertTo-Json -Depth 4) | Out-File -FilePath $OutputPath
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
    }

    return $sorted
}
catch {
    Write-Error "Script execution failed: $_"
    throw
}
