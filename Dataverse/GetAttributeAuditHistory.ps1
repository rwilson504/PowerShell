<#
.SYNOPSIS
    Reports per-attribute audit history (last modified date, change count) for one or more
    Dataverse tables, using the built-in audit log.

.DESCRIPTION
    For each requested table, the script:
      1. Loads attribute metadata (LogicalName, ColumnNumber, IsAuditEnabled, etc.).
      2. Pages through the audit table filtered to that entity's Create + Update operations
         within the configured -DaysBack window, sorted createdon desc.
      3. For every audit row, splits the comma-separated attributemask into ColumnNumbers and
         resolves each back to the attribute's LogicalName.
      4. Aggregates: per attribute, captures the most recent createdon (= LastAuditedOn), the
         operation/user from that row, and the total count of audit entries in the window.

    The output is designed to JOIN with attributeusage_*.csv (from GetFieldFillRateByTable.ps1)
    and relationships_*.csv (from GetTableRelationships.ps1) on the same composite key
    (TableLogicalName + AttributeLogicalName), so you can build a "fields no longer in use"
    report in Excel / Power BI / pandas.

    USE CASE: detecting unused lookup fields
      Combine with attributeusage_*.csv to find columns that have data but haven't been TOUCHED
      recently. A lookup field with FillRatePercent > 0 but DaysSinceLastAudited > 365 means
      records still hold the link, but no one is creating / updating it any more - a strong
      signal that the relationship has fallen out of business use.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Required. One or more table logical names to analyze.

.PARAMETER Attributes
    Optional. Restrict analysis to these attribute logical names. When omitted, every
    attribute on the table is reported.

.PARAMETER DaysBack
    Number of days of audit history to scan. Pass an explicit value (1 to 3650) to override,
    or leave at the default 0 to use AUTO mode:

      AUTO mode picks the largest sensible window for the environment:
        - If organization.auditretentionperiodv2 is set, use that exact value.
        - If retention is unset / 0 (commonly 'Never expire' / 'Forever'), use the number of
          days since the org was created (organization.createdon). Capped at 3650.
        - If both reads fail, fall back to 365.

    Going as far back as the data permits is essentially free (one extra audit page or two
    on tables with low write activity), and gives the cleanest 'this field is genuinely
    unused' signal.

    NOTE: If you DO pass an explicit -DaysBack and it exceeds the org's configured retention,
    the script warns up front. You will only see whatever audit data is still present.

.PARAMETER LookupAttributesOnly
    Only emit rows for attributes whose AttributeType is Lookup. Useful for the "unused
    lookup" detection workflow.

.PARAMETER UnusedOnly
    Only emit rows where no audit entries were found in the window (LastAuditedOn = null).
    Combine with -LookupAttributesOnly for a fast "lookups never touched in the last N days"
    report.

.PARAMETER IncludeAuditDisabledColumns
    By default, columns where IsAuditEnabled = false are omitted (they will never have audit
    entries by design and just add noise). Set this switch to include them so you can audit
    your audit configuration too.

.PARAMETER MaxAuditPageSize
    Maximum audit rows pulled per page. Default 5000 (the Web API maximum). Lower values
    increase round-trip count but reduce per-call latency on slow links.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetAttributeAuditHistory.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account"

    Reports per-attribute last-modified data for every audited attribute on account in the
    last 365 days.

.EXAMPLE
    .\GetAttributeAuditHistory.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account","contact" -LookupAttributesOnly -UnusedOnly -DaysBack 365 -OutputFormat CSV

    Lists lookup fields on account and contact that were never modified in the last year -
    candidates for cleanup or relationship review.

.EXAMPLE
    .\GetAttributeAuditHistory.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_company" -DaysBack 90 -OutputFormat CSV -OutputPath ".\audit.csv"

    Reports the last 90 days of per-attribute audit activity on msf_company; writes to CSV.

.NOTES
    PREREQUISITES
      Auditing must be ENABLED at three levels for column changes to be captured:
        1. Org level   (organization.IsAuditEnabled)
        2. Table level (entity.IsAuditEnabled)
        3. Column level (attribute.IsAuditEnabled)
      The script reports the org-level state up front. Per-column state is included in the
      output (IsAuditEnabledForColumn). If a column is not audited you'll see Status =
      'AuditDisabledForColumn' (or AuditDisabledForOrg / AuditDisabledForTable when those
      higher levels are off). It does NOT mean the field is unused - it means we can't tell.

    CORRELATING WITH attributeusage_*.csv (the typical workflow)
      Both CSVs share TableLogicalName + AttributeLogicalName as a composite join key, plus
      LookupTargets and the same metadata columns. Bring both into Excel as separate sheets
      and use VLOOKUP to add audit-history columns next to fill-rate columns:

        =VLOOKUP(A2&"|"&D2,
                 'audithistory'!$A:$N,    composite key joins on TableLogicalName + AttributeLogicalName
                 11, FALSE)               whatever column you want to pull (LastAuditedOn, etc.)

      Power Query / Power BI / pandas: just merge the two on the composite key.

    "FIELD NO LONGER IN USE" CHECKLIST (combining all three CSVs)
      A lookup field is a strong cleanup candidate if ALL of the following are true:
        - attributeusage : FillRatePercent > 0  (data exists - so it WAS used)
        - audit history  : DaysSinceLastAudited > 365 OR LastAuditedOn is null
        - relationships  : IsCustomRelationship = true AND IsManaged = false
                           (you can actually delete it without breaking a Microsoft solution)
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
    [ValidateRange(0, 3650)]
    [int]$DaysBack = 0,

    [Parameter(Mandatory = $false)]
    [switch]$LookupAttributesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$UnusedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeAuditDisabledColumns,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 5000)]
    [int]$MaxAuditPageSize = 5000,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Content-Type"     = "application/json; charset=utf-8"
    "Prefer"           = "odata.include-annotations=*,odata.maxpagesize=$MaxAuditPageSize"
}

# Operation enum mapping (from Microsoft.Crm.Sdk.Messages.AuditOperation)
$OperationLabels = @{
    1 = 'Create'
    2 = 'Update'
    3 = 'Delete'
    4 = 'Access'
    5 = 'Upsert'
}

function Get-OrgAuditEnabled {
    param ([string]$OrgUrl, [hashtable]$Headers)
    try {
        $r = Invoke-RestMethod -Uri "$OrgUrl/api/data/v9.2/organizations?`$select=isauditenabled,auditretentionperiodv2,createdon" -Headers $Headers
        return [PSCustomObject]@{
            IsAuditEnabled  = [bool]$r.value[0].isauditenabled
            RetentionDays   = $r.value[0].auditretentionperiodv2
            CreatedOn       = $r.value[0].createdon
        }
    }
    catch {
        Write-Warning "Could not read org-level audit setting: $_"
        return $null
    }
}

function Get-EntityMetadata {
    <#
    .SYNOPSIS
        Returns metadata for a single table including IsAuditEnabled at the table level.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)

    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')?" +
        "`$select=LogicalName,SchemaName,EntitySetName,DisplayName,IsAuditEnabled,ObjectTypeCode"

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $displayName = if ($resp.DisplayName.UserLocalizedLabel) {
            $resp.DisplayName.UserLocalizedLabel.Label
        } else {
            $LogicalName
        }
        return [PSCustomObject]@{
            LogicalName       = $resp.LogicalName
            SchemaName        = $resp.SchemaName
            EntitySetName     = $resp.EntitySetName
            DisplayName       = $displayName
            ObjectTypeCode    = $resp.ObjectTypeCode
            IsAuditEnabled    = [bool]$resp.IsAuditEnabled.Value
        }
    }
    catch {
        Write-Warning "Failed to load metadata for table '$LogicalName': $_"
        return $null
    }
}

function Get-TableAttributesForAudit {
    <#
    .SYNOPSIS
        Returns attribute metadata enriched with ColumnNumber, IsAuditEnabled, and
        LookupTargets (for Lookup attrs) so we can resolve audit attributemask integers
        back to attribute logical names and reason about audit configuration.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)

    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes?" +
        "`$select=LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute,IsAuditEnabled,IsValidForRead,IsLogical,AttributeOf,ColumnNumber"

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
                IsAuditEnabled    = [bool]$attr.IsAuditEnabled.Value
                IsValidForRead    = [bool]$attr.IsValidForRead
                IsLogical         = [bool]$attr.IsLogical
                AttributeOf       = $attr.AttributeOf
                ColumnNumber      = $attr.ColumnNumber
                LookupTargets     = $null
            }
        }
        $url = $resp.'@odata.nextLink'
    } while ($url)

    # Enrich Lookup rows with target tables (matches the column added to attributeusage)
    $lookupAttrs = $all | Where-Object { $_.AttributeType -eq 'Lookup' }
    if ($lookupAttrs.Count -gt 0) {
        $lookupUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?" +
            "`$select=LogicalName,Targets"
        try {
            $targetMap = @{}
            do {
                $tResp = Invoke-RestMethod -Uri $lookupUrl -Headers $Headers -Method Get
                foreach ($t in $tResp.value) {
                    $targetMap[$t.LogicalName] = ($t.Targets -join ';')
                }
                $lookupUrl = $tResp.'@odata.nextLink'
            } while ($lookupUrl)

            foreach ($attr in $lookupAttrs) {
                if ($targetMap.ContainsKey($attr.LogicalName)) {
                    $attr.LookupTargets = $targetMap[$attr.LogicalName]
                }
            }
        }
        catch {
            Write-Warning "Failed to load Lookup targets for '$LogicalName': $_"
        }
    }

    return $all
}

function Get-AuditDataForTable {
    <#
    .SYNOPSIS
        Pages through audit rows for a table within the supplied window. Returns an array of
        @{ createdon; operation; attributemask; userid } objects sorted createdon desc.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$LogicalName,
        [datetime]$SinceUtc
    )

    $sinceStr = $SinceUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
    # Operation 1=Create, 2=Update. Excluding 3 (Delete) - delete audit rows have empty
    # attributemask anyway and we don't care about field-level activity for deleted records.
    $filter = "objecttypecode eq '$LogicalName' and createdon ge $sinceStr and (operation eq 1 or operation eq 2)"
    $select = "createdon,operation,attributemask,_userid_value,objectid"
    $url = "$OrgUrl/api/data/v9.2/audits?" +
        "`$filter=$([System.Uri]::EscapeDataString($filter))" +
        "&`$select=$select" +
        "&`$orderby=createdon desc"

    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    do {
        try {
            $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        }
        catch {
            Write-Warning "Audit query failed for '$LogicalName' on page $($page + 1): $_"
            break
        }
        foreach ($row in $resp.value) {
            $all.Add([PSCustomObject]@{
                createdon     = $row.createdon
                operation     = $row.operation
                attributemask = $row.attributemask
                userid        = $row.'_userid_value'
                objectid      = $row.objectid
            })
        }
        $page++
        Write-Progress -Activity "Audit history: $LogicalName" `
            -Status "Page $page - $($all.Count) audit rows fetched" `
            -PercentComplete -1
        $url = $resp.'@odata.nextLink'
    } while ($url)

    Write-Progress -Activity "Audit history: $LogicalName" -Completed
    return $all.ToArray()
}

# Main script execution
try {
    Write-Host "Checking org-level audit setting..." -ForegroundColor Cyan
    $orgAudit = Get-OrgAuditEnabled -OrgUrl $OrganizationUrl -Headers $headers
    $orgAuditOn = $null
    if ($null -eq $orgAudit) {
        Write-Warning "Could not determine org-level audit setting; results may be empty."
    }
    else {
        $orgAuditOn = $orgAudit.IsAuditEnabled
        if (-not $orgAuditOn) {
            Write-Host "Org-level auditing is DISABLED. No audit data will be available." -ForegroundColor Yellow
            Write-Host "  Enable in Power Platform admin center -> Environment -> Settings -> Audit and logs -> Audit settings." -ForegroundColor Yellow
        }
        else {
            $retentionLabel = if ($null -eq $orgAudit.RetentionDays -or $orgAudit.RetentionDays -le 0) {
                "(unset / never expire)"
            } else {
                "$($orgAudit.RetentionDays) day(s)"
            }
            $envAgeLabel = if ($orgAudit.CreatedOn) {
                $age = [int]([math]::Floor(((Get-Date).ToUniversalTime() - ([datetime]$orgAudit.CreatedOn).ToUniversalTime()).TotalDays))
                "$age day(s) old (created $([datetime]$orgAudit.CreatedOn | Get-Date -Format 'yyyy-MM-dd'))"
            } else {
                "(unknown age)"
            }
            Write-Host "Org-level auditing is enabled. Retention: $retentionLabel. Environment: $envAgeLabel." -ForegroundColor Green
        }
    }

    # Auto-resolve -DaysBack when 0: prefer retention, fall back to env age, fall back to 365.
    if ($DaysBack -eq 0) {
        $resolvedDays = 365
        $resolvedSrc  = "fallback default"
        if ($orgAudit) {
            if ($orgAudit.RetentionDays -and $orgAudit.RetentionDays -gt 0) {
                $resolvedDays = [math]::Min([int]$orgAudit.RetentionDays, 3650)
                $resolvedSrc  = "configured retention"
            }
            elseif ($orgAudit.CreatedOn) {
                $envAge = [int]([math]::Ceiling(((Get-Date).ToUniversalTime() - ([datetime]$orgAudit.CreatedOn).ToUniversalTime()).TotalDays))
                $resolvedDays = [math]::Max(1, [math]::Min($envAge + 1, 3650))   # +1 to include creation-day audits
                $resolvedSrc  = "environment age (retention is unset / 'never expire')"
            }
        }
        $DaysBack = $resolvedDays
        Write-Host "AUTO -DaysBack: scanning $DaysBack day(s) of history (source: $resolvedSrc)." -ForegroundColor Cyan
    }
    elseif ($orgAudit -and $orgAudit.IsAuditEnabled -and $orgAudit.RetentionDays -and $orgAudit.RetentionDays -gt 0 -and $DaysBack -gt $orgAudit.RetentionDays) {
        Write-Host "  WARNING: -DaysBack=$DaysBack exceeds the configured retention of $($orgAudit.RetentionDays) day(s). You will only see ~$($orgAudit.RetentionDays) day(s) of data." -ForegroundColor Yellow
    }

    $sinceUtc = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
    Write-Host "Audit window: last $DaysBack day(s) (since $($sinceUtc.ToString('yyyy-MM-dd HH:mm:ss UTC')))" -ForegroundColor Cyan

    $allResults = New-Object System.Collections.Generic.List[object]

    foreach ($logicalName in $Tables) {
        Write-Host "`n=== Processing table: $logicalName ===" -ForegroundColor Cyan

        $meta = Get-EntityMetadata -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        if (-not $meta) {
            Write-Warning "Skipping '$logicalName' (metadata lookup failed)."
            continue
        }
        Write-Host "  Table-level auditing: $(if ($meta.IsAuditEnabled) {'enabled'} else {'DISABLED'})" -ForegroundColor $(if ($meta.IsAuditEnabled) {'Green'} else {'Yellow'})

        $attrs = Get-TableAttributesForAudit -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        Write-Host "  Total attributes: $($attrs.Count)" -ForegroundColor Gray

        # Build attribute filter set if user specified -Attributes
        $attrFilterSet = $null
        if ($Attributes -and $Attributes.Count -gt 0) {
            $attrFilterSet = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]@($Attributes | ForEach-Object { $_.ToLowerInvariant() }),
                [System.StringComparer]::OrdinalIgnoreCase
            )
        }

        # Decide which attributes to emit rows for
        $eligible = @()
        foreach ($attr in $attrs) {
            if ($attrFilterSet -and -not $attrFilterSet.Contains($attr.LogicalName)) { continue }
            if ($attr.IsLogical) { continue }
            if ($attr.AttributeOf) { continue }  # skip lookup-projection sub-attributes
            if ($LookupAttributesOnly -and $attr.AttributeType -ne 'Lookup') { continue }
            if (-not $IncludeAuditDisabledColumns -and -not $attr.IsAuditEnabled) { continue }
            $eligible += $attr
        }
        Write-Host "  Eligible attributes to report: $($eligible.Count)" -ForegroundColor Green

        # Build ColumnNumber -> Attribute map for resolving attributemask
        $columnMap = @{}
        foreach ($attr in $attrs) {
            if ($null -ne $attr.ColumnNumber) {
                $columnMap[[int]$attr.ColumnNumber] = $attr
            }
        }

        # Decide whether to bother pulling audit data
        $skipAuditQuery = (-not $orgAuditOn) -or (-not $meta.IsAuditEnabled) -or ($eligible.Count -eq 0)
        $auditAggregate = @{}  # ColumnNumber -> @{ LastCreatedOn; LastOperation; LastUserId; Count }

        if (-not $skipAuditQuery) {
            Write-Host "  Pulling audit rows since $($sinceUtc.ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan
            $auditRows = Get-AuditDataForTable -OrgUrl $OrganizationUrl -Headers $headers `
                -LogicalName $logicalName -SinceUtc $sinceUtc
            Write-Host "  Audit rows fetched: $($auditRows.Count)" -ForegroundColor Gray

            # Aggregate. Audit rows are sorted createdon desc, so the FIRST time we see a
            # ColumnNumber is its most recent change.
            foreach ($row in $auditRows) {
                if ([string]::IsNullOrWhiteSpace($row.attributemask)) { continue }
                $colNums = $row.attributemask -split ',' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
                foreach ($cn in $colNums) {
                    if (-not $auditAggregate.ContainsKey($cn)) {
                        $auditAggregate[$cn] = [PSCustomObject]@{
                            LastCreatedOn  = $row.createdon
                            LastOperation  = $row.operation
                            LastUserId     = $row.userid
                            LastObjectId   = $row.objectid
                            Count          = 0
                        }
                    }
                    $auditAggregate[$cn].Count++
                }
            }
        }

        # Emit one row per eligible attribute
        $now = (Get-Date).ToUniversalTime()
        foreach ($attr in $eligible) {
            $cn = if ($null -ne $attr.ColumnNumber) { [int]$attr.ColumnNumber } else { -1 }
            $hit = $null
            if ($cn -ge 0 -and $auditAggregate.ContainsKey($cn)) {
                $hit = $auditAggregate[$cn]
            }

            $status =
                if (-not $orgAuditOn)         { 'AuditDisabledForOrg' }
                elseif (-not $meta.IsAuditEnabled) { 'AuditDisabledForTable' }
                elseif (-not $attr.IsAuditEnabled) { 'AuditDisabledForColumn' }
                elseif ($hit)                  { 'Success' }
                else                           { 'NoAuditData' }

            $lastAuditedOn   = if ($hit) { $hit.LastCreatedOn } else { $null }
            $daysSince       = if ($lastAuditedOn) {
                [math]::Max(0, [int]([math]::Floor(($now - ([datetime]$lastAuditedOn).ToUniversalTime()).TotalDays)))
            } else { $null }
            $lastAction      = if ($hit) { $OperationLabels[[int]$hit.LastOperation] } else { $null }
            $lastUserId      = if ($hit) { $hit.LastUserId } else { $null }
            $entriesInWindow = if ($hit) { $hit.Count } else { 0 }

            # -UnusedOnly: omit rows where audit data WAS found (we only want never-touched)
            if ($UnusedOnly -and $status -eq 'Success') { continue }

            $allResults.Add([PSCustomObject][ordered]@{
                TableLogicalName        = $logicalName
                TableDisplayName        = $meta.DisplayName
                TableSchemaName         = $meta.SchemaName
                AttributeLogicalName    = $attr.LogicalName
                AttributeSchemaName     = $attr.SchemaName
                AttributeDisplayName    = $attr.DisplayName
                AttributeType           = $attr.AttributeType
                IsCustomAttribute       = $attr.IsCustomAttribute
                LookupTargets           = $attr.LookupTargets
                IsAuditEnabledForColumn = $attr.IsAuditEnabled
                LastAuditedOn           = $lastAuditedOn
                LastAuditedAction       = $lastAction
                LastAuditedByUserId     = $lastUserId
                AuditEntriesInWindow    = $entriesInWindow
                DaysSinceLastAudited    = $daysSince
                WindowDays              = $DaysBack
                Status                  = $status
            })
        }
    }

    # Sort: by table, then by DaysSinceLastAudited desc (oldest / never-modified first)
    $sorted = $allResults | Sort-Object `
        TableLogicalName,
        @{ Expression = { if ($null -eq $_.DaysSinceLastAudited) { [int]::MaxValue } else { [int]$_.DaysSinceLastAudited } }; Descending = $true },
        AttributeLogicalName

    # Summary
    $rowsSuccess  = ($sorted | Where-Object Status -eq 'Success').Count
    $rowsNoAudit  = ($sorted | Where-Object Status -eq 'NoAuditData').Count
    $rowsAuditOff = ($sorted | Where-Object { $_.Status -like 'AuditDisabled*' }).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total attribute rows: $($sorted.Count)"
    Write-Host "  With audit data:               $rowsSuccess" -ForegroundColor Green
    Write-Host "  No audit entries in $DaysBack-day window: $rowsNoAudit" -ForegroundColor Yellow
    Write-Host "  Auditing disabled (org/table/column):  $rowsAuditOff" -ForegroundColor $(if ($rowsAuditOff -gt 0) { 'Yellow' } else { 'Green' })
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
                $OutputPath = Join-Path (Get-Location) "audithistory_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "audithistory_$timestamp.json"
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
