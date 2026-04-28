<#
.SYNOPSIS
    Gets record counts for specified tables or all readable tables in Dataverse.

.DESCRIPTION
    This script retrieves the record count for each specified table in Dataverse
    using the RetrieveTotalRecordCount API function. This is significantly faster
    than FetchXML aggregate queries as it retrieves counts in batch from system
    indexes rather than counting rows individually.

    Note: The counts returned are approximate as they come from system indexes
    which may be slightly out of date, but are sufficient for most purposes.

    If no tables are specified, it retrieves metadata to find all readable tables
    and gets counts for each one.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    An optional array of table logical names to get counts for.
    If not provided, all readable tables will be queried.

.PARAMETER IncludeSystemTables
    When querying all tables, include system tables (those starting with 'sys').
    Default is $false.

.PARAMETER CustomEntitiesOnly
    When querying all tables (no -Tables parameter), restrict the result to custom entities
    (IsCustomEntity eq true). System/Microsoft tables are excluded. Useful for auditing
    custom-built capabilities without the noise of out-of-the-box tables. Has no effect when
    -Tables is specified - explicit table lists are always honored as-is.

.PARAMETER BatchSize
    The number of tables to include per RetrieveTotalRecordCount API call.
    Default is 20. Lower values reduce the chance of a single unsupported table
    (virtual/elastic/preview entity) poisoning the whole batch and forcing per-table retries;
    higher values reduce total request count when most tables are supported.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results. If not provided, results are written to the console.

.PARAMETER IncludeLastActivity
    When specified, retrieves the last CreatedOn, last ModifiedOn, and oldest CreatedOn timestamps
    for each table by querying the top 1 record sorted on each column. This adds three extra API
    calls per table (skipped for tables with 0 records, unless -ActivityFallback is also set), so
    it can significantly increase runtime in environments with many tables.

    When enabled, the output also includes computed columns:
      - DaysSinceLastCreated  : Whole days since the most recently created record
      - DaysSinceLastModified : Whole days since the most recently modified record
      - UsageBucket           : One of Empty / Active (<=90d) / Dormant (91-365d) / Stale (>365d) / Unknown

    Useful for identifying tables/capabilities that are no longer in active use.

.PARAMETER ActivityFallback
    Only meaningful with -IncludeLastActivity. When set, also runs the activity timestamp queries
    for tables whose RecordCount came back as 0 or N/A. This is useful because the
    RetrieveTotalRecordCount API reads from periodically-refreshed table statistics and can return
    stale 0 values on test/sandbox environments or for tables that don't participate in stats
    collection. If activity queries find records, the row's UsageBucket is updated accordingly.

.PARAMETER IncludeUnsupportedTypes
    By default, tables with TableType = Virtual or Elastic are pre-skipped because
    RetrieveTotalRecordCount does not support them and trying causes batch failures. They appear
    in the output with Status = 'Skipped (Virtual)' / 'Skipped (Elastic)'. Set this switch to
    attempt them anyway (most will still fail).

.NOTES
    The output always includes these metadata columns (no extra API cost beyond the existing
    metadata call): SchemaName, EntitySetName, IsCustomEntity, TableType.

    The RetrieveTotalRecordCount API does not support Virtual or Elastic tables. By default
    these are pre-skipped (Status = 'Skipped (Virtual)' / 'Skipped (Elastic)'). When -IncludeLastActivity
    is set, the activity probe is still attempted for elastic tables (which support OData queries);
    virtual tables vary by data provider.

    The RetrieveTotalRecordCount API also rejects an entire batch payload if any single entity in
    it is unsupported. When a batch fails, the script automatically retries each table in that
    batch individually so one bad apple does not poison the rest. Tables that still fail
    individually are reported with Status = 'Error'.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables @("account", "contact", "lead")

    Gets record counts for the account, contact, and lead tables.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token

    Gets record counts for all readable tables in the environment.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -OutputFormat CSV -OutputPath "C:\temp\recordcounts.csv"

    Gets record counts for all readable tables and exports to CSV.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -IncludeLastActivity -OutputFormat CSV -OutputPath "C:\temp\recordcounts.csv"

    Gets record counts plus the last CreatedOn/ModifiedOn timestamps for each table to help
    identify tables that are no longer in active use.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,
    
    [Parameter(Mandatory = $true)]
    [string]$AccessToken,
    
    [Parameter(Mandatory = $false)]
    [string[]]$Tables,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemTables = $false,

    [Parameter(Mandatory = $false)]
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 200)]
    [int]$BatchSize = 20,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLastActivity,

    [Parameter(Mandatory = $false)]
    [switch]$ActivityFallback,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeUnsupportedTypes
)

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

# Set up headers for API calls
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
    "Accept" = "application/json"
    "Content-Type" = "application/json; charset=utf-8"
    "Prefer" = "odata.include-annotations=*"
}

function Get-AllReadableTables {
    <#
    .SYNOPSIS
        Retrieves all tables that the current user has read access to.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [bool]$IncludeSystem,
        [bool]$CustomOnly
    )

    Write-Host "Retrieving metadata for all tables..." -ForegroundColor Cyan
    
    # Query EntityDefinitions to get all entities that are valid for read operations
    $filter = "IsValidForAdvancedFind eq true and IsIntersect eq false"
    if ($CustomOnly) {
        $filter += " and IsCustomEntity eq true"
    }
    $metadataUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions?" + 
        "`$select=LogicalName,SchemaName,EntitySetName,DisplayName,IsCustomEntity,IsValidForAdvancedFind,TableType" +
        "&`$filter=$filter"
    
    try {
        $response = Invoke-RestMethod -Uri $metadataUrl -Headers $Headers -Method Get
        $entities = $response.value
        
        $readableTables = @()
        
        foreach ($entity in $entities) {
            $logicalName = $entity.LogicalName
            
            # Skip system tables if not requested
            if (-not $IncludeSystem -and $logicalName.StartsWith("sys")) {
                continue
            }
            
            $displayName = if ($entity.DisplayName.UserLocalizedLabel) { 
                $entity.DisplayName.UserLocalizedLabel.Label 
            } else { 
                $logicalName 
            }
            
            $readableTables += [PSCustomObject]@{
                LogicalName     = $logicalName
                DisplayName     = $displayName
                SchemaName      = $entity.SchemaName
                EntitySetName   = $entity.EntitySetName
                IsCustomEntity  = [bool]$entity.IsCustomEntity
                TableType       = $entity.TableType
            }
        }
        
        Write-Host "Found $($readableTables.Count) readable tables." -ForegroundColor Green
        return $readableTables | Sort-Object LogicalName
    }
    catch {
        Write-Error "Failed to retrieve table metadata: $_"
        throw
    }
}

function Get-EntitySetNameMap {
    <#
    .SYNOPSIS
        Retrieves metadata (EntitySetName, SchemaName, IsCustomEntity, DisplayName) for a list
        of table logical names. Required for building OData query URLs and enriching output.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string[]]$LogicalNames
    )

    $map = @{}
    if (-not $LogicalNames -or $LogicalNames.Count -eq 0) {
        return $map
    }

    # Chunk the lookup so the OR-filter URL stays well under typical 16 KB URL length limits.
    # Each "LogicalName eq 'xxx' or " clause is ~25-50 chars; 25 per chunk leaves plenty of headroom.
    $chunkSize = 25
    $totalChunks = [math]::Ceiling($LogicalNames.Count / $chunkSize)

    for ($i = 0; $i -lt $totalChunks; $i++) {
        $startIdx = $i * $chunkSize
        $endIdx   = [math]::Min($startIdx + $chunkSize - 1, $LogicalNames.Count - 1)
        $chunk    = $LogicalNames[$startIdx..$endIdx]

        $filterClauses = ($chunk | ForEach-Object { "LogicalName eq '$_'" }) -join " or "
        $metadataUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions?" +
            "`$select=LogicalName,SchemaName,EntitySetName,DisplayName,IsCustomEntity,TableType" +
            "&`$filter=$filterClauses"

        try {
            $response = Invoke-RestMethod -Uri $metadataUrl -Headers $Headers -Method Get
            foreach ($entity in $response.value) {
                $map[$entity.LogicalName] = [PSCustomObject]@{
                    EntitySetName  = $entity.EntitySetName
                    SchemaName     = $entity.SchemaName
                    IsCustomEntity = [bool]$entity.IsCustomEntity
                    TableType      = $entity.TableType
                    DisplayName    = if ($entity.DisplayName.UserLocalizedLabel) {
                        $entity.DisplayName.UserLocalizedLabel.Label
                    } else {
                        $entity.LogicalName
                    }
                }
            }
        }
        catch {
            Write-Warning "Failed to retrieve entity metadata for chunk $($i + 1) of $totalChunks (tables $startIdx..$endIdx): $_"
        }
    }

    return $map
}

function Get-LastActivityForTable {
    <#
    .SYNOPSIS
        Gets the most recent CreatedOn / ModifiedOn and the oldest CreatedOn timestamps for a
        single table by querying the top 1 record sorted on indexed columns.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$EntitySetName
    )

    $result = [PSCustomObject]@{
        LastCreatedOn         = $null
        LastModifiedOn        = $null
        OldestRecordCreatedOn = $null
    }

    if ([string]::IsNullOrWhiteSpace($EntitySetName)) {
        return $result
    }

    try {
        $createdUrl = "$OrgUrl/api/data/v9.2/$EntitySetName" + "?`$top=1&`$select=createdon&`$orderby=createdon desc"
        $createdResp = Invoke-RestMethod -Uri $createdUrl -Headers $Headers -Method Get
        if ($createdResp.value -and $createdResp.value.Count -gt 0) {
            $result.LastCreatedOn = $createdResp.value[0].createdon
        }
    }
    catch {
        # Table may not have createdon (some virtual/system tables) - leave null
    }

    try {
        $modifiedUrl = "$OrgUrl/api/data/v9.2/$EntitySetName" + "?`$top=1&`$select=modifiedon&`$orderby=modifiedon desc"
        $modifiedResp = Invoke-RestMethod -Uri $modifiedUrl -Headers $Headers -Method Get
        if ($modifiedResp.value -and $modifiedResp.value.Count -gt 0) {
            $result.LastModifiedOn = $modifiedResp.value[0].modifiedon
        }
    }
    catch {
        # Table may not have modifiedon - leave null
    }

    try {
        $oldestUrl = "$OrgUrl/api/data/v9.2/$EntitySetName" + "?`$top=1&`$select=createdon&`$orderby=createdon asc"
        $oldestResp = Invoke-RestMethod -Uri $oldestUrl -Headers $Headers -Method Get
        if ($oldestResp.value -and $oldestResp.value.Count -gt 0) {
            $result.OldestRecordCreatedOn = $oldestResp.value[0].createdon
        }
    }
    catch {
        # Table may not have createdon - leave null
    }

    return $result
}

function Get-RecordCountsBatch {
    <#
    .SYNOPSIS
        Gets record counts for a batch of tables using the RetrieveTotalRecordCount API.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string[]]$EntityNames
    )

    # Build the JSON array of entity names and URL-encode it
    $entityNamesJson = ConvertTo-Json -InputObject $EntityNames -Compress
    $encodedNames = [System.Uri]::EscapeDataString($entityNamesJson)
    
    $apiUrl = "$OrgUrl/api/data/v9.2/RetrieveTotalRecordCount(EntityNames=@EntityNames)?@EntityNames=$encodedNames"
    
    $response = Invoke-RestMethod -Uri $apiUrl -Headers $Headers -Method Get
    
    # Build a hashtable from Keys and Values arrays
    $countMap = @{}
    $keys = $response.EntityRecordCountCollection.Keys
    $values = $response.EntityRecordCountCollection.Values
    
    for ($i = 0; $i -lt $keys.Count; $i++) {
        $countMap[$keys[$i]] = [long]$values[$i]
    }
    
    return $countMap
}

function Get-RecordCounts {
    <#
    .SYNOPSIS
        Main function to get record counts for all specified tables using batched API calls.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [array]$TableList,
        [int]$BatchSize = 50,
        [bool]$IncludeLastActivity = $false,
        [bool]$ActivityFallback = $false,
        [bool]$SkipUnsupportedTypes = $true
    )

    # Build lookups for display names, schema names, entity set names, custom flag, and table type
    $displayNameMap = @{}
    $schemaNameMap = @{}
    $entitySetNameMap = @{}
    $isCustomEntityMap = @{}
    $tableTypeMap = @{}
    $allLogicalNames = @()

    foreach ($table in $TableList) {
        $logicalName = if ($table -is [PSCustomObject]) { $table.LogicalName } else { $table }
        $displayName = if ($table -is [PSCustomObject] -and $table.DisplayName) { $table.DisplayName } else { $logicalName }
        $entitySetName = if ($table -is [PSCustomObject] -and $table.PSObject.Properties['EntitySetName']) { $table.EntitySetName } else { $null }
        $schemaName = if ($table -is [PSCustomObject] -and $table.PSObject.Properties['SchemaName']) { $table.SchemaName } else { $null }
        $isCustom = if ($table -is [PSCustomObject] -and $table.PSObject.Properties['IsCustomEntity']) { [bool]$table.IsCustomEntity } else { $null }
        $tableType = if ($table -is [PSCustomObject] -and $table.PSObject.Properties['TableType']) { $table.TableType } else { $null }

        $displayNameMap[$logicalName] = $displayName
        $entitySetNameMap[$logicalName] = $entitySetName
        $schemaNameMap[$logicalName] = $schemaName
        $isCustomEntityMap[$logicalName] = $isCustom
        $tableTypeMap[$logicalName] = $tableType
        $allLogicalNames += $logicalName
    }

    # Always look up missing metadata (SchemaName, EntitySetName, IsCustomEntity, TableType) so output
    # columns are populated even when the user passed bare logical-name strings via -Tables.
    $missingMetadata = $allLogicalNames | Where-Object {
        -not $entitySetNameMap[$_] -or -not $schemaNameMap[$_] -or $null -eq $isCustomEntityMap[$_] -or -not $tableTypeMap[$_]
    }
    if ($missingMetadata.Count -gt 0) {
        Write-Host "Looking up metadata for $($missingMetadata.Count) table(s)..." -ForegroundColor Cyan
        $lookupMap = Get-EntitySetNameMap -OrgUrl $OrgUrl -Headers $Headers -LogicalNames $missingMetadata
        foreach ($logicalName in $lookupMap.Keys) {
            if (-not $entitySetNameMap[$logicalName]) { $entitySetNameMap[$logicalName] = $lookupMap[$logicalName].EntitySetName }
            if (-not $schemaNameMap[$logicalName])    { $schemaNameMap[$logicalName]    = $lookupMap[$logicalName].SchemaName }
            if ($null -eq $isCustomEntityMap[$logicalName]) { $isCustomEntityMap[$logicalName] = $lookupMap[$logicalName].IsCustomEntity }
            if (-not $tableTypeMap[$logicalName]) { $tableTypeMap[$logicalName] = $lookupMap[$logicalName].TableType }
            # Update display name if we didn't have one (i.e. user passed a bare string)
            if (-not $displayNameMap[$logicalName] -or $displayNameMap[$logicalName] -eq $logicalName) {
                $displayNameMap[$logicalName] = $lookupMap[$logicalName].DisplayName
            }
        }
    }

    # Pre-skip Virtual/Elastic tables (RetrieveTotalRecordCount does not support them)
    $skippedTables = @{}  # logicalName -> reason ("Virtual" / "Elastic")
    $namesToBatch = $allLogicalNames
    if ($SkipUnsupportedTypes) {
        $virtualCount = 0
        $elasticCount = 0
        $namesToBatch = $allLogicalNames | Where-Object {
            $tt = $tableTypeMap[$_]
            if ($tt -eq 'Virtual')  { $skippedTables[$_] = 'Virtual'; $virtualCount++; return $false }
            if ($tt -eq 'Elastic')  { $skippedTables[$_] = 'Elastic'; $elasticCount++; return $false }
            return $true
        }
        if ($skippedTables.Count -gt 0) {
            Write-Host "Pre-skipping $($skippedTables.Count) unsupported table(s): $virtualCount Virtual, $elasticCount Elastic. Use -IncludeUnsupportedTypes to attempt them anyway." -ForegroundColor Yellow
        }
    }

    $totalTables = $namesToBatch.Count
    $totalBatches = if ($totalTables -gt 0) { [math]::Ceiling($totalTables / $BatchSize) } else { 0 }
    $allCounts = @{}
    $failedBatches = @()

    Write-Host "Retrieving record counts for $totalTables table(s) in $totalBatches batch(es) of up to $BatchSize tables..." -ForegroundColor Cyan

    for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
        $startIdx = $batchIndex * $BatchSize
        $endIdx = [math]::Min($startIdx + $BatchSize - 1, $totalTables - 1)
        $batchNames = $namesToBatch[$startIdx..$endIdx]
        $batchNum = $batchIndex + 1

        Write-Progress -Activity "Getting record counts" -Status "Batch $batchNum of $totalBatches ($($batchNames.Count) tables)" -PercentComplete (($batchNum / $totalBatches) * 100)

        try {
            $batchCounts = Get-RecordCountsBatch -OrgUrl $OrgUrl -Headers $Headers -EntityNames $batchNames
            foreach ($key in $batchCounts.Keys) {
                $allCounts[$key] = $batchCounts[$key]
            }
        }
        catch {
            # RetrieveTotalRecordCount rejects the entire batch payload if any single entity is
            # unsupported (virtual tables, elastic tables, etc.). Retry each table individually
            # so one bad apple doesn't poison the rest.
            Write-Warning "Batch $batchNum failed ($($batchNames.Count) tables) - retrying individually..."
            $individualSuccesses = 0
            $individualFailures = 0
            for ($i = 0; $i -lt $batchNames.Count; $i++) {
                $singleName = $batchNames[$i]
                Write-Progress -Activity "Getting record counts" `
                    -Status "Batch $batchNum retry: $($i + 1) of $($batchNames.Count) - $singleName" `
                    -PercentComplete (($batchNum / $totalBatches) * 100) `
                    -CurrentOperation "Retrying $singleName individually"
                try {
                    $singleCounts = Get-RecordCountsBatch -OrgUrl $OrgUrl -Headers $Headers -EntityNames @($singleName)
                    foreach ($key in $singleCounts.Keys) {
                        $allCounts[$key] = $singleCounts[$key]
                    }
                    $individualSuccesses++
                }
                catch {
                    $failedBatches += $singleName
                    $individualFailures++
                }
            }
            Write-Host "  Batch $batchNum retry: $individualSuccesses succeeded, $individualFailures failed individually." -ForegroundColor Yellow
        }
    }

    Write-Progress -Activity "Getting record counts" -Completed

    # Optionally enrich results with last CreatedOn/ModifiedOn timestamps
    $lastActivityMap = @{}
    if ($IncludeLastActivity) {
        # By default only query tables with count > 0 to save API calls.
        # When -ActivityFallback is set, also query tables where the count came back as 0
        # or N/A (the count API can return stale 0 values; the activity probe is authoritative).
        # Skipped Virtual tables: don't probe (their data provider may not support OData filters).
        # Skipped Elastic tables: still probe - they support OData $top/$orderby on createdon/modifiedon.
        $tablesToProbe = $allLogicalNames | Where-Object {
            if (-not $entitySetNameMap[$_]) { return $false }
            if ($skippedTables[$_] -eq 'Virtual') { return $false }
            $hasCount = $allCounts.ContainsKey($_)
            if ($hasCount -and $allCounts[$_] -gt 0) { return $true }
            if ($skippedTables[$_] -eq 'Elastic') { return $true }  # always probe elastic
            if ($ActivityFallback) { return $true }
            return $false
        }

        $totalToQuery = $tablesToProbe.Count
        $modeNote = if ($ActivityFallback) { " (fallback enabled - includes empty/N/A tables)" } else { "" }
        Write-Host "Retrieving last activity timestamps for $totalToQuery table(s)$modeNote..." -ForegroundColor Cyan

        $idx = 0
        foreach ($logicalName in $tablesToProbe) {
            $idx++
            $entitySetName = $entitySetNameMap[$logicalName]
            Write-Progress -Activity "Getting last activity" -Status "$idx of $totalToQuery : $logicalName" -PercentComplete (($idx / [math]::Max($totalToQuery, 1)) * 100)
            $lastActivityMap[$logicalName] = Get-LastActivityForTable -OrgUrl $OrgUrl -Headers $Headers -EntitySetName $entitySetName
        }

        Write-Progress -Activity "Getting last activity" -Completed
    }

    # Build results
    $results = @()
    # Use UTC for the "now" baseline because Dataverse returns CreatedOn/ModifiedOn in UTC.
    # Comparing in UTC avoids timezone drift that can produce slightly-negative deltas.
    $now = (Get-Date).ToUniversalTime()
    foreach ($logicalName in $allLogicalNames) {
        $displayName    = $displayNameMap[$logicalName]
        $schemaName     = $schemaNameMap[$logicalName]
        $entitySetName  = $entitySetNameMap[$logicalName]
        $isCustomEntity = $isCustomEntityMap[$logicalName]
        $tableType      = $tableTypeMap[$logicalName]

        $lastCreated  = $null
        $lastModified = $null
        $oldestCreated = $null
        if ($IncludeLastActivity -and $lastActivityMap.ContainsKey($logicalName)) {
            $lastCreated   = $lastActivityMap[$logicalName].LastCreatedOn
            $lastModified  = $lastActivityMap[$logicalName].LastModifiedOn
            $oldestCreated = $lastActivityMap[$logicalName].OldestRecordCreatedOn
        }

        # Determine status & record count value
        if ($skippedTables.ContainsKey($logicalName)) {
            $status      = "Skipped ($($skippedTables[$logicalName]))"
            $recordCount = "N/A"
        }
        elseif ($allCounts.ContainsKey($logicalName)) {
            $status      = "Success"
            $recordCount = $allCounts[$logicalName]
        }
        elseif ($logicalName -in $failedBatches) {
            $status      = "Error"
            $recordCount = "N/A"
        }
        else {
            $status      = "Not Returned by API"
            $recordCount = "N/A"
        }

        # Compute Days* metrics and UsageBucket
        $daysSinceLastCreated  = $null
        $daysSinceLastModified = $null
        $usageBucket           = $null
        if ($IncludeLastActivity) {
            # Compare in UTC: Dataverse returns timestamps in UTC, $now is also UTC above.
            # Clamp to 0 as a safety net (a record can't legitimately be modified in the future).
            if ($lastCreated)  { $daysSinceLastCreated  = [math]::Max(0, [int]([math]::Floor(($now - ([datetime]$lastCreated).ToUniversalTime()).TotalDays)))  }
            if ($lastModified) { $daysSinceLastModified = [math]::Max(0, [int]([math]::Floor(($now - ([datetime]$lastModified).ToUniversalTime()).TotalDays))) }

            # UsageBucket is based primarily on most recent activity (modified, falling back to created)
            $referenceDays = if ($null -ne $daysSinceLastModified) { $daysSinceLastModified }
                             elseif ($null -ne $daysSinceLastCreated) { $daysSinceLastCreated }
                             else { $null }

            # If we got activity timestamps, the table demonstrably has records - even if the
            # count API returned 0 or N/A (stale stats, or skipped Virtual/Elastic). Trust the
            # activity probe in that case.
            $hasActivityEvidence = ($null -ne $referenceDays)

            if ($status -like 'Skipped*' -and -not $hasActivityEvidence) {
                $usageBucket = "Unsupported"
            }
            elseif ($status -ne "Success" -and $status -notlike 'Skipped*' -and -not $hasActivityEvidence) {
                $usageBucket = "Unknown"
            }
            elseif (-not $hasActivityEvidence -and ($recordCount -eq 0 -or $recordCount -eq "N/A")) {
                # No timestamps and count is 0 or unavailable -> truly empty (or unknowable)
                $usageBucket = if ($recordCount -eq 0) { "Empty" } else { "Unknown" }
            }
            elseif ($null -eq $referenceDays) {
                $usageBucket = "Unknown"
            }
            elseif ($referenceDays -le 90) {
                $usageBucket = "Active (<=90d)"
            }
            elseif ($referenceDays -le 365) {
                $usageBucket = "Dormant (91-365d)"
            }
            else {
                $usageBucket = "Stale (>365d)"
            }
        }

        # Assemble output object (consistent column order across all rows)
        $obj = [ordered]@{
            TableLogicalName = $logicalName
            TableDisplayName = $displayName
            SchemaName       = $schemaName
            EntitySetName    = $entitySetName
            IsCustomEntity   = $isCustomEntity
            TableType        = $tableType
            RecordCount      = $recordCount
            Status           = $status
        }
        if ($IncludeLastActivity) {
            $obj.LastCreatedOn         = $lastCreated
            $obj.LastModifiedOn        = $lastModified
            $obj.OldestRecordCreatedOn = $oldestCreated
            $obj.DaysSinceLastCreated  = $daysSinceLastCreated
            $obj.DaysSinceLastModified = $daysSinceLastModified
            $obj.UsageBucket           = $usageBucket
        }
        $results += [PSCustomObject]$obj
    }

    return $results
}

# Main script execution
try {
    # Determine which tables to query
    if ($Tables -and $Tables.Count -gt 0) {
        Write-Host "Processing $($Tables.Count) specified table(s)..." -ForegroundColor Cyan
        $tablesToQuery = $Tables
    }
    else {
        # Get all readable tables from metadata
        $tablesToQuery = Get-AllReadableTables -OrgUrl $OrganizationUrl -Headers $headers -IncludeSystem $IncludeSystemTables -CustomOnly $CustomEntitiesOnly
    }

    # Get record counts for all tables using RetrieveTotalRecordCount
    $results = Get-RecordCounts -OrgUrl $OrganizationUrl -Headers $headers -TableList $tablesToQuery -BatchSize $BatchSize -IncludeLastActivity:$IncludeLastActivity -ActivityFallback:$ActivityFallback -SkipUnsupportedTypes:(-not $IncludeUnsupportedTypes)

    # Calculate summary statistics
    $successfulResults = $results | Where-Object { $_.Status -eq "Success" }
    $notReturnedTables = $results | Where-Object { $_.Status -eq "Not Returned by API" }
    $errorTables       = $results | Where-Object { $_.Status -eq "Error" }
    $skippedResults    = $results | Where-Object { $_.Status -like 'Skipped*' }
    $totalRecords      = ($successfulResults | Measure-Object -Property RecordCount -Sum).Sum
    $tablesWithData    = ($successfulResults | Where-Object { $_.RecordCount -gt 0 }).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total tables queried: $($results.Count)"
    Write-Host "Tables with counts returned: $($successfulResults.Count)" -ForegroundColor Green
    if ($skippedResults.Count -gt 0) {
        Write-Host "Tables skipped (Virtual/Elastic): $($skippedResults.Count)" -ForegroundColor Yellow
    }
    if ($notReturnedTables.Count -gt 0) {
        Write-Host "Tables not returned by API (virtual/unsupported): $($notReturnedTables.Count)" -ForegroundColor Yellow
    }
    Write-Host "Tables with errors: $($errorTables.Count)" -ForegroundColor $(if ($errorTables.Count -gt 0) { "Red" } else { "Green" })
    Write-Host "Tables with data: $tablesWithData"
    Write-Host "Total records across all tables: $totalRecords"
    Write-Host "Note: Counts are approximate (from system indexes)." -ForegroundColor Yellow
    Write-Host ""

    # Output results based on format
    switch ($OutputFormat) {
        "Table" {
            if ($OutputPath) {
                $results | Format-Table -AutoSize | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $results | Sort-Object @{Expression={if($_.RecordCount -eq "N/A"){-1}else{[long]$_.RecordCount}}; Descending=$true} | Format-Table -AutoSize
            }
        }
        "CSV" {
            if (-not $OutputPath) {
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "recordcounts_$timestamp.csv"
            }
            $results | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            $jsonOutput = $results | ConvertTo-Json -Depth 3
            if (-not $OutputPath) {
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "recordcounts_$timestamp.json"
            }
            $jsonOutput | Out-File -FilePath $OutputPath
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
    }

    # Return results for pipeline use
    return $results
}
catch {
    Write-Error "Script execution failed: $_"
    throw
}
