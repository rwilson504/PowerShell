<#
.SYNOPSIS
    For each table, reports per-user record counts on the standard creator/modifier columns
    plus any user lookups you specify. Designed to identify "who owns this process" across
    audit columns AND business-meaningful approver / reviewer / owner-style lookups.

.DESCRIPTION
    For every requested table the script:
      1. Resolves the table's lookup attributes whose Targets include 'systemuser'.
      2. Picks the set of user-lookup columns to analyze:
           - Standard four (createdby, createdonbehalfby, modifiedby, modifiedonbehalfby)
             unless -ExcludeStandardUserAttributes is set.
           - User-specified columns via -UserLookupAttributes.
           - All systemuser-targeted lookups when -AutoDetectUserLookups is set.
      3. For each (table, user-lookup) pair, issues a single FetchXML aggregate query that
          groupbys the user column and counts records. This returns one row per distinct user
          in a single API call, regardless of table size (subject to the standard ~50,000
          aggregate-record limit).
      4. Optionally restricts records via -Filter (an OData $filter expression applied through
          the FetchXML <filter> element).

    Output rows: one per (Table, AttributeLogicalName, UserId) with the user's display name,
    domain, IsDisabled flag, RecordCount, and a 1-based Rank within the (Table, Attribute)
    group. The composite join key (TableLogicalName + AttributeLogicalName) matches the other
    three CSVs in this folder so you can pivot user activity alongside fill-rate and audit
    data in Excel / Power BI / pandas.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Required. One or more table logical names to analyze.

.PARAMETER UserLookupAttributes
    Optional. Additional lookup attribute logical names (beyond the standard four) to include.
    These must be Lookup-type columns whose Targets include any of the -UserTargetTables
    (default: systemuser). The script verifies each from metadata and warns/skips any that
    don't match.

.PARAMETER UserTargetTables
    Which target table types qualify as 'user lookups'. Default is @('systemuser'). Add
    'contact' to also pick up customer-contact lookups, 'account' for org lookups, or
    any custom person table. The standard four audit columns (createdby, modifiedby, etc.)
    are always systemuser-targeted, so they're always included regardless of this list
    unless -ExcludeStandardUserAttributes is set.

    Example: -UserTargetTables 'systemuser','contact'  (covers internal users + portal
    contacts in the same report).

.PARAMETER AutoDetectUserLookups
    When set, the script auto-discovers EVERY Lookup attribute on each table whose Targets
    include any of the -UserTargetTables and includes them all. Saves you from having to
    know your custom-lookup names ahead of time.

.PARAMETER ExcludeStandardUserAttributes
    Skip the four standard audit lookups (createdby, createdonbehalfby, modifiedby,
    modifiedonbehalfby). Useful when you only want to see custom approver / reviewer activity
    without the high-volume system noise.

.PARAMETER Filter
    Optional OData $filter expression (e.g., "statecode eq 0" or
    "createdon ge 2025-01-01T00:00:00Z") that restricts which records are counted. Applied to
    EVERY user-lookup query so all results reflect the same record subset.

.PARAMETER TopUsersPerAttribute
    When > 0, only emit the top N users per (Table, Attribute) pair (sorted by RecordCount
    descending). Default 0 = emit all distinct users. Useful to keep the CSV compact when an
    attribute has thousands of distinct users.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetUserActivityByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program"

    Reports per-user record counts on createdby/createdonbehalfby/modifiedby/modifiedonbehalfby
    for every msf_program record.

.EXAMPLE
    .\GetUserActivityByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program" -UserLookupAttributes "msf_approver","msf_reviewer" -OutputFormat CSV

    Adds the two custom approver/reviewer lookups so you see who's owning approvals as well
    as who's editing the records.

.EXAMPLE
    .\GetUserActivityByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account","contact","msf_program" -AutoDetectUserLookups -ExcludeStandardUserAttributes -OutputFormat CSV

    Auto-discovers every custom user lookup on each table and reports activity ONLY on those
    (skipping the high-volume standard audit columns). Best for finding business-process owners.

.NOTES
    CORRELATING WITH OTHER CSVs IN THIS FOLDER

    The output joins to attributeusage_*.csv, audithistory_*.csv, and relationships_*.csv on
    TableLogicalName + AttributeLogicalName. In Excel:
      - Pivot useractivity_*.csv by User to find your most-active people across many tables
      - Pivot by (Table, Attribute) to find single-owner processes (Distinct = 1) - they
        disappear when that person leaves
      - Filter UserDisplayName for known service identities ('SYSTEM', '#'-prefixed app users)
        to separate human work from automation churn

    LIMITATIONS

    FetchXML aggregate queries are subject to a server-side ~50,000-record limit. Tables
    larger than that will return an aggregate-limit fault and the row is emitted with
    Status='AggregateLimitExceeded'. A future enhancement could fall back to a paged scan
    when this happens.

    SystemUser display name is the user's 'fullname'. Non-interactive system users (accessmode=4)
    have IsServiceAccount = true so you can filter them out / focus on them.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $false)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [string]$SolutionUniqueName,

    [Parameter(Mandatory = $false)]
    [string[]]$UserLookupAttributes,

    [Parameter(Mandatory = $false)]
    [string[]]$UserTargetTables = @('systemuser'),

    [Parameter(Mandatory = $false)]
    [switch]$AutoDetectUserLookups,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeStandardUserAttributes,

    [Parameter(Mandatory = $false)]
    [string]$Filter,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1000)]
    [int]$TopUsersPerAttribute = 0,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Content-Type"     = "application/json; charset=utf-8"
    "Prefer"           = "odata.include-annotations=*"
}

$StandardUserAttributes = @('createdby','createdonbehalfby','modifiedby','modifiedonbehalfby')

function Get-EntityMetadata {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')?" +
        "`$select=LogicalName,SchemaName,EntitySetName,DisplayName"
    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $displayName = if ($resp.DisplayName.UserLocalizedLabel) {
            $resp.DisplayName.UserLocalizedLabel.Label
        } else { $LogicalName }
        return [PSCustomObject]@{
            LogicalName   = $resp.LogicalName
            SchemaName    = $resp.SchemaName
            EntitySetName = $resp.EntitySetName
            DisplayName   = $displayName
        }
    }
    catch {
        Write-Warning "Failed to load metadata for table '$LogicalName': $_"
        return $null
    }
}

function Get-SystemUserLookups {
    <#
    .SYNOPSIS
        Returns all Lookup attributes on the table whose Targets include any of the supplied
        target tables (default: systemuser). Each item exposes LogicalName, SchemaName,
        DisplayName, Targets, IsCustomAttribute.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName, [string[]]$TargetTables = @('systemuser'))

    # Normalize target list for case-insensitive matching
    $targetSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($TargetTables | ForEach-Object { $_.ToLowerInvariant() }),
        [System.StringComparer]::OrdinalIgnoreCase)

    # Step 1: pull base attribute metadata so we can grab DisplayName, IsCustomAttribute
    $baseAttrUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes?" +
        "`$select=LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute,AttributeOf,IsLogical"
    $baseMap = @{}
    try {
        do {
            $resp = Invoke-RestMethod -Uri $baseAttrUrl -Headers $Headers -Method Get
            foreach ($a in $resp.value) {
                $baseMap[$a.LogicalName] = [PSCustomObject]@{
                    LogicalName       = $a.LogicalName
                    SchemaName        = $a.SchemaName
                    DisplayName       = if ($a.DisplayName.UserLocalizedLabel) { $a.DisplayName.UserLocalizedLabel.Label } else { $a.LogicalName }
                    AttributeType     = $a.AttributeType
                    IsCustomAttribute = [bool]$a.IsCustomAttribute
                    AttributeOf       = $a.AttributeOf
                    IsLogical         = [bool]$a.IsLogical
                }
            }
            $baseAttrUrl = $resp.'@odata.nextLink'
        } while ($baseAttrUrl)
    }
    catch {
        Write-Warning "Failed to load attribute metadata for '$LogicalName': $_"
        return @()
    }

    # Step 2: pull Lookup-cast metadata (Targets only available on the lookup cast)
    $lookupUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?" +
        "`$select=LogicalName,Targets"
    $userLookups = @()
    try {
        do {
            $resp = Invoke-RestMethod -Uri $lookupUrl -Headers $Headers -Method Get
            foreach ($a in $resp.value) {
                # Match if any of this lookup's Targets is in the requested set
                $matchedTargets = @($a.Targets | Where-Object { $targetSet.Contains($_) })
                if ($matchedTargets.Count -gt 0) {
                    $base = $baseMap[$a.LogicalName]
                    if ($base -and -not $base.IsLogical -and -not $base.AttributeOf) {
                        $userLookups += [PSCustomObject]@{
                            LogicalName       = $a.LogicalName
                            SchemaName        = $base.SchemaName
                            DisplayName       = $base.DisplayName
                            Targets           = ($a.Targets -join ';')
                            IsCustomAttribute = $base.IsCustomAttribute
                        }
                    }
                }
            }
            $lookupUrl = $resp.'@odata.nextLink'
        } while ($lookupUrl)
    }
    catch {
        Write-Warning "Failed to load lookup metadata for '$LogicalName': $_"
    }

    return $userLookups
}

function Invoke-FetchXmlAggregate {
    <#
    .SYNOPSIS
        Issues a FetchXML aggregate query (groupby + count) for one user-lookup column on
        one table, optionally with a record filter. Returns rows of @{ UserId; UserName;
        UserLookupLogicalName; Count } - or a Status outcome.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$EntitySetName,
        [string]$EntityLogicalName,
        [string]$UserLookupLogicalName,
        [string]$Filter
    )

    $alias = 'usr'
    $cntAlias = 'cnt'

    # Parse OData filter into FetchXML conditions for accurate parity with other scripts.
    # Supported simple patterns: "<col> <op> <value>" with eq/ne/gt/ge/lt/le and chained 'and'.
    # For unsupported patterns we wrap the entire expression in a <filter type='and'> with a
    # single placeholder condition - that won't parse, so we pass the raw OData filter to a
    # separate request. To keep this self-contained and reliable we instead omit the filter
    # from the FetchXML and apply it via a follow-up if needed. For now: ignore -Filter when
    # FetchXML can't safely express it (we explicitly document that FetchXML is the engine).
    # The user can already get unfiltered counts; advanced filtering is a future enhancement.
    $filterXml = ''
    if (-not [string]::IsNullOrWhiteSpace($Filter)) {
        $filterXml = ConvertTo-FetchXmlFilter -ODataFilter $Filter
    }

    $fetchXml = "<fetch aggregate=`"true`"><entity name=`"$EntityLogicalName`">" +
                "<attribute name=`"$UserLookupLogicalName`" alias=`"$alias`" groupby=`"true`" />" +
                "<attribute name=`"$EntityLogicalName" + "id`" alias=`"$cntAlias`" aggregate=`"count`" />" +
                $filterXml +
                "</entity></fetch>"

    $encoded = [System.Uri]::EscapeDataString($fetchXml)
    $url = "$OrgUrl/api/data/v9.2/$EntitySetName" + "?fetchXml=$encoded"

    try {
        $resp = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
    }
    catch {
        $msg = $_.Exception.Message
        # 0x8004E023 = AggregateQueryRecordLimitExceeded
        if ($msg -match 'AggregateQueryRecordLimit' -or $msg -match '0x8004E023') {
            return [PSCustomObject]@{ Status = 'AggregateLimitExceeded'; Rows = @() }
        }
        Write-Warning "Aggregate query failed for $EntityLogicalName.$UserLookupLogicalName : $msg"
        return [PSCustomObject]@{ Status = 'Error'; Rows = @() }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($g in $resp.value) {
        $userId   = $g.$alias
        $userName = $g."$alias@OData.Community.Display.V1.FormattedValue"
        $targetEntity = $g."$alias@Microsoft.Dynamics.CRM.lookuplogicalname"   # null when group is null
        $count    = [long]$g.$cntAlias
        if ([string]::IsNullOrWhiteSpace($userId)) {
            # Null / not-set bucket - represents records where this user lookup was empty
            $rows.Add([PSCustomObject]@{
                UserId       = ''
                UserName     = '(no value)'
                TargetEntity = ''
                Count        = $count
            }) | Out-Null
        }
        else {
            $rows.Add([PSCustomObject]@{
                UserId       = $userId
                UserName     = if ($userName) { $userName } else { $userId }
                TargetEntity = $targetEntity
                Count        = $count
            }) | Out-Null
        }
    }
    return [PSCustomObject]@{ Status = 'Success'; Rows = $rows.ToArray() }
}

function ConvertTo-FetchXmlFilter {
    <#
    .SYNOPSIS
        Converts a small subset of OData filter syntax into a <filter> FetchXML fragment.
        Supports clauses joined by 'and' using the operators eq/ne/gt/ge/lt/le and quoted
        string / unquoted numeric / ISO datetime values. Any expression we can't translate
        is returned as-is wrapped in a comment for visibility.
    #>
    param ([string]$ODataFilter)
    if ([string]::IsNullOrWhiteSpace($ODataFilter)) { return '' }

    $opMap = @{ 'eq' = 'eq'; 'ne' = 'ne'; 'gt' = 'gt'; 'ge' = 'ge'; 'lt' = 'lt'; 'le' = 'le' }
    $clauses = $ODataFilter -split '(?i)\s+and\s+'
    $conditions = New-Object System.Collections.Generic.List[string]
    foreach ($clause in $clauses) {
        $clean = $clause.Trim().Trim('(',')')
        if ($clean -match "^\s*([a-z0-9_]+)\s+(eq|ne|gt|ge|lt|le)\s+(.+?)\s*$") {
            $col = $matches[1]
            $op  = $matches[2].ToLowerInvariant()
            $valRaw = $matches[3].Trim()
            $val = if ($valRaw -match "^'(.*)'$") { $matches[1] } else { $valRaw }
            $valEsc = $val -replace "'", "&apos;" -replace '<','&lt;' -replace '>','&gt;' -replace '&','&amp;'
            $conditions.Add("<condition attribute=`"$col`" operator=`"$($opMap[$op])`" value=`"$valEsc`" />") | Out-Null
        }
        else {
            Write-Warning "Could not translate OData clause '$clean' into FetchXML; ignoring it."
        }
    }
    if ($conditions.Count -eq 0) { return '' }
    return "<filter type=`"and`">" + ($conditions -join '') + "</filter>"
}

function Get-UserDirectory {
    <#
    .SYNOPSIS
        Looks up display info for the supplied user-record GUIDs across one or more target
        tables (systemuser, contact, account, etc.) and returns
        @{ "<targetEntity>|<guid>" -> @{ DisplayName; DomainName; IsDisabled; IsServiceAccount } }.
        For non-systemuser tables IsDisabled / IsServiceAccount are best-effort and may be empty.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [hashtable]$UserIdsByTarget)
    $map = @{}
    if (-not $UserIdsByTarget -or $UserIdsByTarget.Count -eq 0) { return $map }

    foreach ($target in $UserIdsByTarget.Keys) {
        $ids = $UserIdsByTarget[$target] | Where-Object { $_ } | Select-Object -Unique
        if ($ids.Count -eq 0) { continue }

        # Determine the entity-set name + the relevant select columns for this target type.
        switch ($target) {
            'systemuser' {
                $entitySet = 'systemusers'
                $idCol     = 'systemuserid'
                $selectCols = 'systemuserid,fullname,domainname,isdisabled,accessmode'
            }
            'contact' {
                $entitySet  = 'contacts'
                $idCol      = 'contactid'
                $selectCols = 'contactid,fullname,emailaddress1,statecode'
            }
            'account' {
                $entitySet  = 'accounts'
                $idCol      = 'accountid'
                $selectCols = 'accountid,name,emailaddress1,statecode'
            }
            default {
                # Fall back to plural-by-+s convention; user can override behavior by editing
                # this switch if their custom person table doesn't follow that convention.
                $entitySet  = $target + 's'
                $idCol      = $target + 'id'
                $selectCols = $idCol
            }
        }

        $chunkSize = 25
        $totalChunks = [math]::Ceiling($ids.Count / $chunkSize)
        for ($i = 0; $i -lt $totalChunks; $i++) {
            $start = $i * $chunkSize
            $end   = [math]::Min($start + $chunkSize - 1, $ids.Count - 1)
            $chunk = $ids[$start..$end]
            $filter = ($chunk | ForEach-Object { "$idCol eq $_" }) -join ' or '
            $url = "$OrgUrl/api/data/v9.2/$entitySet`?`$select=$selectCols&`$filter=$filter"
            try {
                $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
                foreach ($u in $r.value) {
                    $id = $u.$idCol
                    $display = $null
                    if ($u.PSObject.Properties['fullname'] -and $u.fullname) { $display = $u.fullname }
                    elseif ($u.PSObject.Properties['name'] -and $u.name)     { $display = $u.name }
                    $domainOrEmail = $null
                    if ($u.PSObject.Properties['domainname'] -and $u.domainname)       { $domainOrEmail = $u.domainname }
                    elseif ($u.PSObject.Properties['emailaddress1'] -and $u.emailaddress1) { $domainOrEmail = $u.emailaddress1 }
                    $isDisabled = $null
                    if ($u.PSObject.Properties['isdisabled'])  { $isDisabled = [bool]$u.isdisabled }
                    elseif ($u.PSObject.Properties['statecode']) { $isDisabled = ([int]$u.statecode -ne 0) }
                    $isService = $false
                    if ($target -eq 'systemuser' -and $u.PSObject.Properties['accessmode']) {
                        $isService = ([int]$u.accessmode -eq 4)
                    }
                    $map["$target|$id"] = [PSCustomObject]@{
                        TargetEntity     = $target
                        DisplayName      = $display
                        DomainName       = $domainOrEmail
                        IsDisabled       = $isDisabled
                        IsServiceAccount = $isService
                    }
                }
            }
            catch {
                Write-Warning "User directory lookup failed for $target chunk $($i + 1) of $totalChunks : $_"
            }
        }
    }
    return $map
}

# Main script execution
try {
    # Load shared solution-filter helper and apply -SolutionUniqueName scope
    . (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "_SolutionFilterHelper.ps1")
    $Tables = Resolve-SolutionScopedTables -OrgUrl $OrganizationUrl -Headers $headers -Tables $Tables -SolutionUniqueName $SolutionUniqueName
    if (-not $Tables -or $Tables.Count -eq 0) {
        Write-Error "No tables to process. Specify -Tables and/or -SolutionUniqueName."
        exit 1
    }

    $allResults = New-Object System.Collections.Generic.List[object]

    if ($Filter) {
        Write-Host "Applying -Filter (translated to FetchXML where possible): $Filter" -ForegroundColor Cyan
    }

    foreach ($logicalName in $Tables) {
        Write-Host "`n=== Processing table: $logicalName ===" -ForegroundColor Cyan

        $meta = Get-EntityMetadata -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        if (-not $meta) {
            Write-Warning "Skipping '$logicalName' (metadata lookup failed)."
            continue
        }

        $userLookups = Get-SystemUserLookups -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName -TargetTables $UserTargetTables
        Write-Host "  user-targeted lookups available (targets: $($UserTargetTables -join ', ')): $($userLookups.Count)" -ForegroundColor Gray

        # Decide which attributes to query
        $selected = New-Object System.Collections.Generic.List[object]
        $availableMap = @{}
        foreach ($u in $userLookups) { $availableMap[$u.LogicalName] = $u }

        if (-not $ExcludeStandardUserAttributes) {
            foreach ($std in $StandardUserAttributes) {
                if ($availableMap.ContainsKey($std)) { $selected.Add($availableMap[$std]) | Out-Null }
            }
        }

        if ($AutoDetectUserLookups) {
            foreach ($u in $userLookups) {
                if (-not $selected.Contains($u)) { $selected.Add($u) | Out-Null }
            }
        }

        if ($UserLookupAttributes) {
            foreach ($a in $UserLookupAttributes) {
                if ($availableMap.ContainsKey($a)) {
                    if (-not $selected.Contains($availableMap[$a])) { $selected.Add($availableMap[$a]) | Out-Null }
                }
                else {
                    Write-Warning "  Requested user-lookup '$a' is not a Lookup on $logicalName whose Targets include any of: $($UserTargetTables -join ', ') - skipping."
                }
            }
        }

        if ($selected.Count -eq 0) {
            Write-Warning "  No user-lookup columns selected for $logicalName. Skipping."
            continue
        }

        Write-Host "  User-lookup columns being analyzed: $(($selected | ForEach-Object { $_.LogicalName }) -join ', ')" -ForegroundColor Green

        # Run aggregate per attribute
        $idx = 0
        $idsByTarget = @{}   # targetEntity -> List[guid]
        $perAttrRows = @{}
        $perAttrStatus = @{}
        foreach ($attr in $selected) {
            $idx++
            Write-Progress -Activity "User activity: $logicalName" `
                -Status "$idx of $($selected.Count) - $($attr.LogicalName)" `
                -PercentComplete (($idx / $selected.Count) * 100)

            $result = Invoke-FetchXmlAggregate -OrgUrl $OrganizationUrl -Headers $headers `
                -EntitySetName $meta.EntitySetName `
                -EntityLogicalName $logicalName `
                -UserLookupLogicalName $attr.LogicalName `
                -Filter $Filter

            $perAttrStatus[$attr.LogicalName] = $result.Status
            $perAttrRows[$attr.LogicalName]   = $result.Rows
            foreach ($r in $result.Rows) {
                if ($r.UserId -and $r.TargetEntity) {
                    if (-not $idsByTarget.ContainsKey($r.TargetEntity)) {
                        $idsByTarget[$r.TargetEntity] = New-Object System.Collections.Generic.List[string]
                    }
                    $idsByTarget[$r.TargetEntity].Add($r.UserId) | Out-Null
                }
            }
        }
        Write-Progress -Activity "User activity: $logicalName" -Completed

        $totalIdRefs = ($idsByTarget.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Write-Host "  Looking up directory data for $totalIdRefs total user-record reference(s) across $($idsByTarget.Count) target table(s)..." -ForegroundColor Gray
        $userDir = Get-UserDirectory -OrgUrl $OrganizationUrl -Headers $headers -UserIdsByTarget $idsByTarget

        # Emit output rows
        foreach ($attr in $selected) {
            $status = $perAttrStatus[$attr.LogicalName]
            $rows = $perAttrRows[$attr.LogicalName]

            if ($status -ne 'Success') {
                $allResults.Add([PSCustomObject][ordered]@{
                    TableLogicalName     = $logicalName
                    TableDisplayName     = $meta.DisplayName
                    TableSchemaName      = $meta.SchemaName
                    AttributeLogicalName = $attr.LogicalName
                    AttributeSchemaName  = $attr.SchemaName
                    AttributeDisplayName = $attr.DisplayName
                    AttributeType        = 'Lookup'
                    IsCustomAttribute    = $attr.IsCustomAttribute
                    LookupTargets        = $attr.Targets
                    UserTargetEntity     = ''
                    UserId               = ''
                    UserDisplayName      = ''
                    UserDomainName       = ''
                    IsDisabled           = ''
                    IsServiceAccount     = ''
                    RecordCount          = 'N/A'
                    Rank                 = ''
                    Status               = $status
                }) | Out-Null
                continue
            }

            # Sort rows desc by Count, place '(no value)' last
            $sorted = $rows | Sort-Object @{Expression = { if ($_.UserId) { 1 } else { 0 } }; Descending = $true},
                                          @{Expression = { [long]$_.Count }; Descending = $true}

            if ($TopUsersPerAttribute -gt 0) {
                $sorted = $sorted | Select-Object -First $TopUsersPerAttribute
            }

            $rank = 0
            foreach ($r in $sorted) {
                $rank++
                $u = $null
                if ($r.UserId -and $r.TargetEntity) {
                    $dirKey = "$($r.TargetEntity)|$($r.UserId)"
                    if ($userDir.ContainsKey($dirKey)) { $u = $userDir[$dirKey] }
                }
                $allResults.Add([PSCustomObject][ordered]@{
                    TableLogicalName     = $logicalName
                    TableDisplayName     = $meta.DisplayName
                    TableSchemaName      = $meta.SchemaName
                    AttributeLogicalName = $attr.LogicalName
                    AttributeSchemaName  = $attr.SchemaName
                    AttributeDisplayName = $attr.DisplayName
                    AttributeType        = 'Lookup'
                    IsCustomAttribute    = $attr.IsCustomAttribute
                    LookupTargets        = $attr.Targets
                    UserTargetEntity     = $r.TargetEntity
                    UserId               = $r.UserId
                    UserDisplayName      = $r.UserName
                    UserDomainName       = if ($u) { $u.DomainName } else { '' }
                    IsDisabled           = if ($u) { $u.IsDisabled } else { '' }
                    IsServiceAccount     = if ($u) { $u.IsServiceAccount } else { '' }
                    RecordCount          = $r.Count
                    Rank                 = $rank
                    Status               = 'Success'
                }) | Out-Null
            }
        }
    }

    # Sort: by table, attribute, rank
    $sorted = $allResults | Sort-Object TableLogicalName, AttributeLogicalName, @{ Expression = { if ($_.Rank -eq '') { 9999 } else { [int]$_.Rank } } }

    # Summary
    $totalRows  = $sorted.Count
    $successRows = ($sorted | Where-Object Status -eq 'Success').Count
    $errorRows   = ($sorted | Where-Object Status -ne 'Success').Count
    $distinctUsers = ($sorted | Where-Object { $_.UserId } | Select-Object -ExpandProperty UserId -Unique).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total output rows: $totalRows"
    Write-Host "  Success rows:    $successRows" -ForegroundColor Green
    Write-Host "  Error rows:      $errorRows" -ForegroundColor $(if ($errorRows -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host "  Distinct users:  $distinctUsers"
    Write-Host ""

    switch ($OutputFormat) {
        "Table" {
            if ($OutputPath) {
                $sorted | Format-Table -AutoSize | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            } else {
                $sorted | Format-Table -AutoSize
            }
        }
        "CSV" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "useractivity_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "useractivity_$timestamp.json"
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
