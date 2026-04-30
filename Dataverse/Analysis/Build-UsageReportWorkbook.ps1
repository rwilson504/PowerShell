<#
.SYNOPSIS
    Combines all CSVs produced by Invoke-DataverseUsageReport.ps1 into pre-computed join
    CSVs (Master / Tables / Cleanup) plus an optional single .xlsx workbook containing
    every sheet.

.DESCRIPTION
    Takes the timestamped folder produced by Invoke-DataverseUsageReport (containing the 8
    analysis CSVs) and writes additional CSVs that do the cross-CSV correlation work for you:

      master.csv    : full per-attribute join. Spine = attributeusage. audithistory,
                      uipresence, top-1 user from useractivity, and solution membership are
                      left-joined onto it on (TableLogicalName, AttributeLogicalName). Plus
                      a DeadFieldScore composite (0-4) implementing the 'is this safe to
                      delete?' recipe.
      tables.csv    : per-table roll-up joining recordcounts + tableusage on
                      TableLogicalName, plus aggregated containing-solutions list.
      cleanup.csv   : pre-filtered Master view (DeadFieldScore >= 2 AND IsCustomAttribute).
      README.md     : column glossary + how-to-use-in-Excel guidance.

    All four files are pure CSV / Markdown - they require no extra modules and work on any
    locked-down workstation. Open them in Excel directly, or use whatever analysis tool you
    prefer.

    OPTIONAL -CombineToXlsx switch: bundles every CSV (8 source + 3 generated) plus the
    README into a single .xlsx workbook with one sheet per file. Uses Excel COM automation
    (no PowerShell modules needed - just Excel installed on the workstation). Skipped
    silently if Excel isn't available.

.PARAMETER InputFolder
    The folder produced by Invoke-DataverseUsageReport.ps1.

.PARAMETER OutputFolder
    Folder to write the joined CSVs and (optionally) the .xlsx into. Defaults to InputFolder.

.PARAMETER CombineToXlsx
    When set, also produces UsageReport.xlsx by Excel COM automation (requires Excel
    installed on the workstation; no PowerShell modules required).

.PARAMETER OpenAfterBuild
    With -CombineToXlsx, opens the .xlsx in Excel after building.

.EXAMPLE
    .\Build-UsageReportWorkbook.ps1 -InputFolder ".\dataverse-usage-orgABC-msf_Core-20260429_120000"

    Produces master.csv, tables.csv, cleanup.csv, and README.md inside the input folder.

.EXAMPLE
    .\Build-UsageReportWorkbook.ps1 -InputFolder ".\dataverse-usage-..." -CombineToXlsx -OpenAfterBuild

    Produces the join CSVs AND a UsageReport.xlsx with one sheet per CSV; opens it in Excel.

.NOTES
    Why not the ImportExcel module? Production / locked-down workstations frequently can't
    Install-Module. Pure CSV output works everywhere. The optional -CombineToXlsx path uses
    Excel COM which is built-in to any machine that has Office installed.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$InputFolder,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [switch]$CombineToXlsx,

    [Parameter(Mandatory = $false)]
    [switch]$OpenAfterBuild
)

if (-not (Test-Path $InputFolder)) {
    Write-Error "Input folder not found: $InputFolder"
    exit 1
}
$InputFolder = (Resolve-Path $InputFolder).Path
if (-not $OutputFolder) { $OutputFolder = $InputFolder }
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$OutputFolder = (Resolve-Path $OutputFolder).Path

Write-Host "Reading CSVs from: $InputFolder" -ForegroundColor Cyan

# Map of "logical report name" -> filename pattern
$reportPatterns = [ordered]@{
    RecordCounts       = 'recordcounts_*.csv'
    TableUsage         = 'tableusage_*.csv'
    AttributeUsage     = 'attributeusage_*.csv'
    AuditHistory       = 'audithistory_*.csv'
    UserActivity       = 'useractivity_*.csv'
    UIPresence         = 'uipresence_*.csv'
    Relationships      = 'relationships_*.csv'
    SolutionMembership = 'solutionmembership_*.csv'
    SitemapPresence    = 'sitemappresence_*.csv'
}

$datasets = @{}
$datasetFiles = @{}
foreach ($name in $reportPatterns.Keys) {
    $match = Get-ChildItem -Path $InputFolder -Filter $reportPatterns[$name] -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($match) {
        Write-Host "  $name -> $($match.Name)" -ForegroundColor Gray
        $datasets[$name]     = @(Import-Csv -Path $match.FullName)
        $datasetFiles[$name] = $match.FullName
    }
    else {
        Write-Host "  $name -> (not found, skipping)" -ForegroundColor DarkGray
        $datasets[$name]     = @()
        $datasetFiles[$name] = $null
    }
}

if ($datasets['AttributeUsage'].Count -eq 0) {
    Write-Warning "No attributeusage CSV found - the Master and Cleanup sheets will be empty."
}

# Helpers -----------------------------------------------------------------
function Get-RowMap   { param([array]$Rows); $m=@{}; foreach($r in $Rows){ if(-not $r){continue}; $k="$($r.TableLogicalName)|$($r.AttributeLogicalName)".ToLowerInvariant(); $m[$k]=$r }; return $m }
function Get-TableMap { param([array]$Rows); $m=@{}; foreach($r in $Rows){ if(-not $r){continue}; $m[$r.TableLogicalName.ToLowerInvariant()]=$r }; return $m }
function Get-TopUserMap {
    param([array]$Rows); $m=@{}
    foreach($r in $Rows){
        if(-not $r){continue}
        if($r.Rank -ne '1'){continue}
        if($r.UserDisplayName -eq '(no value)'){continue}
        $k="$($r.TableLogicalName)|$($r.AttributeLogicalName)".ToLowerInvariant()
        $m[$k]=$r
    }
    return $m
}

Write-Host "Building joins..." -ForegroundColor Cyan
$auditMap   = Get-RowMap     $datasets['AuditHistory']
$uiMap      = Get-RowMap     $datasets['UIPresence']
$topUserMap = Get-TopUserMap $datasets['UserActivity']
$rcMap      = Get-TableMap   $datasets['RecordCounts']
$tuMap      = Get-TableMap   $datasets['TableUsage']

# Solution membership: aggregate to per-attribute and per-table lists
$attrSolMap  = @{}
$tableSolMap = @{}
foreach ($r in $datasets['SolutionMembership']) {
    if ($r.ComponentType -eq 'Attribute') {
        $key = "$($r.TableLogicalName)|$($r.AttributeLogicalName)".ToLowerInvariant()
        if (-not $attrSolMap.ContainsKey($key)) { $attrSolMap[$key] = New-Object System.Collections.Generic.List[object] }
        $attrSolMap[$key].Add($r)
    }
    elseif ($r.ComponentType -eq 'Entity') {
        $key = $r.TableLogicalName.ToLowerInvariant()
        if (-not $tableSolMap.ContainsKey($key)) { $tableSolMap[$key] = New-Object System.Collections.Generic.List[object] }
        $tableSolMap[$key].Add($r)
    }
}

# Sitemap presence: aggregate per-table -> set of distinct apps that surface it.
# Only Entity-bound SubAreas (TableLogicalName non-empty) participate in this aggregate;
# non-entity tabs (Dashboard / Url / WebResource) are kept on the raw sheet for drill-down.
$tableSitemapMap = @{}
foreach ($r in $datasets['SitemapPresence']) {
    if ([string]::IsNullOrWhiteSpace($r.TableLogicalName)) { continue }
    $key = $r.TableLogicalName.ToLowerInvariant()
    if (-not $tableSitemapMap.ContainsKey($key)) { $tableSitemapMap[$key] = New-Object System.Collections.Generic.List[object] }
    $tableSitemapMap[$key].Add($r)
}

# ---- MASTER (one row per attribute) ----
$master = New-Object System.Collections.Generic.List[object]
foreach ($a in $datasets['AttributeUsage']) {
    $key      = "$($a.TableLogicalName)|$($a.AttributeLogicalName)".ToLowerInvariant()
    $tblKey   = $a.TableLogicalName.ToLowerInvariant()
    $aud      = if ($auditMap.ContainsKey($key))    { $auditMap[$key] }    else { $null }
    $ui       = if ($uiMap.ContainsKey($key))       { $uiMap[$key] }       else { $null }
    $topUser  = if ($topUserMap.ContainsKey($key))  { $topUserMap[$key] }  else { $null }
    $tu       = if ($tuMap.ContainsKey($tblKey))    { $tuMap[$tblKey] }    else { $null }

    $solMatches = if ($attrSolMap.ContainsKey($key)) { $attrSolMap[$key] } else { @() }
    $solUniqueNames = ($solMatches | ForEach-Object { $_.SolutionUniqueName } | Sort-Object -Unique) -join ';'
    $allManaged     = if ($solMatches.Count -gt 0) { (($solMatches | Where-Object { $_.SolutionIsManaged -eq 'False' }).Count -eq 0) } else { $null }
    $anyCustom      = if ($solMatches.Count -gt 0) { (($solMatches | Where-Object { $_.IsCustomComponent -eq 'True' }).Count -gt 0) } else { $null }

    # Table-level sitemap presence (computed once per parent table, then projected onto
    # every attribute row of that table). Empty if no SitemapPresence CSV was produced.
    $sitemapMatches = if ($tableSitemapMap.ContainsKey($tblKey)) { $tableSitemapMap[$tblKey] } else { @() }
    $tableInAnyApp  = $sitemapMatches.Count -gt 0
    $tableAppNames  = ($sitemapMatches | ForEach-Object { $_.AppDisplayName } | Where-Object { $_ } | Sort-Object -Unique) -join ';'
    $tableAppCount  = ($sitemapMatches | ForEach-Object { $_.AppUniqueName } | Where-Object { $_ } | Sort-Object -Unique).Count

    # DeadFieldScore (0-5).
    # +1 for each of:
    #   1. FillRatePercent = 0
    #   2. No audit events in window
    #   3. AnyUIPresence = False (forms/views/charts)
    #   4. All containing solutions are unmanaged (you can actually delete it)
    #   5. The parent TABLE isn't surfaced in any model-driven app sitemap
    #
    # FillRatePercent / TotalRecords / PopulatedCount columns are blank when the
    # upstream call could not produce a value (Status != Success in the source CSV).
    # Treat blank-or-non-numeric as "unknown" so the score doesn't credit a +1 for
    # missing data.
    $fillRateNum = if ([string]::IsNullOrWhiteSpace([string]$a.FillRatePercent)) { $null } else { [double]$a.FillRatePercent }
    $hasNoData     = ($null -ne $fillRateNum -and $fillRateNum -eq 0)
    $notTouched    = (-not $aud) -or [string]::IsNullOrWhiteSpace($aud.LastAuditedOn)
    $notInUI       = $ui -and ($ui.AnyUIPresence -eq 'False')
    $notInSitemap  = ($datasets['SitemapPresence'].Count -gt 0) -and (-not $tableInAnyApp)
    $deadScore   = 0
    if ($hasNoData)            { $deadScore++ }
    if ($notTouched)           { $deadScore++ }
    if ($notInUI)              { $deadScore++ }
    if ($allManaged -eq $false){ $deadScore++ }
    if ($notInSitemap)         { $deadScore++ }

    $master.Add([PSCustomObject][ordered]@{
        TableLogicalName        = $a.TableLogicalName
        TableDisplayName        = $a.TableDisplayName
        AttributeLogicalName    = $a.AttributeLogicalName
        AttributeDisplayName    = $a.AttributeDisplayName
        AttributeType           = $a.AttributeType
        IsCustomAttribute       = $a.IsCustomAttribute
        LookupTargets           = $a.LookupTargets
        TotalRecords            = $a.TotalRecords
        PopulatedCount          = $a.PopulatedCount
        FillRatePercent         = $a.FillRatePercent
        LastAuditedOn           = if ($aud) { $aud.LastAuditedOn } else { '' }
        DaysSinceLastAudited    = if ($aud) { $aud.DaysSinceLastAudited } else { '' }
        AuditEntriesInWindow    = if ($aud) { $aud.AuditEntriesInWindow } else { '' }
        DistinctUsersInWindow   = if ($aud) { $aud.DistinctUsersInWindow } else { '' }
        DistinctRecordsTouched  = if ($aud) { $aud.DistinctRecordsTouched } else { '' }
        AuditStatus             = if ($aud) { $aud.Status } else { '' }
        TopUserDisplayName      = if ($topUser) { $topUser.UserDisplayName } else { '' }
        TopUserRecordCount      = if ($topUser) { $topUser.RecordCount }     else { '' }
        TopUserIsServiceAccount = if ($topUser) { $topUser.IsServiceAccount } else { '' }
        OnAnyForm               = if ($ui) { $ui.OnAnyForm }     else { '' }
        FormCount               = if ($ui) { $ui.FormCount }     else { '' }
        OnAnyView               = if ($ui) { $ui.OnAnyView }     else { '' }
        ViewCount               = if ($ui) { $ui.ViewCount }     else { '' }
        OnAnyChart              = if ($ui) { $ui.OnAnyChart }    else { '' }
        AnyUIPresence           = if ($ui) { $ui.AnyUIPresence } else { '' }
        TableInAnyAppSitemap    = $tableInAnyApp
        TableAppCount           = $tableAppCount
        TableAppNames           = $tableAppNames
        SolutionsContainingAttr = $solUniqueNames
        SolutionsAllManaged     = $allManaged
        SolutionsAnyCustom      = $anyCustom
        TableNewestModifiedOn   = if ($tu) { $tu.NewestModifiedOn }    else { '' }
        TableDistinctCreators   = if ($tu) { $tu.DistinctCreators }    else { '' }
        TableRecordsLast365Days = if ($tu) { $tu.RecordsCreatedLast365Days } else { '' }
        DeadFieldScore          = $deadScore
    })
}
Write-Host "  Master rows: $($master.Count)" -ForegroundColor Gray

# ---- TABLES (one row per table) ----
$tablesRollup = New-Object System.Collections.Generic.List[object]
$allTableNames = New-Object System.Collections.Generic.HashSet[string]
foreach ($r in $datasets['RecordCounts']) { [void]$allTableNames.Add($r.TableLogicalName.ToLowerInvariant()) }
foreach ($r in $datasets['TableUsage'])   { [void]$allTableNames.Add($r.TableLogicalName.ToLowerInvariant()) }
foreach ($t in ($allTableNames | Sort-Object)) {
    $rc = if ($rcMap.ContainsKey($t)) { $rcMap[$t] } else { $null }
    $tu = if ($tuMap.ContainsKey($t)) { $tuMap[$t] } else { $null }
    $sols = if ($tableSolMap.ContainsKey($t)) { ($tableSolMap[$t] | ForEach-Object { $_.SolutionUniqueName } | Sort-Object -Unique) -join ';' } else { '' }
    $smMatches = if ($tableSitemapMap.ContainsKey($t)) { $tableSitemapMap[$t] } else { @() }
    $appNames  = ($smMatches | ForEach-Object { $_.AppDisplayName } | Where-Object { $_ } | Sort-Object -Unique) -join ';'
    $appCount  = ($smMatches | ForEach-Object { $_.AppUniqueName } | Where-Object { $_ } | Sort-Object -Unique).Count

    # Derived boolean flags. These are friendlier to filter on in Excel / Power BI than
    # re-typing the numeric comparisons in every PivotTable / slicer. Each one is left
    # blank when the underlying data isn't available, so PivotTables don't bucket
    # "unknown" alongside True/False.
    $rcVal = if ($rc) { [string]$rc.RecordCount } else { '' }
    $hasData = if ([string]::IsNullOrWhiteSpace($rcVal)) { '' } else { ([long]$rcVal -gt 0) }
    $dslmVal = if ($rc) { [string]$rc.DaysSinceLastModified } else { '' }
    $isActive = if ([string]::IsNullOrWhiteSpace($dslmVal)) { '' } else { ([int]$dslmVal -le 90) }
    $isStale  = if ([string]::IsNullOrWhiteSpace($dslmVal)) { '' } else { ([int]$dslmVal -gt 180) }

    $tablesRollup.Add([PSCustomObject][ordered]@{
        TableLogicalName            = if ($rc) { $rc.TableLogicalName }   elseif ($tu) { $tu.TableLogicalName }   else { $t }
        TableDisplayName            = if ($rc) { $rc.TableDisplayName }   elseif ($tu) { $tu.TableDisplayName }   else { '' }
        TableSchemaName             = if ($rc) { $rc.SchemaName }         elseif ($tu) { $tu.TableSchemaName }    else { '' }
        TableType                   = if ($rc) { $rc.TableType }          else { '' }
        IsCustomEntity              = if ($rc) { $rc.IsCustomEntity }     else { '' }
        OwnershipType               = if ($tu) { $tu.OwnershipType }      else { '' }
        RecordCount                 = if ($rc) { $rc.RecordCount }        else { '' }
        HasData                     = $hasData
        UsageBucket                 = if ($rc) { $rc.UsageBucket }        else { '' }
        LastModifiedOn              = if ($rc) { $rc.LastModifiedOn }     else { '' }
        DaysSinceLastModified       = if ($rc) { $rc.DaysSinceLastModified } else { '' }
        IsActive                    = $isActive
        IsStale                     = $isStale
        NewestCreatedOn             = if ($tu) { $tu.NewestCreatedOn }    else { '' }
        RecordsCreatedLast30Days    = if ($tu) { $tu.RecordsCreatedLast30Days }  else { '' }
        RecordsCreatedLast90Days    = if ($tu) { $tu.RecordsCreatedLast90Days }  else { '' }
        RecordsCreatedLast365Days   = if ($tu) { $tu.RecordsCreatedLast365Days } else { '' }
        DistinctCreators            = if ($tu) { $tu.DistinctCreators }   else { '' }
        DistinctModifiers           = if ($tu) { $tu.DistinctModifiers }  else { '' }
        DistinctOwners              = if ($tu) { $tu.DistinctOwners }     else { '' }
        InAnyAppSitemap             = ($smMatches.Count -gt 0)
        AppCount                    = $appCount
        AppNames                    = $appNames
        ContainingSolutions         = $sols
    })
}
Write-Host "  Tables rollup rows: $($tablesRollup.Count)" -ForegroundColor Gray

# ---- CLEANUP CANDIDATES ----
$cleanup = $master |
    Where-Object { [int]$_.DeadFieldScore -ge 2 -and $_.IsCustomAttribute -eq 'True' } |
    Sort-Object @{Expression='DeadFieldScore'; Descending=$true}, TableLogicalName, AttributeLogicalName
Write-Host "  Cleanup candidates: $(@($cleanup).Count)" -ForegroundColor Gray

# ---- SUMMARY (single-screen overview) ----
# Rolls up the Tables and Master sheets into the metrics you'd build first in any
# analysis: tables/columns counted, % custom, total record volume, UsageBucket
# distribution, IsActive/IsStale/HasData distribution, cleanup-candidate counts.
# Ordered as a (Section, Metric, Value) table so it renders cleanly in Excel and
# stays trivially extensible.
Write-Host "Building summary.." -ForegroundColor Cyan

# Helpers: count rows whose specified column equals a target; tolerate blanks
function _CountWhere { param($Rows, [string]$Col, $Value)
    @($Rows | Where-Object { $_.$Col -eq $Value }).Count
}
function _SumLong { param($Rows, [string]$Col)
    $total = 0L
    foreach ($r in $Rows) {
        $v = $r.$Col
        if ([string]::IsNullOrWhiteSpace([string]$v)) { continue }
        $n = 0L
        if ([long]::TryParse([string]$v, [ref]$n)) { $total += $n }
    }
    return $total
}

$tableTotal       = $tablesRollup.Count
$tableCustom      = _CountWhere $tablesRollup 'IsCustomEntity' 'True'
$tableSystem      = _CountWhere $tablesRollup 'IsCustomEntity' 'False'
$tablePctCustom   = if ($tableTotal -gt 0) { [math]::Round(($tableCustom / $tableTotal) * 100, 1) } else { 0 }
$totalRecords     = _SumLong    $tablesRollup 'RecordCount'
$tablesWithData   = _CountWhere $tablesRollup 'HasData' 'True'
$tablesEmpty      = _CountWhere $tablesRollup 'HasData' 'False'
$tablesActive     = _CountWhere $tablesRollup 'IsActive' 'True'
$tablesStale      = _CountWhere $tablesRollup 'IsActive' 'False'   # not active = stale or unknown
$tablesStrictStale= _CountWhere $tablesRollup 'IsStale'  'True'
$tablesInApp      = _CountWhere $tablesRollup 'InAnyAppSitemap' 'True'

# UsageBucket distribution (preserve a meaningful order)
$bucketOrder = @('Active (<=90d)','Dormant (91-365d)','Stale (>365d)','Empty','Unknown','Unsupported')
$bucketRows  = @{}
foreach ($t in $tablesRollup) {
    $b = if ([string]::IsNullOrWhiteSpace([string]$t.UsageBucket)) { '(blank)' } else { [string]$t.UsageBucket }
    if (-not $bucketRows.ContainsKey($b)) { $bucketRows[$b] = 0 }
    $bucketRows[$b]++
}
# Make sure every well-known bucket appears even when its count is 0
foreach ($b in $bucketOrder) { if (-not $bucketRows.ContainsKey($b)) { $bucketRows[$b] = 0 } }

# Master / attribute roll-up
$attrTotal     = $master.Count
$attrCustom    = _CountWhere $master 'IsCustomAttribute' 'True'
$attrPctCustom = if ($attrTotal -gt 0) { [math]::Round(($attrCustom / $attrTotal) * 100, 1) } else { 0 }
$attrZeroFill  = @($master | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.FillRatePercent) -and [double]$_.FillRatePercent -eq 0 }).Count
$attrFullFill  = @($master | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.FillRatePercent) -and [double]$_.FillRatePercent -eq 100 }).Count
$cleanupCount  = @($cleanup).Count
$dead3Plus     = @($master | Where-Object { [int]$_.DeadFieldScore -ge 3 -and $_.IsCustomAttribute -eq 'True' }).Count

$summary = New-Object System.Collections.Generic.List[object]
function _AddSummary { param([string]$Section, [string]$Metric, $Value, [string]$Notes='')
    $summary.Add([PSCustomObject][ordered]@{
        Section = $Section
        Metric  = $Metric
        Value   = $Value
        Notes   = $Notes
    }) | Out-Null
}

_AddSummary 'Overview' 'Generated'                  (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')         'Local time of this workbook build'
_AddSummary 'Overview' 'Source folder'              (Split-Path -Leaf $InputFolder)                  ''

_AddSummary 'Tables'   'Total tables'                $tableTotal                                     ''
_AddSummary 'Tables'   'Custom tables'               $tableCustom                                    "$tablePctCustom% of total"
_AddSummary 'Tables'   'System tables'               $tableSystem                                    ''
_AddSummary 'Tables'   'Tables with data (HasData)'  $tablesWithData                                 ''
_AddSummary 'Tables'   'Empty tables'                $tablesEmpty                                    'RecordCount = 0'
_AddSummary 'Tables'   'Active (<=90d)'              $tablesActive                                   'IsActive = True'
_AddSummary 'Tables'   'Stale (>180d)'               $tablesStrictStale                              'IsStale = True'
_AddSummary 'Tables'   'Surfaced in any app sitemap' $tablesInApp                                    'InAnyAppSitemap = True'
_AddSummary 'Tables'   'Total records (sum)'         $totalRecords                                   'Sum of RecordCount across all Success rows'

_AddSummary 'UsageBucket' 'Active (<=90d)'           $bucketRows['Active (<=90d)']                   ''
_AddSummary 'UsageBucket' 'Dormant (91-365d)'        $bucketRows['Dormant (91-365d)']                ''
_AddSummary 'UsageBucket' 'Stale (>365d)'            $bucketRows['Stale (>365d)']                    ''
_AddSummary 'UsageBucket' 'Empty'                    $bucketRows['Empty']                            ''
_AddSummary 'UsageBucket' 'Unknown'                  $bucketRows['Unknown']                          ''
_AddSummary 'UsageBucket' 'Unsupported'              $bucketRows['Unsupported']                      'Virtual/Elastic - skipped by RetrieveTotalRecordCount'
foreach ($b in ($bucketRows.Keys | Where-Object { $_ -notin $bucketOrder } | Sort-Object)) {
    _AddSummary 'UsageBucket' $b $bucketRows[$b] ''
}

_AddSummary 'Attributes' 'Total attributes scanned'  $attrTotal                                      ''
_AddSummary 'Attributes' 'Custom attributes'         $attrCustom                                     "$attrPctCustom% of total"
_AddSummary 'Attributes' 'Always populated (100%)'   $attrFullFill                                   ''
_AddSummary 'Attributes' 'Never populated (0%)'      $attrZeroFill                                   'Strong dead-field signal'

_AddSummary 'Cleanup'    'Cleanup candidates'        $cleanupCount                                   'DeadFieldScore >= 2 AND IsCustomAttribute'
_AddSummary 'Cleanup'    'High-confidence dead'      $dead3Plus                                      'DeadFieldScore >= 3 AND IsCustomAttribute'
Write-Host "  Summary metrics: $($summary.Count)" -ForegroundColor Gray

# ---- WRITE OUTPUTS ----
$masterPath     = Join-Path $OutputFolder 'master.csv'
$tablesPath     = Join-Path $OutputFolder 'tables.csv'
$cleanupPath    = Join-Path $OutputFolder 'cleanup.csv'
$summaryPath    = Join-Path $OutputFolder 'summary.csv'
$dictionaryPath = Join-Path $OutputFolder 'dictionary.csv'
$readmePath     = Join-Path $OutputFolder 'README.md'

Write-Host "Writing computed CSVs..." -ForegroundColor Cyan
$master       | Export-Csv -Path $masterPath  -NoTypeInformation
$tablesRollup | Export-Csv -Path $tablesPath  -NoTypeInformation
@($cleanup)   | Export-Csv -Path $cleanupPath -NoTypeInformation
$summary      | Export-Csv -Path $summaryPath -NoTypeInformation

# ---- COLUMN DICTIONARY (one row per column in master.csv / tables.csv) ----
# Pure human-readable definitions of every output column so anyone opening the
# workbook cold can interpret the data without re-reading the source scripts.
$dictionary = @(
    # ---- master.csv columns ----
    [PSCustomObject]@{ Sheet='Master'; Column='TableLogicalName';        Source='attributeusage'; Description='Dataverse logical name of the parent table (e.g. msf_servicerequest). Lowercase. Composite join key with AttributeLogicalName.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableDisplayName';        Source='attributeusage'; Description='Friendly display label for the table (e.g. "Service Request"). Localized to the org default language.' }
    [PSCustomObject]@{ Sheet='Master'; Column='AttributeLogicalName';    Source='attributeusage'; Description='Dataverse logical name of the column (e.g. msf_priority). Lowercase. Composite join key with TableLogicalName.' }
    [PSCustomObject]@{ Sheet='Master'; Column='AttributeDisplayName';    Source='attributeusage'; Description='Friendly display label for the column.' }
    [PSCustomObject]@{ Sheet='Master'; Column='AttributeType';           Source='attributeusage'; Description='Dataverse type code: String, Memo, Picklist, Lookup, DateTime, Boolean, Money, Integer, Decimal, Customer, Owner, etc.' }
    [PSCustomObject]@{ Sheet='Master'; Column='IsCustomAttribute';       Source='attributeusage'; Description='True/False. True for custom (publisher-prefixed) columns; False for system columns shipped by Microsoft.' }
    [PSCustomObject]@{ Sheet='Master'; Column='LookupTargets';           Source='attributeusage'; Description='For Lookup-type columns: semicolon-separated list of target table logical names (e.g. "systemuser;team"). Empty for non-lookups.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TotalRecords';            Source='attributeusage'; Description='Total number of records in the table at the time of the scan (the denominator for FillRatePercent). Blank when the row''s Status != Success.' }
    [PSCustomObject]@{ Sheet='Master'; Column='PopulatedCount';          Source='attributeusage'; Description='Number of records where this column has a non-null value. Blank when the row''s Status != Success.' }
    [PSCustomObject]@{ Sheet='Master'; Column='FillRatePercent';         Source='attributeusage'; Description='PopulatedCount / TotalRecords * 100. 0 = field is never populated; 100 = always populated. Blank when the row''s Status != Success.' }
    [PSCustomObject]@{ Sheet='Master'; Column='LastAuditedOn';           Source='audithistory';   Description='Most recent createdon timestamp of any audit row that touched this column. Empty if no audit data in the window.' }
    [PSCustomObject]@{ Sheet='Master'; Column='DaysSinceLastAudited';    Source='audithistory';   Description='Whole days from "now" (UTC) back to LastAuditedOn. Empty if no audit data.' }
    [PSCustomObject]@{ Sheet='Master'; Column='AuditEntriesInWindow';    Source='audithistory';   Description='Total Create+Update audit events touching this column within the configured -DaysBack window.' }
    [PSCustomObject]@{ Sheet='Master'; Column='DistinctUsersInWindow';   Source='audithistory';   Description='Number of distinct users who triggered an audit event on this column. 1 = single-owner risk; high = broad usage.' }
    [PSCustomObject]@{ Sheet='Master'; Column='DistinctRecordsTouched';  Source='audithistory';   Description='Number of distinct records on which this column was modified. Distinguishes "5 records edited 100x" from "500 records edited once".' }
    [PSCustomObject]@{ Sheet='Master'; Column='AuditStatus';             Source='audithistory';   Description='Per-attribute audit status: Success / NoAuditData / AuditDisabledForOrg / AuditDisabledForTable / AuditDisabledForColumn. Tells you whether "no entries" means "field unused" or "we cannot tell because auditing is off".' }
    [PSCustomObject]@{ Sheet='Master'; Column='TopUserDisplayName';      Source='useractivity';   Description='Display name of the #1 user (by record count) for ANY user lookup on this attribute. Excludes the "(no value)" bucket. Useful for finding process owners.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TopUserRecordCount';      Source='useractivity';   Description='How many records the TopUser is the user-lookup value on.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TopUserIsServiceAccount'; Source='useractivity';   Description='True/False. True when the top user is a non-interactive (accessmode=4) systemuser, e.g. "# DataverseSync". Filter these out to focus on real humans.' }
    [PSCustomObject]@{ Sheet='Master'; Column='OnAnyForm';               Source='uipresence';     Description='True/False. True if this column appears on at least one Main / QuickCreate / QuickView / etc. form.' }
    [PSCustomObject]@{ Sheet='Master'; Column='FormCount';               Source='uipresence';     Description='Number of distinct forms (Main + QuickCreate + QuickView) that reference this column.' }
    [PSCustomObject]@{ Sheet='Master'; Column='OnAnyView';               Source='uipresence';     Description='True/False. True if this column appears on at least one system view (savedquery).' }
    [PSCustomObject]@{ Sheet='Master'; Column='ViewCount';               Source='uipresence';     Description='Number of distinct system views that reference this column.' }
    [PSCustomObject]@{ Sheet='Master'; Column='OnAnyChart';              Source='uipresence';     Description='True/False. True if this column appears in at least one chart datadescription.' }
    [PSCustomObject]@{ Sheet='Master'; Column='AnyUIPresence';           Source='uipresence';     Description='True/False composite. True when ANY of OnAnyForm / OnAnyView / OnAnyChart is True. False = the column is invisible everywhere in the UI - the strongest "dead field" signal.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableInAnyAppSitemap';    Source='sitemappresence'; Description='True/False. True when this column''s parent TABLE is surfaced in at least one model-driven app sitemap (Area/Group/SubArea Entity binding). False = no app exposes the table to end users - automation-only / hidden table.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableAppCount';           Source='sitemappresence'; Description='How many distinct model-driven apps surface the parent table.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableAppNames';           Source='sitemappresence'; Description='Semicolon-separated friendly names of the apps surfacing the parent table.' }
    [PSCustomObject]@{ Sheet='Master'; Column='SolutionsContainingAttr'; Source='solutionmembership'; Description='Semicolon-separated unique-names of every solution that ships this column as an Attribute component.' }
    [PSCustomObject]@{ Sheet='Master'; Column='SolutionsAllManaged';     Source='solutionmembership'; Description='True if EVERY containing solution is managed (you cannot delete the column directly). False if at least one is unmanaged. Empty if not in any solution.' }
    [PSCustomObject]@{ Sheet='Master'; Column='SolutionsAnyCustom';      Source='solutionmembership'; Description='True if AT LEAST ONE containing solution component has IsCustomComponent = True. Indicates customer-built vs Microsoft-shipped attribute.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableNewestModifiedOn';   Source='tableusage';     Description='Most recent modifiedon across the entire table (not just this column). Helps you tell "abandoned table" from "active table with abandoned columns".' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableDistinctCreators';   Source='tableusage';     Description='Distinct count of createdby users across the table. 1 = single-owner table.' }
    [PSCustomObject]@{ Sheet='Master'; Column='TableRecordsLast365Days'; Source='tableusage';     Description='How many records on this table were CREATED in the last 365 days. 0 = no new records in a year (likely abandoned table).' }
    [PSCustomObject]@{ Sheet='Master'; Column='DeadFieldScore';          Source='computed';       Description='0-5 composite. Adds 1 for each of: FillRatePercent=0; no audit events in window; AnyUIPresence=False; all containing solutions are unmanaged; parent table not in any model-driven app sitemap. A custom attribute scoring 3+ is a strong cleanup candidate. Cleanup sheet auto-filters to score >= 2.' }

    # ---- tables.csv columns ----
    [PSCustomObject]@{ Sheet='Tables'; Column='TableLogicalName';          Source='recordcounts';   Description='Dataverse logical name of the table.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='TableDisplayName';          Source='recordcounts';   Description='Friendly display label.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='TableSchemaName';           Source='recordcounts';   Description='PascalCase schema name (e.g. msf_ServiceRequest). Used in code-gen and solution exports.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='TableType';                 Source='recordcounts';   Description='Standard / Virtual / Elastic. Virtual + Elastic tables are pre-skipped from RetrieveTotalRecordCount because the API does not support them.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='IsCustomEntity';            Source='recordcounts';   Description='True/False. True for custom (publisher-prefixed) tables; False for system tables shipped by Microsoft.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='OwnershipType';             Source='tableusage';     Description='UserOwned / TeamOwned / OrganizationOwned / BusinessOwned / etc. Determines whether ownerid lookup applies. Blank for system / no-owner tables (Dataverse OwnershipType=None).' }
    [PSCustomObject]@{ Sheet='Tables'; Column='RecordCount';               Source='recordcounts';   Description='Approximate row count from RetrieveTotalRecordCount (system-index based; may be slightly stale on test envs). Blank when Status != Success (Skipped / Error / Stats Not Available); see the source RecordCounts sheet Status column for the reason.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='HasData';                   Source='computed';       Description='True/False derived from RecordCount > 0. Blank when RecordCount is unavailable. Use as a quick filter for "tables with at least one row" without re-typing comparisons in every PivotTable.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='UsageBucket';               Source='recordcounts';   Description='Bucket label: Empty / Active (<=90d) / Dormant (91-365d) / Stale (>365d) / Unsupported / Unknown - based on the most recent activity timestamp.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='LastModifiedOn';            Source='recordcounts';   Description='Most recent modifiedon timestamp from the activity probe. Empty if no records or probe was skipped.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='DaysSinceLastModified';     Source='recordcounts';   Description='Whole days since LastModifiedOn (UTC).' }
    [PSCustomObject]@{ Sheet='Tables'; Column='IsActive';                  Source='computed';       Description='True/False derived from DaysSinceLastModified <= 90. Identifies tables with recent write activity. Blank when DaysSinceLastModified is unavailable.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='IsStale';                   Source='computed';       Description='True/False derived from DaysSinceLastModified > 180. Identifies tables that have not been written to in over six months. Blank when DaysSinceLastModified is unavailable. Note: a table can be both "not IsActive" and "not IsStale" when activity falls between 91 and 180 days.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='NewestCreatedOn';           Source='tableusage';     Description='Most recent createdon - tells you whether new records are still being created.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='RecordsCreatedLast30Days';  Source='tableusage';     Description='Count of records with createdon in the last 30 days.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='RecordsCreatedLast90Days';  Source='tableusage';     Description='Count of records with createdon in the last 90 days.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='RecordsCreatedLast365Days'; Source='tableusage';     Description='Count of records with createdon in the last 365 days. Trend signal alongside the raw record count.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='DistinctCreators';          Source='tableusage';     Description='Distinct createdby user count. 1 = automation-only; high = broad org usage.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='DistinctModifiers';         Source='tableusage';     Description='Distinct modifiedby user count.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='DistinctOwners';            Source='tableusage';     Description='Distinct ownerid user count. Empty for organization-owned tables (no per-record owner).' }
    [PSCustomObject]@{ Sheet='Tables'; Column='InAnyAppSitemap';           Source='sitemappresence'; Description='True/False. True when the table is bound to an Entity SubArea in at least one model-driven app sitemap. False = the table is not user-facing in any app - strong indicator of automation-only / hidden table.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='AppCount';                  Source='sitemappresence'; Description='How many distinct model-driven apps surface this table.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='AppNames';                  Source='sitemappresence'; Description='Semicolon-separated friendly names of the apps surfacing this table.' }
    [PSCustomObject]@{ Sheet='Tables'; Column='ContainingSolutions';       Source='solutionmembership'; Description='Semicolon-separated unique-names of every solution that ships this table as an Entity component.' }

    # ---- SitemapPresence (raw drill-down sheet) columns ----
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='AppUniqueName';     Source='sitemappresence'; Description='Unique name of the model-driven app (appmodule.uniquename).' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='AppDisplayName';    Source='sitemappresence'; Description='Friendly display name of the app (appmodule.name).' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='AppId';             Source='sitemappresence'; Description='appmoduleid GUID.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='SitemapName';       Source='sitemappresence'; Description='Name of the sitemap record associated with the app.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='AreaId';            Source='sitemappresence'; Description='Sitemap Area Id attribute.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='AreaTitle';         Source='sitemappresence'; Description='Sitemap Area title (LCID 1033 preferred).' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='GroupId';           Source='sitemappresence'; Description='Sitemap Group Id attribute.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='GroupTitle';        Source='sitemappresence'; Description='Sitemap Group title (LCID 1033 preferred).' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='SubAreaId';         Source='sitemappresence'; Description='Sitemap SubArea Id attribute.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='SubAreaTitle';      Source='sitemappresence'; Description='Sitemap SubArea title (LCID 1033 preferred).' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='SubAreaType';       Source='sitemappresence'; Description='Entity / Dashboard / Url / WebResource / Unknown - what kind of tab the SubArea represents. Only Entity rows have a TableLogicalName.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='TableLogicalName';  Source='sitemappresence'; Description='Logical name of the table bound to this SubArea (Entity attribute on the SubArea node). Empty for non-Entity tabs. Lowercase. Joins to other sheets.' }
    [PSCustomObject]@{ Sheet='SitemapPresence'; Column='Url';               Source='sitemappresence'; Description='URL the SubArea opens (only set for non-Entity tabs - dashboards, custom URLs, web resources).' }

    # ---- Summary sheet columns ----
    [PSCustomObject]@{ Sheet='Summary'; Column='Section'; Source='computed'; Description='Grouping label for the metric (Overview / Tables / UsageBucket / Attributes / Cleanup). Use to filter or PivotTable.' }
    [PSCustomObject]@{ Sheet='Summary'; Column='Metric';  Source='computed'; Description='Human-readable name of the metric (e.g. "Empty tables", "Active (<=90d)").' }
    [PSCustomObject]@{ Sheet='Summary'; Column='Value';   Source='computed'; Description='The value for the metric. Numeric for counts / sums; date string for "Generated"; folder name for "Source folder".' }
    [PSCustomObject]@{ Sheet='Summary'; Column='Notes';   Source='computed'; Description='Optional context (e.g. percent of total, the underlying filter expression, or what the metric excludes).' }
)
$dictionary | Export-Csv -Path $dictionaryPath -NoTypeInformation

# README -------------------------------------------------------------------
$readme = @"
# Dataverse Usage Report

Generated by ``Build-UsageReportWorkbook.ps1`` on top of CSVs produced by
``Invoke-DataverseUsageReport.ps1``.

## Files in this folder

### Generated join files (this script)
| File | What it is |
|---|---|
| ``summary.csv``    | Single-screen overview: counts per UsageBucket, empty/active/stale tables, total records, % custom entities, cleanup-candidate counts. Three columns: Section / Metric / Value (+ Notes). |
| ``master.csv``     | Full per-attribute join. Spine = attributeusage, with audit / UI / top-user / solution data left-joined on (TableLogicalName, AttributeLogicalName). One row per attribute. |
| ``tables.csv``     | Per-table roll-up joining recordcounts + tableusage on TableLogicalName. One row per table. |
| ``dictionary.csv`` | Column glossary for master.csv and tables.csv. One row per output column with Sheet / Column / Source / Description. |
| ``cleanup.csv``    | Pre-filtered Master view: ``DeadFieldScore >= 2 AND IsCustomAttribute = True``. Sorted by DeadFieldScore desc. |

### Source CSVs (from Invoke-DataverseUsageReport.ps1)
| File | Source script |
|---|---|
| ``recordcounts_*.csv``       | ``GetRecordCountByTable.ps1`` |
| ``relationships_*.csv``      | ``GetTableRelationships.ps1`` |
| ``solutionmembership_*.csv`` | ``GetSolutionMembership.ps1`` |
| ``sitemappresence_*.csv``    | ``GetSitemapEntityPresence.ps1`` |
| ``tableusage_*.csv``         | ``GetTableUsageActivity.ps1`` |
| ``attributeusage_*.csv``     | ``GetFieldFillRateByTable.ps1`` |
| ``uipresence_*.csv``         | ``GetFieldUIPresence.ps1`` |
| ``useractivity_*.csv``       | ``GetUserActivityByTable.ps1`` |
| ``audithistory_*.csv``       | ``GetAttributeAuditHistory.ps1`` |

## DeadFieldScore (0-5)

Composite signal in ``master.csv``. Adds 1 for each of:

1. ``FillRatePercent = 0``                            - the field has no data
2. No audit events found in the audit window          - nothing is touching it
3. ``AnyUIPresence = False``                          - field is not on any form, view, or chart
4. All containing solutions are unmanaged             - you can actually delete it
5. Parent table not in any model-driven app sitemap   - the table isn't user-facing at all

A custom attribute scoring 3+ is a strong cleanup candidate. ``cleanup.csv`` shows
everything scoring 2+.

## Recommended Excel workflow

1. Open ``master.csv`` in Excel.
2. Convert to a Table (Ctrl+T) if not already auto-detected.
3. Sort / filter on ``DeadFieldScore`` to triage cleanup.
4. PivotTable on ``TopUserDisplayName`` to find your most-active and single-owner contributors.
5. PivotTable on ``LookupTargets`` to see which target tables are pointed at most.

If you have **Copilot in Excel**, open it on the Master sheet and ask things like:
- "Which lookup attributes have FillRate=0 and a TopUserDisplayName containing '#' (service accounts)?"
- "Group by TableLogicalName and show average FillRatePercent."
- "Show me attributes where DistinctUsersInWindow is 1 - these are single-owner risks."

## Refresh

Re-run ``Invoke-DataverseUsageReport.ps1`` to produce a new timestamped folder of CSVs,
then re-run this script against it. summary/master/tables/cleanup are rebuilt from scratch.
"@
Set-Content -Path $readmePath -Value $readme -Encoding UTF8

Write-Host "Generated:" -ForegroundColor Green
Write-Host "  $summaryPath    ($($summary.Count) metrics)"
Write-Host "  $masterPath     ($($master.Count) rows)"
Write-Host "  $tablesPath     ($($tablesRollup.Count) rows)"
Write-Host "  $cleanupPath    ($(@($cleanup).Count) rows)"
Write-Host "  $dictionaryPath ($($dictionary.Count) column definitions)"
Write-Host "  $readmePath"

# OPTIONAL: combine into single .xlsx via Excel COM ------------------------
$xlsxPath = $null
if ($CombineToXlsx) {
    Write-Host "`nLaunching Excel to combine into a single .xlsx workbook..." -ForegroundColor Cyan
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    }
    catch {
        Write-Warning "Excel COM not available on this machine. Skipping .xlsx generation."
        Write-Warning "  ($($_.Exception.Message))"
        Write-Host "The CSVs above are still ready for analysis - just open them in Excel manually." -ForegroundColor Yellow
    }

    if ($excel) {
        $previousSheetsInNewWorkbook = $excel.SheetsInNewWorkbook
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        try {
            # Sheet build order: README first, then computed joins, then raw CSVs.
            # Plain [object[]] would force us to use += which can mis-fire as op_Addition
            # against a hashtable; ArrayList.Add returns an int (still fine) and avoids
            # generic-type-literal parsing quirks in PS 7.
            $sheetSpecs = New-Object System.Collections.ArrayList
            [void]$sheetSpecs.Add(@{ Name='Summary';    File=$summaryPath;    IsMarkdown=$false })
            [void]$sheetSpecs.Add(@{ Name='README';     File=$readmePath;     IsMarkdown=$true })
            [void]$sheetSpecs.Add(@{ Name='Master';     File=$masterPath;     IsMarkdown=$false })
            [void]$sheetSpecs.Add(@{ Name='Tables';     File=$tablesPath;     IsMarkdown=$false })
            [void]$sheetSpecs.Add(@{ Name='Dictionary'; File=$dictionaryPath; IsMarkdown=$false })
            [void]$sheetSpecs.Add(@{ Name='Cleanup';    File=$cleanupPath;    IsMarkdown=$false })
            foreach ($name in $reportPatterns.Keys) {
                if ($datasetFiles[$name]) {
                    [void]$sheetSpecs.Add(@{ Name=$name; File=$datasetFiles[$name]; IsMarkdown=$false })
                }
            }

            # Pre-allocate every sheet in one go via SheetsInNewWorkbook. This sidesteps
            # all the failure modes of Worksheets.Add / Worksheets.Move under PS 7 COM
            # interop ("op_Addition", "Unable to get the Move property", etc.) by simply
            # never calling Add or Move at all.
            $sheetCount = $sheetSpecs.Count
            $excel.SheetsInNewWorkbook = $sheetCount
            $wb = $excel.Workbooks.Add()

            # Defensive: if Excel didn't honor SheetsInNewWorkbook (rare, but possible if a
            # template overrides it), append/remove sheets manually one at a time.
            while ($wb.Worksheets.Count -lt $sheetCount) {
                $tail = $wb.Worksheets.Item($wb.Worksheets.Count)
                [void]$wb.Worksheets.Add([System.Reflection.Missing]::Value, $tail)
            }
            while ($wb.Worksheets.Count -gt $sheetCount) {
                $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
            }

            # Name each pre-allocated sheet in order
            for ($i = 0; $i -lt $sheetCount; $i++) {
                $wb.Worksheets.Item($i + 1).Name = $sheetSpecs[$i].Name
            }
            Write-Host "  Allocated $sheetCount sheets." -ForegroundColor DarkGray

            # Populate each sheet
            $buildSw = [System.Diagnostics.Stopwatch]::StartNew()
            for ($i = 0; $i -lt $sheetCount; $i++) {
                $spec  = $sheetSpecs[$i]
                $sheet = $wb.Worksheets.Item($i + 1)
                $sheetSw = [System.Diagnostics.Stopwatch]::StartNew()

                if ($spec.IsMarkdown) {
                    $lines = @(Get-Content $spec.File)
                    Write-Host ("  [{0}/{1}] {2,-22} README markdown ({3} lines)..." -f ($i + 1), $sheetCount, $spec.Name, $lines.Count) -ForegroundColor Cyan
                    for ($j = 0; $j -lt $lines.Count; $j++) {
                        $sheet.Cells.Item([int]($j + 1), 1).Value2 = [string]$lines[$j]
                    }
                    $sheet.Columns.Item(1).ColumnWidth = 120
                    $sheet.Columns.Item(1).WrapText    = $false
                    $sheetSw.Stop()
                    Write-Host ("        done in {0:n1}s" -f $sheetSw.Elapsed.TotalSeconds) -ForegroundColor DarkGray
                }
                else {
                    $rows = @(Import-Csv $spec.File)
                    if ($rows.Count -eq 0) {
                        $sheet.Cells.Item(1, 1).Value2 = "(empty)"
                        $sheetSw.Stop()
                        Write-Host ("  [{0}/{1}] {2,-22} (empty CSV)" -f ($i + 1), $sheetCount, $spec.Name) -ForegroundColor DarkGray
                        continue
                    }
                    $headers  = @($rows[0].PSObject.Properties.Name)
                    $colCount = $headers.Count
                    $rowCount = $rows.Count + 1   # +1 for header
                    $totalCells = $rows.Count * $colCount
                    Write-Host ("  [{0}/{1}] {2,-22} {3,6:n0} rows x {4,3} cols ({5,7:n0} cells)..." -f ($i + 1), $sheetCount, $spec.Name, $rows.Count, $colCount, $totalCells) -ForegroundColor Cyan

                    # Helper: 1->A, 26->Z, 27->AA, etc. (Excel column-letter math).
                    function _ColLetter([int]$n) {
                        $s = ''
                        while ($n -gt 0) {
                            $rem = ($n - 1) % 26
                            $s = [char](65 + $rem) + $s
                            $n = [int][Math]::Floor(($n - 1) / 26)
                        }
                        return $s
                    }
                    $lastColLetter = _ColLetter $colCount

                    # Classify each column by header-name pattern so we can (a) write the
                    # cell value as the right native Excel type (number / date), and
                    # (b) apply a friendly NumberFormat to the whole column afterwards.
                    # Storing strings would force every analyst to re-type the column in
                    # Excel before SUM/AVG/sort would behave correctly.
                    #
                    # Pattern rules (header name):
                    #   Percent$                                  -> percent  (e.g. FillRatePercent)
                    #   (On|Date)$                                -> date     (e.g. LastModifiedOn, NewestCreatedOn)
                    #   Count|Records|Rank|Score|Days|Distinct[A-Z]|Entries  -> int
                    #   anything else                             -> text
                    $colTypes = New-Object 'string[]' $colCount
                    for ($c = 0; $c -lt $colCount; $c++) {
                        $h = $headers[$c]
                        $colTypes[$c] = if     ($h -match 'Percent$')                                          { 'percent' }
                                        elseif ($h -match '(On|Date)$')                                        { 'date' }
                                        elseif ($h -match '(Count|Records|Rank|Score|Days|Distinct[A-Z]|Entries)') { 'int' }
                                        else                                                                   { 'text' }
                    }

                    # Per-cell write via Cells.Item(row, col).Value2 = [string].
                    # We tried two faster paths first - both fail under PS 7 + Excel COM:
                    #   1) Object[,] -> Range.Value2  : "Unable to cast Object[,] to String"
                    #   2) Per-row 1D Object[] wrapped as ', $rowVals' -> same error
                    # The IDispatch resolver mis-picks a String-typed setter overload no
                    # matter how the SAFEARRAY is shaped. Per-cell assignment is the only
                    # reliable path under PS 7. ScreenUpdating/Calculation are already off
                    # so wall-clock cost is acceptable for typical report sizes.
                    #
                    # Even per-cell assignment via $cell.Value2 = $val hits a related
                    # PSAdapter caching bug: PowerShell caches the Value2 setter
                    # signature on first use (typically String, from header writes),
                    # then fails with "Unable to cast Double/Int64 to String" on every
                    # subsequent numeric assignment. The reliable workaround is to
                    # bypass the PSAdapter cache by going through Type.InvokeMember
                    # directly, which uses fresh IDispatch resolution per call.
                    $excel.Calculation = -4135   # xlCalculationManual

                    $setProp = [System.Reflection.BindingFlags]::SetProperty
                    $comType = [System.__ComObject]
                    function _SetCellValue {
                        param($Cell, $Value)
                        # IMPORTANT: pass the value as a single-element [object[]] - this
                        # forces InvokeMember to do fresh IDispatch dispatch and the
                        # COM marshaller picks the correct VARIANT subtype based on the
                        # boxed value's runtime type.
                        [void]$comType.InvokeMember('Value2', $setProp, $null, $Cell, ([object[]]@($Value)))
                    }

                    # Header row
                    for ($c = 0; $c -lt $colCount; $c++) {
                        _SetCellValue $sheet.Cells.Item(1, $c + 1) ([string]$headers[$c])
                    }

                    # Data rows. Print a progress line periodically so the user can see
                    # the script is still alive on big sheets (per-cell COM writes are
                    # unavoidably slow - this is the only reliable path under PS 7).
                    $progressEvery = [Math]::Max(50, [int]($rows.Count / 20))
                    $rowSw = [System.Diagnostics.Stopwatch]::StartNew()
                    for ($r = 0; $r -lt $rows.Count; $r++) {
                        $row = $rows[$r]
                        $excelRow = $r + 2
                        for ($c = 0; $c -lt $colCount; $c++) {
                            $val = $row.($headers[$c])
                            $cell = $sheet.Cells.Item($excelRow, $c + 1)
                            if ($null -eq $val -or [string]::IsNullOrWhiteSpace([string]$val)) {
                                # Leave blank cells truly blank so AVERAGE/COUNT ignore them.
                                _SetCellValue $cell ''
                                continue
                            }
                            switch ($colTypes[$c]) {
                                'int' {
                                    # Excel COM Range.Value2 setter accepts VT_I4 / VT_R8 / String
                                    # but NOT VT_I8 - the IDispatch resolver mis-picks the String
                                    # setter and throws "Unable to cast Int64 to String". Cast to
                                    # [double] to land on VT_R8 (safe up to 2^53). Display formatting
                                    # is applied per-column further down (#,##0).
                                    $n = 0L
                                    if ([long]::TryParse([string]$val, [ref]$n)) { _SetCellValue $cell ([double]$n) }
                                    else { _SetCellValue $cell ([string]$val) }
                                }
                                'percent' {
                                    # Stored as the 0-100 raw number; the column NumberFormat
                                    # below adds a literal "%" suffix without dividing.
                                    $d = 0.0
                                    if ([double]::TryParse([string]$val, [ref]$d)) { _SetCellValue $cell $d }
                                    else { _SetCellValue $cell ([string]$val) }
                                }
                                'date' {
                                    $dt = [datetime]::MinValue
                                    if ([datetime]::TryParse([string]$val, [ref]$dt)) {
                                        # Excel stores dates as OADate doubles; Value2 expects
                                        # the numeric representation, NumberFormat handles display.
                                        _SetCellValue $cell ($dt.ToOADate())
                                    }
                                    else { _SetCellValue $cell ([string]$val) }
                                }
                                default {
                                    # Mixed-type columns (e.g. Summary's Value column carries
                                    # counts on most rows + a date string + a folder name). Auto-promote
                                    # cells that parse cleanly as a long so SUM/AVERAGE work and the
                                    # cell can wear a thousands-separator format. Leaves text alone.
                                    # Cast to [double] (not [long]) - see the 'int' branch comment
                                    # for why VT_I8 fails through the IDispatch String setter.
                                    $sval = [string]$val
                                    $n = 0L
                                    if ($sval -match '^-?\d{1,18}$' -and [long]::TryParse($sval, [ref]$n)) {
                                        _SetCellValue $cell ([double]$n)
                                        $cell.NumberFormat = '#,##0'
                                    } else {
                                        _SetCellValue $cell $sval
                                    }
                                }
                            }
                        }
                        if ((($r + 1) % $progressEvery) -eq 0) {
                            $pct = [int]((($r + 1) / $rows.Count) * 100)
                            $rate = if ($rowSw.Elapsed.TotalSeconds -gt 0) { ($r + 1) / $rowSw.Elapsed.TotalSeconds } else { 0 }
                            $remaining = if ($rate -gt 0) { ($rows.Count - ($r + 1)) / $rate } else { 0 }
                            Write-Host ("        {0,5:n0}/{1,-5:n0} rows ({2,3}%, {3,5:n0} rows/sec, ~{4,3:n0}s left)" -f ($r + 1), $rows.Count, $pct, $rate, $remaining) -ForegroundColor DarkGray
                        }
                    }
                    $rowSw.Stop()

                    $excel.Calculation = -4105   # xlCalculationAutomatic

                    # Apply column-level NumberFormat. Skip the header row (start at row 2).
                    # Without this, every column reads as "General" - 135023 instead of
                    # 135,023 and ISO timestamps instead of friendly dates.
                    for ($c = 0; $c -lt $colCount; $c++) {
                        $fmt = switch ($colTypes[$c]) {
                            'int'     { '#,##0' }
                            'percent' { '0.0"%"' }    # literal %; underlying number stays 0-100
                            'date'    { 'mm/dd/yyyy hh:mm' }
                            default   { $null }
                        }
                        if ($fmt) {
                            $colLetter = _ColLetter ($c + 1)
                            $sheet.Range("${colLetter}2:${colLetter}${rowCount}").NumberFormat = $fmt
                        }
                    }

                    # Re-grab the full range as a single object for ListObjects.Add below
                    $range = $sheet.Range("A1:${lastColLetter}${rowCount}")

                    # Make it a real Excel Table for AutoFilter + Copilot recognition.
                    # Excel table names can't contain dots, spaces or start with a digit -
                    # sanitize. Errors here are non-fatal; sheet is still usable as a range.
                    Write-Host "        promoting to Excel Table + AutoFit..." -ForegroundColor DarkGray
                    try {
                        $tblName = "$($spec.Name)_tbl" -replace '[^A-Za-z0-9_]', '_'
                        if ($tblName -match '^\d') { $tblName = "_$tblName" }
                        [void]$sheet.ListObjects.Add(1, $range, [System.Reflection.Missing]::Value, 1)   # 1 = xlSrcRange, 1 = xlYes
                        $sheet.ListObjects.Item(1).Name = $tblName
                    }
                    catch {
                        Write-Warning "Could not promote $($spec.Name) range to Excel Table (cosmetic only): $($_.Exception.Message)"
                    }
                    $sheet.Columns.AutoFit() | Out-Null
                    $sheetSw.Stop()
                    Write-Host ("        done in {0:n1}s" -f $sheetSw.Elapsed.TotalSeconds) -ForegroundColor DarkGray
                }
            }
            $buildSw.Stop()
            Write-Host ("  All sheets populated in {0:n1}s." -f $buildSw.Elapsed.TotalSeconds) -ForegroundColor DarkGray

            $xlsxPath = Join-Path $OutputFolder 'UsageReport.xlsx'
            if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
            Write-Host "  Saving workbook to $xlsxPath ..." -ForegroundColor Cyan
            $saveSw = [System.Diagnostics.Stopwatch]::StartNew()
            $wb.SaveAs($xlsxPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
            $wb.Close($false)
            $saveSw.Stop()
            Write-Host ("  $xlsxPath ($([math]::Round((Get-Item $xlsxPath).Length/1KB,1)) KB, saved in {0:n1}s)" -f $saveSw.Elapsed.TotalSeconds) -ForegroundColor Green
        }
        catch {
            Write-Warning "Excel COM workbook build failed: $($_.Exception.Message)"
            Write-Warning "  at: $($_.InvocationInfo.PositionMessage)"
        }
        finally {
            try { $excel.SheetsInNewWorkbook = $previousSheetsInNewWorkbook } catch { }
            try { $excel.Calculation = -4105 } catch { }   # xlCalculationAutomatic
            $excel.ScreenUpdating = $true
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
}

if ($OpenAfterBuild -and $xlsxPath -and (Test-Path $xlsxPath)) {
    Write-Host "`nOpening workbook..." -ForegroundColor Cyan
    Start-Process $xlsxPath
}
elseif ($OpenAfterBuild) {
    Write-Host "`nOpening Master.csv..." -ForegroundColor Cyan
    Start-Process $masterPath
}

return [PSCustomObject]@{
    OutputFolder       = $OutputFolder
    SummaryCsv         = $summaryPath
    MasterCsv          = $masterPath
    TablesCsv          = $tablesPath
    CleanupCsv         = $cleanupPath
    DictionaryCsv      = $dictionaryPath
    ReadmePath         = $readmePath
    XlsxPath           = $xlsxPath
    SummaryRowCount    = $summary.Count
    MasterRowCount     = $master.Count
    TablesRowCount     = $tablesRollup.Count
    CleanupRowCount    = @($cleanup).Count
    DictionaryRowCount = $dictionary.Count
}
