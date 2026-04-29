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

    # DeadFieldScore (0-4)
    $fillRateNum = if ($a.FillRatePercent -eq 'N/A') { $null } else { [double]$a.FillRatePercent }
    $hasNoData   = ($null -ne $fillRateNum -and $fillRateNum -eq 0)
    $notTouched  = (-not $aud) -or [string]::IsNullOrWhiteSpace($aud.LastAuditedOn)
    $notInUI     = $ui -and ($ui.AnyUIPresence -eq 'False')
    $deadScore   = 0
    if ($hasNoData)            { $deadScore++ }
    if ($notTouched)           { $deadScore++ }
    if ($notInUI)              { $deadScore++ }
    if ($allManaged -eq $false){ $deadScore++ }

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
    $tablesRollup.Add([PSCustomObject][ordered]@{
        TableLogicalName            = if ($rc) { $rc.TableLogicalName }   elseif ($tu) { $tu.TableLogicalName }   else { $t }
        TableDisplayName            = if ($rc) { $rc.TableDisplayName }   elseif ($tu) { $tu.TableDisplayName }   else { '' }
        TableSchemaName             = if ($rc) { $rc.SchemaName }         elseif ($tu) { $tu.TableSchemaName }    else { '' }
        TableType                   = if ($rc) { $rc.TableType }          else { '' }
        IsCustomEntity              = if ($rc) { $rc.IsCustomEntity }     else { '' }
        OwnershipType               = if ($tu) { $tu.OwnershipType }      else { '' }
        RecordCount                 = if ($rc) { $rc.RecordCount }        else { '' }
        UsageBucket                 = if ($rc) { $rc.UsageBucket }        else { '' }
        LastModifiedOn              = if ($rc) { $rc.LastModifiedOn }     else { '' }
        DaysSinceLastModified       = if ($rc) { $rc.DaysSinceLastModified } else { '' }
        NewestCreatedOn             = if ($tu) { $tu.NewestCreatedOn }    else { '' }
        RecordsCreatedLast30Days    = if ($tu) { $tu.RecordsCreatedLast30Days }  else { '' }
        RecordsCreatedLast90Days    = if ($tu) { $tu.RecordsCreatedLast90Days }  else { '' }
        RecordsCreatedLast365Days   = if ($tu) { $tu.RecordsCreatedLast365Days } else { '' }
        DistinctCreators            = if ($tu) { $tu.DistinctCreators }   else { '' }
        DistinctModifiers           = if ($tu) { $tu.DistinctModifiers }  else { '' }
        DistinctOwners              = if ($tu) { $tu.DistinctOwners }     else { '' }
        ContainingSolutions         = $sols
    })
}
Write-Host "  Tables rollup rows: $($tablesRollup.Count)" -ForegroundColor Gray

# ---- CLEANUP CANDIDATES ----
$cleanup = $master |
    Where-Object { [int]$_.DeadFieldScore -ge 2 -and $_.IsCustomAttribute -eq 'True' } |
    Sort-Object @{Expression='DeadFieldScore'; Descending=$true}, TableLogicalName, AttributeLogicalName
Write-Host "  Cleanup candidates: $(@($cleanup).Count)" -ForegroundColor Gray

# ---- WRITE OUTPUTS ----
$masterPath  = Join-Path $OutputFolder 'master.csv'
$tablesPath  = Join-Path $OutputFolder 'tables.csv'
$cleanupPath = Join-Path $OutputFolder 'cleanup.csv'
$readmePath  = Join-Path $OutputFolder 'README.md'

Write-Host "Writing computed CSVs..." -ForegroundColor Cyan
$master       | Export-Csv -Path $masterPath  -NoTypeInformation
$tablesRollup | Export-Csv -Path $tablesPath  -NoTypeInformation
@($cleanup)   | Export-Csv -Path $cleanupPath -NoTypeInformation

# README -------------------------------------------------------------------
$readme = @"
# Dataverse Usage Report

Generated by ``Build-UsageReportWorkbook.ps1`` on top of CSVs produced by
``Invoke-DataverseUsageReport.ps1``.

## Files in this folder

### Generated join files (this script)
| File | What it is |
|---|---|
| ``master.csv``   | Full per-attribute join. Spine = attributeusage, with audit / UI / top-user / solution data left-joined on (TableLogicalName, AttributeLogicalName). One row per attribute. |
| ``tables.csv``   | Per-table roll-up joining recordcounts + tableusage on TableLogicalName. One row per table. |
| ``cleanup.csv``  | Pre-filtered Master view: ``DeadFieldScore >= 2 AND IsCustomAttribute = True``. Sorted by DeadFieldScore desc. |

### Source CSVs (from Invoke-DataverseUsageReport.ps1)
| File | Source script |
|---|---|
| ``recordcounts_*.csv``       | ``GetRecordCountByTable.ps1`` |
| ``relationships_*.csv``      | ``GetTableRelationships.ps1`` |
| ``solutionmembership_*.csv`` | ``GetSolutionMembership.ps1`` |
| ``tableusage_*.csv``         | ``GetTableUsageActivity.ps1`` |
| ``attributeusage_*.csv``     | ``GetFieldFillRateByTable.ps1`` |
| ``uipresence_*.csv``         | ``GetFieldUIPresence.ps1`` |
| ``useractivity_*.csv``       | ``GetUserActivityByTable.ps1`` |
| ``audithistory_*.csv``       | ``GetAttributeAuditHistory.ps1`` |

## DeadFieldScore (0-4)

Composite signal in ``master.csv``. Adds 1 for each of:

1. ``FillRatePercent = 0``                   - the field has no data
2. No audit events found in the audit window - nothing is touching it
3. ``AnyUIPresence = False``                 - field is not on any form, view, or chart
4. All containing solutions are unmanaged    - you can actually delete it

A custom attribute scoring 3 or 4 is a strong cleanup candidate. ``cleanup.csv`` shows
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
then re-run this script against it. master/tables/cleanup are rebuilt from scratch.
"@
Set-Content -Path $readmePath -Value $readme -Encoding UTF8

Write-Host "Generated:" -ForegroundColor Green
Write-Host "  $masterPath  ($($master.Count) rows)"
Write-Host "  $tablesPath  ($($tablesRollup.Count) rows)"
Write-Host "  $cleanupPath ($(@($cleanup).Count) rows)"
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
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        try {
            # Sheet build order: README first, then computed joins, then raw CSVs
            $sheetSpecs = @(
                @{ Name='README';        File=$readmePath;  IsMarkdown=$true }
                @{ Name='Master';        File=$masterPath;  IsMarkdown=$false }
                @{ Name='Tables';        File=$tablesPath;  IsMarkdown=$false }
                @{ Name='Cleanup';       File=$cleanupPath; IsMarkdown=$false }
            )
            foreach ($name in $reportPatterns.Keys) {
                if ($datasetFiles[$name]) {
                    $sheetSpecs += @{ Name=$name; File=$datasetFiles[$name]; IsMarkdown=$false }
                }
            }

            $wb = $excel.Workbooks.Add()
            # Excel.Application.Add() yields a workbook with one default sheet; rename + reuse
            $defaultSheet = $wb.Worksheets.Item(1)

            for ($i = 0; $i -lt $sheetSpecs.Count; $i++) {
                $spec = $sheetSpecs[$i]
                if ($i -eq 0) {
                    $sheet = $defaultSheet
                } else {
                    $sheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
                }
                $sheet.Name = $spec.Name

                if ($spec.IsMarkdown) {
                    # Drop the README markdown into column A as text rows
                    $lines = Get-Content $spec.File
                    for ($j = 0; $j -lt $lines.Count; $j++) {
                        $sheet.Cells.Item($j + 1, 1).Value2 = $lines[$j]
                    }
                    $sheet.Columns.Item(1).ColumnWidth = 120
                    $sheet.Columns.Item(1).WrapText    = $false
                }
                else {
                    # Read CSV as 2D array and bulk-write to the sheet
                    $rows = @(Import-Csv $spec.File)
                    if ($rows.Count -eq 0) {
                        $sheet.Cells.Item(1, 1).Value2 = "(empty)"
                        continue
                    }
                    $headers = $rows[0].PSObject.Properties.Name
                    $colCount = $headers.Count
                    $rowCount = $rows.Count + 1   # +1 for header

                    # Build a 2D object[,] - much faster than per-cell write
                    $matrix = New-Object 'object[,]' $rowCount, $colCount
                    for ($c = 0; $c -lt $colCount; $c++) { $matrix[0, $c] = $headers[$c] }
                    for ($r = 0; $r -lt $rows.Count; $r++) {
                        $row = $rows[$r]
                        for ($c = 0; $c -lt $colCount; $c++) {
                            $matrix[$r + 1, $c] = $row.($headers[$c])
                        }
                    }

                    $startCell = $sheet.Cells.Item(1, 1)
                    $endCell   = $sheet.Cells.Item($rowCount, $colCount)
                    $range     = $sheet.Range($startCell, $endCell)
                    $range.Value2 = $matrix

                    # Make it a real Excel Table for AutoFilter + Copilot recognition
                    [void]$sheet.ListObjects.Add(1, $range, $null, 1)   # 1 = xlSrcRange, 1 = xlYes (has headers)
                    $sheet.ListObjects.Item(1).Name = "$($spec.Name)_tbl"
                    $sheet.Columns.AutoFit() | Out-Null
                }
            }

            $xlsxPath = Join-Path $OutputFolder 'UsageReport.xlsx'
            if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
            $wb.SaveAs($xlsxPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
            $wb.Close($false)
            Write-Host "  $xlsxPath ($([math]::Round((Get-Item $xlsxPath).Length/1KB,1)) KB)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Excel COM workbook build failed: $_"
        }
        finally {
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
    OutputFolder    = $OutputFolder
    MasterCsv       = $masterPath
    TablesCsv       = $tablesPath
    CleanupCsv      = $cleanupPath
    ReadmePath      = $readmePath
    XlsxPath        = $xlsxPath
    MasterRowCount  = $master.Count
    TablesRowCount  = $tablesRollup.Count
    CleanupRowCount = @($cleanup).Count
}
