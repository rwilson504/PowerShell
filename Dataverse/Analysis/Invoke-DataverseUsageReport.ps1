<#
.SYNOPSIS
    Runs the full Dataverse usage-analysis suite against one environment, scoped to a
    solution and/or table list, and drops every CSV into a timestamped folder.

.DESCRIPTION
    Calls each of the eight analysis scripts in turn and writes their CSV outputs to a
    single output folder so they can be opened together in Excel / Power BI / pandas.
    Every CSV uses the same composite join key (TableLogicalName + AttributeLogicalName,
    or RelationshipSchemaName for the relationships report).

    Scripts invoked, in order:
      1. GetRecordCountByTable.ps1        -> recordcounts_<ts>.csv
      2. GetTableRelationships.ps1        -> relationships_<ts>.csv
      3. GetSolutionMembership.ps1        -> solutionmembership_<ts>.csv
      4. GetSitemapEntityPresence.ps1     -> sitemappresence_<ts>.csv
      5. GetTableUsageActivity.ps1        -> tableusage_<ts>.csv
      6. GetFieldFillRateByTable.ps1      -> attributeusage_<ts>.csv
      7. GetFieldUIPresence.ps1           -> uipresence_<ts>.csv
      8. GetUserActivityByTable.ps1       -> useractivity_<ts>.csv
      9. GetAttributeAuditHistory.ps1     -> audithistory_<ts>.csv

    Use -Skip to omit any subset (e.g. -Skip Audit,UserActivity for a metadata-only run).

.PARAMETER OrganizationUrl
    The Dataverse organization URL.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Optional list of table logical names to scope every report to.

.PARAMETER SolutionUniqueName
    Optional solution unique name to scope every report to. Combines via intersection
    with -Tables when both supplied.

.PARAMETER OutputRoot
    Root folder under which a timestamped subfolder will be created. Default is the
    current directory. Final output path will be:
      <OutputRoot>\dataverse-usage-<env>-<timestamp>\

.PARAMETER Skip
    One or more report names to skip. Valid values:
      RecordCounts, Relationships, SolutionMembership, TableUsage,
      AttributeUsage, UIPresence, UserActivity, Audit

.PARAMETER IncludeLastActivity
    Forwarded to GetRecordCountByTable. Adds last/oldest CreatedOn + ModifiedOn timestamps.

.PARAMETER ActivityFallback
    Forwarded to GetRecordCountByTable. Probe activity even when count is 0/N/A.

.PARAMETER CustomEntitiesOnly
    Forwarded to GetRecordCountByTable when -SolutionUniqueName is NOT supplied (no effect
    when the solution filter already narrows the table set).

.PARAMETER AuditDaysBack
    Forwarded to GetAttributeAuditHistory. 0 = AUTO (default; uses retention or env age).

.PARAMETER AutoDetectUserLookups
    Forwarded to GetUserActivityByTable. Auto-discovers every systemuser-targeted lookup.

.PARAMETER UserLookupAttributes
    Forwarded to GetUserActivityByTable. Custom user lookups beyond the 4 standard ones.

.EXAMPLE
    .\Invoke-DataverseUsageReport.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SolutionUniqueName "msf_Core" -IncludeLastActivity -ActivityFallback -AutoDetectUserLookups

    Runs every report scoped to msf_Core. Output folder gets dropped in the current directory.

.EXAMPLE
    .\Invoke-DataverseUsageReport.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program","msf_contract" -OutputRoot "C:\Reports" -Skip Audit,UIPresence

    Runs the suite for two specific tables, writing into C:\Reports\dataverse-usage-...\, and
    skips the audit + UI-presence reports.
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
    [string]$OutputRoot = (Get-Location).Path,

    [Parameter(Mandatory = $false)]
    [ValidateSet('RecordCounts','Relationships','SolutionMembership','SitemapPresence','TableUsage','AttributeUsage','UIPresence','UserActivity','Audit')]
    [string[]]$Skip,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeLastActivity,

    [Parameter(Mandatory = $false)]
    [switch]$ActivityFallback,

    [Parameter(Mandatory = $false)]
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 3650)]
    [int]$AuditDaysBack = 0,

    [Parameter(Mandatory = $false)]
    [switch]$AutoDetectUserLookups,

    [Parameter(Mandatory = $false)]
    [string[]]$UserLookupAttributes,

    [Parameter(Mandatory = $false)]
    [string[]]$UserTargetTables = @('systemuser'),

    [Parameter(Mandatory = $false)]
    [hashtable]$CustomTargetNameColumns,

    [Parameter(Mandatory = $false)]
    [switch]$BuildWorkbook,

    [Parameter(Mandatory = $false)]
    [switch]$CombineToXlsx,

    [Parameter(Mandatory = $false)]
    [switch]$OpenAfterBuild
)

if (-not $Tables -and -not $SolutionUniqueName) {
    Write-Warning "Neither -Tables nor -SolutionUniqueName supplied. Reports that require an explicit scope (AttributeUsage, UIPresence, UserActivity, Audit, TableUsage) will fail."
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Build the output folder
$envSlug = ($OrganizationUrl -replace 'https?://', '' -replace '\..*$', '' -replace '[^a-zA-Z0-9-]', '-')
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$folderName = if ($SolutionUniqueName) {
    "dataverse-usage-$envSlug-$SolutionUniqueName-$timestamp"
} else {
    "dataverse-usage-$envSlug-$timestamp"
}
$outDir = Join-Path $OutputRoot $folderName
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

Write-Host "`n========================================================================" -ForegroundColor Cyan
Write-Host "  Dataverse Usage Report" -ForegroundColor Cyan
Write-Host "  Org      : $OrganizationUrl"
if ($SolutionUniqueName) { Write-Host "  Solution : $SolutionUniqueName" }
if ($Tables)             { Write-Host "  Tables   : $($Tables -join ', ')" }
Write-Host "  Output   : $outDir"
if ($Skip)               { Write-Host "  Skipping : $($Skip -join ', ')" -ForegroundColor Yellow }
Write-Host "========================================================================`n" -ForegroundColor Cyan

# Build a common parameter set forwarded to every analysis script
$common = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $AccessToken
    OutputFormat    = 'CSV'
}
if ($Tables -and $Tables.Count -gt 0) { $common.Tables             = $Tables }
if ($SolutionUniqueName)              { $common.SolutionUniqueName = $SolutionUniqueName }

function Invoke-Report {
    param (
        [string]$Name,
        [string]$Script,
        [string]$OutputFile,
        [hashtable]$ExtraParams
    )

    if ($Skip -and ($Skip -contains $Name)) {
        Write-Host "[SKIP] $Name" -ForegroundColor DarkGray
        return [PSCustomObject]@{ Name=$Name; Status='Skipped'; Path=$null; ElapsedSec=0; Error=$null }
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $path = Join-Path $outDir $OutputFile
    $callParams = $common.Clone()
    if ($ExtraParams) { foreach ($k in $ExtraParams.Keys) { $callParams[$k] = $ExtraParams[$k] } }
    $callParams.OutputPath = $path

    Write-Host "[RUN ] $Name -> $OutputFile" -ForegroundColor Cyan
    try {
        & (Join-Path $scriptDir $Script) @callParams *>&1 | Out-Null
        $sw.Stop()
        if (Test-Path $path) {
            $sizeKB = [math]::Round((Get-Item $path).Length / 1KB, 1)
            Write-Host "[OK  ] $Name : $sizeKB KB in $([math]::Round($sw.Elapsed.TotalSeconds, 1))s" -ForegroundColor Green
            return [PSCustomObject]@{ Name=$Name; Status='Success'; Path=$path; ElapsedSec=[math]::Round($sw.Elapsed.TotalSeconds,1); Error=$null }
        }
        else {
            Write-Host "[FAIL] $Name : no output produced" -ForegroundColor Yellow
            return [PSCustomObject]@{ Name=$Name; Status='NoOutput'; Path=$path; ElapsedSec=[math]::Round($sw.Elapsed.TotalSeconds,1); Error='No output file produced' }
        }
    }
    catch {
        $sw.Stop()
        Write-Host "[FAIL] $Name : $_" -ForegroundColor Red
        return [PSCustomObject]@{ Name=$Name; Status='Error'; Path=$path; ElapsedSec=[math]::Round($sw.Elapsed.TotalSeconds,1); Error="$_" }
    }
}

$summary = New-Object System.Collections.Generic.List[object]

# ---- 1. Record counts (cheap; do first; useful for understanding scale) ------
$rcParams = @{}
if ($CustomEntitiesOnly -and -not $SolutionUniqueName) { $rcParams.CustomEntitiesOnly = $true }
if ($IncludeLastActivity)                              { $rcParams.IncludeLastActivity = $true }
if ($ActivityFallback)                                 { $rcParams.ActivityFallback    = $true }
$summary.Add((Invoke-Report -Name 'RecordCounts'      -Script 'GetRecordCountByTable.ps1'   -OutputFile "recordcounts_$timestamp.csv"      -ExtraParams $rcParams))

# ---- 2. Relationships (single big metadata call; cheap) ----------------------
$summary.Add((Invoke-Report -Name 'Relationships'     -Script 'GetTableRelationships.ps1'   -OutputFile "relationships_$timestamp.csv"))

# ---- 3. Solution membership (cheap; useful for cleanup-impact analysis) ------
$summary.Add((Invoke-Report -Name 'SolutionMembership' -Script 'GetSolutionMembership.ps1'  -OutputFile "solutionmembership_$timestamp.csv"))

# ---- 4. Sitemap presence (one call per app; cheap; "is this table user-facing?") ----
#       Runs even when -Tables / -SolutionUniqueName aren't supplied (it scans all apps).
$summary.Add((Invoke-Report -Name 'SitemapPresence'   -Script 'GetSitemapEntityPresence.ps1' -OutputFile "sitemappresence_$timestamp.csv"))

# ---- 5. Table-level activity (FetchXML aggregates; medium cost per table) ----
$summary.Add((Invoke-Report -Name 'TableUsage'        -Script 'GetTableUsageActivity.ps1'   -OutputFile "tableusage_$timestamp.csv"))

# ---- 5. Attribute fill rate (one $batch per chunk; medium cost) -------------
$summary.Add((Invoke-Report -Name 'AttributeUsage'    -Script 'GetFieldFillRateByTable.ps1' -OutputFile "attributeusage_$timestamp.csv"))

# ---- 6. UI presence (form/view xml scan; cheap-medium) ----------------------
$summary.Add((Invoke-Report -Name 'UIPresence'        -Script 'GetFieldUIPresence.ps1'      -OutputFile "uipresence_$timestamp.csv"))

# ---- 7. User activity (one FetchXML aggregate per user lookup; medium) -------
$uaParams = @{}
if ($AutoDetectUserLookups)                           { $uaParams.AutoDetectUserLookups = $true }
if ($UserLookupAttributes -and $UserLookupAttributes.Count -gt 0) { $uaParams.UserLookupAttributes = $UserLookupAttributes }
if ($UserTargetTables -and ($UserTargetTables.Count -gt 1 -or $UserTargetTables[0] -ne 'systemuser')) { $uaParams.UserTargetTables = $UserTargetTables }
if ($CustomTargetNameColumns -and $CustomTargetNameColumns.Count -gt 0) { $uaParams.CustomTargetNameColumns = $CustomTargetNameColumns }
$summary.Add((Invoke-Report -Name 'UserActivity'      -Script 'GetUserActivityByTable.ps1'  -OutputFile "useractivity_$timestamp.csv"     -ExtraParams $uaParams))

# ---- 8. Audit history (audit-table scan; can be slow; biggest payload) -------
$ahParams = @{}
if ($AuditDaysBack -gt 0) { $ahParams.DaysBack = $AuditDaysBack }
$summary.Add((Invoke-Report -Name 'Audit'             -Script 'GetAttributeAuditHistory.ps1' -OutputFile "audithistory_$timestamp.csv"     -ExtraParams $ahParams))

Write-Host "`n========================================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "========================================================================" -ForegroundColor Cyan
$summary | Format-Table Name, Status, ElapsedSec, @{N='SizeKB';E={ if ($_.Path -and (Test-Path $_.Path)) { [math]::Round((Get-Item $_.Path).Length/1KB,1) } else { '' } }}, Path -AutoSize

$totalSec = ($summary | Measure-Object ElapsedSec -Sum).Sum
Write-Host "Total elapsed: $([math]::Round($totalSec,1))s" -ForegroundColor Cyan
Write-Host "Output folder: $outDir`n" -ForegroundColor Green

# Optionally build the joined workbook
if ($BuildWorkbook -or $CombineToXlsx -or $OpenAfterBuild) {
    Write-Host "`n--- Building workbook ---" -ForegroundColor Cyan
    $buildScript = Join-Path $scriptDir 'Build-UsageReportWorkbook.ps1'
    $buildParams = @{ InputFolder = $outDir }
    if ($CombineToXlsx)   { $buildParams.CombineToXlsx  = $true }
    if ($OpenAfterBuild)  { $buildParams.OpenAfterBuild = $true }
    & $buildScript @buildParams | Out-Null
}

return [PSCustomObject]@{
    OutputFolder = $outDir
    Reports      = $summary
}
