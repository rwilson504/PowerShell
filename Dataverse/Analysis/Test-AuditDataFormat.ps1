<#
.SYNOPSIS
    Diagnostic helper for GetAttributeAuditHistory.ps1. Dumps sample audit rows so we can
    see the actual attributemask format in your environment.

.DESCRIPTION
    Pulls 5-10 recent audit rows for the supplied table and prints:
      - Total audit row count for the table in the window
      - Row counts per operation (Create, Update, etc.)
      - Sample rows showing attributemask, changedata, operation, action
      - Whether attributemask is consistently empty (which would explain why event columns
        come back blank after script processing)

.PARAMETER OrganizationUrl
    The Dataverse organization URL.

.PARAMETER AccessToken
    Access token for the Web API.

.PARAMETER Table
    Logical name of the table to diagnose (e.g. 'msf_program').

.PARAMETER DaysBack
    How far back to look. Default 365.

.PARAMETER SampleSize
    How many sample rows to dump. Default 5.

.EXAMPLE
    .\Test-AuditDataFormat.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Table "msf_program"
#>

param (
    [Parameter(Mandatory = $true)] [string]$OrganizationUrl,
    [Parameter(Mandatory = $true)] [string]$AccessToken,
    [Parameter(Mandatory = $true)] [string]$Table,
    [Parameter(Mandatory = $false)] [int]$DaysBack = 365,
    [Parameter(Mandatory = $false)] [int]$SampleSize = 5
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Prefer"           = "odata.include-annotations=*"
}

$opLabels = @{ 1='Create'; 2='Update'; 3='Delete'; 4='Access'; 5='Upsert'; 100='UserAccessAuditStarted'; 101='UserAccessAuditEnded' }
$sinceUtc = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
$sinceStr = $sinceUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

Write-Host "`n=== Org audit settings ===" -ForegroundColor Cyan
$org = Invoke-RestMethod -Uri "$OrganizationUrl/api/data/v9.2/organizations?`$select=isauditenabled,auditretentionperiodv2,createdon" -Headers $headers
$org.value | Format-List isauditenabled, auditretentionperiodv2, createdon

Write-Host "=== Table-level audit setting for '$Table' ===" -ForegroundColor Cyan
try {
    $tbl = Invoke-RestMethod -Uri "$OrganizationUrl/api/data/v9.2/EntityDefinitions(LogicalName='$Table')?`$select=LogicalName,IsAuditEnabled,ObjectTypeCode" -Headers $headers
    "  Logical Name : $($tbl.LogicalName)"
    "  ObjectTypeCode: $($tbl.ObjectTypeCode)"
    "  IsAuditEnabled: $($tbl.IsAuditEnabled.Value)"
}
catch {
    Write-Warning "Could not read table metadata: $_"
}

Write-Host "`n=== Total audit rows for '$Table' in last $DaysBack days ===" -ForegroundColor Cyan
try {
    $countResp = Invoke-RestMethod -Uri "$OrganizationUrl/api/data/v9.2/audits?`$filter=objecttypecode eq '$Table' and createdon ge $sinceStr&`$count=true&`$top=1" -Headers $headers
    Write-Host "Total: $($countResp.'@odata.count')"
}
catch {
    Write-Warning "Count query failed: $_"
}

Write-Host "`n=== Per-operation breakdown ===" -ForegroundColor Cyan
foreach ($op in 1..7) {
    try {
        $r = Invoke-RestMethod -Uri "$OrganizationUrl/api/data/v9.2/audits?`$filter=objecttypecode eq '$Table' and createdon ge $sinceStr and operation eq $op&`$count=true&`$top=1" -Headers $headers
        $label = if ($opLabels.ContainsKey($op)) { $opLabels[$op] } else { "Op$op" }
        if ($r.'@odata.count' -gt 0) {
            Write-Host ("  {0,-12} (op={1}): {2}" -f $label, $op, $r.'@odata.count')
        }
    }
    catch { }
}

Write-Host "`n=== Sample of last $SampleSize audit rows ===" -ForegroundColor Cyan
try {
    $sample = Invoke-RestMethod -Uri "$OrganizationUrl/api/data/v9.2/audits?`$filter=objecttypecode eq '$Table' and createdon ge $sinceStr&`$select=auditid,createdon,operation,action,attributemask,changedata,objectid&`$orderby=createdon desc&`$top=$SampleSize" -Headers $headers

    if ($sample.value.Count -eq 0) {
        Write-Host "  No audit rows returned." -ForegroundColor Yellow
        Write-Host "  This usually means: org audit OFF, OR table audit OFF, OR no Create/Update activity in the window." -ForegroundColor Yellow
    }
    else {
        $i = 0
        foreach ($row in $sample.value) {
            $i++
            $opLabel = if ($opLabels.ContainsKey([int]$row.operation)) { $opLabels[[int]$row.operation] } else { "Op$($row.operation)" }
            $maskLen = if ($row.attributemask) { ($row.attributemask -as [string]).Length } else { 0 }
            $maskPreview = if ($row.attributemask) { $row.attributemask } else { '(empty)' }
            $changeLen = if ($row.changedata) { ($row.changedata -as [string]).Length } else { 0 }
            Write-Host "`n--- Row $i ---" -ForegroundColor Yellow
            Write-Host "  createdon         : $($row.createdon)"
            Write-Host "  operation         : $opLabel ($($row.operation))"
            Write-Host "  attributemask len : $maskLen"
            Write-Host "  attributemask val : $maskPreview"
            Write-Host "  changedata len    : $changeLen"
            if ($changeLen -gt 0 -and $changeLen -lt 600) {
                Write-Host "  changedata        : $($row.changedata)"
            }
            elseif ($changeLen -ge 600) {
                Write-Host "  changedata (first 400 chars): $(($row.changedata -as [string]).Substring(0, 400))..."
            }
        }
    }
}
catch {
    Write-Warning "Sample query failed: $_"
}

Write-Host "`n=== Diagnosis hints ===" -ForegroundColor Cyan
Write-Host "1. If 'No audit rows returned' but auditing is on, check that THIS specific table has table-level audit enabled (not just org-level)." -ForegroundColor Gray
Write-Host "2. If rows are returned but attributemask is consistently '(empty)', the script will skip them - that's the bug. Send me a few sample attributemask values so we can fix the parsing." -ForegroundColor Gray
Write-Host "3. If attributemask values look like comma-separated integers ('5,12,17'), the script SHOULD work - re-run with -Verbose to see what's happening per row." -ForegroundColor Gray
Write-Host "4. If attributemask values look like GUIDs or some other format, the script needs an update for your org's audit format." -ForegroundColor Gray
