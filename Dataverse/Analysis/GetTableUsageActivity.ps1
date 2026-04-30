<#
.SYNOPSIS
    Reports table-level activity signals: newest record date, records created in trend
    windows, distinct creators, distinct owners. The table-level analog of the per-attribute
    GetAttributeAuditHistory.ps1.

.DESCRIPTION
    For each requested table the script:
      1. Resolves EntitySetName via metadata.
      2. Issues a small set of FetchXML aggregate queries to compute:
           - NewestCreatedOn / NewestModifiedOn
           - RecordsCreatedLast30 / 90 / 365 days
           - DistinctCreators (count of unique createdby users)
           - DistinctModifiers (count of unique modifiedby users)
           - DistinctOwners (count of unique ownerid users; null on tables without ownership)

    These signals work on tables of any size (FetchXML aggregates use indexes), require
    NO audit configuration, and complement the existing scripts:
      - GetRecordCountByTable          : how many records / when last touched
      - GetAttributeAuditHistory       : per-FIELD activity (needs audit on)
      - GetTableUsageActivity (this)   : per-TABLE activity (no audit needed)
      - GetUserActivityByTable         : WHO did the activity

    Output joins to the other CSVs on TableLogicalName.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.
.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.
.PARAMETER Tables
    Required. One or more table logical names.
.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".
.PARAMETER OutputPath
    Optional file path for export.

.EXAMPLE
    .\GetTableUsageActivity.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "account","msf_program"

    Reports table-level activity for account and msf_program.

.NOTES
    Cross-CSV joins: TableLogicalName matches every other CSV in this folder. Pivot in Excel
    to find tables with: high RecordCount, low DistinctCreators (single-owner risk), no
    activity in the last 90 days (likely abandoned).
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

function Get-EntityMetaSimple {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')?" +
        "`$select=LogicalName,SchemaName,EntitySetName,DisplayName,OwnershipType,PrimaryIdAttribute"
    try {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        return [PSCustomObject]@{
            LogicalName        = $r.LogicalName
            SchemaName         = $r.SchemaName
            EntitySetName      = $r.EntitySetName
            DisplayName        = if ($r.DisplayName.UserLocalizedLabel) { $r.DisplayName.UserLocalizedLabel.Label } else { $LogicalName }
            OwnershipType      = $r.OwnershipType
            PrimaryIdAttribute = $r.PrimaryIdAttribute
        }
    }
    catch {
        Write-Warning "Failed to load metadata for '$LogicalName': $_"
        return $null
    }
}

function Invoke-FetchAgg {
    <#
    .SYNOPSIS
        Issues a FetchXML aggregate query and returns the parsed value array.
        Returns @{ Status; Value; Error } - Status is 'Success' / 'AggregateLimitExceeded' / 'Error'.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$EntitySetName, [string]$FetchXml)
    $encoded = [System.Uri]::EscapeDataString($FetchXml)
    $url = "$OrgUrl/api/data/v9.2/$EntitySetName" + "?fetchXml=$encoded"
    try {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        return [PSCustomObject]@{ Status='Success'; Value=$r.value; Error=$null }
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -match 'AggregateQueryRecordLimit' -or $msg -match '0x8004E023') {
            return [PSCustomObject]@{ Status='AggregateLimitExceeded'; Value=@(); Error=$msg }
        }
        return [PSCustomObject]@{ Status='Error'; Value=@(); Error=$msg }
    }
}

function Format-DateForFetch {
    param ([datetime]$Dt)
    return $Dt.ToString("yyyy-MM-ddTHH:mm:ssZ")
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
    $nowUtc      = (Get-Date).ToUniversalTime()
    $threshold30 = Format-DateForFetch -Dt $nowUtc.AddDays(-30)
    $threshold90 = Format-DateForFetch -Dt $nowUtc.AddDays(-90)
    $threshold365= Format-DateForFetch -Dt $nowUtc.AddDays(-365)

    foreach ($logicalName in $Tables) {
        Write-Host "`n=== Processing table: $logicalName ===" -ForegroundColor Cyan

        $meta = Get-EntityMetaSimple -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        if (-not $meta) { continue }

        $pk     = $meta.PrimaryIdAttribute
        $set    = $meta.EntitySetName
        $hasOwner = ($meta.OwnershipType -in @('UserOwned','TeamOwned'))   # avoid DistinctOwners on Org-owned

        # Newest dates (single aggregate query)
        $maxFx = "<fetch aggregate=`"true`"><entity name=`"$logicalName`">" +
                 "<attribute name=`"createdon`"  alias=`"maxc`" aggregate=`"max`" />" +
                 "<attribute name=`"modifiedon`" alias=`"maxm`" aggregate=`"max`" />" +
                 "</entity></fetch>"
        $maxRes = Invoke-FetchAgg -OrgUrl $OrganizationUrl -Headers $headers -EntitySetName $set -FetchXml $maxFx
        $newestCreated  = $null
        $newestModified = $null
        $maxRows = @($maxRes.Value)
        if ($maxRes.Status -eq 'Success' -and $maxRows.Count -gt 0) {
            $newestCreated  = $maxRows[0].maxc
            $newestModified = $maxRows[0].maxm
        }

        # Records created in trend windows (one query per window)
        $created30 = $null; $created90 = $null; $created365 = $null
        foreach ($pair in @(
            @{ Var='created30';  Since=$threshold30  },
            @{ Var='created90';  Since=$threshold90  },
            @{ Var='created365'; Since=$threshold365 }
        )) {
            $fx = "<fetch aggregate=`"true`"><entity name=`"$logicalName`">" +
                  "<attribute name=`"$pk`" alias=`"cnt`" aggregate=`"count`" />" +
                  "<filter><condition attribute=`"createdon`" operator=`"ge`" value=`"$($pair.Since)`" /></filter>" +
                  "</entity></fetch>"
            $r = Invoke-FetchAgg -OrgUrl $OrganizationUrl -Headers $headers -EntitySetName $set -FetchXml $fx
            $rows = @($r.Value)
            if ($r.Status -eq 'Success' -and $rows.Count -gt 0) {
                Set-Variable -Name $pair.Var -Value ([long]$rows[0].cnt)
            }
        }

        # Distinct creators / modifiers / owners (one query per column)
        $distinctCreators = $null; $distinctModifiers = $null; $distinctOwners = $null
        $cols = @('createdby','modifiedby')
        if ($hasOwner) { $cols += 'ownerid' }
        $varMap = @{ 'createdby' = 'distinctCreators'; 'modifiedby' = 'distinctModifiers'; 'ownerid' = 'distinctOwners' }
        foreach ($col in $cols) {
            $fx = "<fetch aggregate=`"true`"><entity name=`"$logicalName`">" +
                  "<attribute name=`"$col`" alias=`"d`" aggregate=`"countcolumn`" distinct=`"true`" />" +
                  "</entity></fetch>"
            $r = Invoke-FetchAgg -OrgUrl $OrganizationUrl -Headers $headers -EntitySetName $set -FetchXml $fx
            $rows = @($r.Value)
            if ($r.Status -eq 'Success' -and $rows.Count -gt 0) {
                Set-Variable -Name $varMap[$col] -Value ([long]$rows[0].d)
            }
        }

        $daysSinceCreated  = if ($newestCreated)  { [math]::Max(0, [int]([math]::Floor(($nowUtc - ([datetime]$newestCreated).ToUniversalTime()).TotalDays))) } else { $null }
        $daysSinceModified = if ($newestModified) { [math]::Max(0, [int]([math]::Floor(($nowUtc - ([datetime]$newestModified).ToUniversalTime()).TotalDays))) } else { $null }

        # OwnershipType: Dataverse returns one of None / UserOwned / TeamOwned /
        # BusinessOwned / OrganizationOwned / BusinessParented. 'None' is a real enum
        # value (system / no-owner tables) but feels like a sentinel when mixed into
        # filters and PivotTables alongside the meaningful categories. Emit blank in
        # that case so PivotTable category buckets only contain real ownership types.
        $ownershipType = $meta.OwnershipType
        if ($ownershipType -and $ownershipType -ieq 'None') { $ownershipType = '' }

        $allResults.Add([PSCustomObject][ordered]@{
            TableLogicalName            = $logicalName
            TableDisplayName            = $meta.DisplayName
            TableSchemaName             = $meta.SchemaName
            OwnershipType               = $ownershipType
            NewestCreatedOn             = $newestCreated
            NewestModifiedOn            = $newestModified
            DaysSinceNewestCreated      = $daysSinceCreated
            DaysSinceNewestModified     = $daysSinceModified
            RecordsCreatedLast30Days    = $created30
            RecordsCreatedLast90Days    = $created90
            RecordsCreatedLast365Days   = $created365
            DistinctCreators            = $distinctCreators
            DistinctModifiers           = $distinctModifiers
            DistinctOwners              = $distinctOwners
        }) | Out-Null
    }

    $sorted = $allResults | Sort-Object TableLogicalName

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Tables analyzed: $($sorted.Count)"
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
                $OutputPath = Join-Path (Get-Location) "tableusage_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "tableusage_$timestamp.json"
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
