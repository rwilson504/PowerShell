<#
.SYNOPSIS
    Reports which tables are surfaced in the sitemaps of model-driven apps in the
    environment - the strongest "is this table user-facing?" signal in the suite.

.DESCRIPTION
    Enumerates every published model-driven app (appmodule) in the environment and walks
    its sitemap XML (Area -> Group -> SubArea), emitting one row per Entity-bound SubArea.

    A table can have data, audit activity, and a populated record count and STILL be
    invisible to end users (created/maintained entirely by Power Automate / plug-ins /
    integrations, with no app surfacing it). Sitemap presence answers "is this table
    actually exposed to a user via at least one model-driven app, and where?".

    The output joins on TableLogicalName to:
      - tables.csv / tableusage.csv : add user-facing-app context to per-table activity.
      - master.csv (per-attribute)  : transitively, since every attribute belongs to a table.

    The Build-UsageReportWorkbook builder also produces a per-table aggregate
    (InAnyAppSitemap, AppCount, AppNames) on the Tables sheet from this CSV.

    SubAreas without an Entity binding (URL tabs, Dashboard tabs, WebResource tabs,
    PowerBI/PowerApps embeds) are emitted as separate rows with TableLogicalName = empty
    and SubAreaType describing the tab type, so you can still see "the app has 4 dashboard
    tabs" without polluting the per-table aggregate.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Optional. Restricts output to SubAreas whose Entity matches one of these table logical
    names. Non-entity SubAreas are still emitted (for context) only when no -Tables filter
    is supplied.

.PARAMETER SolutionUniqueName
    Optional. Restricts output to SubAreas whose Entity is a member of the named solution
    (resolved via the standard solution-scope helper). Applied alongside -Tables.

.PARAMETER AppUniqueNames
    Optional. Restricts the appmodules scanned to a specific list of unique names.

.PARAMETER IncludeUnpublished
    Switch. Include appmodules with componentstate other than 0 (Published). Off by default.

.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".

.PARAMETER OutputPath
    Optional file path for export.

.EXAMPLE
    .\GetSitemapEntityPresence.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token

    Reports every Entity-bound SubArea across every published model-driven app.

.EXAMPLE
    .\GetSitemapEntityPresence.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SolutionUniqueName "msf_Core" -OutputFormat CSV -OutputPath ".\sitemap.csv"

    Reports only those SubAreas that surface tables belonging to the msf_Core solution.

.NOTES
    Cross-CSV joins:
      - TableLogicalName -> every other CSV in the suite (per-attribute and per-table).
      - AppUniqueName / AppDisplayName -> the AppModule (no other report joins on this today).
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
    [string[]]$AppUniqueNames,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeUnpublished,

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

# ---- Resolve scope (Tables + SolutionUniqueName) ------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "_SolutionFilterHelper.ps1")

$scopedTables = $null
if ($SolutionUniqueName -or ($Tables -and $Tables.Count -gt 0)) {
    $scopedTables = Resolve-SolutionScopedTables -OrgUrl $OrganizationUrl -Headers $headers -Tables $Tables -SolutionUniqueName $SolutionUniqueName
    if ($SolutionUniqueName -and (-not $scopedTables -or $scopedTables.Count -eq 0)) {
        Write-Warning "No tables in scope after applying -SolutionUniqueName / -Tables. No rows will be emitted."
        $scopedTables = @()
    }
}
# Lowercase set for fast lookup
$tableFilterSet = $null
if ($scopedTables -and $scopedTables.Count -gt 0) {
    $tableFilterSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($scopedTables | ForEach-Object { $_.ToLowerInvariant() }),
        [System.StringComparer]::OrdinalIgnoreCase)
}

# ---- Enumerate model-driven apps ---------------------------------------------
Write-Host "Enumerating model-driven apps..." -ForegroundColor Cyan
$appFilter = if ($IncludeUnpublished) { "" } else { "&`$filter=componentstate eq 0" }
$appUrl = "$OrganizationUrl/api/data/v9.2/appmodules?" +
    "`$select=appmoduleid,appmoduleidunique,name,uniquename,formfactor,componentstate,description,publishedon,_solutionid_value" +
    $appFilter

$appList = New-Object System.Collections.Generic.List[object]
try {
    $next = $appUrl
    do {
        $r = Invoke-RestMethod -Uri $next -Headers $headers -Method Get
        foreach ($a in $r.value) { $appList.Add($a) }
        $next = $r.'@odata.nextLink'
    } while ($next)
}
catch {
    Write-Error "Failed to enumerate appmodules: $_"
    exit 1
}

if ($AppUniqueNames -and $AppUniqueNames.Count -gt 0) {
    $wantedSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($AppUniqueNames | ForEach-Object { $_.ToLowerInvariant() }),
        [System.StringComparer]::OrdinalIgnoreCase)
    $appList = [System.Collections.Generic.List[object]]@($appList | Where-Object { $_.uniquename -and $wantedSet.Contains($_.uniquename) })
}

Write-Host "  Found $($appList.Count) app(s) to scan." -ForegroundColor Green
if ($appList.Count -eq 0) {
    Write-Warning "No appmodules matched the filters. If your environment has only draft/unpublished apps, re-run with -IncludeUnpublished."
}
else {
    # Surface componentstate distribution + a few app names so an empty result is
    # easy to debug. componentstate: 0=Published, 1=Unpublished, 2=Deleted, 3=Deleted Unpublished.
    $byState = $appList | Group-Object componentstate | Sort-Object Name
    $stateSummary = ($byState | ForEach-Object { "$($_.Count) state=$($_.Name)" }) -join ', '
    Write-Host "  componentstate breakdown: $stateSummary" -ForegroundColor Gray
    $sample = ($appList | Select-Object -First 5 | ForEach-Object { "$($_.uniquename) [state=$($_.componentstate)]" }) -join '; '
    Write-Host "  Sample apps: $sample" -ForegroundColor Gray
}

# ---- For each app, fetch its sitemap XML and walk Area/Group/SubArea ---------
$rows = New-Object System.Collections.Generic.List[object]
$appsWithSitemap = 0
$appsWithoutSitemap = 0
$appsExpandFailed  = 0
$appsEmptyXml      = 0
$appsXmlParseFail  = 0
$appSitemapDetail  = New-Object System.Collections.Generic.List[object]   # for end-of-run drill-down on failures

foreach ($app in $appList) {
    $appLabel = if ($app.uniquename) { $app.uniquename } else { $app.name }
    # Pull the associated appsitemap(s). The relationship name on appmodule is
    # 'appmodule_appsitemap' (M:N to appsitemap). An appmodule typically has exactly one.
    $sitemapUrl = "$OrganizationUrl/api/data/v9.2/appmodules($($app.appmoduleid))?" +
        "`$select=appmoduleid,name,uniquename" +
        "&`$expand=appmodule_appsitemap(`$select=appsitemapid,sitemapxml,sitemapname,sitemapnameunique)"
    try {
        $detail = Invoke-RestMethod -Uri $sitemapUrl -Headers $headers -Method Get
    }
    catch {
        Write-Warning "  [$appLabel] Failed to retrieve sitemap: $($_.Exception.Message)"
        $appsExpandFailed++
        $appsWithoutSitemap++
        $appSitemapDetail.Add([PSCustomObject]@{ App=$appLabel; Outcome='ExpandFailed'; Detail=$_.Exception.Message }) | Out-Null
        continue
    }

    $sitemaps = @()
    if ($detail.appmodule_appsitemap) { $sitemaps = @($detail.appmodule_appsitemap) }
    if ($sitemaps.Count -eq 0) {
        $appsWithoutSitemap++
        $appSitemapDetail.Add([PSCustomObject]@{ App=$appLabel; Outcome='NoSitemapRows'; Detail='appmodule_appsitemap expand returned 0 rows' }) | Out-Null
        continue
    }
    $appsWithSitemap++

    foreach ($sm in $sitemaps) {
        if ([string]::IsNullOrWhiteSpace($sm.sitemapxml)) {
            $appsEmptyXml++
            $appSitemapDetail.Add([PSCustomObject]@{ App=$appLabel; Outcome='EmptyXml'; Detail="sitemapname='$($sm.sitemapname)'" }) | Out-Null
            continue
        }
        try {
            [xml]$xml = $sm.sitemapxml
        }
        catch {
            Write-Warning "  [$appLabel] sitemap '$($sm.sitemapname)' XML parse failed: $($_.Exception.Message)"
            $appsXmlParseFail++
            $appSitemapDetail.Add([PSCustomObject]@{ App=$appLabel; Outcome='XmlParseFail'; Detail=$_.Exception.Message }) | Out-Null
            continue
        }

        # Walk Area -> Group -> SubArea. Match the lcid="1033" titling convention used by
        # the existing ConvertSitemapToCSV.ps1 - fall back to first available title if 1033
        # isn't present.
        function _Title($node) {
            if (-not $node -or -not $node.Titles -or -not $node.Titles.Title) { return '' }
            $t = $node.Titles.Title | Where-Object { $_.LCID -eq '1033' } | Select-Object -First 1
            if (-not $t) { $t = $node.Titles.Title | Select-Object -First 1 }
            if ($t) { return [string]$t.Title } else { return '' }
        }

        if (-not $xml.SiteMap) { continue }
        foreach ($area in @($xml.SiteMap.Area)) {
            if (-not $area) { continue }
            $areaId    = [string]$area.Id
            $areaTitle = _Title $area
            foreach ($group in @($area.Group)) {
                if (-not $group) { continue }
                $groupId    = [string]$group.Id
                $groupTitle = _Title $group
                foreach ($sub in @($group.SubArea)) {
                    if (-not $sub) { continue }
                    $subId    = [string]$sub.Id
                    $subTitle = _Title $sub
                    $entity   = if ($sub.HasAttribute('Entity'))    { [string]$sub.Entity }    else { '' }
                    $url      = if ($sub.HasAttribute('Url'))       { [string]$sub.Url }       else { '' }
                    $type     = if ($entity)                        { 'Entity' }
                                elseif ($url -match 'dashboards')   { 'Dashboard' }
                                elseif ($url -match 'webresource')  { 'WebResource' }
                                elseif ($url)                       { 'Url' }
                                else                                { 'Unknown' }

                    $tableLn = $entity.ToLowerInvariant()

                    # Apply table filter if supplied. Non-entity tabs are emitted only when
                    # no filter is in effect (they have no TableLogicalName to match against).
                    if ($tableFilterSet) {
                        if (-not $entity) { continue }
                        if (-not $tableFilterSet.Contains($tableLn)) { continue }
                    }

                    $rows.Add([PSCustomObject][ordered]@{
                        AppUniqueName     = $app.uniquename
                        AppDisplayName    = $app.name
                        AppId             = $app.appmoduleid
                        SitemapName       = $sm.sitemapname
                        AreaId            = $areaId
                        AreaTitle         = $areaTitle
                        GroupId           = $groupId
                        GroupTitle        = $groupTitle
                        SubAreaId         = $subId
                        SubAreaTitle      = $subTitle
                        SubAreaType       = $type
                        TableLogicalName  = $tableLn
                        Url               = $url
                    })
                }
            }
        }
    }
}

Write-Host "  Apps with a sitemap: $appsWithSitemap   Apps without (canvas/legacy/empty): $appsWithoutSitemap" -ForegroundColor Cyan
if ($appsExpandFailed -gt 0 -or $appsEmptyXml -gt 0 -or $appsXmlParseFail -gt 0) {
    Write-Host "  Sitemap fetch issues: ExpandFailed=$appsExpandFailed  EmptyXml=$appsEmptyXml  XmlParseFail=$appsXmlParseFail" -ForegroundColor Yellow
}
Write-Host "  Sitemap entries emitted: $($rows.Count)" -ForegroundColor Green

# Empty-result hint: when 0 rows come out, point the operator at the most likely
# causes rather than leaving them to guess.
if ($rows.Count -eq 0 -and $appList.Count -gt 0) {
    Write-Warning "No sitemap entries were emitted. Most common causes:"
    Write-Warning "  1) The signed-in user lacks Read on the appsitemap table -> the appmodule_appsitemap expand returns 0 rows silently."
    Write-Warning "     Fix: have a System Administrator run this, or grant Read on appsitemap to the user/team."
    Write-Warning "  2) All apps are Unpublished (componentstate != 0) -> re-run with -IncludeUnpublished."
    Write-Warning "  3) -Tables / -SolutionUniqueName filtered every entity-bound SubArea out."
    Write-Warning "     Fix: re-run without those filters to confirm sitemap data exists, then narrow."
    Write-Warning "  4) The environment uses Modern app designer apps without an appsitemap row (rare)."
    if ($appSitemapDetail.Count -gt 0) {
        Write-Host "  First 5 per-app outcomes:" -ForegroundColor DarkGray
        $appSitemapDetail | Select-Object -First 5 | ForEach-Object {
            Write-Host "    [$($_.App)] $($_.Outcome): $($_.Detail)" -ForegroundColor DarkGray
        }
    }
}

# ---- Output -------------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]@($rows |
    Sort-Object @{Expression='TableLogicalName'; Descending=$false}, AppUniqueName, AreaTitle, GroupTitle, SubAreaTitle)

switch ($OutputFormat) {
    "CSV" {
        if (-not $OutputPath) { $OutputPath = ".\sitemappresence_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" }
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported to $OutputPath" -ForegroundColor Green
    }
    "JSON" {
        if (-not $OutputPath) { $OutputPath = ".\sitemappresence_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" }
        $results | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8
        Write-Host "Exported to $OutputPath" -ForegroundColor Green
    }
    default {
        $results | Format-Table -AutoSize
    }
}

return $results
