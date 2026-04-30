<#
.SYNOPSIS
    Reports which tables are surfaced in the sitemaps of model-driven apps in the
    environment - the strongest "is this table user-facing?" signal in the suite.

.DESCRIPTION
    Enumerates every published model-driven app (appmodule) in the environment and walks
    its sitemap XML (Area -> Group -> SubArea), emitting one row per Entity-bound SubArea
    plus rows for non-entity tabs (Dashboard / Url / WebResource).

    Modern app sitemaps are stored as `sitemap` rows with `isappaware=true` and linked
    from the app via `appmodulecomponent` rows of `componenttype=62`. The /sitemaps
    endpoint defaults to filtering app-aware sitemaps OUT, so this script issues an
    explicit `?$filter=isappaware eq true` to retrieve them.

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

    FALLBACK: When an app has no linked sitemap row (rare; some settings-only or
    canvas-wrapped apps), one row per entity is synthesized from the entity list embedded
    in `appmodule.descriptor` (`appInfo.AppComponents.Entities[]`) or `appmodule.configxml`
    (`AppModuleComponent[type=1]/@schemaName`). In that case AreaId / AreaTitle /
    GroupId / GroupTitle / SubAreaTitle / Url are emitted blank and SubAreaType is 'Entity'.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Optional. Restricts output to (App, Entity) pairs whose Entity matches one of these
    table logical names.

.PARAMETER SolutionUniqueName
    Optional. Restricts output to entities that are members of the named solution
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

    Reports every (App, Entity) pair across every published model-driven app.

.EXAMPLE
    .\GetSitemapEntityPresence.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -SolutionUniqueName "msf_Core" -OutputFormat CSV -OutputPath ".\sitemap.csv"

    Reports only those (App, Entity) pairs whose Entity belongs to the msf_Core solution.

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
    "`$select=appmoduleid,appmoduleidunique,name,uniquename,formfactor,componentstate,description,publishedon,descriptor,configxml" +
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

# ---- Pre-fetch the appmodule -> sitemap link rows + the sitemap XMLs --------
#
# Modern model-driven apps store their navigation in a real `sitemap` row whose
# id is linked from the app via an `appmodulecomponent` row of componenttype 62.
# The /sitemaps endpoint defaults to filtering out app-aware sitemaps, so we
# can't just enumerate them - we have to follow the link rows.
#
#   appmodulecomponent (componenttype=62)
#       ._appmoduleidunique_value  -> appmodule.appmoduleidunique
#       .objectid                  -> sitemap.sitemapid
#
# We bulk-fetch all type=62 link rows once, group by app, then bulk-fetch the
# referenced sitemaps using `?$filter=isappaware eq true` (the only way to make
# the /sitemaps endpoint return modern-app sitemaps).
#
# When an app has no linked sitemap (very rare; some settings/canvas-only apps),
# we fall back to the entity list embedded in appmodule.descriptor /
# appmodule.configxml so the app still contributes to the per-table aggregate
# even without Area/Group/SubArea hierarchy.

Write-Host "Loading sitemap link rows..." -ForegroundColor Cyan
$linkRows = @()
try {
    $resp = Invoke-RestMethod -Uri ("$OrganizationUrl/api/data/v9.2/appmodulecomponents?" +
        "`$select=componenttype,objectid,_appmoduleidunique_value&`$filter=componenttype eq 62") `
        -Headers $headers -Method Get
    $linkRows = @($resp.value)
}
catch {
    Write-Warning "Failed to load appmodulecomponent link rows (componenttype=62): $($_.Exception.Message)"
}
Write-Host "  Link rows: $($linkRows.Count)" -ForegroundColor Gray

# Build map: appmoduleidunique -> list of sitemap-ids
$siteMapIdsByAppUnique = @{}
foreach ($r in $linkRows) {
    $appUnique = [string]$r.'_appmoduleidunique_value'
    if (-not $appUnique -or -not $r.objectid) { continue }
    if (-not $siteMapIdsByAppUnique.ContainsKey($appUnique)) {
        $siteMapIdsByAppUnique[$appUnique] = New-Object System.Collections.Generic.List[string]
    }
    $siteMapIdsByAppUnique[$appUnique].Add([string]$r.objectid) | Out-Null
}

# Bulk-fetch every app-aware sitemap row's XML (one trip; each row is large but
# this is still cheaper than one round-trip per app)
Write-Host "Loading app-aware sitemap rows..." -ForegroundColor Cyan
$siteMapById = @{}
try {
    $next = "$OrganizationUrl/api/data/v9.2/sitemaps?" +
        "`$select=sitemapid,sitemapname,sitemapnameunique,isappaware,sitemapxml,sitemapxmlmanaged&" +
        "`$filter=isappaware eq true"
    do {
        $resp = Invoke-RestMethod -Uri $next -Headers $headers -Method Get
        foreach ($sm in $resp.value) { $siteMapById[[string]$sm.sitemapid] = $sm }
        $next = $resp.'@odata.nextLink'
    } while ($next)
}
catch {
    Write-Warning "Failed to load app-aware sitemap rows: $($_.Exception.Message)"
}
Write-Host "  Sitemap rows: $($siteMapById.Count)" -ForegroundColor Gray

# ---- For each app, walk its sitemap XML (Area/Group/SubArea) ----------------
$rows = New-Object System.Collections.Generic.List[object]
$appsWithSitemap = 0
$appsFallbackEntities = 0
$appsNoData = 0
$appsParseFail = 0
$appOutcomes = New-Object System.Collections.Generic.List[object]

# Title-from-Titles helper (LCID 1033 preferred, first available otherwise).
function _SmTitle($node) {
    if (-not $node -or -not $node.Titles -or -not $node.Titles.Title) { return '' }
    $t = $node.Titles.Title | Where-Object { $_.LCID -eq '1033' } | Select-Object -First 1
    if (-not $t) { $t = $node.Titles.Title | Select-Object -First 1 }
    if ($t) { return [string]$t.Title } else { return '' }
}

foreach ($app in $appList) {
    $appLabel = if ($app.uniquename) { $app.uniquename } else { $app.name }
    $appUnique = [string]$app.appmoduleidunique

    # Path 1 (preferred): walk the linked sitemap XML for full Area/Group/SubArea.
    $smIds = if ($siteMapIdsByAppUnique.ContainsKey($appUnique)) { $siteMapIdsByAppUnique[$appUnique] } else { @() }
    $emittedFromSitemap = 0

    foreach ($smId in $smIds) {
        if (-not $siteMapById.ContainsKey($smId)) {
            $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='SitemapRowMissing'; Detail="link points to sitemapid=$smId but row was not returned by /sitemaps" }) | Out-Null
            continue
        }
        $sm = $siteMapById[$smId]
        $xmlText = if (-not [string]::IsNullOrWhiteSpace($sm.sitemapxml)) { $sm.sitemapxml }
                   elseif (-not [string]::IsNullOrWhiteSpace($sm.sitemapxmlmanaged)) { $sm.sitemapxmlmanaged }
                   else { $null }
        if (-not $xmlText) {
            $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='EmptySitemapXml'; Detail="sitemapid=$smId" }) | Out-Null
            continue
        }
        try {
            [xml]$xml = $xmlText
        }
        catch {
            Write-Warning "  [$appLabel] sitemap '$($sm.sitemapname)' XML parse failed: $($_.Exception.Message)"
            $appsParseFail++
            $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='XmlParseFail'; Detail=$_.Exception.Message }) | Out-Null
            continue
        }
        if (-not $xml.SiteMap) { continue }

        foreach ($area in @($xml.SiteMap.Area)) {
            if (-not $area) { continue }
            $areaId    = [string]$area.Id
            $areaTitle = _SmTitle $area
            foreach ($group in @($area.Group)) {
                if (-not $group) { continue }
                $groupId    = [string]$group.Id
                $groupTitle = _SmTitle $group
                foreach ($sub in @($group.SubArea)) {
                    if (-not $sub) { continue }
                    $subId    = [string]$sub.Id
                    $subTitle = _SmTitle $sub
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
                    $emittedFromSitemap++
                }
            }
        }
    }

    if ($emittedFromSitemap -gt 0) {
        $appsWithSitemap++
        Write-Host "  [$appLabel] $emittedFromSitemap sitemap entr$(if($emittedFromSitemap -eq 1){'y'}else{'ies'}) (sitemap)" -ForegroundColor Gray
        continue
    }

    # Path 2 (fallback): no sitemap data - synthesize one row per entity from the
    # descriptor / configxml entity list. AreaId/AreaTitle/GroupId/GroupTitle/Url
    # stay blank in this case; SubAreaId carries the metadata id when available.
    $entityMap  = @{}
    $sourceUsed = $null

    if (-not [string]::IsNullOrWhiteSpace($app.descriptor)) {
        try {
            $d = $app.descriptor | ConvertFrom-Json
            if ($d -and $d.appInfo -and $d.appInfo.AppComponents -and $d.appInfo.AppComponents.Entities) {
                foreach ($e in $d.appInfo.AppComponents.Entities) {
                    if ($e.LogicalName) {
                        $ln = ([string]$e.LogicalName).ToLowerInvariant()
                        if (-not $entityMap.ContainsKey($ln)) { $entityMap[$ln] = $e.Id }
                    }
                }
                if ($entityMap.Count -gt 0) { $sourceUsed = 'descriptor' }
            }
        }
        catch {
            Write-Warning "  [$appLabel] descriptor JSON parse failed: $($_.Exception.Message)"
            $appsParseFail++
            $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='DescriptorParseFail'; Detail=$_.Exception.Message }) | Out-Null
        }
    }

    if ($entityMap.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($app.configxml)) {
        try {
            $hits = [regex]::Matches($app.configxml, 'AppModuleComponent\s+type="1"\s+schemaName="([^"]+)"')
            foreach ($m in $hits) {
                $ln = $m.Groups[1].Value.ToLowerInvariant()
                if (-not $entityMap.ContainsKey($ln)) { $entityMap[$ln] = '' }
            }
            if ($entityMap.Count -gt 0) { $sourceUsed = 'configxml' }
        }
        catch {
            Write-Warning "  [$appLabel] configxml regex parse failed: $($_.Exception.Message)"
            $appsParseFail++
            $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='ConfigXmlParseFail'; Detail=$_.Exception.Message }) | Out-Null
        }
    }

    if ($entityMap.Count -eq 0) {
        $appsNoData++
        $appOutcomes.Add([PSCustomObject]@{ App=$appLabel; Outcome='NoData'; Detail='no linked sitemap, descriptor, or configxml entities'}) | Out-Null
        continue
    }

    $appsFallbackEntities++
    Write-Host "  [$appLabel] $($entityMap.Count) entit$(if($entityMap.Count -eq 1){'y'}else{'ies'}) (fallback: $sourceUsed)" -ForegroundColor DarkGray

    foreach ($tableLn in ($entityMap.Keys | Sort-Object)) {
        if ($tableFilterSet) {
            if (-not $tableFilterSet.Contains($tableLn)) { continue }
        }
        $rows.Add([PSCustomObject][ordered]@{
            AppUniqueName     = $app.uniquename
            AppDisplayName    = $app.name
            AppId             = $app.appmoduleid
            SitemapName       = ''
            AreaId            = ''
            AreaTitle         = ''
            GroupId           = ''
            GroupTitle        = ''
            SubAreaId         = $entityMap[$tableLn]
            SubAreaTitle      = ''
            SubAreaType       = 'Entity'
            TableLogicalName  = $tableLn
            Url               = ''
        })
    }
}

Write-Host "  Apps with sitemap XML: $appsWithSitemap   Apps using entity-list fallback: $appsFallbackEntities   Apps with no data: $appsNoData" -ForegroundColor Cyan
if ($appsParseFail -gt 0) {
    Write-Host "  Parse failures: $appsParseFail (see warnings above)" -ForegroundColor Yellow
}
Write-Host "  Sitemap entries emitted: $($rows.Count)" -ForegroundColor Green

# Empty-result hint: when 0 rows come out, point the operator at the most likely
# causes rather than leaving them to guess.
if ($rows.Count -eq 0 -and $appList.Count -gt 0) {
    Write-Warning "No sitemap entries were emitted. Most common causes:"
    Write-Warning "  1) -Tables / -SolutionUniqueName filtered every entity out (check the filter scope)."
    Write-Warning "     Fix: re-run without those filters to confirm app-entity data exists, then narrow."
    Write-Warning "  2) The signed-in user lacks Read on the appmodule descriptor/configxml columns."
    Write-Warning "     Fix: have a System Customizer or System Administrator run this."
    Write-Warning "  3) All apps are Unpublished (componentstate != 0) -> re-run with -IncludeUnpublished."
    Write-Warning "  4) The appmodules in scope have no entity components (canvas-only or settings apps)."
    if ($appOutcomes.Count -gt 0) {
        Write-Host "  First 5 per-app outcomes:" -ForegroundColor DarkGray
        $appOutcomes | Select-Object -First 5 | ForEach-Object {
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
