<#
.SYNOPSIS
    Reports whether each Dataverse attribute appears on any form, view, or chart - the
    strongest "field is dead in the UI" signal.

.DESCRIPTION
    For each requested table the script:
      1. Loads attribute metadata.
      2. Pulls every systemform (forms), savedquery (system views), userquery (personal
         views), and savedqueryvisualization (charts) for the table.
      3. Scans each artifact's xml (formxml / fetchxml / layoutxml / presentationxml) for
         attribute references. Counts and per-form-type breakdowns are recorded.
      4. Emits one row per attribute with OnAnyForm / OnAnyView / OnAnyChart booleans plus
         counts.

    Key signal: an attribute with FillRate=0 (from attributeusage_*.csv), no recent audit
    (audithistory_*.csv), AND OnAnyForm=false / OnAnyView=false is essentially dead - safe
    to delete with high confidence.

    Output composite key TableLogicalName + AttributeLogicalName joins to the other
    *_*.csv files in this folder.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.
.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.
.PARAMETER Tables
    Required. One or more table logical names to analyze.
.PARAMETER IncludeUserQueries
    When set, also scan personal views (userquery). Default off because personal views are
    user-scoped and generally noisier.
.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".
.PARAMETER OutputPath
    Optional file path for export.

.EXAMPLE
    .\GetFieldUIPresence.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program"

    Reports per-attribute UI presence on the msf_program table.

.NOTES
    Form types (systemform.type):
      2  = Main         (the primary record form)
      6  = Quick Create
      7  = Quick View
      8  = Dialog
      11 = Card
      12 = Main - Interactive Experience

    Cross-CSV "dead field" recipe:
      attributeusage   : FillRatePercent = 0
      audithistory     : LastAuditedOn = null OR DaysSinceLastAudited > 365
      uipresence       : OnAnyForm = false AND OnAnyView = false
      solutionmembership : ALL containing solutions are unmanaged AND not 'System'/'Default'
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
    [switch]$IncludeUserQueries,

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

$FormTypeLabels = @{
    2  = 'Main'
    6  = 'QuickCreate'
    7  = 'QuickView'
    8  = 'Dialog'
    11 = 'Card'
    12 = 'MainInteractive'
}

function Get-EntityMetadataAndAttributes {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)

    $entityUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')?" +
        "`$select=LogicalName,SchemaName,DisplayName,ObjectTypeCode"
    $entity = $null
    try {
        $resp = Invoke-RestMethod -Uri $entityUrl -Headers $Headers -Method Get
        $entity = [PSCustomObject]@{
            LogicalName    = $resp.LogicalName
            SchemaName     = $resp.SchemaName
            ObjectTypeCode = $resp.ObjectTypeCode
            DisplayName    = if ($resp.DisplayName.UserLocalizedLabel) { $resp.DisplayName.UserLocalizedLabel.Label } else { $LogicalName }
        }
    }
    catch {
        Write-Warning "Failed to load entity '$LogicalName': $_"
        return $null
    }

    $attrUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$LogicalName')/Attributes?" +
        "`$select=LogicalName,SchemaName,DisplayName,AttributeType,IsCustomAttribute,IsLogical,AttributeOf"
    $attrs = @()
    do {
        $resp = Invoke-RestMethod -Uri $attrUrl -Headers $Headers -Method Get
        foreach ($a in $resp.value) {
            if ($a.IsLogical -or $a.AttributeOf) { continue }
            $attrs += [PSCustomObject]@{
                LogicalName       = $a.LogicalName
                SchemaName        = $a.SchemaName
                DisplayName       = if ($a.DisplayName.UserLocalizedLabel) { $a.DisplayName.UserLocalizedLabel.Label } else { $a.LogicalName }
                AttributeType     = $a.AttributeType
                IsCustomAttribute = [bool]$a.IsCustomAttribute
            }
        }
        $attrUrl = $resp.'@odata.nextLink'
    } while ($attrUrl)

    return [PSCustomObject]@{ Entity = $entity; Attributes = $attrs }
}

function Get-FormsForEntity {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/systemforms?" +
        "`$select=formid,name,type,formxml" +
        "&`$filter=objecttypecode eq '$LogicalName'"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-SavedQueriesForEntity {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/savedqueries?" +
        "`$select=savedqueryid,name,querytype,fetchxml,layoutxml" +
        "&`$filter=returnedtypecode eq '$LogicalName'"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-UserQueriesForEntity {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/userqueries?" +
        "`$select=userqueryid,name,querytype,fetchxml,layoutxml" +
        "&`$filter=returnedtypecode eq '$LogicalName'"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-ChartsForEntity {
    param ([string]$OrgUrl, [hashtable]$Headers, [string]$LogicalName)
    $url = "$OrgUrl/api/data/v9.2/savedqueryvisualizations?" +
        "`$select=savedqueryvisualizationid,name,datadescription,presentationdescription" +
        "&`$filter=primaryentitytypecode eq '$LogicalName'"
    $all = @()
    try {
        do {
            $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
            $all += $r.value
            $url = $r.'@odata.nextLink'
        } while ($url)
    }
    catch {
        Write-Warning "Failed to load charts for '$LogicalName': $_"
    }
    return $all
}

function Get-AttributeReferences {
    <#
    .SYNOPSIS
        Returns a HashSet[string] of attribute logical names referenced anywhere in the
        supplied xml string. Uses two regex patterns that catch the canonical form/view
        attribute references:
          datafieldname="logicalname"   (formxml)
          name="logicalname"            (fetchxml + layoutxml + chart datadescription)
    #>
    param ([string]$Xml)
    $set = New-Object System.Collections.Generic.HashSet[string]
    if ([string]::IsNullOrWhiteSpace($Xml)) { return $set }

    foreach ($m in [regex]::Matches($Xml, 'datafieldname="([^"]+)"')) {
        [void]$set.Add($m.Groups[1].Value.ToLowerInvariant())
    }
    # FetchXML <attribute name="..."/> and <condition attribute="..."/>
    foreach ($m in [regex]::Matches($Xml, '<attribute\s+name="([^"]+)"')) {
        [void]$set.Add($m.Groups[1].Value.ToLowerInvariant())
    }
    foreach ($m in [regex]::Matches($Xml, 'condition\s+attribute="([^"]+)"')) {
        [void]$set.Add($m.Groups[1].Value.ToLowerInvariant())
    }
    # Chart datadescription embedded JSON or XML usually references via "alias" or "name"
    foreach ($m in [regex]::Matches($Xml, '"alias"\s*:\s*"([^"]+)"')) {
        [void]$set.Add($m.Groups[1].Value.ToLowerInvariant())
    }

    return $set
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

    foreach ($logicalName in $Tables) {
        Write-Host "`n=== Processing table: $logicalName ===" -ForegroundColor Cyan

        $meta = Get-EntityMetadataAndAttributes -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        if (-not $meta) { continue }
        $entity = $meta.Entity
        $attrs  = $meta.Attributes
        Write-Host "  Eligible attributes: $($attrs.Count)" -ForegroundColor Gray

        Write-Host "  Loading forms..." -ForegroundColor Gray
        $forms = Get-FormsForEntity -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        Write-Host "  Loading system views..." -ForegroundColor Gray
        $views = Get-SavedQueriesForEntity -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        $userViews = @()
        if ($IncludeUserQueries) {
            Write-Host "  Loading personal views..." -ForegroundColor Gray
            $userViews = Get-UserQueriesForEntity -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName
        }
        Write-Host "  Loading charts..." -ForegroundColor Gray
        $charts = Get-ChartsForEntity -OrgUrl $OrganizationUrl -Headers $headers -LogicalName $logicalName

        Write-Host "  Forms=$($forms.Count) Views=$($views.Count) UserViews=$($userViews.Count) Charts=$($charts.Count)" -ForegroundColor Green

        # Pre-compute attribute reference sets per artifact
        $attrCounts = @{}  # logicalName -> @{ FormCount; FormTypes(set); ViewCount; UserViewCount; ChartCount }
        foreach ($a in $attrs) {
            $attrCounts[$a.LogicalName] = [PSCustomObject]@{
                FormCount     = 0
                FormTypes     = New-Object System.Collections.Generic.HashSet[string]
                ViewCount     = 0
                UserViewCount = 0
                ChartCount    = 0
            }
        }

        foreach ($f in $forms) {
            $refs = Get-AttributeReferences -Xml $f.formxml
            $typeLabel = if ($FormTypeLabels.ContainsKey([int]$f.type)) { $FormTypeLabels[[int]$f.type] } else { "Type$($f.type)" }
            foreach ($a in $refs) {
                if ($attrCounts.ContainsKey($a)) {
                    $attrCounts[$a].FormCount++
                    [void]$attrCounts[$a].FormTypes.Add($typeLabel)
                }
            }
        }
        foreach ($v in $views) {
            $refs = Get-AttributeReferences -Xml ($v.fetchxml + ' ' + $v.layoutxml)
            foreach ($a in $refs) { if ($attrCounts.ContainsKey($a)) { $attrCounts[$a].ViewCount++ } }
        }
        foreach ($v in $userViews) {
            $refs = Get-AttributeReferences -Xml ($v.fetchxml + ' ' + $v.layoutxml)
            foreach ($a in $refs) { if ($attrCounts.ContainsKey($a)) { $attrCounts[$a].UserViewCount++ } }
        }
        foreach ($c in $charts) {
            $refs = Get-AttributeReferences -Xml ($c.datadescription + ' ' + $c.presentationdescription)
            foreach ($a in $refs) { if ($attrCounts.ContainsKey($a)) { $attrCounts[$a].ChartCount++ } }
        }

        foreach ($a in $attrs) {
            $c = $attrCounts[$a.LogicalName]
            $allResults.Add([PSCustomObject][ordered]@{
                TableLogicalName     = $logicalName
                TableDisplayName     = $entity.DisplayName
                TableSchemaName      = $entity.SchemaName
                AttributeLogicalName = $a.LogicalName
                AttributeSchemaName  = $a.SchemaName
                AttributeDisplayName = $a.DisplayName
                AttributeType        = $a.AttributeType
                IsCustomAttribute    = $a.IsCustomAttribute
                OnAnyForm            = ($c.FormCount -gt 0)
                FormCount            = $c.FormCount
                FormTypes            = ($c.FormTypes | Sort-Object) -join ';'
                OnAnyView            = ($c.ViewCount -gt 0)
                ViewCount            = $c.ViewCount
                OnAnyUserView        = ($c.UserViewCount -gt 0)
                UserViewCount        = $c.UserViewCount
                OnAnyChart           = ($c.ChartCount -gt 0)
                ChartCount           = $c.ChartCount
                AnyUIPresence        = (($c.FormCount + $c.ViewCount + $c.UserViewCount + $c.ChartCount) -gt 0)
            }) | Out-Null
        }
    }

    $sorted = $allResults | Sort-Object TableLogicalName, AttributeLogicalName

    $totalRows  = $sorted.Count
    $deadFields = ($sorted | Where-Object { -not $_.AnyUIPresence }).Count
    $onForms    = ($sorted | Where-Object { $_.OnAnyForm }).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total attribute rows: $totalRows"
    Write-Host "  Visible on any form: $onForms" -ForegroundColor Green
    Write-Host "  Not on any form/view/chart: $deadFields" -ForegroundColor $(if ($deadFields -gt 0) { 'Yellow' } else { 'Green' })
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
                $OutputPath = Join-Path (Get-Location) "uipresence_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "uipresence_$timestamp.json"
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
