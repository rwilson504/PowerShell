<#
.SYNOPSIS
    Reports which Dataverse solution(s) ship each table, attribute, and relationship - so
    you can answer "is anything still shipping this field?" before deleting it.

.DESCRIPTION
    Joins the solutioncomponents table (componenttype 1=Entity, 2=Attribute, 10=Relationship)
    against EntityDefinitions and AttributeDefinitions metadata to produce one row per
    (component, solution) membership. A field shipped only by long-deleted unmanaged
    solutions is a clean cleanup candidate; one shipped by an active managed solution is
    likely still required.

    Output columns are aligned with the other CSVs in this folder for cross-correlation:
    TableLogicalName + AttributeLogicalName composite key joins to attributeusage_*.csv,
    audithistory_*.csv, useractivity_*.csv, and recordcounts_*.csv. RelationshipSchemaName
    joins to relationships_*.csv.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Optional. One or more table logical names to filter to. When omitted, every solution
    component for every entity/attribute/relationship is returned.

.PARAMETER ComponentTypes
    Which solution component types to include. Valid values: "Entity" (table-level
    membership), "Attribute" (column-level), "Relationship" (1:N / N:1 / M:M membership).
    Default is all three.

.PARAMETER UnmanagedOnly
    When set, only emit rows where the containing solution is unmanaged. Useful for finding
    org-built artifacts that you actually have permission to delete.

.PARAMETER ManagedOnly
    When set, only emit rows where the containing solution is managed. Mutually exclusive
    with -UnmanagedOnly.

.PARAMETER ExcludeSystemSolutions
    Skip the well-known system solutions ('System', 'Active', 'Default', 'Cr03482' Common
    Data Services Default Solution, plus anything starting with 'msdyn'/'msft'/'mscrm') so
    the output focuses on customer-built and ISV solutions.

.PARAMETER OutputFormat
    "Table" / "CSV" / "JSON". Default "Table".

.PARAMETER OutputPath
    Optional file path to export the results.

.EXAMPLE
    .\GetSolutionMembership.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token

    Lists every Entity / Attribute / Relationship membership across every solution.

.EXAMPLE
    .\GetSolutionMembership.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables "msf_program","msf_contract" -ComponentTypes Attribute,Relationship -UnmanagedOnly -OutputFormat CSV

    Lists only attribute and relationship memberships on msf_program/msf_contract that ship
    in unmanaged solutions - cleanup candidates.

.NOTES
    Component type discriminators (Microsoft.Crm.Sdk.Messages.solutioncomponent_componenttype):
        1  = Entity
        2  = Attribute
        10 = Relationship (covers OneToMany / ManyToOne / ManyToMany)

    Cross-CSV joins:
      attributeusage / audithistory / useractivity / recordcounts:
        TableLogicalName + AttributeLogicalName  (when ComponentType = 'Attribute')
        TableLogicalName                         (when ComponentType = 'Entity')
      relationships:
        RelationshipSchemaName                   (when ComponentType = 'Relationship')

    "Cleanup candidate" recipe:
      attributeusage : FillRatePercent = 0
      audithistory   : DaysSinceLastAudited > 365 OR LastAuditedOn = null
      solutionmembership : ALL containing solutions are unmanaged AND not 'System' / 'Active' / 'Default'
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $false)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Entity","Attribute","Relationship")]
    [string[]]$ComponentTypes = @("Entity","Attribute","Relationship"),

    [Parameter(Mandatory = $false)]
    [switch]$UnmanagedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ManagedOnly,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeSystemSolutions,

    [Parameter(Mandatory = $false)]
    [string]$SolutionUniqueName,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

if ($UnmanagedOnly -and $ManagedOnly) {
    Write-Error "-UnmanagedOnly and -ManagedOnly are mutually exclusive."
    exit 1
}

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Content-Type"     = "application/json; charset=utf-8"
    "Prefer"           = "odata.include-annotations=*"
}

$ComponentTypeMap = @{
    1  = 'Entity'
    2  = 'Attribute'
    10 = 'Relationship'
}
$ComponentTypeReverseMap = @{}
foreach ($k in $ComponentTypeMap.Keys) { $ComponentTypeReverseMap[$ComponentTypeMap[$k]] = $k }

# Well-known system solutions to skip with -ExcludeSystemSolutions
$SystemSolutionUniqueNames = @('System','Active','Default','Cr03482','ActiveCustomizationDefaultSolution')

function Get-AllSolutions {
    param ([string]$OrgUrl, [hashtable]$Headers)
    $url = "$OrgUrl/api/data/v9.2/solutions?" +
        "`$select=solutionid,uniquename,friendlyname,ismanaged,version,installedon,_publisherid_value"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-AllEntityDefinitions {
    param ([string]$OrgUrl, [hashtable]$Headers)
    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions?" +
        "`$select=MetadataId,LogicalName,SchemaName,DisplayName,IsCustomEntity"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-AllRelationshipDefinitions {
    param ([string]$OrgUrl, [hashtable]$Headers)
    $url = "$OrgUrl/api/data/v9.2/RelationshipDefinitions"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

function Get-AttributeDefinitionsForEntities {
    <#
    .SYNOPSIS
        Returns array of @{ MetadataId; LogicalName; SchemaName; DisplayName; IsCustomAttribute;
        EntityLogicalName } for all attributes of the supplied entities.
    #>
    param ([string]$OrgUrl, [hashtable]$Headers, [array]$Entities)
    $all = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($entity in $Entities) {
        $i++
        Write-Progress -Activity "Loading attribute metadata" `
            -Status "$i of $($Entities.Count) - $($entity.LogicalName)" `
            -PercentComplete (($i / $Entities.Count) * 100)
        $url = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$($entity.LogicalName)')/Attributes?" +
            "`$select=MetadataId,LogicalName,SchemaName,DisplayName,IsCustomAttribute,AttributeOf,IsLogical"
        try {
            do {
                $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
                foreach ($a in $r.value) {
                    if ($a.IsLogical -or $a.AttributeOf) { continue }
                    $all.Add([PSCustomObject]@{
                        MetadataId        = $a.MetadataId
                        LogicalName       = $a.LogicalName
                        SchemaName        = $a.SchemaName
                        DisplayName       = if ($a.DisplayName.UserLocalizedLabel) { $a.DisplayName.UserLocalizedLabel.Label } else { $a.LogicalName }
                        IsCustomAttribute = [bool]$a.IsCustomAttribute
                        EntityLogicalName = $entity.LogicalName
                    }) | Out-Null
                }
                $url = $r.'@odata.nextLink'
            } while ($url)
        }
        catch {
            Write-Warning "Failed to load attributes for $($entity.LogicalName): $_"
        }
    }
    Write-Progress -Activity "Loading attribute metadata" -Completed
    return $all.ToArray()
}

function Get-SolutionComponents {
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [int[]]$ComponentTypeNumbers
    )
    $typeFilter = ($ComponentTypeNumbers | ForEach-Object { "componenttype eq $_" }) -join ' or '
    $url = "$OrgUrl/api/data/v9.2/solutioncomponents?" +
        "`$select=componenttype,objectid,_solutionid_value,rootcomponentbehavior" +
        "&`$filter=$typeFilter"
    $all = @()
    do {
        $r = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $r.value
        $url = $r.'@odata.nextLink'
    } while ($url)
    return $all
}

# Main script execution
try {
    $wantTypeNumbers = $ComponentTypes | ForEach-Object { $ComponentTypeReverseMap[$_] }

    Write-Host "Loading solution catalog..." -ForegroundColor Cyan
    $solutions = Get-AllSolutions -OrgUrl $OrganizationUrl -Headers $headers
    $solutionMap = @{}
    foreach ($s in $solutions) { $solutionMap[$s.solutionid] = $s }
    Write-Host "  Solutions in environment: $($solutions.Count)" -ForegroundColor Gray

    Write-Host "Loading entity catalog..." -ForegroundColor Cyan
    $entities = Get-AllEntityDefinitions -OrgUrl $OrganizationUrl -Headers $headers
    $entityByMetadataId = @{}
    foreach ($e in $entities) { $entityByMetadataId[$e.MetadataId] = $e }
    Write-Host "  Entities in environment: $($entities.Count)" -ForegroundColor Gray

    # Tables filter (case-insensitive)
    $tablesFilterSet = $null
    if ($Tables -and $Tables.Count -gt 0) {
        $tablesFilterSet = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@($Tables | ForEach-Object { $_.ToLowerInvariant() }),
            [System.StringComparer]::OrdinalIgnoreCase)
        $entities = $entities | Where-Object { $tablesFilterSet.Contains($_.LogicalName) }
        Write-Host "  Filtered to $($entities.Count) entities" -ForegroundColor Gray
    }

    $attributesByMetadataId = @{}
    if ($wantTypeNumbers -contains 2) {
        Write-Host "Loading attribute catalog (this is the slow step)..." -ForegroundColor Cyan
        $attributes = Get-AttributeDefinitionsForEntities -OrgUrl $OrganizationUrl -Headers $headers -Entities $entities
        foreach ($a in $attributes) { $attributesByMetadataId[$a.MetadataId] = $a }
        Write-Host "  Attributes loaded: $($attributes.Count)" -ForegroundColor Gray
    }

    $relationshipsByMetadataId = @{}
    if ($wantTypeNumbers -contains 10) {
        Write-Host "Loading relationship catalog..." -ForegroundColor Cyan
        $rels = Get-AllRelationshipDefinitions -OrgUrl $OrganizationUrl -Headers $headers
        foreach ($r in $rels) { $relationshipsByMetadataId[$r.MetadataId] = $r }
        Write-Host "  Relationships loaded: $($rels.Count)" -ForegroundColor Gray
    }

    Write-Host "Loading solution-component memberships..." -ForegroundColor Cyan
    $components = Get-SolutionComponents -OrgUrl $OrganizationUrl -Headers $headers -ComponentTypeNumbers $wantTypeNumbers
    Write-Host "  Solution-component rows: $($components.Count)" -ForegroundColor Gray

    $allResults = New-Object System.Collections.Generic.List[object]

    foreach ($c in $components) {
        $solutionId = $c.'_solutionid_value'
        $sol        = if ($solutionMap.ContainsKey($solutionId)) { $solutionMap[$solutionId] } else { $null }
        if (-not $sol) { continue }

        if ($UnmanagedOnly -and $sol.ismanaged) { continue }
        if ($ManagedOnly -and -not $sol.ismanaged) { continue }
        if ($ExcludeSystemSolutions -and ($SystemSolutionUniqueNames -contains $sol.uniquename -or $sol.uniquename -match '^(msdyn|msft|mscrm)')) { continue }
        if ($SolutionUniqueName -and $sol.uniquename -ne $SolutionUniqueName) { continue }

        $typeName = $ComponentTypeMap[[int]$c.componenttype]
        if (-not $typeName) { continue }

        switch ([int]$c.componenttype) {
            1 {
                # Entity
                if (-not $entityByMetadataId.ContainsKey($c.objectid)) { continue }
                $e = $entityByMetadataId[$c.objectid]
                if ($tablesFilterSet -and -not $tablesFilterSet.Contains($e.LogicalName)) { continue }

                $allResults.Add([PSCustomObject][ordered]@{
                    ComponentType            = 'Entity'
                    TableLogicalName         = $e.LogicalName
                    TableDisplayName         = if ($e.DisplayName.UserLocalizedLabel) { $e.DisplayName.UserLocalizedLabel.Label } else { $e.LogicalName }
                    TableSchemaName          = $e.SchemaName
                    AttributeLogicalName     = ''
                    AttributeSchemaName      = ''
                    AttributeDisplayName     = ''
                    RelationshipSchemaName   = ''
                    IsCustomComponent        = [bool]$e.IsCustomEntity
                    SolutionUniqueName       = $sol.uniquename
                    SolutionFriendlyName     = $sol.friendlyname
                    SolutionVersion          = $sol.version
                    SolutionIsManaged        = [bool]$sol.ismanaged
                    SolutionInstalledOn      = $sol.installedon
                    RootComponentBehavior    = $c.'rootcomponentbehavior@OData.Community.Display.V1.FormattedValue'
                }) | Out-Null
            }
            2 {
                # Attribute
                if (-not $attributesByMetadataId.ContainsKey($c.objectid)) { continue }
                $a = $attributesByMetadataId[$c.objectid]
                if ($tablesFilterSet -and -not $tablesFilterSet.Contains($a.EntityLogicalName)) { continue }
                $e = $entityByMetadataId.Values | Where-Object { $_.LogicalName -eq $a.EntityLogicalName } | Select-Object -First 1

                $allResults.Add([PSCustomObject][ordered]@{
                    ComponentType            = 'Attribute'
                    TableLogicalName         = $a.EntityLogicalName
                    TableDisplayName         = if ($e -and $e.DisplayName.UserLocalizedLabel) { $e.DisplayName.UserLocalizedLabel.Label } else { $a.EntityLogicalName }
                    TableSchemaName          = if ($e) { $e.SchemaName } else { '' }
                    AttributeLogicalName     = $a.LogicalName
                    AttributeSchemaName      = $a.SchemaName
                    AttributeDisplayName     = $a.DisplayName
                    RelationshipSchemaName   = ''
                    IsCustomComponent        = $a.IsCustomAttribute
                    SolutionUniqueName       = $sol.uniquename
                    SolutionFriendlyName     = $sol.friendlyname
                    SolutionVersion          = $sol.version
                    SolutionIsManaged        = [bool]$sol.ismanaged
                    SolutionInstalledOn      = $sol.installedon
                    RootComponentBehavior    = $c.'rootcomponentbehavior@OData.Community.Display.V1.FormattedValue'
                }) | Out-Null
            }
            10 {
                # Relationship
                if (-not $relationshipsByMetadataId.ContainsKey($c.objectid)) { continue }
                $r = $relationshipsByMetadataId[$c.objectid]
                $isM2M    = ($r.'@odata.type' -eq '#Microsoft.Dynamics.CRM.ManyToManyRelationshipMetadata')
                $tableLN  = if ($isM2M) { $r.Entity1LogicalName } else { $r.ReferencedEntity }
                if ($tablesFilterSet -and -not $tablesFilterSet.Contains($tableLN) -and -not ($isM2M -and $tablesFilterSet.Contains($r.Entity2LogicalName))) { continue }

                $allResults.Add([PSCustomObject][ordered]@{
                    ComponentType            = 'Relationship'
                    TableLogicalName         = $tableLN
                    TableDisplayName         = ''
                    TableSchemaName          = ''
                    AttributeLogicalName     = ''
                    AttributeSchemaName      = ''
                    AttributeDisplayName     = ''
                    RelationshipSchemaName   = $r.SchemaName
                    IsCustomComponent        = [bool]$r.IsCustomRelationship
                    SolutionUniqueName       = $sol.uniquename
                    SolutionFriendlyName     = $sol.friendlyname
                    SolutionVersion          = $sol.version
                    SolutionIsManaged        = [bool]$sol.ismanaged
                    SolutionInstalledOn      = $sol.installedon
                    RootComponentBehavior    = $c.'rootcomponentbehavior@OData.Community.Display.V1.FormattedValue'
                }) | Out-Null
            }
        }
    }

    $sorted = $allResults | Sort-Object ComponentType, TableLogicalName, AttributeLogicalName, RelationshipSchemaName, SolutionUniqueName

    $byType = $sorted | Group-Object ComponentType
    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total membership rows: $($sorted.Count)"
    foreach ($g in $byType) {
        Write-Host "  $($g.Name): $($g.Count)"
    }
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
                $OutputPath = Join-Path (Get-Location) "solutionmembership_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "solutionmembership_$timestamp.json"
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
