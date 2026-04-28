<#
.SYNOPSIS
    Gets the 1:N, N:1, and M:M relationships between tables in Dataverse.

.DESCRIPTION
    Queries the Dataverse RelationshipDefinitions metadata endpoint to enumerate
    every relationship in the environment, then emits one row per relationship
    direction so each table's outbound relationships are easy to review.

    For each relationship the script returns:
      - The "this side" table (logical + display name)
      - The relationship type (1:N, N:1, M:M)
      - The related table on the other side (logical + display name)
      - The relationship schema name
      - For 1:N / N:1: the lookup attribute (referencing column on the N side)
      - For M:M: the intersect (link) table name
      - IsCustomRelationship and IsManaged flags

    A 1:N relationship is emitted twice in the output: once as "1:N" from the
    parent (referenced) table's perspective and once as "N:1" from the child
    (referencing) table's perspective. M:M relationships are also emitted twice,
    once for each side.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    Optional array of table logical names. When supplied, only relationships in
    which at least one side is one of these tables are returned. When omitted,
    every relationship in the environment is returned.

.PARAMETER RelationshipTypes
    Which relationship types to include. Valid values are "OneToMany" (1:N),
    "ManyToOne" (N:1), and "ManyToMany" (M:M). Default is all three.

.PARAMETER CustomEntitiesOnly
    When set, restrict the output to relationships where BOTH sides are custom
    entities (IsCustomEntity eq true). Useful for auditing relationships between
    your own tables without out-of-the-box noise.

.PARAMETER CustomRelationshipsOnly
    When set, only include relationships where IsCustomRelationship is true.
    This is independent of -CustomEntitiesOnly: a custom relationship can exist
    between two system tables (e.g., a custom N:N you added to account/contact).

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results. If not provided, results are
    written to the console (or a timestamped file in the current directory for
    CSV/JSON when no path is supplied).

.EXAMPLE
    .\GetTableRelationships.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token

    Lists every relationship in the environment.

.EXAMPLE
    .\GetTableRelationships.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables @("account","contact") -OutputFormat CSV

    Lists every relationship that involves account or contact and exports to CSV.

.EXAMPLE
    .\GetTableRelationships.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -CustomRelationshipsOnly -RelationshipTypes ManyToMany

    Lists only custom many-to-many relationships in the environment.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $false)]
    [string[]]$Tables,

    [Parameter(Mandatory = $false)]
    [ValidateSet("OneToMany", "ManyToOne", "ManyToMany")]
    [string[]]$RelationshipTypes = @("OneToMany", "ManyToOne", "ManyToMany"),

    [Parameter(Mandatory = $false)]
    [switch]$CustomEntitiesOnly,

    [Parameter(Mandatory = $false)]
    [switch]$CustomRelationshipsOnly,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    "Accept"           = "application/json"
    "Content-Type"     = "application/json; charset=utf-8"
    "Prefer"           = "odata.include-annotations=*"
}

function Get-EntityDisplayNameMap {
    <#
    .SYNOPSIS
        Returns a hashtable mapping LogicalName -> @{DisplayName; IsCustomEntity}.
        Used to enrich relationship output with friendly table names.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers
    )

    Write-Host "Loading entity display names..." -ForegroundColor Cyan
    $url = "$OrgUrl/api/data/v9.2/EntityDefinitions?" +
        "`$select=LogicalName,DisplayName,IsCustomEntity"

    $map = @{}
    do {
        $response = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        foreach ($entity in $response.value) {
            $displayName = if ($entity.DisplayName.UserLocalizedLabel) {
                $entity.DisplayName.UserLocalizedLabel.Label
            } else {
                $entity.LogicalName
            }
            $map[$entity.LogicalName] = [PSCustomObject]@{
                DisplayName    = $displayName
                IsCustomEntity = [bool]$entity.IsCustomEntity
            }
        }
        $url = $response.'@odata.nextLink'
    } while ($url)

    Write-Host "Loaded $($map.Count) entity name mappings." -ForegroundColor Green
    return $map
}

function Get-AllRelationships {
    <#
    .SYNOPSIS
        Pulls every relationship in the environment from the RelationshipDefinitions
        endpoint, following pagination links. Returns the raw metadata records
        (mixed OneToMany and ManyToMany shapes).
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [bool]$IncludeOneToMany,
        [bool]$IncludeManyToMany
    )

    Write-Host "Loading relationship definitions..." -ForegroundColor Cyan
    $all = @()
    $url = "$OrgUrl/api/data/v9.2/RelationshipDefinitions"

    do {
        $response = Invoke-RestMethod -Uri $url -Headers $Headers -Method Get
        $all += $response.value
        $url = $response.'@odata.nextLink'
    } while ($url)

    # Filter by relationship kind based on the @odata.type discriminator
    $filtered = $all | Where-Object {
        $type = $_.'@odata.type'
        ($IncludeOneToMany -and $type -eq '#Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata') -or
        ($IncludeManyToMany -and $type -eq '#Microsoft.Dynamics.CRM.ManyToManyRelationshipMetadata')
    }

    Write-Host "Loaded $($all.Count) total relationship(s) ($($filtered.Count) match the requested types)." -ForegroundColor Green
    return $filtered
}

function Get-DisplayName {
    param (
        [hashtable]$NameMap,
        [string]$LogicalName
    )
    if ($NameMap.ContainsKey($LogicalName)) {
        return $NameMap[$LogicalName].DisplayName
    }
    return $LogicalName
}

function Test-IsCustomEntity {
    param (
        [hashtable]$NameMap,
        [string]$LogicalName
    )
    if ($NameMap.ContainsKey($LogicalName)) {
        return [bool]$NameMap[$LogicalName].IsCustomEntity
    }
    return $false
}

# Main script execution
try {
    $includeOneToMany  = ($RelationshipTypes -contains "OneToMany") -or ($RelationshipTypes -contains "ManyToOne")
    $includeManyToMany = ($RelationshipTypes -contains "ManyToMany")
    $emit1ToN          = ($RelationshipTypes -contains "OneToMany")
    $emitNTo1          = ($RelationshipTypes -contains "ManyToOne")
    $emitMtoM          = ($RelationshipTypes -contains "ManyToMany")

    $nameMap       = Get-EntityDisplayNameMap -OrgUrl $OrganizationUrl -Headers $headers
    $relationships = Get-AllRelationships -OrgUrl $OrganizationUrl -Headers $headers `
        -IncludeOneToMany $includeOneToMany -IncludeManyToMany $includeManyToMany

    # Normalize -Tables filter (case-insensitive)
    $tablesFilterSet = $null
    if ($Tables -and $Tables.Count -gt 0) {
        $tablesFilterSet = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@($Tables | ForEach-Object { $_.ToLowerInvariant() }),
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    Write-Host "Building output rows..." -ForegroundColor Cyan
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($rel in $relationships) {
        $type = $rel.'@odata.type'

        if ($type -eq '#Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata') {
            $referenced  = $rel.ReferencedEntity      # parent / "1" side
            $referencing = $rel.ReferencingEntity     # child  / "N" side
            $fkAttribute = $rel.ReferencingAttribute  # FK column on the child

            $isCustomRel = [bool]$rel.IsCustomRelationship
            $isManaged   = [bool]$rel.IsManaged

            if ($CustomRelationshipsOnly -and -not $isCustomRel) { continue }

            $referencedIsCustom  = Test-IsCustomEntity -NameMap $nameMap -LogicalName $referenced
            $referencingIsCustom = Test-IsCustomEntity -NameMap $nameMap -LogicalName $referencing

            if ($CustomEntitiesOnly -and (-not $referencedIsCustom -or -not $referencingIsCustom)) { continue }

            # Emit 1:N from the parent's perspective
            if ($emit1ToN) {
                if (-not $tablesFilterSet -or $tablesFilterSet.Contains($referenced) -or $tablesFilterSet.Contains($referencing)) {
                    $results.Add([PSCustomObject][ordered]@{
                        TableLogicalName        = $referenced
                        TableDisplayName        = Get-DisplayName -NameMap $nameMap -LogicalName $referenced
                        RelationshipType        = '1:N'
                        RelatedTableLogicalName = $referencing
                        RelatedTableDisplayName = Get-DisplayName -NameMap $nameMap -LogicalName $referencing
                        RelationshipSchemaName  = $rel.SchemaName
                        LookupAttribute         = $fkAttribute
                        IntersectTable          = $null
                        IsCustomRelationship    = $isCustomRel
                        IsManaged               = $isManaged
                    })
                }
            }

            # Emit N:1 from the child's perspective
            if ($emitNTo1) {
                if (-not $tablesFilterSet -or $tablesFilterSet.Contains($referencing) -or $tablesFilterSet.Contains($referenced)) {
                    $results.Add([PSCustomObject][ordered]@{
                        TableLogicalName        = $referencing
                        TableDisplayName        = Get-DisplayName -NameMap $nameMap -LogicalName $referencing
                        RelationshipType        = 'N:1'
                        RelatedTableLogicalName = $referenced
                        RelatedTableDisplayName = Get-DisplayName -NameMap $nameMap -LogicalName $referenced
                        RelationshipSchemaName  = $rel.SchemaName
                        LookupAttribute         = $fkAttribute
                        IntersectTable          = $null
                        IsCustomRelationship    = $isCustomRel
                        IsManaged               = $isManaged
                    })
                }
            }
        }
        elseif ($type -eq '#Microsoft.Dynamics.CRM.ManyToManyRelationshipMetadata' -and $emitMtoM) {
            $entity1     = $rel.Entity1LogicalName
            $entity2     = $rel.Entity2LogicalName
            $intersect   = $rel.IntersectEntityName
            $isCustomRel = [bool]$rel.IsCustomRelationship
            $isManaged   = [bool]$rel.IsManaged

            if ($CustomRelationshipsOnly -and -not $isCustomRel) { continue }

            $entity1IsCustom = Test-IsCustomEntity -NameMap $nameMap -LogicalName $entity1
            $entity2IsCustom = Test-IsCustomEntity -NameMap $nameMap -LogicalName $entity2
            if ($CustomEntitiesOnly -and (-not $entity1IsCustom -or -not $entity2IsCustom)) { continue }

            # Emit one row from each side
            if (-not $tablesFilterSet -or $tablesFilterSet.Contains($entity1) -or $tablesFilterSet.Contains($entity2)) {
                $results.Add([PSCustomObject][ordered]@{
                    TableLogicalName        = $entity1
                    TableDisplayName        = Get-DisplayName -NameMap $nameMap -LogicalName $entity1
                    RelationshipType        = 'M:M'
                    RelatedTableLogicalName = $entity2
                    RelatedTableDisplayName = Get-DisplayName -NameMap $nameMap -LogicalName $entity2
                    RelationshipSchemaName  = $rel.SchemaName
                    LookupAttribute         = $null
                    IntersectTable          = $intersect
                    IsCustomRelationship    = $isCustomRel
                    IsManaged               = $isManaged
                })
                $results.Add([PSCustomObject][ordered]@{
                    TableLogicalName        = $entity2
                    TableDisplayName        = Get-DisplayName -NameMap $nameMap -LogicalName $entity2
                    RelationshipType        = 'M:M'
                    RelatedTableLogicalName = $entity1
                    RelatedTableDisplayName = Get-DisplayName -NameMap $nameMap -LogicalName $entity1
                    RelationshipSchemaName  = $rel.SchemaName
                    LookupAttribute         = $null
                    IntersectTable          = $intersect
                    IsCustomRelationship    = $isCustomRel
                    IsManaged               = $isManaged
                })
            }
        }
    }

    # Sort for stable output: by table, then type, then related table
    $sorted = $results | Sort-Object TableLogicalName, RelationshipType, RelatedTableLogicalName

    # Summary
    $oneToManyCount  = ($sorted | Where-Object RelationshipType -eq '1:N').Count
    $manyToOneCount  = ($sorted | Where-Object RelationshipType -eq 'N:1').Count
    $manyToManyCount = ($sorted | Where-Object RelationshipType -eq 'M:M').Count
    $distinctTables  = ($sorted | Select-Object -ExpandProperty TableLogicalName -Unique).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total relationship rows: $($sorted.Count)"
    Write-Host "  1:N rows: $oneToManyCount"
    Write-Host "  N:1 rows: $manyToOneCount"
    Write-Host "  M:M rows: $manyToManyCount"
    Write-Host "Distinct tables represented: $distinctTables"
    Write-Host ""

    # Output
    switch ($OutputFormat) {
        "Table" {
            if ($OutputPath) {
                $sorted | Format-Table -AutoSize | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $sorted | Format-Table -AutoSize
            }
        }
        "CSV" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "relationships_$timestamp.csv"
            }
            $sorted | Export-Csv -Path $OutputPath -NoTypeInformation
            Write-Host "Results exported to $OutputPath" -ForegroundColor Green
        }
        "JSON" {
            if (-not $OutputPath) {
                $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
                $OutputPath = Join-Path (Get-Location) "relationships_$timestamp.json"
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
