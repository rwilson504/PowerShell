<#
.SYNOPSIS
    Shared helper for resolving a Dataverse solution unique name to its set of in-scope
    table logical names. Dot-source from any analysis script that wants to add a
    -SolutionUniqueName scope filter.

.DESCRIPTION
    Returns ALL tables that have an Entity component (componenttype=1) inside the named
    solution. This is the "filter to my solution" use case: when you add a custom table
    to a solution, the table itself is added as an Entity component (typically with
    'Include all subcomponents'). Filtering by Entity-typed components captures that case
    cleanly with just three paged metadata calls.

    NOTE: If a solution contains ONLY individual Attribute or Relationship components
    (without the parent Entity), those tables are NOT returned by this helper. Use
    GetSolutionMembership.ps1 directly for full attribute / relationship-level membership.
#>

function Resolve-SolutionTables {
    <#
    .SYNOPSIS
        Resolves a solution unique name to the set of table logical names whose Entity
        component is inside the solution.

    .OUTPUTS
        PSCustomObject with:
          Solution : the solution record (solutionid, uniquename, friendlyname, ismanaged, version)
          Tables   : sorted array of distinct logical names
    #>
    param (
        [Parameter(Mandatory = $true)] [string]$OrgUrl,
        [Parameter(Mandatory = $true)] [hashtable]$Headers,
        [Parameter(Mandatory = $true)] [string]$UniqueName
    )

    # 1. Look up the solution
    $solUrl = "$OrgUrl/api/data/v9.2/solutions?" +
        "`$select=solutionid,uniquename,friendlyname,ismanaged,version,installedon" +
        "&`$filter=uniquename eq '$UniqueName'"
    try {
        $sol = Invoke-RestMethod -Uri $solUrl -Headers $Headers -Method Get
    }
    catch {
        Write-Error "Failed to query solutions: $_"
        return $null
    }
    if (-not $sol.value -or $sol.value.Count -eq 0) {
        Write-Error "Solution '$UniqueName' not found in environment."
        return $null
    }
    $solution = $sol.value[0]

    # 2. Get Entity components only (componenttype=1)
    $compUrl = "$OrgUrl/api/data/v9.2/solutioncomponents?" +
        "`$select=objectid" +
        "&`$filter=_solutionid_value eq $($solution.solutionid) and componenttype eq 1"
    $entityIds = New-Object System.Collections.Generic.HashSet[string]
    try {
        do {
            $r = Invoke-RestMethod -Uri $compUrl -Headers $Headers -Method Get
            foreach ($c in $r.value) { [void]$entityIds.Add([string]$c.objectid) }
            $compUrl = $r.'@odata.nextLink'
        } while ($compUrl)
    }
    catch {
        Write-Warning "Failed to query solutioncomponents for '$UniqueName': $_"
        return [PSCustomObject]@{ Solution = $solution; Tables = @() }
    }

    if ($entityIds.Count -eq 0) {
        Write-Warning "Solution '$UniqueName' contains no Entity components. (Add the table directly to the solution to scope by it. For attribute/relationship-only solutions, use GetSolutionMembership.ps1 to enumerate components.)"
        return [PSCustomObject]@{ Solution = $solution; Tables = @() }
    }

    # 3. Resolve entity MetadataIds to LogicalNames via a single paged metadata scan
    $tables = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $entUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions?`$select=MetadataId,LogicalName"
    try {
        do {
            $r = Invoke-RestMethod -Uri $entUrl -Headers $Headers -Method Get
            foreach ($e in $r.value) {
                if ($entityIds.Contains([string]$e.MetadataId)) {
                    [void]$tables.Add($e.LogicalName)
                }
            }
            $entUrl = $r.'@odata.nextLink'
        } while ($entUrl)
    }
    catch {
        Write-Warning "Failed to load EntityDefinitions while resolving '$UniqueName': $_"
    }

    return [PSCustomObject]@{
        Solution = $solution
        Tables   = ($tables | Sort-Object)
    }
}

function Resolve-SolutionScopedTables {
    <#
    .SYNOPSIS
        Helper for analysis scripts. Takes the user-supplied -Tables array and the
        -SolutionUniqueName (either may be empty/null) and returns the resolved final
        list of tables to process. Writes informational console messages.

    .OUTPUTS
        Final array of table logical names. May be empty if intersection finds nothing.
    #>
    param (
        [Parameter(Mandatory = $true)] [string]$OrgUrl,
        [Parameter(Mandatory = $true)] [hashtable]$Headers,
        [string[]]$Tables,
        [string]$SolutionUniqueName
    )

    if ([string]::IsNullOrWhiteSpace($SolutionUniqueName)) {
        return $Tables
    }

    Write-Host "Resolving solution scope for '$SolutionUniqueName'..." -ForegroundColor Cyan
    $scope = Resolve-SolutionTables -OrgUrl $OrgUrl -Headers $Headers -UniqueName $SolutionUniqueName
    if (-not $scope) { return @() }

    $solLabel = "$($scope.Solution.uniquename) ($($scope.Solution.friendlyname))"
    Write-Host "  Solution '$solLabel' contains $($scope.Tables.Count) Entity component(s)." -ForegroundColor Green

    if (-not $Tables -or $Tables.Count -eq 0) {
        return @($scope.Tables)
    }

    # Intersect user's -Tables with the solution's tables
    $solSet = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($scope.Tables | ForEach-Object { $_.ToLowerInvariant() }),
        [System.StringComparer]::OrdinalIgnoreCase)
    $intersected = @($Tables | Where-Object { $solSet.Contains($_) })
    Write-Host "  Intersection with -Tables ($($Tables.Count) requested): $($intersected.Count) remain." -ForegroundColor Cyan
    return $intersected
}
