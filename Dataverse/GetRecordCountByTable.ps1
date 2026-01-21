<#
.SYNOPSIS
    Gets record counts for specified tables or all readable tables in Dataverse.

.DESCRIPTION
    This script retrieves the record count for each specified table in Dataverse.
    It uses FetchXML aggregate queries to bypass the 5,000 record limit.
    If no tables are specified, it retrieves metadata to find all readable tables
    and gets counts for each one.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER Tables
    An optional array of table logical names to get counts for.
    If not provided, all readable tables will be queried.

.PARAMETER IncludeSystemTables
    When querying all tables, include system tables (those starting with 'sys').
    Default is $false.

.PARAMETER OutputFormat
    The output format. Valid values are "Table", "CSV", "JSON". Default is "Table".

.PARAMETER OutputPath
    Optional file path to export the results. If not provided, results are written to the console.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -Tables @("account", "contact", "lead")

    Gets record counts for the account, contact, and lead tables.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token

    Gets record counts for all readable tables in the environment.

.EXAMPLE
    .\GetRecordCountByTable.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -OutputFormat CSV -OutputPath "C:\temp\recordcounts.csv"

    Gets record counts for all readable tables and exports to CSV.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,
    
    [Parameter(Mandatory = $true)]
    [string]$AccessToken,
    
    [Parameter(Mandatory = $false)]
    [string[]]$Tables,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemTables = $false,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Table", "CSV", "JSON")]
    [string]$OutputFormat = "Table",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

# Set up headers for API calls
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
    "Accept" = "application/json"
    "Content-Type" = "application/json; charset=utf-8"
    "Prefer" = "odata.include-annotations=*"
}

function Get-AllReadableTables {
    <#
    .SYNOPSIS
        Retrieves all tables that the current user has read access to.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [bool]$IncludeSystem
    )

    Write-Host "Retrieving metadata for all tables..." -ForegroundColor Cyan
    
    # Query EntityDefinitions to get all entities with their read privileges
    # We filter for entities that are valid for read operations
    # Include DataProviderId and TableType to identify virtual tables that don't support aggregates
    # Include PrimaryIdAttribute and EntitySetName for accurate aggregate queries
    $metadataUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions?" + 
        "`$select=LogicalName,DisplayName,IsValidForAdvancedFind,CanCreateViews,IsCustomizable,IsActivity,IsActivityParty,DataProviderId,TableType,ExternalName,PrimaryIdAttribute,EntitySetName" +
        "&`$filter=IsValidForAdvancedFind eq true and IsIntersect eq false"
    
    try {
        $response = Invoke-RestMethod -Uri $metadataUrl -Headers $Headers -Method Get
        $entities = $response.value
        
        $readableTables = @()
        
        foreach ($entity in $entities) {
            $logicalName = $entity.LogicalName
            
            # Skip system tables if not requested
            if (-not $IncludeSystem -and $logicalName.StartsWith("sys")) {
                continue
            }
            
            # Skip certain internal tables that typically cause issues
            $skipTables = @(
                "abortedsystemjob",
                "actioncardusersettings", 
                "activityfileattachment",
                "attributeimageconfig",
                "canvasappextendedmetadata"
            )
            
            if ($logicalName -in $skipTables) {
                continue
            }
            
            # Tables with plugins that don't support aggregate queries
            # These have QueryExpressionConverter plugins or other restrictions
            $noAggregatePluginTables = @(
                "datalakefolder",
                "datalakefolderpermission",
                "datalakeworkspace",
                "datalakeworkspacepermission",
                "delegatedauthorization",
                "archivecleanupinfo",
                "governanceconfiguration",
                "msdyn_aimodelcatalog",
                "msdyn_aioptimizationprivatedata",
                "componentversion",
                "componentversiondatasource",
                "gitbranch",
                "gitorganization",
                "gitproject",
                "gitrepository",
                "gitsolution",
                "gitconfigurationretrievaldatasource"
            )
            
            $displayName = if ($entity.DisplayName.UserLocalizedLabel) { 
                $entity.DisplayName.UserLocalizedLabel.Label 
            } else { 
                $logicalName 
            }
            
            # Check if this is a virtual table (has DataProviderId) or special table type
            # Virtual tables and data source tables typically don't support FetchXML aggregates
            $isVirtual = $null -ne $entity.DataProviderId -and $entity.DataProviderId -ne [Guid]::Empty
            $tableType = $entity.TableType
            $hasExternalName = -not [string]::IsNullOrEmpty($entity.ExternalName)
            
            # Tables that likely don't support aggregates:
            # - Virtual tables (have DataProviderId)
            # - Tables with TableType = "Virtual" or "Elastic"
            # - Data source tables (usually end with 'ds' or 'datasource')
            # - Tables with known plugin restrictions
            $supportsAggregate = -not $isVirtual -and $tableType -ne "Virtual" -and $tableType -ne "Elastic"
            
            # Additional check for known data source table patterns
            if ($logicalName -match '(datasource|^datalake|nrddatasource)$') {
                $supportsAggregate = $false
            }
            
            # Check against known tables with plugin restrictions
            if ($logicalName -in $noAggregatePluginTables) {
                $supportsAggregate = $false
            }
            
            # Check for mspp_ (Power Pages) tables - many have aggregate restrictions
            if ($logicalName -match '^mspp_') {
                $supportsAggregate = $false
            }
            
            $readableTables += [PSCustomObject]@{
                LogicalName = $logicalName
                DisplayName = $displayName
                IsVirtual = $isVirtual
                TableType = $tableType
                SupportsAggregate = $supportsAggregate
                PrimaryIdAttribute = $entity.PrimaryIdAttribute
                EntitySetName = $entity.EntitySetName
            }
        }
        
        Write-Host "Found $($readableTables.Count) readable tables." -ForegroundColor Green
        return $readableTables | Sort-Object LogicalName
    }
    catch {
        Write-Error "Failed to retrieve table metadata: $_"
        throw
    }
}

function Get-TableRecordCount {
    <#
    .SYNOPSIS
        Gets the record count for a single table using FetchXML aggregate query.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [string]$TableLogicalName,
        [string]$PrimaryIdAttribute,
        [string]$EntitySetName
    )

    # Use provided PrimaryIdAttribute or fall back to convention
    if ([string]::IsNullOrEmpty($PrimaryIdAttribute)) {
        $PrimaryIdAttribute = "$($TableLogicalName)id"
    }
    
    # Build FetchXML aggregate query to get count
    # Using aggregate with count bypasses the 5000 record limit
    $fetchXml = @"
<fetch aggregate="true">
    <entity name="$TableLogicalName">
        <attribute name="$PrimaryIdAttribute" alias="recordcount" aggregate="count"/>
    </entity>
</fetch>
"@

    # URL encode the FetchXML
    $encodedFetch = [System.Web.HttpUtility]::UrlEncode($fetchXml)
    
    # Use provided EntitySetName or fetch it from metadata
    if ([string]::IsNullOrEmpty($EntitySetName)) {
        $entitySetUrl = "$OrgUrl/api/data/v9.2/EntityDefinitions(LogicalName='$TableLogicalName')?`$select=EntitySetName,PrimaryIdAttribute"
        
        try {
            $entityDef = Invoke-RestMethod -Uri $entitySetUrl -Headers $Headers -Method Get
            $EntitySetName = $entityDef.EntitySetName
            # Also update PrimaryIdAttribute if we had to fetch metadata anyway
            if ($PrimaryIdAttribute -eq "$($TableLogicalName)id" -and $entityDef.PrimaryIdAttribute) {
                $PrimaryIdAttribute = $entityDef.PrimaryIdAttribute
                # Rebuild FetchXML with correct attribute
                $fetchXml = @"
<fetch aggregate="true">
    <entity name="$TableLogicalName">
        <attribute name="$PrimaryIdAttribute" alias="recordcount" aggregate="count"/>
    </entity>
</fetch>
"@
                $encodedFetch = [System.Web.HttpUtility]::UrlEncode($fetchXml)
            }
        }
        catch {
            # If we can't get the entity set name, try common pluralization
            $EntitySetName = $TableLogicalName + "s"
        }
    }
    
    $fetchUrl = "$OrgUrl/api/data/v9.2/$EntitySetName`?fetchXml=$encodedFetch"
    
    try {
        $response = Invoke-RestMethod -Uri $fetchUrl -Headers $Headers -Method Get
        
        if ($response.value -and $response.value.Count -gt 0) {
            # The count is returned in the first record
            $count = $response.value[0].recordcount
            return [long]$count
        }
        return 0
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        $errorBody = $_.ErrorDetails.Message
        
        # Try to parse error details for better categorization
        $errorCode = $null
        $errorMessage = $null
        if ($errorBody) {
            try {
                $errorJson = $errorBody | ConvertFrom-Json
                $errorCode = $errorJson.error.code
                $errorMessage = $errorJson.error.message
            } catch { }
        }
        
        if ($statusCode -eq 403) {
            # User doesn't have read permission
            return -1
        }
        elseif ($errorMessage -match "aggregates aren't supported" -or $errorMessage -match "aggregate operation is requested") {
            # Table has plugin that doesn't support aggregates
            Write-Warning "Table '$TableLogicalName' does not support aggregate queries (plugin restriction)"
            return -4
        }
        elseif ($errorMessage -match "Resource not found for the segment") {
            # EntitySetName doesn't match actual API endpoint
            Write-Warning "Table '$TableLogicalName' has invalid EntitySetName in metadata"
            return -5
        }
        elseif ($statusCode -eq 400) {
            # Bad request - possibly invalid entity or attribute
            Write-Warning "Unable to query table '$TableLogicalName': Bad request"
            return -2
        }
        else {
            Write-Warning "Error querying table '$TableLogicalName': $errorMessage"
            return -3
        }
    }
}

function Get-RecordCounts {
    <#
    .SYNOPSIS
        Main function to get record counts for all specified tables.
    #>
    param (
        [string]$OrgUrl,
        [hashtable]$Headers,
        [array]$TableList
    )

    $results = @()
    $totalTables = $TableList.Count
    $currentIndex = 0
    
    foreach ($table in $TableList) {
        $currentIndex++
        $logicalName = if ($table -is [PSCustomObject]) { $table.LogicalName } else { $table }
        $displayName = if ($table -is [PSCustomObject] -and $table.DisplayName) { $table.DisplayName } else { $logicalName }
        $supportsAggregate = if ($table -is [PSCustomObject] -and $null -ne $table.SupportsAggregate) { $table.SupportsAggregate } else { $true }
        $isVirtual = if ($table -is [PSCustomObject] -and $null -ne $table.IsVirtual) { $table.IsVirtual } else { $false }
        $tableType = if ($table -is [PSCustomObject] -and $table.TableType) { $table.TableType } else { "Standard" }
        $primaryIdAttribute = if ($table -is [PSCustomObject] -and $table.PrimaryIdAttribute) { $table.PrimaryIdAttribute } else { $null }
        $entitySetName = if ($table -is [PSCustomObject] -and $table.EntitySetName) { $table.EntitySetName } else { $null }
        
        Write-Progress -Activity "Getting record counts" -Status "Processing $logicalName ($currentIndex of $totalTables)" -PercentComplete (($currentIndex / $totalTables) * 100)
        
        # Skip virtual tables that don't support aggregates
        if (-not $supportsAggregate) {
            $results += [PSCustomObject]@{
                TableLogicalName = $logicalName
                TableDisplayName = $displayName
                RecordCount = "N/A"
                Status = "Virtual/DataSource (No Aggregate)"
                TableType = $tableType
                IsVirtual = $isVirtual
            }
            continue
        }
        
        $count = Get-TableRecordCount -OrgUrl $OrgUrl -Headers $Headers -TableLogicalName $logicalName -PrimaryIdAttribute $primaryIdAttribute -EntitySetName $entitySetName
        
        $status = switch ($count) {
            -1 { "No Read Permission" }
            -2 { "Invalid Table/Attribute" }
            -3 { "Error" }
            -4 { "No Aggregate Support (Plugin)" }
            -5 { "Invalid EntitySetName" }
            default { "Success" }
        }
        
        $results += [PSCustomObject]@{
            TableLogicalName = $logicalName
            TableDisplayName = $displayName
            RecordCount = if ($count -lt 0) { "N/A" } else { $count }
            Status = $status
            TableType = $tableType
            IsVirtual = $isVirtual
        }
    }
    
    Write-Progress -Activity "Getting record counts" -Completed
    return $results
}

# Main script execution
try {
    # Add System.Web assembly for URL encoding
    Add-Type -AssemblyName System.Web

    # Determine which tables to query
    if ($Tables -and $Tables.Count -gt 0) {
        Write-Host "Processing $($Tables.Count) specified table(s)..." -ForegroundColor Cyan
        $tablesToQuery = $Tables
    }
    else {
        # Get all readable tables from metadata
        $tablesToQuery = Get-AllReadableTables -OrgUrl $OrganizationUrl -Headers $headers -IncludeSystem $IncludeSystemTables
    }

    # Get record counts for all tables
    $results = Get-RecordCounts -OrgUrl $OrganizationUrl -Headers $headers -TableList $tablesToQuery

    # Calculate summary statistics
    $successfulResults = $results | Where-Object { $_.Status -eq "Success" }
    $virtualTables = $results | Where-Object { $_.Status -eq "Virtual/DataSource (No Aggregate)" }
    $pluginRestricted = $results | Where-Object { $_.Status -eq "No Aggregate Support (Plugin)" }
    $invalidEntitySet = $results | Where-Object { $_.Status -eq "Invalid EntitySetName" }
    $errorTables = $results | Where-Object { $_.Status -in @("Error", "Invalid Table/Attribute", "No Read Permission") }
    $totalRecords = ($successfulResults | Measure-Object -Property RecordCount -Sum).Sum
    $tablesWithData = ($successfulResults | Where-Object { $_.RecordCount -gt 0 }).Count

    Write-Host "`n=== Summary ===" -ForegroundColor Green
    Write-Host "Total tables found: $($results.Count)"
    Write-Host "Standard tables queried successfully: $($successfulResults.Count)" -ForegroundColor Green
    Write-Host "Virtual/DataSource/Elastic tables (skipped): $($virtualTables.Count)" -ForegroundColor Yellow
    if ($pluginRestricted.Count -gt 0) {
        Write-Host "Tables with plugin restrictions (no aggregate): $($pluginRestricted.Count)" -ForegroundColor Yellow
    }
    if ($invalidEntitySet.Count -gt 0) {
        Write-Host "Tables with invalid EntitySetName: $($invalidEntitySet.Count)" -ForegroundColor Yellow
    }
    Write-Host "Tables with other errors: $($errorTables.Count)" -ForegroundColor $(if ($errorTables.Count -gt 0) { "Red" } else { "Green" })
    Write-Host "Tables with data: $tablesWithData"
    Write-Host "Total records across all tables: $totalRecords"
    Write-Host ""

    # Output results based on format
    switch ($OutputFormat) {
        "Table" {
            if ($OutputPath) {
                $results | Format-Table -AutoSize | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $results | Sort-Object @{Expression={if($_.RecordCount -eq "N/A"){-1}else{[long]$_.RecordCount}}; Descending=$true} | Format-Table -AutoSize
            }
        }
        "CSV" {
            if ($OutputPath) {
                $results | Export-Csv -Path $OutputPath -NoTypeInformation
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $results | ConvertTo-Csv -NoTypeInformation
            }
        }
        "JSON" {
            $jsonOutput = $results | ConvertTo-Json -Depth 3
            if ($OutputPath) {
                $jsonOutput | Out-File -FilePath $OutputPath
                Write-Host "Results exported to $OutputPath" -ForegroundColor Green
            }
            else {
                $jsonOutput
            }
        }
    }

    # Return results for pipeline use
    return $results
}
catch {
    Write-Error "Script execution failed: $_"
    throw
}
