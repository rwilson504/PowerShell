<#
.SYNOPSIS
    Converts a Power Platform XML sitemap into a CSV file that can be opened in Excel.

.DESCRIPTION
    This script reads an XML sitemap file and extracts data from its hierarchical structure
    (Area → Group → SubArea). It retrieves attributes such as Area ID, Area Title, Group ID,
    Group Title, SubArea ID, SubArea Title, Entity, and URL. The extracted information is then
    exported to a CSV file. This CSV file can be opened with Excel for further analysis or reporting.

.PARAMETER XmlPath
    The full file path of the XML sitemap to be processed.

.PARAMETER CsvPath
    The full file path where the CSV output should be saved.

.EXAMPLE
    PS C:\> .\ConvertSitemapToCsv.ps1 -XmlPath "C:\Data\sitemap.xml" -CsvPath "C:\Data\sitemap.csv"
    Export complete. CSV file saved to: C:\Data\sitemap.csv

.NOTES
    Author: Rick Wilson
    Date  : 2025-02-04
    Version: 1.1
    Requirements: PowerShell 3.0 or later.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$XmlPath,

    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

# Load the XML file.
[xml]$sitemap = Get-Content -Path $XmlPath

# Create an empty array to hold the extracted records.
$rows = @()

# Loop through each Area node.
foreach ($area in $sitemap.SiteMap.Area) {

    # Get the Area Id and its title (using LCID=1033).
    $areaId = $area.Id
    $areaTitle = ($area.Titles.Title | Where-Object { $_.LCID -eq "1033" }).Title

    # Loop through each Group node within the Area.
    foreach ($group in $area.Group) {

        $groupId = $group.Id
        $groupTitle = ($group.Titles.Title | Where-Object { $_.LCID -eq "1033" }).Title

        # Loop through each SubArea node within the Group.
        foreach ($subArea in $group.SubArea) {

            $subAreaId = $subArea.Id

            # Some SubArea nodes have Titles; if not, leave blank.
            $subAreaTitle = ""
            if ($subArea.Titles -and $subArea.Titles.Title) {
                $subAreaTitle = ($subArea.Titles.Title | Where-Object { $_.LCID -eq "1033" }).Title
            }

            # Get the Entity and Url attributes if they exist.
            $entity = $subArea.Entity
            $url    = $subArea.Url

            # Create a custom object for this record.
            $rows += [PSCustomObject]@{
                AreaId       = $areaId
                AreaTitle    = $areaTitle
                GroupId      = $groupId
                GroupTitle   = $groupTitle
                SubAreaId    = $subAreaId
                SubAreaTitle = $subAreaTitle
                Entity       = $entity
                URL          = $url
            }
        }
    }
}

# Export the data to a CSV file.
$rows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Export complete. CSV file saved to: $CsvPath"
