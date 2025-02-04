<#
.SYNOPSIS
    Extracts help page content from a Dynamics solution file's customizations.xml,
    including the table name (Entity) and record type (RecordType) from the msdyn_path,
    and optionally converting the HTML content to Rich Text (RTF) and/or plain text.

.DESCRIPTION
    This script accepts the path to a Dynamics solution file (a .zip archive) and an
    output CSV file path as parameters. It performs the following steps:
      1. Extracts the solution ZIP to a temporary folder.
      2. Searches for the customizations.xml file within the extracted content.
      3. Loads customizations.xml as XML.
      4. Finds all <msdyn_helppage> nodes.
      5. HTML‑decodes the content (from the <msdyn_content> element) using
         [System.Net.WebUtility]::HtmlDecode.
      6. Extracts the msdyn_path value and, if it contains "Entities", extracts:
           - The table name (the folder following "Entities/") into a new column called **Entity**.
           - The next folder (e.g. "Forms" or "Views") into a new column called **RecordType**.
      7. If the -ConvertToRichText switch is provided, converts the decoded HTML into RTF
         using Word automation and adds the result to a new column called **RtfContent**.
      8. If the -ConvertToPlainText switch is provided, converts the decoded HTML into plain text
         (adding a "- " prefix for list items and ensuring a space after each paragraph) into a new column
         called **PlainTextContent**.
      9. Exports the HelpPageId, DisplayName, msdyn_path, Entity, RecordType, decoded Content,
         and (if converted) RtfContent and/or PlainTextContent to a CSV file.

.PARAMETER SolutionFile
    The full path to the Dynamics solution file (ZIP archive).

.PARAMETER OutputFile
    The full path where the CSV output will be saved.

.PARAMETER ConvertToRichText
    Switch. If specified, the script attempts to convert the decoded HTML content to RTF.

.PARAMETER ConvertToPlainText
    Switch. If specified, the script converts the decoded HTML content to plain text by stripping HTML tags
    (adding a "- " prefix for list items and ensuring a space after each paragraph).

.EXAMPLE
    PS C:\> .\ExtractHelpPages.ps1 -SolutionFile "C:\Solutions\MySolution.zip" -OutputFile "C:\Output\HelpPages.csv" -ConvertToRichText -ConvertToPlainText

.NOTES
    Author: Your Name
    Date  : YYYY-MM-DD
    Version: 1.4
    Requirements:
        - PowerShell 5.0+ (for [System.IO.Compression.ZipFile])
        - Microsoft Word (if using -ConvertToRichText)
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SolutionFile,

    [Parameter(Mandatory = $true)]
    [string]$OutputFile,

    [switch]$ConvertToRichText,

    [switch]$ConvertToPlainText
)

# Function to convert HTML content to RTF using Word COM automation.
function Convert-HtmlToRtf {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HtmlContent,
        [Parameter(Mandatory = $true)]
        $WordApp
    )
    # Create temporary file names.
    $tempHtml = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ".htm")
    $tempRtf  = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName() + ".rtf")

    try {
        # Write the HTML content to a temporary HTML file.
        Set-Content -Path $tempHtml -Value $HtmlContent -Encoding UTF8

        # Open the HTML file in Word (read-only).
        $doc = $WordApp.Documents.Open($tempHtml, [Type]::Missing, $true)
        
        # Define the constant for RTF format.
        $wdFormatRTF = 6

        # Save the document as RTF using explicit type casting.
        $doc.SaveAs([object]$tempRtf, [object]$wdFormatRTF)
        
        # Close and release the document.
        $doc.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        $doc = $null
        
        # Force garbage collection and wait a moment to release file locks.
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Start-Sleep -Seconds 1

        # Read the RTF file as a single string.
        $rtf = Get-Content -Path $tempRtf -Raw
        return $rtf
    }
    catch {
        Write-Warning "Failed to convert HTML to RTF: $_"
        return ""
    }
    finally {
        # Attempt to remove temporary files.
        try {
            if (Test-Path $tempHtml) { Remove-Item $tempHtml -Force }
        }
        catch { Write-Warning "Failed to remove temp HTML file: $_" }
        try {
            if (Test-Path $tempRtf) { Remove-Item $tempRtf -Force }
        }
        catch { Write-Warning "Failed to remove temp RTF file: $_" }
    }
}

# Function to convert HTML to plain text by stripping tags and applying formatting.
function Convert-HtmlToPlainText {
    param (
        [Parameter(Mandatory = $true)]
        [string]$HtmlContent
    )
    $plain = $HtmlContent

    # Replace <li> tags with "- " and </li> tags with newline.
    $plain = $plain -replace "<li[^>]*>", "- "
    $plain = $plain -replace "</li>", "`n"

    # Replace <p> tags with nothing and </p> with newline and a space.
    $plain = $plain -replace "<p[^>]*>", ""
    $plain = $plain -replace "</p>", "`n "

    # Remove any remaining HTML tags.
    $plain = $plain -replace "<[^>]+>", ""

    # Decode HTML entities.
    $plain = [System.Net.WebUtility]::HtmlDecode($plain)

    # Normalize newlines and trim extra spaces.
    $plain = $plain -replace "`r", ""
    return $plain.Trim()
}

# Verify the solution file exists.
if (!(Test-Path -Path $SolutionFile)) {
    Write-Error "Solution file '$SolutionFile' does not exist."
    exit 1
}

# Create a temporary folder for extraction.
$tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $tempFolder | Out-Null

try {
    # Extract the solution ZIP into the temporary folder.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($SolutionFile, $tempFolder)
}
catch {
    Write-Error "Failed to extract the solution file: $_"
    exit 1
}

# Locate the customizations.xml file.
$customizationsPath = Get-ChildItem -Path $tempFolder -Filter "customizations.xml" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $customizationsPath) {
    Write-Error "customizations.xml not found in the extracted solution."
    Remove-Item -Recurse -Force $tempFolder
    exit 1
}

# Load customizations.xml into an XML document.
[xml]$xmlContent = Get-Content -Path $customizationsPath.FullName

# Select all msdyn_helppage nodes.
$helpPages = $xmlContent.SelectNodes("//msdyn_helppage")

if (-not $helpPages) {
    Write-Output "No msdyn_helppage nodes found in customizations.xml."
    Remove-Item -Recurse -Force $tempFolder
    exit 0
}

# If rich text conversion is requested, start Word COM automation.
if ($ConvertToRichText) {
    try {
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
    }
    catch {
        Write-Warning "Unable to start Word COM automation. Rich text conversion will be skipped."
        $ConvertToRichText = $false
    }
}

# Prepare an array to hold the output.
$output = @()

# Process each msdyn_helppage node.
foreach ($page in $helpPages) {
    # Retrieve attributes.
    $pageId      = $page.GetAttribute("msdyn_helppageid")
    $displayName = $page.msdyn_displayname

    # Get and HTML-decode the content.
    $rawContent     = $page.msdyn_content
    $decodedContent = [System.Net.WebUtility]::HtmlDecode($rawContent)

    # Extract msdyn_path.
    $path = $page.msdyn_path

    # Determine the Entity from the path if it contains "Entities".
    $entity = ""
    if ($path -match "Entities/([^/]+)/") {
        $entity = $matches[1]
    }

    # Determine the RecordType from the path (e.g., "Views" or "Forms").
    $recordType = ""
    if ($path -match "Entities/[^/]+/([^/]+)/") {
        $recordType = $matches[1]
    }

    # Initialize variables for RTF and PlainText.
    $rtfContent   = $null
    $plainContent = $null

    # Convert to RTF if requested.
    if ($ConvertToRichText -and $wordApp) {
        $rtfContent = Convert-HtmlToRtf -HtmlContent $decodedContent -WordApp $wordApp
    }

    # Convert to plain text if requested.
    if ($ConvertToPlainText) {
        $plainContent = Convert-HtmlToPlainText -HtmlContent $decodedContent
    }

    # Append the collected information.
    $output += [PSCustomObject]@{
        HelpPageId       = $pageId
        DisplayName      = $displayName
        Path             = $path
        Entity           = $entity
        RecordType       = $recordType
        Content          = $decodedContent
        RtfContent       = $rtfContent
        PlainTextContent = $plainContent
    }
}

# Export the results to the CSV file.
$output | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

Write-Output "Extraction complete. CSV file saved to: $OutputFile"

# Cleanup: Close Word if used and remove the temporary folder.
if ($ConvertToRichText -and $wordApp) {
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    $wordApp = $null
}
Remove-Item -Path $tempFolder -Recurse -Force
