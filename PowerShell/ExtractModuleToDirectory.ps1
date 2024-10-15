<#
.SYNOPSIS
Extracts a PowerShell module ZIP file to a selected PowerShell module directory.

.DESCRIPTION
This script is used on an offline system to extract the contents of a PowerShell module ZIP file to a specified PowerShell module directory. 
The script lists all the available PowerShell module directories, allows the user to select one, and then extracts the ZIP file contents to that directory. 
If the -Force parameter is specified, existing files will be overwritten.

.PARAMETER ZipFilePath
The path to the ZIP file containing the PowerShell module.

.PARAMETER Force
If specified, existing files in the target directory will be overwritten.

.EXAMPLE
.\ExtractModuleToDirectory.ps1 -ZipFilePath "C:\Modules\PSReadline-2.2.0.zip"

Lists available PowerShell module directories, prompts the user to select one, and extracts the contents of PSReadline-2.2.0.zip to the selected directory.

.EXAMPLE
.\ExtractModuleToDirectory.ps1 -ZipFilePath "C:\Modules\PSReadline-2.2.0.zip" -Force

Lists available PowerShell module directories, prompts the user to select one, and extracts the contents of PSReadline-2.2.0.zip to the selected directory, overwriting existing files if they already exist.
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "The path to the ZIP file containing the PowerShell module.")]
    [string]$ZipFilePath,

    [Parameter(HelpMessage = "If specified, existing files in the target directory will be overwritten.")]
    [switch]$Force
)

# Function to get all available PowerShell module directories
function Get-ModulePaths {
    $modulePaths = $env:PSModulePath -split ';'
    return $modulePaths
}

# Function to display a menu and get user selection
function Show-Menu {
    param (
        [string[]]$MenuItems
    )

    for ($i = 0; $i -lt $MenuItems.Length; $i++) {
        Write-Host ("[{0}] {1}" -f ($i + 1), $MenuItems[$i])
    }

    $selection = Read-Host "Please select a directory (enter the number)"
    return [int]$selection - 1
}

# Verify that the ZIP file exists
if (-Not (Test-Path -Path $ZipFilePath)) {
    Write-Error "The ZIP file '$ZipFilePath' does not exist."
    exit 1
}

# Get available PowerShell module directories
$modulePaths = Get-ModulePaths

# Display the menu and get user selection
Write-Output "Available PowerShell module directories:"
$selectionIndex = Show-Menu -MenuItems $modulePaths

# Validate the selection
if ($selectionIndex -lt 0 -or $selectionIndex -ge $modulePaths.Length) {
    Write-Error "Invalid selection. Exiting."
    exit 1
}

$targetDir = $modulePaths[$selectionIndex]

# Ensure the target directory exists
if (-Not (Test-Path -Path $targetDir)) {
    Write-Output "The directory '$targetDir' does not exist. Creating it..."
    New-Item -ItemType Directory -Path $targetDir -Force
}

# Extract the ZIP file to the selected directory
Write-Output "Extracting the contents of '$ZipFilePath' to '$targetDir'..."
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFilePath, $targetDir, $Force)
    Write-Output "Extraction complete."
} catch {
    Write-Error "An error occurred while extracting the ZIP file: $_"
    if ($_.Exception.Message -match "already exists") {
        Write-Output "Consider using the -Force parameter to overwrite existing files."
    }
}
