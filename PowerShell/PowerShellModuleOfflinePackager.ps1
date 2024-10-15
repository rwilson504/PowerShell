<#
.SYNOPSIS
Downloads the most current version of a PowerShell NuGet package or a specified version and extracts it for manual installation.

.DESCRIPTION
This script downloads the specified or latest version of a NuGet package from the PowerShell Gallery and extracts it to a specified or default output directory. 
If no version is specified, the latest version of the package is downloaded. If no output directory is specified, the script 
creates a default output folder in the script's directory. After extracting the packages, the script can optionally compress 
the package folders into a single ZIP file for easy transfer to an offline system. If the zipping is done, it also deletes the 
folders created by Save-Package unless the SkipZip parameter is specified.

.PARAMETER Name
The ID of the NuGet package to download.

.PARAMETER Version
The version of the NuGet package to download. If not specified, the latest version will be downloaded.

.PARAMETER OutputDir
The directory where the package will be downloaded and extracted. If not specified, a folder named 'NuGetPackages' will be 
created in the directory where the script is being run.

.PARAMETER SkipZip
If specified, the script will not compress the package folders into a ZIP file.

.EXAMPLE
.\PowerShellModuleOfflinePackager.ps1 -Name "PSReadline"

Downloads the latest version of the PSReadline package, extracts it to the default output directory, and compresses the extracted files into a ZIP file.

.EXAMPLE
.\PowerShellModuleOfflinePackager.ps1 -Name "PSReadline" -Version "2.2.0"

Downloads version 2.2.0 of the PSReadline package, extracts it to the default output directory, and compresses the extracted files into a ZIP file.

.EXAMPLE
.\PowerShellModuleOfflinePackager.ps1 -Name "PSReadline" -OutputDir "C:\MyPackages"

Downloads the latest version of the PSReadline package, extracts it to C:\MyPackages, and compresses the extracted files into a ZIP file.

.EXAMPLE
.\PowerShellModuleOfflinePackager.ps1 -Name "PSReadline" -Version "2.2.0" -OutputDir "C:\MyPackages"

Downloads version 2.2.0 of the PSReadline package, extracts it to C:\MyPackages, and compresses the extracted files into a ZIP file.

.EXAMPLE
.\PowerShellModuleOfflinePackager.ps1 -Name "PSReadline" -SkipZip

Downloads the latest version of the PSReadline package, extracts it to the default output directory, but does not compress the extracted files into a ZIP file.
#>

param (
    [Parameter(Mandatory = $true, HelpMessage = "The ID of the NuGet package to download.")]
    [string]$Name,

    [Parameter(HelpMessage = "The version of the NuGet package to download. If not specified, the latest version will be downloaded.")]
    [string]$Version,

    [Parameter(HelpMessage = "The directory where the package will be downloaded and extracted. If not specified, a folder named 'NuGetPackages' will be created in the script's directory.")]
    [string]$OutputDir = "$PSScriptRoot",

    [Parameter(HelpMessage = "If specified, the script will not compress the package folders into a ZIP file.")]
    [switch]$SkipZip
)

$packageDir = "$($OutputDir)\NuGetPackages"

# Create the output directory if it doesn't exist
if (-Not (Test-Path -Path $PackageDir)) {
    New-Item -ItemType Directory -Path $PackageDir -Force
}

# If no version is specified, find the latest version
if (-Not $Version) {
    Write-Output "Fetching the latest version of $Name from the PowerShell Gallery..."
    $latestPackage = Find-Package -Name $Name -Source PSGallery | Sort-Object -Property Version -Descending | Select-Object -First 1
    if ($latestPackage) {
        $Version = $latestPackage.Version
        Write-Output "Latest version of $Name is $Version"
    } else {
        Write-Error "Package $Name not found in the PowerShell Gallery."
        exit 1
    }
}

# Download and extract the specified version of the package
Write-Output "Downloading and extracting $Name version $Version..."
Save-Package -Name $Name -RequiredVersion $Version -Path $packageDir -Source PSGallery

# Compress the extracted package folders into a single ZIP file, if SkipZip is not specified
if (-Not $SkipZip) {
    $zipFileName = "$Name-$Version.zip"
    $zipFilePath = Join-Path -Path $OutputDir -ChildPath $zipFileName

    Write-Output "Compressing the package folders into $zipFilePath..."
    Compress-Archive -Path "$packageDir\*" -DestinationPath $zipFilePath -Force

    Write-Output "Packages compressed into $zipFilePath"

    # Remove the extracted folders after zipping
    Write-Output "Cleaning up extracted folders..."
    Remove-Item -Path $packageDir -Recurse -Force

    Write-Output "Cleanup complete."
} else {
    Write-Output "Skipping compression and cleanup of the package folders."
}
