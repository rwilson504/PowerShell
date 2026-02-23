<#
.SYNOPSIS
    Searches for blocked files in a directory and optionally unblocks them.

.DESCRIPTION
    This script scans files in a specified directory (or the current directory if none is specified)
    to identify files that are blocked due to being downloaded from the internet (Zone.Identifier).
    By default, the script will unblock these files. Use the -ReportOnly switch to only report
    blocked files without removing the block.

.PARAMETER Path
    The directory path to search for blocked files. Defaults to the current directory.

.PARAMETER Recurse
    If specified, searches subdirectories recursively.

.PARAMETER ReportOnly
    If specified, only reports blocked files without unblocking them.

.PARAMETER Filter
    File filter pattern. Defaults to "*" (all files).

.EXAMPLE
    .\Unblock-Files.ps1
    Searches the current directory for blocked files and unblocks them.

.EXAMPLE
    .\Unblock-Files.ps1 -Path "C:\Downloads" -Recurse
    Searches C:\Downloads and all subdirectories for blocked files and unblocks them.

.EXAMPLE
    .\Unblock-Files.ps1 -Path "C:\Downloads" -ReportOnly
    Reports all blocked files in C:\Downloads without unblocking them.

.EXAMPLE
    .\Unblock-Files.ps1 -Path "C:\Downloads" -Filter "*.ps1" -Recurse
    Searches for blocked PowerShell scripts in C:\Downloads and subdirectories.

.NOTES
    Author: Auto-generated
    Date: 2026-01-26
    Files are blocked when Windows adds a Zone.Identifier alternate data stream to mark
    files downloaded from the internet or other untrusted sources.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [ValidateScript({
        if (Test-Path $_ -PathType Container) {
            $true
        } else {
            throw "Path '$_' does not exist or is not a directory."
        }
    })]
    [string]$Path = (Get-Location).Path,

    [Parameter()]
    [switch]$Recurse,

    [Parameter()]
    [switch]$ReportOnly,

    [Parameter()]
    [string]$Filter = "*"
)

# Initialize counters
$totalFiles = 0
$blockedFiles = 0
$unblockedFiles = 0
$errorFiles = 0

# Build Get-ChildItem parameters
$gciParams = @{
    Path   = $Path
    File   = $true
    Filter = $Filter
}

if ($Recurse) {
    $gciParams.Recurse = $true
}

Write-Host "`nScanning for blocked files in: $Path" -ForegroundColor Cyan
if ($Recurse) {
    Write-Host "Including subdirectories..." -ForegroundColor Cyan
}
Write-Host ""

# Get all files and check for blocked status
$files = Get-ChildItem @gciParams -ErrorAction SilentlyContinue

foreach ($file in $files) {
    $totalFiles++
    
    try {
        # Check if file has Zone.Identifier stream (indicates blocked)
        $zoneId = Get-Item -Path $file.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue
        
        if ($zoneId) {
            $blockedFiles++
            
            if ($ReportOnly) {
                # Just report the blocked file
                Write-Host "[BLOCKED] $($file.FullName)" -ForegroundColor Yellow
            } else {
                # Attempt to unblock the file
                if ($PSCmdlet.ShouldProcess($file.FullName, "Unblock file")) {
                    try {
                        Unblock-File -Path $file.FullName -ErrorAction Stop
                        Write-Host "[UNBLOCKED] $($file.FullName)" -ForegroundColor Green
                        $unblockedFiles++
                    } catch {
                        Write-Host "[ERROR] Failed to unblock: $($file.FullName)" -ForegroundColor Red
                        Write-Host "        $($_.Exception.Message)" -ForegroundColor Red
                        $errorFiles++
                    }
                }
            }
        }
    } catch {
        # Error accessing file - skip silently unless verbose
        Write-Verbose "Could not access file: $($file.FullName)"
    }
}

# Display summary
Write-Host "`n" + ("=" * 60) -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "Total files scanned:  $totalFiles"
Write-Host "Blocked files found:  $blockedFiles" -ForegroundColor $(if ($blockedFiles -gt 0) { "Yellow" } else { "Green" })

if (-not $ReportOnly) {
    Write-Host "Files unblocked:      $unblockedFiles" -ForegroundColor $(if ($unblockedFiles -gt 0) { "Green" } else { "White" })
    if ($errorFiles -gt 0) {
        Write-Host "Errors encountered:   $errorFiles" -ForegroundColor Red
    }
}

Write-Host ("=" * 60) -ForegroundColor Cyan

# Return object for pipeline usage
[PSCustomObject]@{
    Path           = $Path
    TotalFiles     = $totalFiles
    BlockedFiles   = $blockedFiles
    UnblockedFiles = if ($ReportOnly) { 0 } else { $unblockedFiles }
    Errors         = $errorFiles
    ReportOnly     = $ReportOnly
}
