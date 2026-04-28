<#
.SYNOPSIS
    Creates a portable zip of the repo for moving to disconnected systems.

.DESCRIPTION
    Uses `git archive` against HEAD so the zip contains only tracked files
    (respects .gitignore automatically; excludes the .git folder, *.csv outputs,
    *.zip artifacts, etc.).

    Called automatically by the post-commit hook in .githooks/post-commit.
    Can also be run manually any time.

.PARAMETER OutputPath
    Where to write the zip. Defaults to one level above the repo root
    (e.g., ..\PowerShell.zip) so the file does not pollute the working tree.

.EXAMPLE
    .\.githooks\Build-Zip.ps1
    Writes ..\PowerShell.zip relative to the repo root.

.EXAMPLE
    .\.githooks\Build-Zip.ps1 -OutputPath C:\transfer\PowerShell.zip
#>

param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

# Find repo root
$repoRoot = (& git rev-parse --show-toplevel 2>$null)
if (-not $repoRoot -or $LASTEXITCODE -ne 0) {
    Write-Error "Not inside a git repository. Run from within the repo."
    exit 1
}
$repoRoot = $repoRoot.Trim()

# Default output: <parent of repo>\<repo-name>.zip
if (-not $OutputPath) {
    $repoName    = Split-Path -Leaf $repoRoot
    $parentDir   = Split-Path -Parent $repoRoot
    $OutputPath  = Join-Path $parentDir "$repoName.zip"
}

# Ensure the output directory exists
$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Force
}

Push-Location $repoRoot
try {
    # git archive only includes tracked files - no .git dir, no ignored outputs
    & git archive --format=zip --output="$OutputPath" HEAD
    if ($LASTEXITCODE -ne 0) {
        Write-Error "git archive failed (exit code $LASTEXITCODE)"
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}

$sizeKB = [math]::Round((Get-Item $OutputPath).Length / 1KB, 1)
Write-Host "Repo zipped: $OutputPath ($sizeKB KB)" -ForegroundColor Green
