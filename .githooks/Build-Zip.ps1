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
    Where to write the zip. Defaults to the repo root (same directory as README.md),
    e.g., <repo>\PowerShell.zip. The zip is excluded from git tracking via .gitignore (*.zip).

.EXAMPLE
    .\.githooks\Build-Zip.ps1
    Writes <repo>\PowerShell.zip in the repo root.

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

# Default output: <repo-root>\<repo-name>.zip (kept out of git via *.zip in .gitignore)
if (-not $OutputPath) {
    $repoName    = Split-Path -Leaf $repoRoot
    $OutputPath  = Join-Path $repoRoot "$repoName.zip"
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
