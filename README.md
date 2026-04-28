# PowerShell

A collection of re-usable powershell scripts for multiple technologies.

## Portable zip (for moving the repo to another system)

Every commit triggers a `post-commit` hook that runs [.githooks/Build-Zip.ps1](.githooks/Build-Zip.ps1) and writes a fresh `PowerShell.zip` in the repo root (next to this README). The zip is built with `git archive HEAD`, so it contains only tracked files (no `.git` directory, no ignored CSV / zip outputs). The zip itself is excluded from git tracking via `.gitignore` (`*.zip`).

### One-time setup on a fresh clone

```powershell
git config core.hooksPath .githooks
```

(Required because Git does not honor a tracked hooks directory by default.)

### Manual rebuild without committing

```powershell
.\.githooks\Build-Zip.ps1                          # writes .\PowerShell.zip in repo root
.\.githooks\Build-Zip.ps1 -OutputPath C:\xfer.zip  # custom destination
```
