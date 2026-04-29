# PowerShell

A collection of re-usable PowerShell scripts for multiple technologies.

## Categories

- **[Dataverse/](Dataverse/README.md)** — scripts for the Dataverse Web API: usage analysis (record counts, fill rates, audit history, etc.), schema operations (relationships, help-page extraction), user provisioning, app admin, sitemap tooling.
- **[EntraID/](EntraID/README.md)** — Microsoft Entra ID auth helpers and queries (device-code token acquisition, first-party service-principal listing).
- **[Files/](Files/README.md)** — local filesystem utilities.
- **[PowerShell/](PowerShell/README.md)** — meta-tooling: download / unpack / install PowerShell modules for offline machines.
- **[SharePoint/](SharePoint/README.md)** — SharePoint Online tooling.
- **[SharePoint On-Prem/](SharePoint%20On-Prem/README.md)** — SharePoint Server (on-premises) tooling.

When a script calls an authenticated API, it usually exists as a pair: a **base script** (`ScriptName.ps1`) that takes an existing `-AccessToken`, plus a **WithAuth wrapper** (`ScriptNameWithAuth.ps1`) that acquires a token via [EntraID/GetAccessTokenDeviceCode.ps1](EntraID/GetAccessTokenDeviceCode.ps1) and calls the base script.

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
