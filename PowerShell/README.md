# PowerShell

Meta-tooling: download, unpack, and install PowerShell modules — especially useful for offline / locked-down machines that can't run `Install-Module` directly against the PowerShell Gallery.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [PowerShellModuleOfflinePackager.ps1](PowerShellModuleOfflinePackager.ps1) | Downloads the most current (or a specified) version of a PowerShell module's NuGet package and extracts it into a folder ready to copy onto an offline machine. | n/a |
| [ExtractModuleToDirectory.ps1](ExtractModuleToDirectory.ps1) | Extracts a PowerShell module ZIP/NuGet package into a selected PowerShell modules directory so the module becomes importable. Pairs with the offline packager. | n/a |
