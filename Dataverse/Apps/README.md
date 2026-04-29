# Dataverse / Apps

Model-driven and canvas app administration.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [bypass-permissions-for-canvas-apps.ps1](bypass-permissions-for-canvas-apps.ps1) | Configures bypass-consent permissions for canvas apps in bulk, driven by `bypassPermissionsConfig.json`. | n/a — uses Power Platform admin module auth |

## Configuration

| File | Purpose |
|---|---|
| [bypassPermissionsConfig.json](bypassPermissionsConfig.json) | Input config consumed by `bypass-permissions-for-canvas-apps.ps1` — defines which apps and which connection references should be configured for bypass-consent. |
