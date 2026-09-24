# PowerShell Repository — Copilot Instructions

## Overview

This repository is a collection of reusable PowerShell scripts organized by technology area (Dataverse, EntraID, Files, PowerShell, SharePoint, etc.). Scripts follow a consistent structure and, when they call authenticated APIs, are split into two versions: a **base script** and a **WithAuth wrapper**.

Before changing a script, read its folder README, its paired base/wrapper script when one exists, and any underscore-prefixed helper it dot-sources. Preserve compatibility with the PowerShell edition already supported by the nearby code; do not modernize an entire script solely for style consistency.

This is a public utility repository. Do not create or maintain `SESSION_LOG.md` files or `decisions/` records here.

---

## Project Structure

```
<TechnologyArea>/
    README.md                     # Index of every script in this folder (REQUIRED)
    <Subcategory>/                # Optional - large categories (e.g. Dataverse) split further
        README.md                 # Index of every script in this subfolder (REQUIRED)
        ScriptName.ps1            # Base script (accepts AccessToken or handles its own auth)
        ScriptNameWithAuth.ps1    # Auth wrapper that acquires a token then calls the base script
        _SharedHelper.ps1         # Internal dot-sourced helper; not a standalone command
EntraID/
    GetAccessTokenDeviceCode.ps1  # Shared device-code / interactive browser auth helper
README.md                          # Top-level index linking to every category README
```

- Group scripts into folders by technology/service (e.g., `Dataverse/`, `SharePoint/`, `Files/`, `PowerShell/`).
- Each folder should contain related scripts that target that technology.
- When a category gets large enough that its scripts split into purposes (e.g. `Dataverse/Analysis/`, `Dataverse/Schema/`), promote those into subfolders and give each a `README.md`.

---

## Documentation Conventions (READMEs)

Every folder that contains scripts must have a `README.md` listing those scripts. The repo also has a top-level `README.md` linking to each category's README. **Whenever a script is added, removed, renamed, or its purpose changes, update the relevant READMEs in the same commit.**

### Top-level `README.md`

- Brief one-line description of the repo.
- A "Categories" section: bulleted list of each top-level folder with a one-line summary and a link to its `README.md`.
- Anything genuinely repo-wide (zip-build hook setup, fresh-clone steps, etc.).

Example:

```markdown
## Categories

- **[Dataverse/](Dataverse/README.md)** - scripts for Dataverse Web API operations, schema, analysis, and admin.
- **[EntraID/](EntraID/README.md)** - Microsoft Entra ID auth helpers and queries.
- **[Files/](Files/README.md)** - filesystem utilities.
- **[SharePoint/](SharePoint/README.md)** - SharePoint Online tooling.
```

### Per-folder `README.md`

Required sections:

1. **Folder purpose** - one or two sentences explaining what this folder is for.
2. **Sub-categories** (if the folder has subfolders) - linked list with a short summary of each.
3. **Scripts** - a markdown table with one row per script in this folder, containing at minimum:

| Column | Content |
|---|---|
| Script | Name + relative-path link, e.g. `[GetRecordCountByTable.ps1](GetRecordCountByTable.ps1)` |
| Purpose | One-line summary (mirrors the script's `.SYNOPSIS`) |
| WithAuth pair? | Yes/No - link to the WithAuth wrapper if separate |

For scripts that work together (e.g. an orchestrator + helpers), group them under a sub-heading explaining the relationship before the table.

### When to update READMEs

- **Adding a new script** → add a row to the relevant folder's `README.md` table; if the folder is brand-new, also add a link to the new folder in the top-level README.
- **Renaming or moving a script** → update both the source folder's README (remove row) and the destination folder's README (add row).
- **Significant scope change to a script** → update its row's Purpose column.
- **Deleting a script** → remove the row.
- **Adding a new category folder** → create that folder's `README.md` AND add a link to it from the top-level README.

Keep README rows brief (one line of purpose); detailed behavior belongs in the script's comment-based help, not in the README. The README is for "what scripts exist and at a glance, what each one does"; the help block is for "how do I actually call this one."

---

## Script Conventions

### Comment-Based Help

Every script **must** start with a `<# ... #>` comment-based help block containing at minimum:

| Section | Required | Notes |
|---|---|---|
| `.SYNOPSIS` | Yes | One-line summary of what the script does. |
| `.DESCRIPTION` | Yes | Detailed explanation of behavior, parameters, and any prerequisites. |
| `.PARAMETER` | Yes | One entry per parameter — describe purpose, valid values, and defaults. |
| `.EXAMPLE` | Yes | At least one realistic usage example with explanation text. |
| `.NOTES` | Optional | Prerequisites, module installs, security notes, version info, and `Author: Rick Wilson` attribution when included. PowerShell does not support a standalone `.AUTHOR` help keyword. |

### Parameters

- Define parameters using a `param( )` block immediately after the help comment.
- Use `[Parameter(Mandatory = $true)]` for required parameters.
- Use `[ValidateSet()]` for parameters with a fixed set of values (e.g., `Environment`).
- Provide sensible defaults where appropriate (e.g., `$Environment = "Public"`).
- Use strong typing (`[string]`, `[string[]]`, `[switch]`, `[bool]`).
- For Dataverse/API scripts, the base script **always** takes `[string]$OrganizationUrl` and `[string]$AccessToken` as parameters.

### Code Style

- Use descriptive names. Public parameter names use PascalCase (for example, `$AccessToken` and `$OrganizationUrl`); local variables follow the casing and style of the surrounding script.
- Use `Write-Host` with `-ForegroundColor` for user-facing status messages (Cyan for info, Green for success, Yellow for warnings, Red for errors).
- Use `Write-Error` for fatal errors and `Write-Warning` for non-fatal issues.
- Use `Write-Output` or `return` for pipeline-friendly output.
- Extract reusable logic into local functions within the script when needed.
- Use hashtable splatting (`@params`) for calls with many parameters.
- Name shared, non-standalone helper files with a leading underscore and dot-source them relative to `$PSScriptRoot` or the current script directory.
- Use `[CmdletBinding(SupportsShouldProcess)]` and `$PSCmdlet.ShouldProcess()` for commands that make destructive or broad local changes when practical.
- Never write access tokens, secrets, full authorization headers, or device codes to output, logs, examples, or committed files.

---

## Authentication Pattern — Two-Script Approach

When a script calls an authenticated API (Dataverse Web API, Microsoft Graph, etc.), create **two scripts**:

### 1. Base Script (`ScriptName.ps1`)

- Accepts `$AccessToken` as a parameter — it does **not** handle authentication itself.
- Contains all the business logic (API calls, data processing, output).
- Can be called standalone by anyone who already has a token (e.g., from another auth flow, a pipeline, or a service principal).
- Sets up HTTP headers using the provided token:

```powershell
$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "Content-Type"     = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
}
```

- Uses `Invoke-RestMethod` for API calls.

### 2. Auth Wrapper Script (`ScriptNameWithAuth.ps1`)

- Handles token acquisition, then delegates to the base script.
- Accepts the **same parameters** as the base script **except** `$AccessToken`, and **adds** these auth parameters:

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `$TenantId` | `[string]` | Yes | — | Azure AD tenant ID |
| `$ClientId` | `[string]` | Yes | — | App registration client ID |
| `$Environment` | `[string]` | No | `"Public"` | Azure cloud: `Public`, `GCC`, `GCCH`, `DoD` |
| `$AuthenticationMode` | `[string]` | No | `"DeviceCode"` | Authentication flow: `DeviceCode` or browser-based `Interactive` |

- Keep the base script and wrapper parameter surfaces in sync whenever parameters, defaults, validation attributes, or forwarding behavior change.

- Acquires a token by calling the shared auth helper:

```powershell
$accessToken = & ..\EntraID\GetAccessTokenDeviceCode.ps1 `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -Scope "$OrganizationUrl/user_impersonation" `
    -Environment $Environment `
    -AuthenticationMode $AuthenticationMode
```

- Then calls the base script, passing the token and all other parameters:

```powershell
$response = & .\ScriptName.ps1 -OrganizationUrl $OrganizationUrl -AccessToken $accessToken -OtherParam $OtherParam
Write-Output $response
```

- For scripts with many parameters, use splatting to forward them cleanly:

```powershell
$scriptParams = @{
    OrganizationUrl = $OrganizationUrl
    AccessToken     = $accessToken
    OutputFormat    = $OutputFormat
}
# conditionally add optional params
if ($Tables) { $scriptParams.Tables = $Tables }

$mainScript = Join-Path $scriptDir "ScriptName.ps1"
$results = & $mainScript @scriptParams
return $results
```

- Validates the token was acquired before proceeding:

```powershell
if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}
```

### Path Resolution for Auth Script

Two patterns are used to reference the auth helper from a WithAuth script:

- **Relative path** (simpler scripts): `& ..\EntraID\GetAccessTokenDeviceCode.ps1`
- **`$PSScriptRoot` / `Split-Path`** (more robust):

```powershell
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment -AuthenticationMode $AuthenticationMode
```

Prefer the `Split-Path` approach for new scripts as it is more reliable when the working directory differs from the script location.

---

## Shared Auth Helper — `EntraID/GetAccessTokenDeviceCode.ps1`

This is the **single shared authentication script** used by all WithAuth wrappers. It supports OAuth 2.0 device-code authentication and browser-based authorization code with PKCE. Interactive mode requires a public-client app registration with `http://localhost` configured as a Mobile and desktop applications redirect URI.

### Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `$TenantId` | `[string]` | — | Azure AD tenant ID |
| `$ClientId` | `[string]` | — | App registration client/application ID |
| `$Scope` | `[string]` | `"https://your-org.crm.dynamics.com/.default"` | OAuth scope for the target resource |
| `$Environment` | `[string]` | `"Public"` | Azure cloud environment (`Public`, `GCC`, `GCCH`, `DoD`) |
| `$AuthenticationMode` | `[string]` | `"DeviceCode"` | `DeviceCode` or browser-based `Interactive` |

### How It Works

1. Selects the correct login endpoint based on `$Environment`:
   - **Public**: `https://login.microsoftonline.com`
   - **GCC / GCCH / DoD**: `https://login.microsoftonline.us`
2. Uses a valid cached token or refresh token when available.
3. With `DeviceCode`, requests a device code and polls until sign-in completes.
4. With `Interactive`, opens the system browser and receives an authorization-code response on a localhost loopback listener using PKCE.
5. Returns **only the access token string** via `Write-Output`.

### Scope Conventions

- **Dataverse**: `"$OrganizationUrl/user_impersonation"` (e.g., `https://your-org.crm.dynamics.com/user_impersonation`)
- **Microsoft Graph**: `"https://graph.microsoft.com/.default"`
- Other APIs: Use the appropriate resource URI with `/user_impersonation` or `/.default`.

---

## When to Create Auth vs. No-Auth Versions

| Scenario | What to Create |
|---|---|
| Script calls an API requiring a Bearer token | Create **both** `ScriptName.ps1` (base) and `ScriptNameWithAuth.ps1` (wrapper) |
| Script uses PowerShell modules with built-in auth (e.g., `Connect-AzureAD`, `Add-PowerAppsAccount`) | Create a **single script** — the module handles its own interactive auth prompts |
| Script is purely local (file operations, XML processing, etc.) | Create a **single script** — no auth needed |

---

## Dataverse API Conventions

When calling the Dataverse Web API:

- Use API version `v9.2`: `$OrganizationUrl/api/data/v9.2/`
- Always include these headers:

```powershell
$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "Content-Type"     = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
}
```

- For queries that need annotations: add `"Prefer" = "odata.include-annotations=*"` and `"Accept" = "application/json"`.
- Use `Invoke-RestMethod` for JSON API calls. `Invoke-WebRequest` is appropriate when raw response content or headers are required, such as multipart OData `$batch` parsing.
- Convert payloads to JSON with `ConvertTo-Json -Depth 10` to handle nested objects.
- Remove trailing slashes from URLs: `$OrganizationUrl = $OrganizationUrl.TrimEnd('/')`.

---

## Validation and Change Hygiene

- Keep changes scoped to the requested script family. Do not reformat unrelated scripts or generated output.
- After editing a script, parse every touched `.ps1` file with the PowerShell parser so syntax errors are caught without executing authenticated or destructive behavior:

```powershell
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    (Resolve-Path '.\Path\To\Script.ps1'),
    [ref]$null,
    [ref]$errors
)
if ($errors) { $errors | Format-List; exit 1 }
```

- Run the narrowest offline behavior check available. Do not call a live Dataverse, Graph, or SharePoint environment merely to validate syntax.
- When changing a base script with a `WithAuth` pair, parse both files and verify that the wrapper forwards every applicable parameter.
- When adding, removing, renaming, or materially changing a script, update the nearest README in the same change. Update the root README only when a top-level category or repo-wide workflow changes.
- Do not commit generated report folders, CSV/JSON exports, Excel workbooks, access tokens, or `PowerShell.zip`. The portable zip is built from tracked files with `.githooks/Build-Zip.ps1`.

---

## Template: New Base Script (API)

```powershell
<#
.SYNOPSIS
    Brief description of what the script does.

.DESCRIPTION
    Detailed description of the script's behavior and purpose.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (e.g., https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER YourParam
    Description of additional parameter.

.EXAMPLE
    .\ScriptName.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -YourParam "value"

    Description of what this example does.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $true)]
    [string]$YourParam
)

# Remove trailing slash from URL if present
$OrganizationUrl = $OrganizationUrl.TrimEnd('/')

# Set up headers for API calls
$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "Content-Type"     = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
}

# API endpoint
$apiUrl = "$OrganizationUrl/api/data/v9.2/YourEndpoint"

# Make the API call
$response = Invoke-RestMethod -Method Get -Uri $apiUrl -Headers $headers

# Output the response
$response
```

---

## Template: New WithAuth Wrapper Script

```powershell
<#
.SYNOPSIS
    Acquires an access token and calls ScriptName to perform the operation.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token
    and then calls the ScriptName script to perform the operation.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER AuthenticationMode
    Authentication flow. Valid values are "DeviceCode" and "Interactive". Default is "DeviceCode".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER YourParam
    Description of additional parameter.

.EXAMPLE
    .\ScriptNameWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -YourParam "value"

    Description of what this example does.
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Public", "GCC", "GCCH", "DoD")]
    [string]$Environment = "Public",

    [Parameter(Mandatory = $false)]
    [ValidateSet("DeviceCode", "Interactive")]
    [string]$AuthenticationMode = "DeviceCode",

    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$YourParam
)

# Get the access token
Write-Host "Acquiring access token..." -ForegroundColor Cyan
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$authScript = Join-Path $scriptDir "..\EntraID\GetAccessTokenDeviceCode.ps1"
$accessToken = & $authScript -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment -AuthenticationMode $AuthenticationMode

if (-not $accessToken) {
    Write-Error "Failed to acquire access token."
    exit 1
}

Write-Host "Access token acquired successfully." -ForegroundColor Green

# Call the base script
$mainScript = Join-Path $scriptDir "ScriptName.ps1"
$response = & $mainScript -OrganizationUrl $OrganizationUrl -AccessToken $accessToken -YourParam $YourParam

# Output the response
Write-Output $response
```

---

## Template: New Standalone Script (No Auth)

```powershell
<#
.SYNOPSIS
    Brief description.

.DESCRIPTION
    Detailed description.

.PARAMETER YourParam
    Description.

.EXAMPLE
    .\ScriptName.ps1 -YourParam "value"

    Description of what this example does.

.NOTES
    Author: Rick Wilson
    Date  : YYYY-MM-DD
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$YourParam
)

# Script logic here
```
