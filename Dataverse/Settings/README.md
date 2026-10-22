# Dataverse / Settings

Read or update individual organization-wide OrgDBOrgSettings XML values. Adapted
from MARCY ETL's organization-settings utility to this repository's token/base and
interactive-auth wrapper pattern; no ETL modules or project configuration are required.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [SetOrgDbOrgSetting.ps1](SetOrgDbOrgSetting.ps1) | Reads or sets one case-sensitive OrgDBOrgSettings XML value with concurrency protection and read-back verification. | Yes - [SetOrgDbOrgSettingWithAuth.ps1](SetOrgDbOrgSettingWithAuth.ps1) |
| [SetOrgDbOrgSettingWithAuth.ps1](SetOrgDbOrgSettingWithAuth.ps1) | Acquires a token and reads or sets one OrgDBOrgSettings XML value. | Wrapper for [SetOrgDbOrgSetting.ps1](SetOrgDbOrgSetting.ps1) |

Omit `-Value` to read only. `IsExplicit = false` with `Value = null` means the
XML key is absent, not that the effective setting is false. Consult the setting's
documentation for its default. No setting names or defaults are guessed.

Writes add or replace only the named XML element and preserve other XML settings.
They require confirmation by default, an organization ETag, and organization
read/write privileges. Concurrent changes fail rather than silently overwriting
another administrator's edits. Verification failures after PATCH require manual
inspection; there is no automatic retry or rollback.

Both scripts support `-WhatIf` without authentication or API calls. They do not
change direct organization columns or remove settings.

```powershell
.\SetOrgDbOrgSettingWithAuth.ps1 `
    -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" `
    -OrganizationUrl "https://your-org.crm.dynamics.com" `
    -Setting AllowRoleAssignmentOnDisabledUsers -AuthenticationMode Interactive
```

To preview a write, add `-Value "true" -WhatIf`. Remove `-WhatIf` only after
reviewing the organization-wide impact and confirming that the key and value
are supported in the target environment.

## Stub-user settings are not interchangeable

Use `AllowRoleAssignmentOnDisabledUsers` (uppercase A) for assigning security
roles to disabled users. It does **not** enable Stub User Impersonation.
Live verification on 2026-10-22 confirmed that the uppercase key with XML text
`true` is accepted and persists. The lowercase spelling shown in Microsoft's
support table and MARCY ETL's examples was rejected with `0x80048514`, with both
`true` and `1`. Both scripts preserve unrelated XML settings; the ETag guard was
not the cause. Check exact key casing against the OrgDbOrgSettings utility's
setting definitions before changing formats or retrying invalid-setting errors.
`CanCreateApplicationStubUser` concerns application stubs and must not be assumed
to enable impersonable person stubs either.

Creating `systemmanagedusertype = 2` failed in the tested environment because
Stub User Impersonation was not enabled. No supported enablement key has been
verified here. This utility does not resolve or bypass that feature gate, and
successful XML storage alone is not proof that a feature works.

Reference: [Microsoft OrgDBOrgSettings tool and supported settings](https://support.microsoft.com/en-us/topic/orgdborgsettings-tool-for-microsoft-dynamics-crm-20a10f46-2a24-a156-7144-365d49b842ba).
