# Dataverse / Users

User-provisioning utilities — sync Entra users into an environment or create and maintain disabled placeholder (stub) users through the Dataverse Web API.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [SyncUserByEmail.ps1](SyncUserByEmail.ps1) | Syncs a single user (looked up by email address) into the Dataverse environment. | n/a — uses the Power Platform PowerShell module's built-in auth |
| [SyncUsersFromSecurityGroup.ps1](SyncUsersFromSecurityGroup.ps1) | Syncs every member of a Microsoft Entra security group into the environment. | n/a — uses the Power Platform PowerShell module's built-in auth |
| [NewStubUser.ps1](NewStubUser.ps1) | Creates and verifies a disabled, unlicensed, unsynchronized stub user through the Web API. | Yes — [NewStubUserWithAuth.ps1](NewStubUserWithAuth.ps1) |
| [NewStubUserWithAuth.ps1](NewStubUserWithAuth.ps1) | Acquires a token and creates a disabled stub user through the Web API. | Wrapper for [NewStubUser.ps1](NewStubUser.ps1) |
| [UpdateStubUser.ps1](UpdateStubUser.ps1) | Updates selected name, title, email or user-name fields on a verified existing stub user. | Yes — [UpdateStubUserWithAuth.ps1](UpdateStubUserWithAuth.ps1) |
| [UpdateStubUserWithAuth.ps1](UpdateStubUserWithAuth.ps1) | Acquires a token and updates selected profile fields on an existing stub user. | Wrapper for [UpdateStubUser.ps1](UpdateStubUser.ps1) |

## Stub users

Use fictitious email/user names (for example, `alex.former@stub.invalid`), not identities
that real Entra users might need later. Stubs cannot sign in or be converted into
regular synchronized users. The creation scripts reject existing IDs/identities rather than
updating them, default to the root business unit, and accept optional `-Title`
(up to 128 characters), `-BusinessUnitId`, `-DomainName`, and `-SystemUserId` overrides. Duplicate checks are
preflight checks, not a concurrency lock: do not run parallel creates for the same identity.
None of the stub scripts assign roles, change organization settings, or guarantee ownership
eligibility. All support `-WhatIf` without API calls; the wrappers also skip sign-in.

The [Microsoft admin documentation](https://learn.microsoft.com/en-us/power-platform/admin/create-users#how-stub-users-are-created)
describes data import. The Web API approach is described by
[The Apps Bakery](https://theappsbakery.com/migrating-non-existing-users-to-dataverse/),
and the unlicensed/unsynchronized field settings by
[KingswaySoft](https://www.kingswaysoft.com/blog/2023/01/24/Manage-Unlicensed-Users-in-CRM-Using-KingswaySoft).
Validate this approach in a non-production environment before migration use.
If creation succeeds but verification fails, inspect the reported user ID before
retrying; the script does not automatically delete the record.

### Updating a stub profile

Use `UpdateStubUser.ps1` with an existing token, or `UpdateStubUserWithAuth.ps1`
for sign-in. Supply the existing `-SystemUserId` and at least one of `-FirstName`,
`-LastName`, `-Title`, `-EmailAddress`, or `-DomainName`. Only supplied, changed
fields are patched; omitted values are preserved. Use `-Title ""` to clear a
title. Names and identity fields cannot be cleared. `-EmailAddress` changes only
the primary email: pass `-DomainName` too if the fictitious user name should change.
The GUID is never changed, so existing lookup references continue to identify
the same record.

The update scripts verify the built-in Stub Users criteria, reject conflicting
email/user names on other users, and use the retrieved ETag to prevent upserts
and stale updates. Optimistic concurrency must be supported by the target table;
failed updates are not retried automatically. Status, licensing, directory IDs,
access mode and business unit are preserved and checked during read-back.
Identity-conflict checks are preflight checks, not a cross-user lock; avoid
parallel identity changes using the same addresses. Profile changes may trigger
environment-specific automation.

```powershell
.\UpdateStubUserWithAuth.ps1 `
    -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" `
    -OrganizationUrl "https://your-org.crm.dynamics.com" `
    -SystemUserId "11111111-1111-1111-1111-111111111111" `
    -Title "Senior External Approver" -AuthenticationMode Interactive -WhatIf
```

Remove `-WhatIf` only after reviewing the intended target and changes. If PATCH
succeeds but read-back verification fails, inspect the reported ID before retrying;
no automatic rollback is attempted.
