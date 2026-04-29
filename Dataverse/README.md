# Dataverse

Reusable scripts for Dataverse Web API operations, organized by purpose into subcategories.

## Sub-categories

- **[Analysis/](Analysis/README.md)** — usage analysis: record counts, fill rates, audit history, relationships, solution membership, UI presence, user activity, plus an orchestrator that runs them all and a workbook builder that combines the CSVs into a single Excel file.
- **[Schema/](Schema/README.md)** — schema operations: add/delete polymorphic relationships, extract help-page content from solution exports.
- **[Users/](Users/README.md)** — user provisioning: sync individual users or all members of a security group into the environment.
- **[Apps/](Apps/README.md)** — model-driven / canvas app administration.
- **[SiteMap/](SiteMap/README.md)** — sitemap XML tooling.

## Authentication pattern

Every script that calls the Dataverse Web API exists as a pair:

- `ScriptName.ps1` — base script that takes `-OrganizationUrl` and an existing `-AccessToken`. Easy to call from CI or other scripts that already have a token.
- `ScriptNameWithAuth.ps1` — wrapper that takes `-TenantId`, `-ClientId`, `-Environment` instead, calls [../EntraID/GetAccessTokenDeviceCode.ps1](../EntraID/GetAccessTokenDeviceCode.ps1) for an interactive device-code login, and forwards everything to the base script.

For long-running ad-hoc runs from your own workstation, use the `WithAuth` variant. For batch/CI scenarios, get a token once and call the base scripts directly.
