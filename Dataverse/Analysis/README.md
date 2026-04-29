# Dataverse / Analysis

Usage-analysis suite. Eight individual analysis scripts plus an orchestrator that runs them all and a workbook builder that combines the CSV outputs into a single Excel file with pre-computed joins. Designed to answer questions like "is this field still being used", "who owns this process", "which custom relationships are dead code".

## How the scripts work together

1. **[Invoke-DataverseUsageReport.ps1](Invoke-DataverseUsageReport.ps1)** (orchestrator) runs all eight analysis scripts in order, dropping each CSV into a single timestamped folder named `dataverse-usage-<env>-<solution>-<timestamp>/`.
2. **[Build-UsageReportWorkbook.ps1](Build-UsageReportWorkbook.ps1)** then takes that folder, computes three additional join CSVs (`master.csv`, `tables.csv`, `cleanup.csv`) plus a column `dictionary.csv` and a `README.md`, and optionally bundles everything into a single `UsageReport.xlsx` workbook via Excel COM.

You can also run any individual analysis script standalone if you only need one slice of data.

## Orchestrator + workbook builder

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [Invoke-DataverseUsageReport.ps1](Invoke-DataverseUsageReport.ps1) | Runs all 8 analysis scripts in sequence, drops every CSV into a timestamped folder, prints a tabular summary. Forwards all common scope and per-script options (-Tables, -SolutionUniqueName, -IncludeLastActivity, -AutoDetectUserLookups, -UserTargetTables, -CustomTargetNameColumns, etc.). Has -Skip for partial runs and -BuildWorkbook / -CombineToXlsx / -OpenAfterBuild to chain into the workbook builder. | Yes — [Invoke-DataverseUsageReportWithAuth.ps1](Invoke-DataverseUsageReportWithAuth.ps1) |
| [Build-UsageReportWorkbook.ps1](Build-UsageReportWorkbook.ps1) | Reads the orchestrator's output folder and produces master.csv (full per-attribute join with a DeadFieldScore composite), tables.csv (per-table roll-up), cleanup.csv (pre-filtered cleanup candidates), dictionary.csv (column glossary), and README.md. With -CombineToXlsx, bundles every CSV into a single .xlsx via Excel COM. | n/a — file-based, no auth needed |

## Individual analysis scripts (each runnable standalone)

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [GetRecordCountByTable.ps1](GetRecordCountByTable.ps1) | Per-table record counts via the RetrieveTotalRecordCount API. Optional last-activity probes (LastCreatedOn / LastModifiedOn / OldestRecordCreatedOn) and UsageBucket classification. Pre-skips Virtual/Elastic tables; per-table retry on batch failure. | Yes — [GetRecordCountByTableWithAuth.ps1](GetRecordCountByTableWithAuth.ps1) |
| [GetTableRelationships.ps1](GetTableRelationships.ps1) | Lists every 1:N / N:1 / M:M relationship per table, with the relationship's schema name, lookup attribute (or intersect table for M:M), and IsCustomRelationship / IsManaged flags. | Yes — [GetTableRelationshipsWithAuth.ps1](GetTableRelationshipsWithAuth.ps1) |
| [GetFieldFillRateByTable.ps1](GetFieldFillRateByTable.ps1) | Per-attribute fill rate (PopulatedCount / TotalRecords / FillRatePercent) for one or more tables. Uses OData `$batch` for performance. Includes a LookupTargets column for cross-correlation with the relationships report. | Yes — [GetFieldFillRateByTableWithAuth.ps1](GetFieldFillRateByTableWithAuth.ps1) |
| [GetAttributeAuditHistory.ps1](GetAttributeAuditHistory.ps1) | Per-attribute audit history (LastAuditedOn, AuditEntriesInWindow, DistinctUsersInWindow, DistinctRecordsTouched, CreateEvents, UpdateEvents, EventsLast 30/90/365 Days) by scanning the audit table. Parses both `attributemask` and `changedata` so it catches Updates that the API leaves with empty attributemask. | Yes — [GetAttributeAuditHistoryWithAuth.ps1](GetAttributeAuditHistoryWithAuth.ps1) |
| [GetUserActivityByTable.ps1](GetUserActivityByTable.ps1) | For each table, per-user record counts on the standard creator/modifier columns AND any user lookups you specify (or auto-discover). Supports systemuser, contact, account, and custom person tables via -UserTargetTables / -CustomTargetNameColumns. | Yes — [GetUserActivityByTableWithAuth.ps1](GetUserActivityByTableWithAuth.ps1) |
| [GetFieldUIPresence.ps1](GetFieldUIPresence.ps1) | Per-attribute booleans + counts for OnAnyForm / OnAnyView / OnAnyChart by scanning systemform / savedquery / savedqueryvisualization XML. The strongest "dead field" signal — a field with FillRate=0 AND AnyUIPresence=False is essentially unused. | Yes — [GetFieldUIPresenceWithAuth.ps1](GetFieldUIPresenceWithAuth.ps1) |
| [GetTableUsageActivity.ps1](GetTableUsageActivity.ps1) | Table-level activity rollup using FetchXML aggregates: NewestCreatedOn, NewestModifiedOn, RecordsCreatedLast30/90/365Days, DistinctCreators / Modifiers / Owners. Works without auditing being enabled. | Yes — [GetTableUsageActivityWithAuth.ps1](GetTableUsageActivityWithAuth.ps1) |
| [GetSolutionMembership.ps1](GetSolutionMembership.ps1) | Maps each Entity / Attribute / Relationship component to the solution(s) that ship it. Tells you "is anything still shipping this field" before you delete it. Filters: -ComponentTypes, -UnmanagedOnly / -ManagedOnly, -ExcludeSystemSolutions, -SolutionUniqueName. | Yes — [GetSolutionMembershipWithAuth.ps1](GetSolutionMembershipWithAuth.ps1) |
| [GetSitemapEntityPresence.ps1](GetSitemapEntityPresence.ps1) | Walks every model-driven app sitemap (Area / Group / SubArea) and emits one row per Entity-bound SubArea. Answers "is this table actually surfaced to a user, and in which apps?" — the strongest *user-facing* signal in the suite (a table can have data + audit activity and still be invisible to end users). Non-entity tabs (Dashboard / Url / WebResource) are emitted too for drill-down context. | Yes — [GetSitemapEntityPresenceWithAuth.ps1](GetSitemapEntityPresenceWithAuth.ps1) |

## Diagnostic helpers

| Script | Purpose |
|---|---|
| [Test-AuditDataFormat.ps1](Test-AuditDataFormat.ps1) | Dumps sample audit rows + per-operation breakdown for one table. Useful when investigating why GetAttributeAuditHistory returns blank columns on a new environment. |

## Internal shared helpers (dot-sourced)

| File | Purpose |
|---|---|
| [_ODataBatchHelper.ps1](_ODataBatchHelper.ps1) | `Invoke-ODataBatch` — posts a multipart `$batch` of GET sub-requests and parses the responses. Used by the field fill-rate and record-count scripts. |
| [_SolutionFilterHelper.ps1](_SolutionFilterHelper.ps1) | `Resolve-SolutionScopedTables` — resolves a solution unique name to its set of contained tables. Used by every script that supports `-SolutionUniqueName`. |
