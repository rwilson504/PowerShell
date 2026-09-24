# Dataverse / Attachments

Scripts for downloading files stored as note (`annotation`) attachments on Dataverse records.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [Save-RecordAttachments.ps1](Save-RecordAttachments.ps1) | Downloads every note attachment directly related to a record GUID, with paging, safe filenames, duplicate-name handling, and optional overwrite behavior. | Yes - [Save-RecordAttachmentsWithAuth.ps1](Save-RecordAttachmentsWithAuth.ps1) |

## Scope

These scripts download note attachments where `annotation.isdocument` is true and the note's object lookup matches the supplied record GUID. They do not download Dataverse file/image columns or attachments on related email activities.

The WithAuth wrapper supports `-AuthenticationMode DeviceCode` (default) and
`-AuthenticationMode Interactive`. Interactive opens the system browser and requires the
app registration to allow public client flows with `http://localhost` registered as a
Mobile and desktop applications redirect URI.
