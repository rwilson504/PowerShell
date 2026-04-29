# Dataverse / Schema

Schema-modification utilities — relationship management and help-page extraction.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [AddPolymorphicRelationship.ps1](AddPolymorphicRelationship.ps1) | Creates a new polymorphic-lookup entity relationship in Dataverse (the kind where one lookup can point to multiple target tables). | Yes — [AddPolymorphicRelationshipWithAuth.ps1](AddPolymorphicRelationshipWithAuth.ps1) |
| [DeleteRelationship.ps1](DeleteRelationship.ps1) | Deletes an existing entity relationship by SchemaName. | Yes — [DeleteRelationshipWithAuth.ps1](DeleteRelationshipWithAuth.ps1) |
| [ExtractHelpPageContentToCSV.ps1](ExtractHelpPageContentToCSV.ps1) | Extracts help-page content from a Dynamics solution `customizations.xml` and writes it to CSV (with optional rich-text → plain-text conversion). | n/a — operates on local solution zip files, no API auth needed |
