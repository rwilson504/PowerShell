<#
.SYNOPSIS
    Creates a new entity relationship for a polymorphic lookup in Dataverse.

.DESCRIPTION
    This script sends an HTTP POST request to the Dataverse API to create a new entity relationship for a polymorphic lookup.
    The relationship details are provided as parameters to the script.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse API.

.PARAMETER SchemaName
    The schema name for the new relationship.

.PARAMETER ReferencedEntity
    The logical name of the entity that you now want to include as part of the polymorphic lookup.

.PARAMETER ReferencingEntity
    The logical name of the entity which the existing polymorphic lookup exists on.

.PARAMETER LookupSchemaName
    The schema name for the existing lookup attribute.

.PARAMETER LookupDisplayName
    The display name for the existing lookup attribute.

.PARAMETER LookupDescription
    The description for the existing lookup attribute.

.EXAMPLE
    .\AddPolymorphicRelationship.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken "YOUR_ACCESS_TOKEN" -SchemaName "raw_existingentitywithpolylookup_raw_tablebeingaddedtopolylookup_raw_relatedto" -ReferencedEntity "new_tablebeingaddedtopolylookup" -ReferencingEntity "raw_existingentitywithpolylookup" -LookupSchemaName "raw_RelatedTo" -LookupDisplayName "Related To" -LookupDescription ""

    This example creates a new entity relationship for a polymorphic lookup with the specified parameters.
#>

param (
    [string]$OrganizationUrl,
    [string]$AccessToken,
    [string]$SchemaName,
    [string]$ReferencedEntity,
    [string]$ReferencingEntity,
    [string]$LookupSchemaName,
    [string]$LookupDisplayName,
    [string]$LookupDescription
)

# Payload
$payload = @{
    SchemaName = $SchemaName
    '@odata.type' = 'Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata'
    CascadeConfiguration = @{
        Assign = 'NoCascade'
        Delete = 'RemoveLink'
        Merge = 'NoCascade'
        Reparent = 'NoCascade'
        Share = 'NoCascade'
        Unshare = 'NoCascade'
    }
    ReferencedEntity = $ReferencedEntity
    ReferencingEntity = $ReferencingEntity
    Lookup = @{
        AttributeType = 'Lookup'
        AttributeTypeName = @{
            Value = 'LookupType'
        }
        Description = @{
            '@odata.type' = 'Microsoft.Dynamics.CRM.Label'
            LocalizedLabels = @(
                @{
                    '@odata.type' = 'Microsoft.Dynamics.CRM.LocalizedLabel'
                    Label = $LookupDescription
                    LanguageCode = 1033
                }
            )
            UserLocalizedLabel = @{
                '@odata.type' = 'Microsoft.Dynamics.CRM.LocalizedLabel'
                Label = $LookupDescription
                LanguageCode = 1033
            }
        }
        DisplayName = @{
            '@odata.type' = 'Microsoft.Dynamics.CRM.Label'
            LocalizedLabels = @(
                @{
                    '@odata.type' = 'Microsoft.Dynamics.CRM.LocalizedLabel'
                    Label = $LookupDisplayName
                    LanguageCode = 1033
                }
            )
            UserLocalizedLabel = @{
                '@odata.type' = 'Microsoft.Dynamics.CRM.LocalizedLabel'
                Label = $LookupDisplayName
                LanguageCode = 1033
            }
        }
        SchemaName = $LookupSchemaName
        '@odata.type' = 'Microsoft.Dynamics.CRM.LookupAttributeMetadata'
    }
}

# Convert payload to JSON
$jsonPayload = $payload | ConvertTo-Json -Depth 10

# API endpoint
$apiUrl = "$OrganizationUrl/api/data/v9.2/RelationshipDefinitions"

# HTTP request headers
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

# Make the HTTP POST request
$response = Invoke-RestMethod -Method Post -Uri $apiUrl -Headers $headers -Body $jsonPayload

# Output the response
$response
