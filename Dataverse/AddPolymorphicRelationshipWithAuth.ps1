<#
.SYNOPSIS
    Acquires an access token and creates a new entity relationship for a polymorphic lookup in Dataverse.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token and then calls the AddPolymorphicRelationship script to create a new entity relationship.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

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
    .\AddPolymorphicRelationshipWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -SchemaName "raw_existingentitywithpolylookup_raw_tablebeingaddedtopolylookup_raw_relatedto" -ReferencedEntity "new_tablebeingaddedtopolylookup" -ReferencingEntity "raw_existingentitywithpolylookup" -LookupSchemaName "raw_RelatedTo" -LookupDisplayName "Related To" -LookupDescription ""

    This example acquires an access token for the specified tenant and client ID using the default scope in the Public environment, then creates a new entity relationship for a polymorphic lookup with the specified parameters.
#>

param (
    [string]$TenantId,
    [string]$ClientId,    
    [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public",
    [string]$OrganizationUrl,
    [string]$SchemaName,
    [string]$ReferencedEntity,
    [string]$ReferencingEntity,
    [string]$LookupSchemaName,
    [string]$LookupDisplayName,
    [string]$LookupDescription
)

# Get the access token
$accessToken = & ..\EntraID\GetAccessTokenDeviceCode.ps1 -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

# Create the relationship
$response = & .\AddPolymorphicRelationship.ps1 -OrganizationUrl $OrganizationUrl -AccessToken $accessToken -SchemaName $SchemaName -ReferencedEntity $ReferencedEntity -ReferencingEntity $ReferencingEntity -LookupSchemaName $LookupSchemaName -LookupDisplayName $LookupDisplayName -LookupDescription $LookupDescription

# Output the response
Write-Output $response