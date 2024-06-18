<#
.SYNOPSIS
    Acquires an access token and creates a new entity relationship for a polymorphic lookup in Dataverse.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token and then calls the AddPolymorphicRelationship script to create a new entity relationship.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Scope
    The scope for the access token. For Dataverse, this is usually in the form https://your-org.crm.dynamics.com/.default.
    Default value is "https://your-org.crm.dynamics.com/.default".

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER SchemaName
    The schema name for the new relationship.

.PARAMETER ReferencedEntity
    The name of the referenced entity in the relationship.

.PARAMETER ReferencingEntity
    The name of the referencing entity in the relationship.

.PARAMETER LookupSchemaName
    The schema name for the lookup attribute.

.PARAMETER LookupDisplayName
    The display name for the lookup attribute.

.PARAMETER LookupDescription
    The description for the lookup attribute.

.EXAMPLE
    .\AddPolymorphicRelationshipWithAuth.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -SchemaName "new_checkout_poly_new_researchresource" -ReferencedEntity "new_researchresource" -ReferencingEntity "new_checkout" -LookupSchemaName "new_CheckedoutItem" -LookupDisplayName "Checkout item" -LookupDescription "Checkout Polymorphic Lookup Attribute"

    This example acquires an access token for the specified tenant and client ID using the default scope in the Public environment, then creates a new entity relationship for a polymorphic lookup with the specified parameters.
#>

param (
    [string]$TenantId,
    [string]$ClientId,
    [string]$Scope = "https://your-org.crm.dynamics.com/.default",
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
$accessToken = & ..\EntraID\GetAccessTokenDeviceCode.ps1 -TenantId $TenantId -ClientId $ClientId -Scope $Scope -Environment $Environment

# Create the relationship
$response = & .\AddPolymorphicRelationship.ps1 -OrganizationUrl $OrganizationUrl -AccessToken $accessToken -SchemaName $SchemaName -ReferencedEntity $ReferencedEntity -ReferencingEntity $ReferencingEntity -LookupSchemaName $LookupSchemaName -LookupDisplayName $LookupDisplayName -LookupDescription $LookupDescription

# Output the response
Write-Output $response