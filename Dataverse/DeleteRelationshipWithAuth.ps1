<#
.SYNOPSIS
    Acquires an access token and deletes an existing relationship in Dataverse.

.DESCRIPTION
    This script calls the GetAccessTokenDeviceCode script to acquire an access token and then calls the DeleteRelationship script to delete an existing relationship in Dataverse.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER RelationshipName
    The name of the relationship to delete.

.EXAMPLE
    .\RunDelete.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -OrganizationUrl "https://your-org.crm.dynamics.com" -RelationshipName "new_checkout_poly_new_researchresource"

    This example acquires an access token and deletes the specified relationship from the Dataverse organization.
#>

param (
    [string]$TenantId,
    [string]$ClientId,
    [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public",
    [string]$OrganizationUrl,
    [string]$RelationshipName
)

# Get the access token
$accessToken = & ..\EntraID\GetAccessTokenDeviceCode.ps1 -TenantId $TenantId -ClientId $ClientId -Scope "$OrganizationUrl/user_impersonation" -Environment $Environment

# Delete the relationship
$response = & .\DeleteRelationship.ps1 -OrganizationUrl $OrganizationUrl -AccessToken $accessToken -RelationshipName $RelationshipName

# Output the response
Write-Output $response
