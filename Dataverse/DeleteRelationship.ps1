<#
.SYNOPSIS
    Deletes an existing entity relationship from a polymorphic attribute in Dataverse.

.DESCRIPTION
    This script sends an HTTP DELETE request to the Dataverse API to delete an existing entity relationship for a polymorphic attribute.
    The relationship name and details are provided as parameters to the script.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization.

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse API.

.PARAMETER RelationshipName
    The name of the relationship to delete.

.EXAMPLE
    .\DeleteRelationship.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken "YOUR_ACCESS_TOKEN" -RelationshipName "new_checkout_poly_new_researchresource"

    This example deletes the specified relationship from the Dataverse organization.
#>

param (
    [string]$OrganizationUrl,
    [string]$AccessToken,
    [string]$RelationshipName
)

# API endpoint
$apiUrl = "$OrganizationUrl/api/data/v9.2/RelationshipDefinitions(SchemaName='$RelationshipName')"

# HTTP request headers
$headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

# Make the HTTP DELETE request
$response = Invoke-RestMethod -Method Delete -Uri $apiUrl -Headers $headers

# Output the response
$response
