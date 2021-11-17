<#
.SYNOPSIS
    Syncs security group members to environment

.DESCRIPTION
    Retrieves users from an Azure AD security group and synchronizes those users
    as System Users with a Dataverse environment.

.AUTHOR
    Rick Wilson

.PARAMETER OrganizationId
    You can find this by opening the Power Platform Admin site (admin.powerplatform.com) and copying the Organization ID from the details page for your environment.

.PARAMETER SecurityGroupId
    The Object Id of the group in the Azure Portal (portal.azure.com)

.EXAMPLE
    ./SyncUsersFromSecurityGroup.ps1 -OrganizationId 02c201b0-db76-4a6a-b3e1-a69202b479e6 -SecurityGroupId e25a94b2-3111-468e-9125-3d3db3938f13

.PREREQUISITES
    
    .INSTALLS
    Import the necessary modules using the following commands:
    Install-Module -Name Microsoft.PowerApps.Administration.PowerShell
    Install-Module -Name AzureAD

    Alternatively, if you don't have admin rights on your computer, you can use the following to use these modules:
    Save-Module -Name Microsoft.PowerApps.Administration.PowerShell -Path
    Save-Module -Name AzureAD -Path

    .SCRIPT SECURITY
    To run these commands you may need to enable your machine to allow remote signed scripts
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
    
    After you are done you can reset these permissions back to default settings using
    Set-ExecutionPolicy Restricted
#>

# Import the modules to make sure you can access the cmdlets within this PowerShell session.
Import-Module -Name Microsoft.PowerApps.Administration.PowerShell
Import-Module -Name AzureAD

[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
[string]$OrganizationId,

[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
[string]$SecurityGroupId

# Conenct to Azure AD. Will prompt for Azure AD authentication
# For information on how to utilize a username/password or application user
# credentials see here: https://docs.microsoft.com/en-us/powershell/module/azuread/connect-azuread?view=azureadps-2.0
Connect-AzureAD -Confirm

# Connect to Power Platform. Will prompt Power Platform authentication
# For information on how to utilize a username/password or application user
# credentials see here: https://docs.microsoft.com/en-us/powershell/module/microsoft.powerapps.administration.powershell/add-powerappsaccount?view=pa-ps-latest
Add-PowerAppsAccount

# Get the specific environment we are going to use for sync
$Environment = Get-AdminPowerAppEnvironment | Where-Object {$_.OrganizationId -eq $OrganizationId} | Select -First 1

# Retrieves the members of the Azure AD Security Group
$GroupMembers = Get-AzureADGroupMember -ObjectId $SecurityGroupId

# Loop through all of the Users listed in the AD Security Group
ForEach($User in $GroupMembers | Where-Object {$_.ObjectType -eq 'User'})
{
    # Synce each user to the Power Platform Environment.
    Add-AdminPowerAppsSyncUser -EnvironmentName $Environment.EnvironmentName -PrincipalObjectId $User.ObjectId
}
