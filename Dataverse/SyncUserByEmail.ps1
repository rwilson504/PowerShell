<#
.SYNOPSIS
    Syncs a user to Dataverse environment

.DESCRIPTION
    Retrieves user from an Azure AD using email address and synchronizes the user
    as a System Users within a Dataverse environment.

.AUTHOR
    Rick Wilson

.PARAMETER OrganizationId
    You can find this by opening the Power Platform Admin site (admin.powerplatform.com) and copying the Organization ID from the details page for your environment.

.PARAMETER UserEmail
    The email address of the user to sync

.EXAMPLE
    ./SyncUserByEmail.ps1 -OrganizationId 02c201b0-db76-4a6a-b3e1-a69202b479e6 -UserEmail tom@test.com

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

param(
[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
[string]$OrganizationId,

[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
[string]$UserEmail
)

# Import the modules to make sure you can access the cmdlets within this PowerShell session.
Import-Module -Name Microsoft.PowerApps.Administration.PowerShell
Import-Module -Name AzureAD

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

# Retrieves the ObjectId of the user
$User = Get-AzureADUser -ObjectId "$UserEmail"

# Sync the user to Dataverse organization
Add-AdminPowerAppsSyncUser -EnvironmentName $Environment.EnvironmentName -PrincipalObjectId $User.ObjectId

