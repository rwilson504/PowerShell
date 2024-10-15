<#
.SYNOPSIS
    This script disables the consent prompt for specified Power Apps in an environment.

.DESCRIPTION
    The script provides three ways to disable the consent prompt:
    1. Using a configuration file that contains environment and app details.
    2. Using a provided environment ID and app ID.
    3. Allowing the user to select an app from a given environment.

.AUTHOR
    Rick Wilson

.PARAMETER configFilePath
    The path to the configuration JSON file containing environment and app details.

.PARAMETER configEnvironmentName
    The name of the environment (e.g., dev, prod, qa) as specified in the configuration file.

.PARAMETER environmentId
    The ID of the environment to be used for disabling consent if a configuration file is not provided.

.PARAMETER appId
    The ID of the app to be used for disabling consent if a configuration file is not provided.

.EXAMPLE
    ./DisableAppConsent.ps1 -configFilePath "./config.json" -configEnvironmentName "dev"
    This example uses the configuration file to disable consent for all apps in the "dev" environment.

.EXAMPLE
    ./DisableAppConsent.ps1 -environmentId "env-id" -appId "app-id"
    This example disables consent for a specific app using the provided environment ID and app ID.

.EXAMPLE
    ./DisableAppConsent.ps1 -environmentId "env-id"
    This example allows the user to select an app from the specified environment to disable consent.

#>

# Parameter definitions
param (
    [Parameter(Mandatory=$false, HelpMessage="The path to the configuration JSON file containing environment and app details.")]
    [string]$configFilePath = "",

    [Parameter(Mandatory=$false, HelpMessage="The name of the environment (e.g., dev, prod, qa) as specified in the configuration file.")]
    [string]$configEnvironmentName = "",

    [Parameter(Mandatory=$false, HelpMessage="The ID of the environment to be used for disabling consent if a configuration file is not provided.")]
    [string]$environmentId = "",

    [Parameter(Mandatory=$false, HelpMessage="The ID of the app to be used for disabling consent if a configuration file is not provided.")]
    [string]$appId = ""
)

# Function to check if a module is installed and import it if necessary
function Import-ModuleIfNeeded {
    param (
        [Parameter(Mandatory=$true)] [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Output "Module $ModuleName not found. Installing..."
        Install-Module -Name $ModuleName -Force -AllowClobber
    }
    Import-Module -Name $ModuleName
}

# Import required module if it is not already installed
Import-ModuleIfNeeded -ModuleName "Microsoft.PowerApps.Administration.PowerShell"

# Function to Disable Consent on a given App
function Disable-ConsentForApp {
    param (
        [Parameter(Mandatory=$true)] [string]$EnvironmentId,
        [Parameter(Mandatory=$true)] [string]$AppId
    )
    
    # Disable consent
    Set-AdminPowerAppApisToBypassConsent -EnvironmentName $EnvironmentId -AppName $AppId
}

# Main logic to determine which flow to execute
if ($configFilePath -ne "" -and $configEnvironmentName -ne "") {
    # Load configuration file
    if (Test-Path $configFilePath) {
        $config = Get-Content -Raw -Path $configFilePath | ConvertFrom-Json
        
        $environment = $config.environments | Where-Object { $_.name -eq $configEnvironmentName }
        if ($null -ne $environment) {
            foreach ($appName in $environment.appNames) {
                Write-Output "Retrieving App ID for app: $appName in environment: $environment.name"
                $app = Get-AdminPowerApp $appName -EnvironmentName $environment.environmentId
                if ($null -ne $app) {
                    Write-Output "Processing app: $appName with App ID: $($app.AppName) in environment: $environment.name"
                    Disable-ConsentForApp -EnvironmentId $environment.environmentId -AppId $app.AppName
                } else {
                    Write-Output "App: $appName not found in environment: $environment.name"
                }
            }
        } else {
            Write-Output "Environment: $configEnvironmentName not found in configuration file."
            exit
        }
    } else {
        Write-Output "Configuration file not found at path: $configFilePath"
        exit
    }
} elseif ($environmentId -ne "" -and $appId -ne "") {
    # Use provided environment and app ID
    Write-Output "Processing app: $appId in environment: $environmentId"
    Disable-ConsentForApp -EnvironmentId $environmentId -AppId $appId
} elseif ($environmentId -ne "") {
    # Prompt user to select an app from the provided environment
    Write-Output "Loading Power Apps from environment: $environmentId..."
    $apps = Get-AdminPowerApp -EnvironmentName $environmentId
    $appChoices = $apps | ForEach-Object { "$_.DisplayName ($_.AppName)" }

    $selectedAppIndex = $appChoices | Out-GridView -Title "Select a Power App to disable consent" -OutputMode Single
    $selectedApp = $apps[$selectedAppIndex]

    if ($null -ne $selectedApp) {
        Disable-ConsentForApp -EnvironmentId $environmentId -AppId $selectedApp.AppId
    } else {
        Write-Output "No app selected. Exiting."
    }
} else {
    Write-Output "Please provide either a configuration file path and config file environment name, or an environment ID."
    exit
}

Write-Output "Operation completed."
