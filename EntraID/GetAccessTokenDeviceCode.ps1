<#
.SYNOPSIS
    Acquires an access token using the device code flow for Azure AD, including support for GCC, GCCH, and DoD environments.

.DESCRIPTION
    This script uses the device code flow to authenticate a user and acquire an access token from Azure AD.
    The user will be prompted to visit a URL and enter a code to complete the authentication.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Scope
    The scope for the access token. For Dataverse, this is usually in the form https://your-org.crm.dynamics.com/.default.
    Default value is "https://your-org.crm.dynamics.com/.default".

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.EXAMPLE
    .\GetAccessTokenDeviceCode.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID"

    This example acquires an access token for the specified tenant and client ID using the default scope in the Public environment.
#>

param (
    [string]$TenantId,
    [string]$ClientId,
    [string]$Scope = "https://your-org.crm.dynamics.com/.default",
    [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public"
)

# Function to get the access token using device code flow
function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$Scope,
        [string]$Environment
    )

    # Set the correct login endpoints based on the environment
    switch ($Environment) {
        "Public" {
            $loginEndpoint = "https://login.microsoftonline.com"
        }
        "GCC" {
            $loginEndpoint = "https://login.microsoftonline.us"
        }
        "GCCH" {
            $loginEndpoint = "https://login.microsoftonline.us"
        }
        "DoD" {
            $loginEndpoint = "https://login.microsoftonline.us"
        }
        default {
            throw "Invalid environment specified."
        }
    }

    # Device code endpoint
    $deviceCodeEndpoint = "$loginEndpoint/$TenantId/oauth2/v2.0/devicecode"
    
    # Token endpoint
    $tokenEndpoint = "$loginEndpoint/$TenantId/oauth2/v2.0/token"

    # Prepare device code request body
    $deviceCodeRequestBody = @{
        client_id = $ClientId
        scope     = $Scope
    }

    # Request device code
    $deviceCodeResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeEndpoint -ContentType "application/x-www-form-urlencoded" -Body $deviceCodeRequestBody

    # Display instructions to the user
    Write-Host "To sign in, use a web browser to open the page $($deviceCodeResponse.verification_uri) and enter the code $($deviceCodeResponse.user_code) to authenticate."

    # Prepare token request body
    $tokenRequestBody = @{
        client_id     = $ClientId
        grant_type    = "urn:ietf:params:oauth:grant-type:device_code"
        device_code   = $deviceCodeResponse.device_code
        scope         = $Scope
    }

    # Poll the token endpoint until we get the token or an error
    while ($true) {
        Start-Sleep -Seconds $deviceCodeResponse.interval
        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $tokenRequestBody
            return $tokenResponse.access_token
        }
        catch {
            if ($_.Exception.Response.StatusCode -ne 400) {
                throw $_
            }
            $errorResponse = ($_ | ConvertFrom-Json)
            if ($errorResponse.error -ne "authorization_pending") {
                throw $_
            }
        }
    }
}

# Main script
$accessToken = Get-AccessToken -TenantId $TenantId -ClientId $ClientId -Scope $Scope -Environment $Environment

# Output the access token
Write-Output $accessToken
