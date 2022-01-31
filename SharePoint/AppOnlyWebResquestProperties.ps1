<#
.SYNOPSIS
    Produced the URL and Request body for SharePoint App-only Authentication Call
.DESCRIPTION
    When authenticating SharePoint using the App-Only method the URL and Body of the request
    must be correctly set.  This script will allow you to generate data based upon the SharePoint
    site Uri you supply.

    These values can also be used within a Power Automate Desktop action to provide authentication
    to SharePoint using the Invoke web service action.

    To generate a service principal for app-only authentication follow the instructions here:
    https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs
.AUTHOR
    Rick Wilson
.PARAMETER SPSiteUri
    Enter the sharepoint site uri which you are attempting to authenticate to
.EXAMPLE
    PS> ./AppOnlyWebResquestProperties.ps1 -SPSiteUri https://yoursharepoint.sharepoint.com/sites/Clients
#>

Param(
    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [string]
    $SPSiteUri
)

try { 
    $response = Invoke-WebRequest ($SPSiteUri + '/_vti_bin/client.svc') -Headers @{'Accept' = 'application/json'; 'Authorization' = 'Bearer'} 
} 
catch {
    if ($_.Exception.Response.StatusCode -eq 'Unauthorized'){
	    $authHeader = $_.Exception.Response.Headers['WWW-Authenticate'];
        $tenantHost = $_.Exception.Response.ResponseUri.Host;
        $client_id = (Select-String '(?:client_id=)("([^""]+)")' -inputobject $authHeader).Matches[0].Groups[2].Value;
        $tenant = (Select-String '(?:realm=)("([^""]+)")' -inputobject $authHeader).Matches[0].Groups[2].Value
        Write-Host "URL: Copy the output below to the URL field of the Invoke web service action.`n" -ForegroundColor gray; 
        Write-Host ("https://accounts.accesscontrol.windows.net/{0}/tokens/OAuth/2`n" -f $tenant)

        Write-Host "Request Body: Copy the output below to the Request body of the Invoke web service action and replace the <your app id> and <your client secret> values.`n" -ForegroundColor gray;
        Write-Host ("client_id=<your app id>@{1}`n&grant_type=client_credentials`n&resource={0}/{2}@{1}`n&client_secret=<your client secret>" -f $client_id,$tenant,$tenantHost)
    }
    else{
        Write-Host "ERROR: The SharePoint Site URL you entered could not be found.`nPlease check that it is a valid Url and try again."
    }
}