<#
.SYNOPSIS
    Acquires an access token using device-code or interactive browser authentication.

.DESCRIPTION
    This script authenticates a user and acquires an access token from Azure AD. DeviceCode
    prompts the user to visit a URL and enter a code. Interactive opens the system browser,
    uses the OAuth authorization-code flow with PKCE, and receives the response on a localhost
    loopback redirect.

    Interactive requires the app registration to allow public client flows and to have
    http://localhost registered as a Mobile and desktop applications redirect URI.

    Token caching is supported: on first authentication the token response (including refresh token) is
    cached to a local file. On subsequent runs the cached access token is returned if still valid, or
    silently refreshed via the refresh token. Device code flow is only triggered when no valid cache exists.

    Use the -NoCache switch to skip caching and always perform device code authentication.
    Use the -ClearCache switch to remove any cached token for the given parameters before authenticating.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER ClientId
    The client ID (application ID) of your registered Azure AD app.

.PARAMETER Scope
    The scope for the access token. For Dataverse, this is usually in the form https://your-org.crm.dynamics.com/.default.
    Default value is "https://your-org.crm.dynamics.com/.default".

.PARAMETER Environment
    The Azure environment. Valid values are "Public", "GCC", "GCCH", "DoD". Default value is "Public".

.PARAMETER AuthenticationMode
    Authentication flow to use when no valid cached token is available. Valid values are
    "DeviceCode" and "Interactive". Default is "DeviceCode".

.PARAMETER NoCache
    When specified, skips token caching and always performs device code authentication.

.PARAMETER ClearCache
    When specified, removes the cached token file for the given parameters before authenticating.

.EXAMPLE
    .\GetAccessTokenDeviceCode.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID"

    This example acquires an access token for the specified tenant and client ID using the default scope in the Public environment.

.EXAMPLE
    .\GetAccessTokenDeviceCode.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -ClearCache

    Clears any cached token and performs a fresh device code authentication.

.EXAMPLE
    .\GetAccessTokenDeviceCode.ps1 -TenantId "YOUR_TENANT_ID" -ClientId "YOUR_CLIENT_ID" -Scope "https://your-org.crm.dynamics.com/user_impersonation" -AuthenticationMode Interactive

    Opens the system browser for interactive sign-in using authorization code with PKCE.
#>

param (
    [string]$TenantId,
    [string]$ClientId,
    [string]$Scope = "https://your-org.crm.dynamics.com/.default",
    [ValidateSet("Public", "GCC", "GCCH", "DoD")] [string]$Environment = "Public",
    [ValidateSet("DeviceCode", "Interactive")] [string]$AuthenticationMode = "DeviceCode",
    [switch]$NoCache,
    [switch]$ClearCache
)

# --- Token Cache Helpers ---

# Load DPAPI assembly for token encryption
Add-Type -AssemblyName System.Security

function Get-TokenCachePath {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$Scope
    )
    # Build a deterministic cache key from tenant + client + scope
    $cacheDir = Join-Path ([System.IO.Path]::GetTempPath()) ".ps-token-cache"
    if (-not (Test-Path $cacheDir)) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
    }
    $raw = "$TenantId|$ClientId|$Scope"
    $hash = [System.BitConverter]::ToString(
        [System.Security.Cryptography.SHA256]::Create().ComputeHash(
            [System.Text.Encoding]::UTF8.GetBytes($raw)
        )
    ).Replace("-", "").Substring(0, 16)
    return Join-Path $cacheDir "token_$hash.json"
}

function Get-CachedToken {
    param ([string]$CachePath)
    if (-not (Test-Path $CachePath)) { return $null }
    try {
        # Read DPAPI-encrypted base64 and decrypt
        $encryptedBase64 = Get-Content -Path $CachePath -Raw
        $encryptedBytes = [Convert]::FromBase64String($encryptedBase64)
        $decryptedBytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $encryptedBytes, $null,
            [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        $json = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
        $cached = $json | ConvertFrom-Json
        return $cached
    }
    catch {
        # Corrupt or inaccessible cache file — remove it
        Remove-Item -Path $CachePath -Force -ErrorAction SilentlyContinue
        return $null
    }
}

function Save-TokenToCache {
    param (
        [string]$CachePath,
        [psobject]$TokenResponse
    )
    $cacheEntry = @{
        access_token  = $TokenResponse.access_token
        refresh_token = $TokenResponse.refresh_token
        expires_on    = [DateTimeOffset]::UtcNow.AddSeconds([int]$TokenResponse.expires_in).ToUnixTimeSeconds()
        scope         = $TokenResponse.scope
    }
    # Encrypt with DPAPI (CurrentUser scope) and store as base64
    $json = $cacheEntry | ConvertTo-Json -Depth 3
    $plainBytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
        $plainBytes, $null,
        [System.Security.Cryptography.DataProtectionScope]::CurrentUser
    )
    [Convert]::ToBase64String($encryptedBytes) | Set-Content -Path $CachePath -Encoding UTF8 -Force
}

function Test-TokenValid {
    param ([psobject]$CachedToken)
    if (-not $CachedToken -or -not $CachedToken.access_token -or -not $CachedToken.expires_on) {
        return $false
    }
    # Consider valid if more than 5 minutes remain
    $nowUnix = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    return ($CachedToken.expires_on - $nowUnix) -gt 300
}

# --- Auth Functions ---
function Get-AccessTokenByDeviceCode {
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
            return $tokenResponse
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

function ConvertTo-Base64Url {
    param ([byte[]]$Bytes)

    return [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-AccessTokenInteractively {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$Scope,
        [string]$Environment
    )

    switch ($Environment) {
        "Public" { $loginEndpoint = "https://login.microsoftonline.com" }
        "GCC"    { $loginEndpoint = "https://login.microsoftonline.us" }
        "GCCH"   { $loginEndpoint = "https://login.microsoftonline.us" }
        "DoD"    { $loginEndpoint = "https://login.microsoftonline.us" }
    }

    $tcpListener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $tcpListener.Start()
    $port = ([System.Net.IPEndPoint]$tcpListener.LocalEndpoint).Port
    $tcpListener.Stop()

    $redirectUri = "http://localhost:$port/"
    $stateBytes = New-Object byte[] 32
    $verifierBytes = New-Object byte[] 64
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $random.GetBytes($stateBytes)
        $random.GetBytes($verifierBytes)
    }
    finally {
        $random.Dispose()
    }

    $state = ConvertTo-Base64Url -Bytes $stateBytes
    $codeVerifier = ConvertTo-Base64Url -Bytes $verifierBytes
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $challengeBytes = $sha256.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($codeVerifier))
    }
    finally {
        $sha256.Dispose()
    }
    $codeChallenge = ConvertTo-Base64Url -Bytes $challengeBytes
    $authorizationScope = "$Scope offline_access openid profile"

    $authorizeEndpoint = "$loginEndpoint/$TenantId/oauth2/v2.0/authorize"
    $authorizationParameters = @{
        client_id             = $ClientId
        response_type         = 'code'
        redirect_uri          = $redirectUri
        response_mode         = 'query'
        scope                 = $authorizationScope
        state                 = $state
        code_challenge        = $codeChallenge
        code_challenge_method = 'S256'
    }
    $authorizationQuery = ($authorizationParameters.GetEnumerator() | ForEach-Object {
        "$([Uri]::EscapeDataString($_.Key))=$([Uri]::EscapeDataString([string]$_.Value))"
    }) -join '&'
    $authorizationUri = "$authorizeEndpoint`?$authorizationQuery"

    if (-not $authorizationUri.Contains("client_id=$([Uri]::EscapeDataString($ClientId))") -or
        ($authorizationUri -split '&').Count -ne $authorizationParameters.Count) {
        throw "Interactive authorization URL could not be constructed correctly."
    }

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($redirectUri)
    try {
        $listener.Start()
        Write-Host "Opening the system browser for interactive sign-in..." -ForegroundColor Cyan
        Start-Process $authorizationUri

        $contextTask = $listener.GetContextAsync()
        if (-not $contextTask.Wait([TimeSpan]::FromMinutes(5))) {
            throw "Interactive sign-in timed out after 5 minutes."
        }
        $context = $contextTask.Result

        $responseMessage = "Authentication completed. You can close this browser window."
        $responseBytes = [System.Text.Encoding]::UTF8.GetBytes($responseMessage)
        $context.Response.ContentType = 'text/plain; charset=utf-8'
        $context.Response.ContentLength64 = $responseBytes.Length
        $context.Response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
        $context.Response.OutputStream.Close()

        $query = $context.Request.QueryString
        if ($query['error']) {
            throw "Interactive authentication failed: $($query['error']) - $($query['error_description'])"
        }
        if ($query['state'] -ne $state) {
            throw "Interactive authentication returned an invalid state value."
        }
        if (-not $query['code']) {
            throw "Interactive authentication did not return an authorization code."
        }

        $tokenEndpoint = "$loginEndpoint/$TenantId/oauth2/v2.0/token"
        $tokenRequestBody = @{
            client_id     = $ClientId
            grant_type    = 'authorization_code'
            code          = $query['code']
            redirect_uri  = $redirectUri
            code_verifier = $codeVerifier
            scope         = $authorizationScope
        }

        try {
            return Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $tokenRequestBody
        }
        catch {
            $errorDetail = [string]$_.ErrorDetails.Message
            $errorText = "$($_.Exception.Message) $errorDetail"
            if ($errorText -match 'AADSTS7000218|client_assertion|client secret') {
                throw "Interactive authentication requires a public-client app registration. In Entra ID, add http://localhost under Authentication > Mobile and desktop applications and set Allow public client flows to Yes. Do not add a client secret to this local script. Original error: $($_.Exception.Message)"
            }
            throw
        }
    }
    finally {
        if ($listener.IsListening) {
            $listener.Stop()
        }
        $listener.Close()
    }
}

function Get-AccessTokenByRefreshToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$Scope,
        [string]$RefreshToken,
        [string]$Environment
    )

    switch ($Environment) {
        "Public" { $loginEndpoint = "https://login.microsoftonline.com" }
        "GCC"    { $loginEndpoint = "https://login.microsoftonline.us" }
        "GCCH"   { $loginEndpoint = "https://login.microsoftonline.us" }
        "DoD"    { $loginEndpoint = "https://login.microsoftonline.us" }
    }

    $tokenEndpoint = "$loginEndpoint/$TenantId/oauth2/v2.0/token"

    $tokenRequestBody = @{
        client_id     = $ClientId
        grant_type    = "refresh_token"
        refresh_token = $RefreshToken
        scope         = $Scope
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType "application/x-www-form-urlencoded" -Body $tokenRequestBody
    return $tokenResponse
}

# --- Main Script ---

$cachePath = Get-TokenCachePath -TenantId $TenantId -ClientId $ClientId -Scope $Scope

# Handle -ClearCache
if ($ClearCache -and (Test-Path $cachePath)) {
    Remove-Item -Path $cachePath -Force
    Write-Host "Token cache cleared." -ForegroundColor Yellow
}

$tokenResponse = $null

if (-not $NoCache) {
    $cached = Get-CachedToken -CachePath $cachePath

    if ($cached -and (Test-TokenValid -CachedToken $cached)) {
        # Cached access token is still valid
        Write-Host "Using cached access token." -ForegroundColor Green
        Write-Output $cached.access_token
        return
    }

    if ($cached -and $cached.refresh_token) {
        # Try silent refresh
        Write-Host "Access token expired. Refreshing silently..." -ForegroundColor Cyan
        try {
            $tokenResponse = Get-AccessTokenByRefreshToken -TenantId $TenantId -ClientId $ClientId -Scope $Scope -RefreshToken $cached.refresh_token -Environment $Environment
            Write-Host "Token refreshed successfully." -ForegroundColor Green
        }
        catch {
            Write-Warning "Silent refresh failed. Falling back to $AuthenticationMode authentication."
            $tokenResponse = $null
        }
    }
}

# Fall back to the requested interactive flow
if (-not $tokenResponse) {
    if ($AuthenticationMode -eq 'Interactive') {
        $tokenResponse = Get-AccessTokenInteractively -TenantId $TenantId -ClientId $ClientId -Scope $Scope -Environment $Environment
    }
    else {
        $tokenResponse = Get-AccessTokenByDeviceCode -TenantId $TenantId -ClientId $ClientId -Scope $Scope -Environment $Environment
    }
}

# Cache the token response
if (-not $NoCache -and $tokenResponse) {
    Save-TokenToCache -CachePath $cachePath -TokenResponse $tokenResponse
}

# Output the access token
Write-Output $tokenResponse.access_token
