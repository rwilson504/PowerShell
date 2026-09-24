# EntraID

Microsoft Entra ID authentication helpers and queries.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [GetAccessTokenDeviceCode.ps1](GetAccessTokenDeviceCode.ps1) | Acquires an access token against a user-specified resource using device-code authentication or interactive browser authorization code with PKCE. Interactive mode requires a public-client app registration with a `http://localhost` redirect URI. Supports Public, GCC, GCCH, and DoD Azure clouds. | n/a — this IS the auth helper |
| [GetFirstPartyServicePrincipals.ps1](GetFirstPartyServicePrincipals.ps1) | Lists all first-party (Microsoft-owned) service principals in the tenant. Useful for identifying which Microsoft apps have been consented to in your environment. | Yes — [GetFirstPartyServicePrincipalsWithAuth.ps1](GetFirstPartyServicePrincipalsWithAuth.ps1) |
