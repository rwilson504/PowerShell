# EntraID

Microsoft Entra ID authentication helpers and queries.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [GetAccessTokenDeviceCode.ps1](GetAccessTokenDeviceCode.ps1) | Acquires an access token using the device-code flow against a user-specified resource (Dataverse, Graph, etc.). Supports Public, GCC, GCCH, and DoD Azure clouds. Used as the auth helper by every `*WithAuth.ps1` script in the repo. | n/a — this IS the auth helper |
| [GetFirstPartyServicePrincipals.ps1](GetFirstPartyServicePrincipals.ps1) | Lists all first-party (Microsoft-owned) service principals in the tenant. Useful for identifying which Microsoft apps have been consented to in your environment. | Yes — [GetFirstPartyServicePrincipalsWithAuth.ps1](GetFirstPartyServicePrincipalsWithAuth.ps1) |
