# EntraID

Microsoft Entra ID authentication helpers and queries.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [GetAccessTokenDeviceCode.ps1](GetAccessTokenDeviceCode.ps1) | Acquires an access token using the device-code flow against a user-specified resource (Dataverse, Graph, etc.). Supports Public, GCC, GCCH, and DoD Azure clouds. Used as the auth helper by every `*WithAuth.ps1` script in the repo. | n/a — this IS the auth helper |
| [GetFirstPartyServicePrincipals.ps1](GetFirstPartyServicePrincipals.ps1) | Lists all first-party (Microsoft-owned) service principals in the tenant. Useful for identifying which Microsoft apps have been consented to in your environment. | Yes — [GetFirstPartyServicePrincipalsWithAuth.ps1](GetFirstPartyServicePrincipalsWithAuth.ps1) |

## Public-client app registration

These local PowerShell scripts authenticate on behalf of a user as a public
client, without a client secret. In **Entra ID > App registrations**, select the
application whose **Application (client) ID** you pass as `-ClientId`.

For the device-code helper in this checkout, open **Authentication > Advanced
settings**, set **Allow public client flows** to **Yes**, and save. Device-code
authentication does not require a redirect URI.

### Interactive browser sign-in

The configuration below was confirmed working with an interactive-login version
of the helper. The helper currently checked into this repository supports only
device-code authentication; it does not expose `-AuthenticationMode Interactive`.

For a version that opens the system browser and listens on localhost:

1. Open the app registration's **Authentication** page.
2. Remove this script's `http://localhost` redirect URI, including equivalent
	localhost entries, from **Web** and **Single-page application (SPA)**. Preserve
	unrelated callbacks used by other applications.
3. Select **Add a platform > Mobile and desktop applications**.
4. Add `http://localhost` as a custom redirect URI and select **Configure**. If
	your helper uses a different callback path or host, match its redirect URI.
5. Keep **Allow public client flows** set to **Yes** to support device-code
	authentication as well, and save.
6. Rerun the interactive command to start a fresh sign-in. Do not reuse the
	authorization code from a failed attempt.

Opening a browser does not make the script a Web or SPA application: PowerShell,
not browser JavaScript, redeems the authorization code.

The platform dialog also offers suggested desktop redirect URIs. If selected,
these can appear alongside `http://localhost`. Additional desktop callbacks
generally do not interfere because the script specifies its callback. Keep only
those needed by the application's clients; their exact purpose depends on the
URLs selected.

### Dataverse URL and permissions

The Dataverse organization URL is the API resource, not this script's callback.
Pass it as `-OrganizationUrl` to the Dataverse scripts; their auth wrappers use
it to construct the `user_impersonation` scope. Do not add the organization URL
as a Web redirect URI solely for these scripts. Configure the app's delegated
Dataverse `user_impersonation` permission and obtain consent as required by your
tenant. The signed-in user still needs appropriate Dataverse access.

### Troubleshooting platform errors

| Symptom | Meaning and action |
|---|---|
| `AADSTS9002327: Tokens issued for the 'Single-Page Application' client-type may only be redeemed via cross-origin requests` | The callback is registered as SPA, but PowerShell is redeeming the code. Move the localhost callback to **Mobile and desktop applications**, save, and start a fresh sign-in. Do not add an `Origin` header to work around the mismatch. |
| The helper reports that interactive authentication requires a public-client app registration after localhost was configured under **Web** | Check the underlying Entra error, then correct the callback platform to **Mobile and desktop applications**. A generic HTTP 401 alone does not establish the cause. Do not add a client secret to this local public-client flow. |
| Changing **Allow public client flows** does not fix the SPA error | That switch does not change a redirect URI's platform type. Correct the localhost entry itself. |

Keep MFA and Conditional Access protections in place; disabling them is not a
fix for a redirect-platform mismatch.

References:

- [Microsoft: Desktop app authentication configuration](https://learn.microsoft.com/entra/identity-platform/scenario-desktop-app-configuration)
- [Microsoft: Redirect URIs for single-page apps](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-auth-code-flow#redirect-uris-for-single-page-apps-spas)
- [Microsoft: Register an app for Dataverse](https://learn.microsoft.com/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory#public-client-app-registration)
