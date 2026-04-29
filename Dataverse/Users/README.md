# Dataverse / Users

User-provisioning utilities — sync individual users or all members of a security group into the environment.

## Scripts

| Script | Purpose | WithAuth pair? |
|---|---|---|
| [SyncUserByEmail.ps1](SyncUserByEmail.ps1) | Syncs a single user (looked up by email address) into the Dataverse environment. | n/a — uses the Power Platform PowerShell module's built-in auth |
| [SyncUsersFromSecurityGroup.ps1](SyncUsersFromSecurityGroup.ps1) | Syncs every member of a Microsoft Entra security group into the environment. | n/a — uses the Power Platform PowerShell module's built-in auth |
