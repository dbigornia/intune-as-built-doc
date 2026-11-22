# Requirements

## Runtime

| Component | Version | Needed for | Install |
|---|---|---|---|
| PowerShell | 7.0+ (`pwsh.exe`) | Both scripts | `winget install --id Microsoft.PowerShell --source winget` |
| Microsoft.Graph.Authentication | 2.x | Collect | `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser` |
| Microsoft Word (desktop, Windows) | Any current | PDF export, preferred engine | - |
| LibreOffice | 7.x+ | PDF export fallback / servers | `winget install --id TheDocumentFoundation.LibreOffice` |

Build has no module dependencies; the .docx is written directly.

## Microsoft Graph permissions (all read-only)

| Permission | Used for |
|---|---|
| DeviceManagementConfiguration.Read.All | Policies, settings catalog, compliance, update profiles |
| DeviceManagementApps.Read.All | Apps, assignments, relationships |
| DeviceManagementServiceConfig.Read.All | Autopilot profiles/devices, ESP, enrollment configs |
| DeviceManagementScripts.Read.All | Platform scripts and remediations (required since July 2025) |
| DeviceManagementRBAC.Read.All | Scope tag names |
| Group.Read.All | Group names in assignments |
| Organization.Read.All | Tenant name |

### App registration 1 - unattended (certificate)
- The permissions above as **Application** permissions, with admin consent (Global Admin or Privileged Role Administrator, one time).
- Certificate credential; keep the private key on the collection server only (non-exportable).

### App registration 2 - team members (interactive, optional)
- The same permissions as **Delegated** permissions, with admin consent.
- Authentication: Mobile and desktop applications, redirect URI `http://localhost`.
- Enterprise app: **Assignment required = Yes**, assign the team's group.
- Users need an Intune read role (Read Only Operator, or a custom read-only role). Scope tags limit what they see.

## Platform notes
- Word PDF export is not supported in unattended scheduled tasks; use `-PdfEngine LibreOffice` there.
- Uses Graph **beta** endpoints for Intune configuration.
