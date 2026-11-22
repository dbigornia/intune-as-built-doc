# IntuneAsBuilt

Generates a Word as-built document for Intune (Windows enrollment, policies, apps & scripts) in two stages:

| Stage | Script | Needs | Output |
|---|---|---|---|
| 1. Collect | `Invoke-IntuneCollect.ps1` | Graph (read-only) | `snapshots\<yyyy-MM-dd_HHmmZ>\*.json` |
| 2. Build | `Build-IntuneAsBuilt.ps1` | Nothing - works offline (PDF: Word or LibreOffice) | `Intune-AsBuilt_<tenant>_<snapshot>.docx` (+ `.pdf`) |

Snapshots are your point-in-time baseline and audit evidence. You can rebuild the document from any old
snapshot, and two snapshots can be diffed (`git diff --no-index` or VS Code compare) to see what changed.

## Requirements

Full list (modules, Graph permissions, app registrations): [REQUIREMENTS.md](REQUIREMENTS.md)

- PowerShell 7+
- `Microsoft.Graph.Authentication` (Collect only): `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- No Word, PSWriteOffice or OpenXML SDK - the .docx is written directly.

## App registration (unattended runs)

Create a dedicated app registration with a certificate and these **application** permissions (admin consent):

| Permission | Used for |
|---|---|
| DeviceManagementConfiguration.Read.All | Policies, settings catalog, compliance, update profiles |
| DeviceManagementApps.Read.All | Apps, assignments, relationships |
| DeviceManagementServiceConfig.Read.All | Autopilot profiles/devices, ESP, enrollment configs |
| DeviceManagementScripts.Read.All | Platform scripts and remediations (required since July 2025) |
| DeviceManagementRBAC.Read.All | Scope tag names |
| Group.Read.All | Group names in assignments |
| Organization.Read.All | Tenant name |

All calls are `GET`. Nothing is written to the tenant.

## Running it without Global Admin

- **Build** needs no permissions at all. Anyone can build documents from snapshot folders you share.
- **Collect, certificate (app-only):** after one-time admin consent, the app's permissions are used, not the
  user's. Whoever has the certificate's private key can collect, so run it from one management server or
  service account rather than handing the certificate out.
- **Collect, interactive (delegated) for team members:** create a second app registration (public client,
  redirect URI `http://localhost`) with the same permissions as **delegated**. Grant admin consent once, set
  **Assignment required = Yes** on its enterprise app, and assign the team's group. Users then run
  `-Interactive -ClientId <appId>`. What they can read is limited by their own role: give them the Intune
  **Read Only Operator** role or a custom read-only Intune role. Intune scope tags still apply, so a scoped
  user gets a partial document.

## Usage

```powershell
# 1. Collect (certificate auth)
.\Invoke-IntuneCollect.ps1 -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumb>

#    or interactively in the lab
.\Invoke-IntuneCollect.ps1 -Interactive -IncludeAutopilotDevices

# 2. Build from the newest snapshot, and create the design-notes CSV
.\Build-IntuneAsBuilt.ps1 -ExportNotesTemplate .\notes.csv

# 3. Fill in notes.csv, then rebuild with notes and your company template
.\Build-IntuneAsBuilt.ps1 -NotesPath .\notes.csv -ExportNotesTemplate .\notes.csv -TemplatePath .\templates\Company-Template.docx
```

Word asks to update fields when you open the document - choose **Yes** to build the table of contents.

### Collect options
- `-IncludeScriptContent` - store script bodies in the snapshot and print them in Appendix A. Without it, only a SHA-256 hash is stored (changes are still detectable). Review scripts for embedded secrets before turning this on.
- `-IncludeAutopilotDevices` - adds the device register (counts by group tag and model in the document; serials in the JSON only).

### Build options
- `-SnapshotPath` - a specific snapshot (default: newest).
- `-AllPlatforms` - include iOS/Android/macOS objects (default: Windows only).
- `-IncludeEmptyValues` - show unset properties too.
- `-HideMissingNotes` - drop the "Design notes: not yet documented" line.
- `-Pdf` - also save a PDF next to the .docx.
  - **Word** (used automatically on Windows when installed): builds the table of contents, adds heading bookmarks to the PDF, and re-saves the .docx with the TOC filled in.
  - **LibreOffice** (fallback, or `-PdfEngine LibreOffice`): no Office needed, but the TOC isn't built in the PDF.
  - Microsoft doesn't support automating Word from an unattended scheduled task. For scheduled runs on a server, install LibreOffice and use `-PdfEngine LibreOffice`.

## Design notes (the "why")

`notes.csv` lives in the repo root and is committed, so the history of every design decision is kept in git.

`notes.csv` is keyed by Intune object ID, so renaming a policy doesn't lose its notes:

`ObjectId, Category, Name, Purpose, Owner, DependsOn, LastValidated, Rollback`

Re-running with `-ExportNotesTemplate` adds new objects and keeps everything already written, including
notes for objects that were deleted from the tenant (marked `(not in this snapshot)`).

## Company template

A generic template is included at `templates/Company-Template.docx`: header, footer with page numbers, a cover
page with document control and revision history, and the `{{ASBUILT}}` marker. Replace `COMPANY NAME` and
`[COMPANY LOGO]` with your own (Insert > Pictures for the logo), adjust colours in the styles, and commit it.

You can also pass your own `.docx` or `.dotx`. Its headers, footers, logo, page size, margins and styles are kept.
- Put a paragraph containing `{{ASBUILT}}` where the generated content should go (for example after your cover page). Without it, the template body is replaced.
- Make sure the template defines **Title, Heading 1, Heading 2 and Heading 3** (apply each once, then delete the text). The script warns if they're missing.
- Style IDs are English (`Heading1`...). Templates created in a non-English Word may use different IDs.

## Scheduling

Run Collect and Build together weekly or nightly (Task Scheduler on a management server, or Azure Automation with
the certificate in the account), then upload the .docx to SharePoint. Keep the `snapshots` folder - it's the history.

## Known limitations

- Uses Graph **beta** endpoints (the only place most Intune config lives). Microsoft changes these occasionally.
- **Autopilot device preparation (v2)** policies are detected by template family `enrollmentConfiguration`. The device
  security group chosen in the policy isn't returned with the settings - check it in the admin center.
- **Legacy intents** (old endpoint security/baseline templates) show setting IDs rather than friendly names.
- Encrypted custom OMA-URI values and secret settings show as `(encrypted)` / `(secret)`.
- Co-management workload settings live in ConfigMgr, not Graph - document those by hand.
