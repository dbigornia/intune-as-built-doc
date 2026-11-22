# IntuneAsBuilt

Generates a Word (and optional PDF) as-built document for Microsoft Intune: Windows enrollment, device policies,
and apps & scripts, with your own design notes merged in.

| Stage | Script | Needs | Output |
|---|---|---|---|
| 1. Collect | `Invoke-IntuneCollect.ps1` | Graph, read-only | `<workspace>\snapshots\<yyyy-MM-dd_HHmmZ>\*.json` |
| 2. Build | `Build-IntuneAsBuilt.ps1` | Nothing - works offline (PDF: Word or LibreOffice) | `<workspace>\output\Intune-AsBuilt_<tenant>_<snapshot>.docx` (+ `.pdf`) |

Snapshots are point-in-time baselines and audit evidence. You can rebuild the document from any old snapshot,
and two snapshots can be diffed (`git diff --no-index` or VS Code compare) to see what changed.

## Download

1. Download the latest release zip from the **Releases** page (or clone the repo) and extract it.
2. Windows marks downloaded scripts as coming from the internet. Unblock them once, from the extracted folder:
   ```powershell
   Get-ChildItem -Recurse | Unblock-File
   ```
3. Install the prerequisites in [REQUIREMENTS.md](REQUIREMENTS.md): PowerShell 7+ and
   `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`.

Run the scripts with `pwsh`, not Windows PowerShell 5.1.

## Workspace (keep your tenant data private)

Everything tenant-specific lives in a **workspace** folder that you keep separate from the tool, ideally in a
private git repo:

```
intune-as-built-config\        your private folder or repo
├─ snapshots\                   Collect writes here; Build reads the newest
├─ output\                      generated .docx and .pdf
├─ notes.csv                    your design notes (created on first run)
└─ template.docx (or .dotx)     optional company template, used automatically
```

Pass `-Workspace <path>` to both scripts, or set it once:

```powershell
[Environment]::SetEnvironmentVariable('INTUNEASBUILT_WORKSPACE', 'D:\intune-as-built-config', 'User')
```

Explicit parameters (`-NotesPath`, `-TemplatePath`, `-OutputPath`, `-SnapshotPath`) override the workspace defaults.
Without a workspace, the scripts warn and fall back to the tool folder.

## Usage

```powershell
# 1. Collect - unattended with a certificate
.\Invoke-IntuneCollect.ps1 -Workspace D:\intune-as-built-config -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumb>

#    or interactively (lab, or team members through a delegated app registration)
.\Invoke-IntuneCollect.ps1 -Workspace D:\intune-as-built-config -Interactive -ClientId <delegatedAppId>

# 2. Build - newest snapshot, notes and template picked up from the workspace
.\Build-IntuneAsBuilt.ps1 -Workspace D:\intune-as-built-config -Pdf
```

Word asks to update fields when you open the .docx - choose **Yes** to build the table of contents
(not needed if the PDF was made with Word; see `-Pdf` below).

### Collect options
- `-IncludeScriptContent` - store script bodies in the snapshot and print them in Appendix A. Without it, only a SHA-256 hash is stored (changes are still detectable). Review scripts for embedded secrets before turning this on.
- `-IncludeAutopilotDevices` - adds the device register (counts by group tag and model in the document; serials in the JSON only).
- `-OutputRoot` - write snapshots somewhere other than `<workspace>\snapshots`.

### Build options
- `-SnapshotPath` - a specific snapshot (default: newest in the workspace).
- `-AllPlatforms` - include iOS/Android/macOS objects (default: Windows only).
- `-IncludeEmptyValues` - show unset properties too.
- `-HideMissingNotes` - drop the "Design notes: not yet documented" line.
- `-Pdf` - also save a PDF next to the .docx.
  - **Word** (used automatically on Windows when installed): builds the table of contents, adds heading bookmarks, and re-saves the .docx with the TOC filled in.
  - **LibreOffice** (fallback, or `-PdfEngine LibreOffice`): no Office needed, but the TOC isn't built in the PDF.
  - Microsoft doesn't support automating Word from unattended scheduled tasks. On a server, use LibreOffice and `-PdfEngine LibreOffice`.

## Permissions

All Graph calls are `GET`; nothing is written to the tenant. The app registration needs these read-only
permissions (application permissions for certificate auth, delegated for interactive):

| Permission | Used for |
|---|---|
| DeviceManagementConfiguration.Read.All | Policies, settings catalog, compliance, update profiles |
| DeviceManagementApps.Read.All | Apps, assignments, relationships |
| DeviceManagementServiceConfig.Read.All | Autopilot profiles/devices, ESP, enrollment configs |
| DeviceManagementScripts.Read.All | Platform scripts and remediations (required since July 2025) |
| DeviceManagementRBAC.Read.All | Scope tag names |
| Group.Read.All | Group names in assignments |
| Organization.Read.All | Tenant name |

### Running it without Global Admin
- **Build** needs no permissions at all. Anyone can build documents from snapshot folders you share.
- **Collect, certificate (app-only):** after one-time admin consent, the app's permissions are used, not the
  user's. Whoever has the certificate's private key can collect, so run it from one management server or
  service account rather than handing the certificate out.
- **Collect, interactive (delegated):** create a second app registration (public client, redirect URI
  `http://localhost`) with the same permissions as **delegated**. Grant admin consent once, set
  **Assignment required = Yes** on its enterprise app, and assign the team's group. Users run
  `-Interactive -ClientId <appId>`. What they can read is limited by their own Intune role (Read Only Operator
  or a custom read-only role), and scope tags still apply, so a scoped user gets a partial document.

## Design notes (the "why")

`notes.csv` in your workspace holds the hand-written context for each object, keyed by Intune object ID so renaming
a policy doesn't lose its notes:

`ObjectId, Category, Name, Purpose, Owner, DependsOn, LastValidated, Rollback`

See [notes.example.csv](notes.example.csv) for the format. Every build adds new objects to the file and keeps what's
already written, including notes for objects deleted from the tenant (marked `(not in this snapshot)`). Commit the
workspace's `notes.csv` to your private repo to keep the history of design decisions.

## Company template

A generic template is included at [templates/Company-Template.docx](templates/Company-Template.docx): header, footer
with page numbers, a cover page with document control and revision history, and the `{{ASBUILT}}` marker.
Copy it to your workspace as `template.docx`, replace `COMPANY NAME` and `[COMPANY LOGO]`, and adjust the styles.

Any `.docx` or `.dotx` works. Its headers, footers, logo, page size, margins and styles are kept.
- Put a paragraph containing `{{ASBUILT}}` where the generated content should go. Without it, the template body is replaced.
- The template must define **Title, Heading 1, Heading 2 and Heading 3** styles. The script warns if they're missing.
- Style IDs are English (`Heading1`...). Templates created in a non-English Word may use different IDs.

## Scheduling

Run Collect and Build together weekly or nightly (Task Scheduler on a management server, or Azure Automation with
the certificate), set `INTUNEASBUILT_WORKSPACE` for the account, then upload the output to SharePoint.
Keep the `snapshots` folder - it's the history.

## Known limitations

- Uses Graph **beta** endpoints (the only place most Intune config lives). Microsoft changes these occasionally.
- **Autopilot device preparation (v2)** policies are detected by template family `enrollmentConfiguration`. The device
  security group chosen in the policy isn't returned with the settings - check it in the admin center.
- **Legacy intents** (old endpoint security/baseline templates) show setting IDs rather than friendly names.
- Encrypted custom OMA-URI values and secret settings show as `(encrypted)` / `(secret)`.
- Co-management workload settings live in ConfigMgr, not Graph - document those by hand.

## Security, license and disclaimer

See [SECURITY.md](SECURITY.md) for how the tool handles tenant data and how to report a vulnerability.
Released under the [MIT License](LICENSE). Changes are listed in [CHANGELOG.md](CHANGELOG.md).

This is an independent project and is not affiliated with, endorsed by, or supported by Microsoft.
Microsoft, Intune, Entra, Windows and Word are trademarks of the Microsoft group of companies.
