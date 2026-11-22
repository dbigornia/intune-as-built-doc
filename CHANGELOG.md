# Changelog

All notable changes to this project are listed here. Versions follow [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-11-22

First public release.

### Added
- `Invoke-IntuneCollect.ps1`: read-only Graph collection of Windows enrollment (Autopilot v1, device preparation v2,
  ESP, enrollment configs), policies (settings catalog, endpoint security, baselines, templates, ADMX, compliance,
  Windows Update) and apps & scripts into dated JSON snapshots. Certificate or interactive auth.
- `Build-IntuneAsBuilt.ps1`: offline Word document generation with no module dependencies; optional PDF via Word or LibreOffice.
- Design notes merged from `notes.csv` (keyed by object ID) and refreshed on every build.
- Company Word template support with an `{{ASBUILT}}` insertion marker; generic template in `templates/`.
- `-Workspace` / `INTUNEASBUILT_WORKSPACE`: keeps all tenant data in a separate private folder.
- Script content stored as SHA-256 hashes by default; full content only with `-IncludeScriptContent`.
