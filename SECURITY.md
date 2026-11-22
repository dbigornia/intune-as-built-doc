# Security

## What this tool does with your tenant

- **Read-only.** `Invoke-IntuneCollect.ps1` only sends `GET` requests to Microsoft Graph. It never creates,
  changes or deletes anything in Intune or Entra ID.
- **Least privilege.** It needs only the `*.Read.All` permissions listed in [REQUIREMENTS.md](REQUIREMENTS.md).
  Don't grant `ReadWrite` permissions to its app registration.
- **Local output only.** Snapshots and documents are written to your workspace folder. Nothing is sent anywhere else.
- **Script bodies are not stored by default.** Platform script and remediation content is kept only as a SHA-256
  hash unless you use `-IncludeScriptContent`. Review scripts for embedded secrets before turning that on.

## Protecting your data

Snapshots, `notes.csv` and generated documents describe your security configuration in detail
(policy names, group names, object IDs, settings). Keep them in a **private** workspace (see README > Workspace),
never in a fork or clone of this public repository. Keep the app registration's certificate private key on the
collection server only, and mark it non-exportable.

## Reporting a vulnerability

Please don't open a public issue for security problems. Use GitHub's private
**Report a vulnerability** option on this repository's Security tab instead.
