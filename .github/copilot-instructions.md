## Purpose

This repository is a small collection of Windows PowerShell scripts for EUC (end-user computing) tasks — profile scanning, RSOP/GPO modeling and remote application execution. These instructions give an AI coding agent the immediate, actionable context needed to make safe, correct edits or add features.

## Big picture
- The repo contains stand-alone PowerShell scripts (root) that are run directly on Windows hosts. Major scripts:
  - `RSOP Model.ps1` — generates a full HTML GPO configuration report (requires GroupPolicy module & GPMC COM / RSAT). Outputs an HTML file to `%TEMP%` and opens it.
  - `UPM Profile Scanner.ps1` — reads a plain-text list of profile folders and writes a CSV summary to `C:\output.csv` (configurable at top of script).
  - `Remotely Execute Application.ps1` (currently located under `.git/` in this repo) — loops servers from `C:\Servername.txt`, uses `Invoke-Command` to start processes remotely and emails results. NOTE: file placement is unusual; consider moving it to repo root.

## Key patterns & conventions
- Config-by-top-variables: scripts define environment-specific values (paths, domain, OU, SMTP server) as top-level variables — change these instead of editing logic where possible.
- Hardcoded Windows paths and list files: scripts expect files such as `C:\Profile\output_file.txt`, `C:\Servername.txt` and `C:\Credential.xt`. If modifying or running a script, either create these files or parameterize the script.
- Output targets: CSV/HTML outputs are written to `C:\` or `$env:TEMP`. Keep this behavior unless you intentionally refactor for portability.
- Error handling: scripts use try/catch with Write-Host/Write-Warning/Write-Error and may `exit 1` on fatal failures (see `RSOP Model.ps1`). Maintain the existing exit semantics when changing control flow.

## Integration points & external requirements
- `RSOP Model.ps1` requires RSAT/GPMC installed and a domain-joined host with rights to query GPOs. It uses the `GroupPolicy` module, GPMC COM (`GPMgmt.GPM`) and `Get-GPOReport`.
- `Remotely Execute Application.ps1` requires remote PowerShell remoting enabled on target servers and an SMTP server for `Send-MailMessage` (SMTP is configured inside the script).
- `UPM Profile Scanner.ps1` expects a profile store (example: `E:\ProfileStore`) and a filename list input. It inspects `UPM_Profile\NTUSER.dat` for LastWriteTime.

## How to run (examples)
- Run with the system PowerShell `pwsh.exe` (Windows) and allow script execution when needed:

  pwsh -NoProfile -ExecutionPolicy Bypass -File "c:\Temp\PersoGit\UPM Profile Scanner.ps1"

  pwsh -NoProfile -ExecutionPolicy Bypass -File "c:\Temp\PersoGit\RSOP Model.ps1"

Notes: run `RSOP Model.ps1` on a domain-joined host that has RSAT/GPMC installed and with an account that can read GPOs.

## Safe-edit guidance for AI agents
- Prefer small, localizable changes: parameterize top variables by converting them into a `param()` block. Example: replace hardcoded `$outputCsvPath` with a `[string]$OutputCsvPath = 'C:\output.csv'` param and update callers.
- Do not remove try/catch blocks; keep existing logging semantics (Write-Host/Write-Warning) unless you add a consistent logging layer.
- When adding new files, avoid committing secrets. The repo currently reads credentials from `C:\Credential.xt` — do not create or commit plaintext credentials.
- If relocating `Remotely Execute Application.ps1` out of `.git/`, update README and keep commit history clear (move rather than copy+delete where possible).

## Files to inspect when making changes
- `RSOP Model.ps1` — GPO retrieval, `Get-GPOReport` embedding and HTML generation.
- `UPM Profile Scanner.ps1` — file-list ingestion and CSV export (useful example for bulk profile checks).
- `.git/Remotely Execute Application.ps1` — remote execution + email; check credential handling and server list source.
- `README.md` — short project description; update if you add new scripts or change run instructions.

## Final notes
- The codebase is small and procedural; most value comes from making scripts parameter-driven and documenting required environment (RSAT, domain membership, file inputs). If you need to change runtime behavior, add a `-WhatIf` or `-Verbose` switch, and include a short usage header in the script top comments.

If any section above is unclear or you'd like me to expand specific examples (e.g., param blocks for each script, or a safe refactor to move the `.git/` script), tell me which script to focus on and I'll update the instructions and implement the changes.
