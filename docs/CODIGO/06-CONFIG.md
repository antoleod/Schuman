# Configuration

## `configs/appsettings.json`
- `ServiceNow`:
  - `BaseUrl`, `LoginUrl`, `WebViewProfileRoot`, timeouts, retries.
- `Excel`:
  - default file/sheet and scan limits.
- `Output`:
  - root `system`, subfolders `runs/logs/db`.
- `Documents`:
  - default template and output folder.

## Recommendations
- Do not hardcode paths in scripts; use config.
- Keep `Schuman List.xlsx` as the single source of truth.
- Keep `Reception_ITequipment.docx` in root or update config accordingly.
