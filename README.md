# Schuman Automation (Refactored)

## Architecture

- `Invoke-Schuman.ps1`: single enterprise entrypoint with subcommands.
- `src/Schuman.Automation/Core`: config, logging, path/bootstrap utilities.
- `src/Schuman.Automation/Integrations`: isolated adapters for Excel COM and ServiceNow WebView2/JSONv2.
- `src/Schuman.Automation/Workflows`: business workflows (`Export`, `Dashboard`, `Documents`).
- `configs/appsettings.json`: runtime configuration overrides.
- Legacy wrappers (`auto-excel.ps1`, `dashboard-checkin-checkout.ps1`, `Generate-pdf.ps1`) keep backwards compatibility.

## Commands

### Main UI Dashboard (single launcher)

```powershell
.\Schuman-Dashboard.ps1
```

### Export tickets to JSON + Excel write-back

```powershell
.\Invoke-Schuman.ps1 -Operation Export -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -ProcessingScope RitmOnly
```

### Dashboard search

```powershell
.\Invoke-Schuman.ps1 -Operation DashboardSearch -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -SearchText "john"
```

### Dashboard check-in

```powershell
.\Invoke-Schuman.ps1 -Operation DashboardCheckIn -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
```

### Dashboard check-out

```powershell
.\Invoke-Schuman.ps1 -Operation DashboardCheckOut -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
```

### Generate Word (and optional PDF)

```powershell
.\Invoke-Schuman.ps1 -Operation DocsGenerate -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -TemplatePath ".\Reception_ITequipment.docx" -OutputDirectory ".\WORD files" -ExportPdf
```

## Security and Operations

- No passwords are stored; login is interactive SSO in WebView2 with a local profile cache.
- All runs are isolated under `system\runs\<command>_<timestamp>`.
- Every run writes a dedicated `run.log.txt` and appends to `system\logs\history.log`.
- ServiceNow calls use retries and in-memory caches for state labels, users, and CI names.
