# Schuman Automation

Unified PowerShell 5.1 system for ServiceNow, Excel, and document workflows.

## Start UI
```powershell
.\Start-Schuman.ps1
```

## CLI Operations
```powershell
.\Invoke-Schuman.ps1 -Operation Export -ExcelPath ".\Schuman List.xlsx" -SheetName BRU
.\Invoke-Schuman.ps1 -Operation DashboardSearch -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -SearchText "john"
.\Invoke-Schuman.ps1 -Operation DashboardCheckIn -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
.\Invoke-Schuman.ps1 -Operation DashboardCheckOut -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
.\Invoke-Schuman.ps1 -Operation DocsGenerate -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -TemplatePath ".\Reception_ITequipment.docx" -OutputDirectory ".\WORD files"
```

## Architecture
- `src/Schuman.Automation/Main.ps1`: main UI launcher.
- `src/Schuman.Automation/UI`: Dashboard and Generate WinForms UI.
- `src/Schuman.Automation/Workflows`: workflow use cases.
- `src/Schuman.Automation/Integrations`: Excel and ServiceNow integrations.
- `src/Schuman.Automation/Core`: bootstrap, logging, and utility helpers.

## Technical Documentation
- `docs/GUIA-RAPIDA.md`
- `docs/CODIGO/00-MAPA.md`
- `docs/CODIGO/*.md`

## Note
- Legacy scripts are removed. Only `src/Schuman.Automation` is supported.
