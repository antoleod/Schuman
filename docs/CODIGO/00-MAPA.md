# Code Map

## Entry Points
- `Start-Schuman.ps1`: starts the main UI.
- `Invoke-Schuman.ps1`: CLI automation by operation (`Export`, `DashboardSearch`, `DashboardCheckIn`, `DashboardCheckOut`, `DocsGenerate`).

## Modules
- `src/Schuman.Automation/Core`: bootstrap, logging, paths, and text utilities.
- `src/Schuman.Automation/Integrations`: Excel COM and ServiceNow (WebView2 + JSONv2).
- `src/Schuman.Automation/Workflows`: business workflow use cases.
- `src/Schuman.Automation/UI`: WinForms UI (`Main`, `Dashboard`, `Generate`, `Theme`).

## Status
- Legacy scripts removed: `auto-excel.ps1`, `dashboard-checkin-checkout.ps1`, `Generate-pdf.ps1` are no longer used.
