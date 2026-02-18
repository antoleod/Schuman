# Mapa De Codigo

## Entry points
- `Start-Schuman.ps1`: arranca la UI principal.
- `Invoke-Schuman.ps1`: CLI/automatizacion por operacion (`Export`, `DashboardSearch`, `DashboardCheckIn`, `DashboardCheckOut`, `DocsGenerate`).

## Modulos
- `src/Schuman.Automation/Core`: bootstrap, logging, paths y utilidades de texto.
- `src/Schuman.Automation/Integrations`: acceso a Excel COM y ServiceNow (WebView2 + JSONv2).
- `src/Schuman.Automation/Workflows`: casos de uso de negocio.
- `src/Schuman.Automation/UI`: UI WinForms (`Main`, `Dashboard`, `Generate`, `Theme`).

## Estado
- Legacy eliminado: no se usan `auto-excel.ps1`, `dashboard-checkin-checkout.ps1`, `Generate-pdf.ps1`.
