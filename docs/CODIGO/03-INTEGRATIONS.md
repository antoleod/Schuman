# Integrations

## `src/Schuman.Automation/Integrations/Excel.ps1`
- Lectura de tickets desde Excel (`Read-TicketsFromExcel`).
- Escritura de resultados (`Write-TicketResultsToExcel`).
- Busqueda dashboard (`Search-DashboardRows`).
- Update de fila dashboard (`Update-DashboardRow`).
- Todo con COM cleanup estricto (release + GC).

## `src/Schuman.Automation/Integrations/WebView2ServiceNow.ps1`
- Inicializa WebView2 desde Teams Add-in (sin instalar nada).
- Gestion de sesion SSO (`New-ServiceNowSession`, `Close-ServiceNowSession`).
- Consultas JSONv2 (`Invoke-ServiceNowJsonv2Query`).
- Lectura de tickets/tareas y resolucion de labels/cache.
- Cambio de estado SCTASK en UI (`Set-ServiceNowTaskState`).
