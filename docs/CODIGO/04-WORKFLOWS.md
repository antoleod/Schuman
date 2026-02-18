# Workflows

## `src/Schuman.Automation/Workflows/TicketExport.ps1`
- Orquesta export completo:
  - lee tickets de Excel,
  - consulta ServiceNow,
  - genera JSON,
  - write-back opcional a Excel.

## `src/Schuman.Automation/Workflows/Dashboard.ps1`
- `Invoke-DashboardSearchWorkflow`
- `Invoke-DashboardCheckInWorkflow`
- `Invoke-DashboardCheckOutWorkflow`
- Usa `Invoke-DashboardActionCore` para logica comun.

## `src/Schuman.Automation/Workflows/Documents.ps1`
- Genera documentos Word desde plantilla.
- Export PDF opcional.
- Sustitucion de placeholders y guardado por fila RITM.
