# Workflows

## `src/Schuman.Automation/Workflows/TicketExport.ps1`
- Orchestrates full export:
  - reads tickets from Excel,
  - queries ServiceNow,
  - generates JSON,
  - optional write-back to Excel.

## `src/Schuman.Automation/Workflows/Dashboard.ps1`
- `Invoke-DashboardSearchWorkflow`
- `Invoke-DashboardCheckInWorkflow`
- `Invoke-DashboardCheckOutWorkflow`
- Uses `Invoke-DashboardActionCore` for shared logic.

## `src/Schuman.Automation/Workflows/Documents.ps1`
- Generates Word documents from template.
- Optional PDF export.
- Placeholder replacement and per-RITM file output.
