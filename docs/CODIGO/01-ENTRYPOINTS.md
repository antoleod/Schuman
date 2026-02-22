# Entry Points

## `Start-Schuman.ps1`
- Purpose: launch the UI application.
- Flow: validates `src/Schuman.Automation/Main.ps1` and executes it.

## `Invoke-Schuman.ps1`
- Purpose: execute operations in command mode.
- Main parameters: `-Operation`, `-ExcelPath`, `-SheetName`.
- Supported operations:
  - `Export`
  - `DashboardSearch`
  - `DashboardCheckIn`
  - `DashboardCheckOut`
  - `DocsGenerate`
- Error handling: logging + optional popup (`-NoPopups`).
