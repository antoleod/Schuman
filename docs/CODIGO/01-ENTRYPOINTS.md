# Entry Points

## `Start-Schuman.ps1`
- Objetivo: lanzar la aplicacion UI.
- Flujo: valida `src/Schuman.Automation/Main.ps1` y lo ejecuta.

## `Invoke-Schuman.ps1`
- Objetivo: ejecutar operaciones en modo comando.
- Parametros principales: `-Operation`, `-ExcelPath`, `-SheetName`.
- Operaciones:
  - `Export`
  - `DashboardSearch`
  - `DashboardCheckIn`
  - `DashboardCheckOut`
  - `DocsGenerate`
- Manejo de errores: logging + popup opcional (`-NoPopups`).
