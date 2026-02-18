# Core

## `src/Schuman.Automation/Core/Bootstrap.ps1`
- Carga configuracion por defecto y mezcla con `configs/appsettings.json`.
- Normaliza rutas (relativas y `%ENV%`).

## `src/Schuman.Automation/Core/Logging.ps1`
- `New-RunContext`: crea carpeta por ejecucion en `system/runs`.
- `Write-RunLog`: escribe en log de run + historico.

## `src/Schuman.Automation/Core/Paths.ps1`
- Helpers de paths para output y archivos de trabajo.

## `src/Schuman.Automation/Core/Text.ps1`
- Helpers de texto/tickets/estado.
- Incluye reglas de `Test-ClosedState` y utilidades de normalizacion.
