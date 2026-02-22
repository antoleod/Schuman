# Core

## `src/Schuman.Automation/Core/Bootstrap.ps1`
- Loads default configuration and merges with `configs/appsettings.json`.
- Normalizes paths (relative and `%ENV%`).

## `src/Schuman.Automation/Core/Logging.ps1`
- `New-RunContext`: creates a per-run folder in `system/runs`.
- `Write-RunLog`: writes to run log + history log.

## `src/Schuman.Automation/Core/Paths.ps1`
- Path helpers for output and working files.

## `src/Schuman.Automation/Core/Text.ps1`
- Text/ticket/state helper functions.
- Includes `Test-ClosedState` rules and normalization utilities.
