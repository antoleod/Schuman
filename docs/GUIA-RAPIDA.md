# Quick Guide (Schuman)

## 1) Start UI
```powershell
.\Start-Schuman.ps1
```
- ServiceNow SSO login/verification is required at startup.
- The session/token is reused across the app to avoid repeated login prompts.

## 2) Main Window
- `Dashboard`: opens the operational dashboard.
- `Generate`: opens document generation.
- `Close Code Apps`: closes `Code` / `Code Insiders` / `Cursor` if frozen.
- `Close Documents`: closes `Word` and `Excel` to release locked files.

## 3) Dashboard
- Live filtering while typing.
- Row click opens ServiceNow.
- Double-click runs task selection + Check-In/Check-Out.
- Updates both ServiceNow and Excel.
- Emergency buttons: `Close Code Apps` and `Close Documents`.

## 4) Generate
- Footer buttons:
  - `Generate Documents`
  - `Dashboard (Check-in / Check-out)`
  - `Open Output Folder`
  - `Close Code Apps`
  - `Close Documents`
  - `Show Log`
- `Generate Documents` executes real `DocsGenerate` workflow.

## 5) CLI
```powershell
.\Invoke-Schuman.ps1 -Operation Export -ExcelPath ".\Schuman List.xlsx" -SheetName BRU
```

## 6) Logs
- Runs: `system/runs/<operation_timestamp>/`
- History: `system/logs/history.log`

## 7) Component Docs
See `docs/CODIGO/`.
