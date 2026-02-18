# Guia Rapida Schuman

## 1) Inicio UI
```powershell
.\Start-Schuman.ps1
```

## 2) Main
- `Dashboard`: abre el dashboard operativo.
- `Generate`: abre modulo de documentos.

## 3) Dashboard
- Filtro live al escribir.
- Click fila: abre ServiceNow.
- Doble click: seleccionar tarea + Check-In/Check-Out.
- Actualiza ServiceNow y Excel.

## 4) Generate
- Footer limpio:
  - `Generate Documents`
  - `Dashboard (Check-in / Check-out)`
  - `Open Output Folder`
  - `Show Log`
- `Generate Documents` ejecuta `DocsGenerate` real.

## 5) CLI
```powershell
.\Invoke-Schuman.ps1 -Operation Export -ExcelPath ".\Schuman List.xlsx" -SheetName BRU
```

## 6) Logs
- Runs: `system/runs/<operation_timestamp>/`
- Historico: `system/logs/history.log`

## 7) Documentacion por componente
Ver `docs/CODIGO/`.
