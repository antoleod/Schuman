# Guia Rapida Schuman

## 1) Inicio UI
```powershell
.\Start-Schuman.ps1
```
- El login/verificacion SSO de ServiceNow es obligatorio al iniciar.
- La sesion/token se reutiliza durante la app para evitar pedir login repetido.

## 2) Main
- `Dashboard`: abre el dashboard operativo.
- `Generate`: abre modulo de documentos.
- `Cerrar codigo`: cierra procesos `Code`/`Code Insiders`/`Cursor` si se congelan.
- `Cerrar documentos`: cierra procesos `Word` y `Excel` para liberar archivos bloqueados.

## 3) Dashboard
- Filtro live al escribir.
- Click fila: abre ServiceNow.
- Doble click: seleccionar tarea + Check-In/Check-Out.
- Actualiza ServiceNow y Excel.
- Botones de emergencia: `Cerrar codigo` y `Cerrar documentos`.

## 4) Generate
- Footer limpio:
  - `Generate Documents`
  - `Dashboard (Check-in / Check-out)`
  - `Open Output Folder`
  - `Cerrar codigo`
  - `Cerrar documentos`
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
