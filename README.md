# Schuman Automation

Sistema unificado en PowerShell 5.1 para operaciones con ServiceNow, Excel y documentos.

## Inicio
```powershell
.\Start-Schuman.ps1
```

## Operaciones CLI
```powershell
.\Invoke-Schuman.ps1 -Operation Export -ExcelPath ".\Schuman List.xlsx" -SheetName BRU
.\Invoke-Schuman.ps1 -Operation DashboardSearch -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -SearchText "john"
.\Invoke-Schuman.ps1 -Operation DashboardCheckIn -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
.\Invoke-Schuman.ps1 -Operation DashboardCheckOut -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -Row 25
.\Invoke-Schuman.ps1 -Operation DocsGenerate -ExcelPath ".\Schuman List.xlsx" -SheetName BRU -TemplatePath ".\Reception_ITequipment.docx" -OutputDirectory ".\WORD files"
```

## Arquitectura
- `src/Schuman.Automation/Main.ps1`: launcher UI.
- `src/Schuman.Automation/UI`: Dashboard y Generate.
- `src/Schuman.Automation/Workflows`: casos de uso.
- `src/Schuman.Automation/Integrations`: Excel + ServiceNow.
- `src/Schuman.Automation/Core`: bootstrap/logging/helpers.

## Documentacion tecnica
- `docs/GUIA-RAPIDA.md`
- `docs/CODIGO/00-MAPA.md`
- `docs/CODIGO/*.md`

## Nota
- Legacy eliminado. Solo se soporta la estructura `src/Schuman.Automation`.
