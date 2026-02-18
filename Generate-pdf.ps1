#Requires -Version 5.1
<##
Compatibility wrapper for legacy Word/PDF generator.
The implementation moved to Invoke-Schuman.ps1 (DocsGenerate command).
##>

param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$TemplatePath = (Join-Path $PSScriptRoot 'Reception_ITequipment.docx'),
  [string]$OutDir = (Join-Path $PSScriptRoot 'WORD files'),
  [string]$PreferredSheet = 'BRU',
  [switch]$ExportPdf
)

$entry = Join-Path $PSScriptRoot 'Invoke-Schuman.ps1'
if (-not (Test-Path -LiteralPath $entry)) {
  throw "Invoke-Schuman.ps1 not found: $entry"
}

& $entry -Operation DocsGenerate -ExcelPath $ExcelPath -SheetName $PreferredSheet -TemplatePath $TemplatePath -OutputDirectory $OutDir -ExportPdf:$ExportPdf
exit $LASTEXITCODE
