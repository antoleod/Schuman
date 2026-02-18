#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$main = Join-Path $PSScriptRoot 'src\Schuman.Automation\Main.ps1'
if (-not (Test-Path -LiteralPath $main)) {
  throw "Main UI not found: $main"
}

& $main -ExcelPath $ExcelPath -SheetName $SheetName
