#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$moduleRoot = Join-Path $PSScriptRoot 'src\Schuman.Automation'
$importModulesPath = Join-Path $moduleRoot 'Import-SchumanModules.ps1'
$themePath = Join-Path $moduleRoot 'UI\Theme.ps1'
$dashboardUiPath = Join-Path $moduleRoot 'UI\DashboardUI.ps1'

foreach ($p in @($importModulesPath, $themePath, $dashboardUiPath)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required file not found: $p" }
}

. $importModulesPath
. $themePath
. $dashboardUiPath

$config = Initialize-SchumanEnvironment -ProjectRoot $PSScriptRoot
$runContext = New-RunContext -Config $config -RunName 'dashboardui'

$form = New-DashboardUI -ExcelPath $ExcelPath -SheetName $SheetName -Config $config -RunContext $runContext
if (-not ($form -is [System.Windows.Forms.Form])) {
  throw 'New-DashboardUI did not return a WinForms Form.'
}

[void]$form.ShowDialog()
