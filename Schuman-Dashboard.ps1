#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

try {
  [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
} catch {}

$script:DashboardThreadExceptionHandler = [System.Threading.ThreadExceptionEventHandler]{
  param($sender, $eventArgs)
  $msg = 'Unexpected UI error.'
  try {
    if ($eventArgs -and $eventArgs.Exception) { $msg = "" + $eventArgs.Exception.Message }
  } catch {}
  try { [System.Windows.Forms.MessageBox]::Show("Dashboard UI error.`r`n`r`n$msg", 'Schuman Dashboard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
}
[System.Windows.Forms.Application]::add_ThreadException($script:DashboardThreadExceptionHandler)

$script:DashboardDomainExceptionHandler = [System.UnhandledExceptionEventHandler]{
  param($sender, $eventArgs)
  $msg = 'Unexpected fatal error.'
  try {
    $ex = $eventArgs.ExceptionObject -as [System.Exception]
    if ($ex) { $msg = "" + $ex.Message }
  } catch {}
  try { [System.Windows.Forms.MessageBox]::Show("Unhandled error.`r`n`r`n$msg", 'Schuman Dashboard', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
}
[AppDomain]::CurrentDomain.add_UnhandledException($script:DashboardDomainExceptionHandler)

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
