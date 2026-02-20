Set-StrictMode -Version Latest

$moduleRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$script:SchumanModuleFiles = @(
  'Core\Bootstrap.ps1',
  'Core\Logging.ps1',
  'Core\Paths.ps1',
  'Core\Text.ps1',
  'Integrations\Excel.ps1',
  'Integrations\WebView2ServiceNow.ps1',
  'Workflows\TicketExport.ps1',
  'Workflows\Dashboard.ps1',
  'Workflows\Documents.ps1'
)

$script:SchumanRequiredCommands = @(
  'Initialize-SchumanEnvironment',
  'New-RunContext',
  'Write-RunLog',
  'Resolve-ExcelPath',
  'Test-ClosedState',
  'Search-DashboardRows',
  'Update-DashboardRow',
  'Get-ServiceNowTasksForRitm',
  'Set-ServiceNowTaskState',
  'Close-ServiceNowSession'
)

if (-not (Get-Variable -Name SchumanRuntimeInitialized -Scope Script -ErrorAction SilentlyContinue)) {
  [bool]$script:SchumanRuntimeInitialized = $false
}

function global:Assert-SchumanRuntime {
  param(
    [string[]]$RequiredCommands = $script:SchumanRequiredCommands
  )

  $missing = New-Object System.Collections.Generic.List[string]
  foreach ($name in @($RequiredCommands)) {
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    $cmd = Get-Command -Name $name -ErrorAction SilentlyContinue
    if (-not $cmd) { $missing.Add($name) | Out-Null }
  }

  if ($missing.Count -gt 0) {
    $list = ($missing -join ', ')
    throw "Schuman runtime is incomplete. Missing command(s): $list"
  }
}

function global:Initialize-SchumanRuntime {
  param(
    [switch]$ForceReload,
    [switch]$RevalidateOnly
  )

  if ($ForceReload) {
    throw 'ForceReload is not supported from function scope. Re-dot-source Import-SchumanModules.ps1 to reload runtime in caller scope.'
  }

  Assert-SchumanRuntime
  if (-not $RevalidateOnly) {
    $script:SchumanRuntimeInitialized = $true
  }
}

# Bootstrap runtime in caller scope when this file is dot-sourced.
foreach ($rel in $script:SchumanModuleFiles) {
  $path = Join-Path $moduleRoot $rel
  if (-not (Test-Path -LiteralPath $path)) {
    throw "Required module file not found: $path"
  }
  . $path
}
Assert-SchumanRuntime
$script:SchumanRuntimeInitialized = $true
