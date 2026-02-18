#Requires -Version 5.1
$dashboard = Join-Path $PSScriptRoot 'Schuman-Dashboard.ps1'
if (-not (Test-Path -LiteralPath $dashboard)) {
  throw "Schuman-Dashboard.ps1 not found: $dashboard"
}
& $dashboard
