#Requires -Version 5.1
$main = Join-Path $PSScriptRoot 'src\Schuman.Automation\Main.ps1'
if (-not (Test-Path -LiteralPath $main)) {
  throw "Main UI not found: $main"
}
& $main
