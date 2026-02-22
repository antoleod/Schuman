# =============================================================================
# Dashboard Launcher (UIX)
# =============================================================================
# Dedicated entry point for Check-in / Check-out dashboard mode.
# It reuses auto-excel.ps1 with -DashboardMode so export flow is not triggered.

param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot "Schuman List.xlsx"),
  [string]$SheetName = "BRU"
)

$target = Join-Path $PSScriptRoot "auto-excel.ps1"
if (-not (Test-Path -LiteralPath $target)) {
  throw "auto-excel.ps1 not found at: $target"
}

& $target -DashboardMode -ExcelPath $ExcelPath -SheetName $SheetName
