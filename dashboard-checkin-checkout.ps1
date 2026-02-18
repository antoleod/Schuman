# =============================================================================
# Dashboard Launcher (UIX)
# =============================================================================
# Dedicated entry point for Check-in / Check-out dashboard mode.
# It reuses auto-excel.ps1 with -DashboardMode so export flow is not triggered.

param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot "Schuman List.xlsx"),
  [string]$SheetName = "BRU",
  [ValidateSet("Turbo","Smart","Quick","Full")]
  [string]$PreSyncMode = "Turbo"
)

$target = Join-Path $PSScriptRoot "auto-excel.ps1"
if (-not (Test-Path -LiteralPath $target)) {
  throw "auto-excel.ps1 not found at: $target"
}

Write-Host "Running mandatory pre-sync before dashboard (mode=$PreSyncMode)..."
$preArgs = @("-NoPopups", "-ExcelPath", $ExcelPath, "-SheetName", $SheetName, "-ProcessingScope", "RitmOnly")
switch ($PreSyncMode) {
  "Quick" { $preArgs += "-QuickMode" }
  "Smart" { $preArgs += "-SmartMode" }
  "Turbo" { $preArgs += "-TurboMode" }
  default { }
}
& $target @preArgs
if (($LASTEXITCODE -ne $null) -and ($LASTEXITCODE -ne 0)) {
  throw "Mandatory pre-sync failed (exit code $LASTEXITCODE). Dashboard was not started."
}

Write-Host "Pre-sync OK. Launching dashboard..."
& $target -DashboardMode -ExcelPath $ExcelPath -SheetName $SheetName
