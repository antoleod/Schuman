#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU',
  [ValidateSet('Turbo','Smart','Quick','Full')]
  [string]$PreSyncMode = 'Turbo'
)

$entry = Join-Path $PSScriptRoot 'Invoke-Schuman.ps1'
if (-not (Test-Path -LiteralPath $entry)) {
  throw "Invoke-Schuman.ps1 not found: $entry"
}

# Keep mandatory pre-sync behavior.
$maxTickets = 0
$scope = 'RitmOnly'
if ($PreSyncMode -eq 'Quick') { $maxTickets = 30 }

& $entry -Operation Export -ExcelPath $ExcelPath -SheetName $SheetName -ProcessingScope $scope -MaxTickets $maxTickets -NoPopups
if (($LASTEXITCODE -ne $null) -and ($LASTEXITCODE -ne 0)) {
  throw "Pre-sync failed (exit code $LASTEXITCODE)."
}

& $entry -Operation DashboardSearch -ExcelPath $ExcelPath -SheetName $SheetName
