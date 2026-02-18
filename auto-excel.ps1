#Requires -Version 5.1
<##
Compatibility wrapper for legacy entrypoint.
The implementation moved to Invoke-Schuman.ps1.
##>

param(
  [string]$ExcelPath,
  [string]$SheetName = 'BRU',
  [string]$TicketHeader = 'Number',
  [int]$TicketColumn = 4,
  [string]$NameHeader = 'Name',
  [string]$PhoneHeader = 'PI',
  [string]$ActionHeader = 'Estado de RITM',
  [string]$SCTasksHeader = 'SCTasks',
  [switch]$DashboardMode,
  [switch]$NoPopups,
  [ValidateSet('Auto','RitmOnly','IncAndRitm','All')]
  [string]$ProcessingScope = 'Auto',
  [int]$MaxTickets = 0,
  [switch]$QuickMode,
  [switch]$SmartMode,
  [switch]$TurboMode,
  [switch]$NoWriteBack
)

$entry = Join-Path $PSScriptRoot 'Invoke-Schuman.ps1'
if (-not (Test-Path -LiteralPath $entry)) {
  throw "Invoke-Schuman.ps1 not found: $entry"
}

if ($DashboardMode) {
  & $entry -Operation DashboardSearch -ExcelPath $ExcelPath -SheetName $SheetName -NoPopups:$NoPopups
  exit $LASTEXITCODE
}

# Legacy mode switches are accepted for backwards compatibility.
if ($QuickMode -and $MaxTickets -le 0) { $MaxTickets = 30 }
if ($QuickMode -or $TurboMode) { $ProcessingScope = 'RitmOnly' }

& $entry -Operation Export -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
  -NameHeader $NameHeader -PhoneHeader $PhoneHeader -ActionHeader $ActionHeader -SCTasksHeader $SCTasksHeader `
  -ProcessingScope $ProcessingScope -MaxTickets $MaxTickets -NoWriteBack:$NoWriteBack -NoPopups:$NoPopups

exit $LASTEXITCODE
