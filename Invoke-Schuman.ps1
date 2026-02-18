#Requires -Version 5.1
param(
  [ValidateSet('Export','DashboardSearch','DashboardCheckIn','DashboardCheckOut','DocsGenerate')]
  [string]$Operation = 'Export',

  [string]$ExcelPath,
  [string]$SheetName = 'BRU',

  [string]$TicketHeader = 'Number',
  [int]$TicketColumn = 4,
  [string]$NameHeader = 'Name',
  [string]$PhoneHeader = 'PI',
  [string]$ActionHeader = 'Estado de RITM',
  [string]$SCTasksHeader = 'SCTasks',
  [ValidateSet('Auto','RitmOnly','IncAndRitm','All')]
  [string]$ProcessingScope = 'Auto',
  [int]$MaxTickets = 0,
  [switch]$NoWriteBack,

  [string]$SearchText = '',
  [int]$Row = 0,
  [string]$WorkNote,

  [string]$TemplatePath,
  [string]$OutputDirectory,
  [switch]$ExportPdf,

  [switch]$NoPopups
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = $PSScriptRoot
. (Join-Path $projectRoot 'src\Schuman.Automation\Import-SchumanModules.ps1')

$config = Initialize-SchumanEnvironment -ProjectRoot $projectRoot
$runContext = New-RunContext -Config $config -RunName $Operation.ToLowerInvariant()

try {
  $resolvedExcel = Resolve-ExcelPath -Config $config -ExcelPath $ExcelPath -StartDirectory $projectRoot

  switch ($Operation) {
    'Export' {
      $res = Invoke-TicketExportWorkflow -Config $config -RunContext $runContext -ExcelPath $resolvedExcel -SheetName $SheetName `
        -TicketHeader $TicketHeader -TicketColumn $TicketColumn -NameHeader $NameHeader -PhoneHeader $PhoneHeader `
        -ActionHeader $ActionHeader -SCTasksHeader $SCTasksHeader -ProcessingScope $ProcessingScope -MaxTickets $MaxTickets -NoWriteBack:$NoWriteBack

      Write-RunLog -RunContext $runContext -Level INFO -Message "Export completed. JSON: $($res.CombinedJsonPath)"
      if (-not $NoPopups) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("Export complete.`r`nJSON: $($res.CombinedJsonPath)", 'Schuman Automation') | Out-Null
      }
    }

    'DashboardSearch' {
      $rows = Invoke-DashboardSearchWorkflow -RunContext $runContext -ExcelPath $resolvedExcel -SheetName $SheetName -SearchText $SearchText
      $rows | Format-Table Row, RequestedFor, RITM, SCTASK, DashboardStatus, PresentTime, ClosedTime -AutoSize
    }

    'DashboardCheckIn' {
      if ($Row -le 0) { throw 'Provide -Row for DashboardCheckIn.' }
      $msg = if ($WorkNote) { $WorkNote } else { 'Deliver all credentials to the new user' }
      $out = Invoke-DashboardCheckInWorkflow -Config $config -RunContext $runContext -ExcelPath $resolvedExcel -SheetName $SheetName -Row $Row -WorkNote $msg
      Write-Host ($out | ConvertTo-Json -Depth 5)
    }

    'DashboardCheckOut' {
      if ($Row -le 0) { throw 'Provide -Row for DashboardCheckOut.' }
      $msg = if ($WorkNote) { $WorkNote } else { "Laptop has been delivered.`r`nFirst login made with the user.`r`nOutlook, Teams and Jabber successfully tested." }
      $out = Invoke-DashboardCheckOutWorkflow -Config $config -RunContext $runContext -ExcelPath $resolvedExcel -SheetName $SheetName -Row $Row -WorkNote $msg
      Write-Host ($out | ConvertTo-Json -Depth 5)
    }

    'DocsGenerate' {
      $files = Invoke-DocumentGenerationWorkflow -Config $config -RunContext $runContext -ExcelPath $resolvedExcel -SheetName $SheetName `
        -TemplatePath $TemplatePath -OutputDirectory $OutputDirectory -ExportPdf:$ExportPdf
      $files | Format-Table Row, RITM, DocxPath, PdfPath -AutoSize
    }
  }
}
catch {
  Write-RunLog -RunContext $runContext -Level ERROR -Message $_.Exception.Message
  if (-not $NoPopups) {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Schuman Automation Error') | Out-Null
  }
  throw
}
