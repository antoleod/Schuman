Set-StrictMode -Version Latest

function Invoke-TicketExportWorkflow {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [string]$SheetName,
    [string]$TicketHeader,
    [int]$TicketColumn,
    [string]$NameHeader,
    [string]$PhoneHeader,
    [string]$ActionHeader,
    [string]$SCTasksHeader,
    [ValidateSet('Auto','RitmOnly','IncAndRitm','All')][string]$ProcessingScope = 'Auto',
    [int]$MaxTickets = 0,
    [switch]$NoWriteBack
  )

  if (-not $SheetName) { $SheetName = $Config.Excel.DefaultSheet }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Reading tickets from Excel: $ExcelPath"
  $tickets = Read-TicketsFromExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
    -StopAfterEmptyRows $Config.Excel.StopScanAfterEmptyRows -MaxRowsAfterFirstTicket $Config.Excel.MaxRowsAfterFirstTicket

  if ($tickets.Count -eq 0) {
    Write-RunLog -RunContext $RunContext -Level WARN -Message 'No tickets found in Excel.'
    return [pscustomobject]@{
      Results = @()
      CombinedJsonPath = ''
    }
  }

  $filtered = Apply-ProcessingScope -Tickets $tickets -Scope $ProcessingScope
  if ($MaxTickets -gt 0 -and $filtered.Count -gt $MaxTickets) {
    $filtered = @($filtered | Select-Object -First $MaxTickets)
  }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Tickets queued: total=$($tickets.Count), filtered=$($filtered.Count), scope=$ProcessingScope"

  $session = $null
  $results = New-Object System.Collections.Generic.List[object]
  try {
    $session = New-ServiceNowSession -Config $Config -RunContext $RunContext

    $i = 0
    foreach ($ticket in $filtered) {
      $i++
      $swTicket = Start-PerfStopwatch
      Write-RunLog -RunContext $RunContext -Level INFO -Message ("[{0}/{1}] Extracting {2}" -f $i, $filtered.Count, $ticket)

      $res = Get-ServiceNowTicket -Session $session -Ticket $ticket
      if ($res.ok -eq $false) {
        Write-RunLog -RunContext $RunContext -Level WARN -Message ("{0} extraction failed: {1}" -f $ticket, $res.reason)
      } else {
        Write-RunLog -RunContext $RunContext -Level INFO -Message ("{0} status='{1}' open_tasks={2}" -f $ticket, $res.status, $res.open_tasks)
      }

      [void]$results.Add($res)
      [void](Stop-PerfStopwatch -Stopwatch $swTicket -RunContext $RunContext -Label ("ticket_{0}" -f $ticket))
    }
  }
  finally {
    Close-ServiceNowSession -Session $session
  }

  $combinedJsonPath = Join-Path $RunContext.RunDir 'tickets_export.json'
  ($results | ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $combinedJsonPath -Encoding UTF8
  Write-RunLog -RunContext $RunContext -Level INFO -Message "Combined JSON generated: $combinedJsonPath"

  if (-not $NoWriteBack) {
    $resultMap = @{}
    foreach ($r in $results) {
      if ($r.ticket) { $resultMap[$r.ticket] = $r }
    }

    Write-RunLog -RunContext $RunContext -Level INFO -Message 'Writing extracted data back to Excel.'
    Write-TicketResultsToExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
      -ResultByTicket $resultMap -NameHeader $NameHeader -PhoneHeader $PhoneHeader -ActionHeader $ActionHeader -SCTasksHeader $SCTasksHeader
  }

  return [pscustomobject]@{
    Results = @($results)
    CombinedJsonPath = $combinedJsonPath
  }
}

function Apply-ProcessingScope {
  param(
    [Parameter(Mandatory = $true)][string[]]$Tickets,
    [Parameter(Mandatory = $true)][ValidateSet('Auto','RitmOnly','IncAndRitm','All')][string]$Scope
  )

  switch ($Scope) {
    'RitmOnly' { return @($Tickets | Where-Object { $_ -like 'RITM*' }) }
    'IncAndRitm' { return @($Tickets | Where-Object { $_ -like 'RITM*' -or $_ -like 'INC*' }) }
    'All' { return @($Tickets) }
    default {
      $ritm = @($Tickets | Where-Object { $_ -like 'RITM*' })
      if ($ritm.Count -gt 0) { return $ritm }
      return @($Tickets)
    }
  }
}
