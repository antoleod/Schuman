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
    [ValidateSet('Auto','RitmOnly','IncOnly','IncAndRitm','All')][string]$ProcessingScope = 'Auto',
    [ValidateSet('Auto','ConfigurationItemOnly','CommentsOnly','CommentsAndCI')][string]$PiSearchMode = 'Auto',
    [int]$MaxTickets = 0,
    [switch]$NoWriteBack,
    [switch]$SkipLegalNameFallback
  )

  if (-not $SheetName) { $SheetName = $Config.Excel.DefaultSheet }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Reading tickets from Excel: $ExcelPath"
  $tickets = Read-TicketsFromExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
    -StopAfterEmptyRows $Config.Excel.StopScanAfterEmptyRows -MaxRowsAfterFirstTicket $Config.Excel.MaxRowsAfterFirstTicket

  $ticketList = @($tickets)
  if ($ticketList.Count -eq 0) {
    Write-RunLog -RunContext $RunContext -Level WARN -Message 'No tickets found in Excel.'
    return [pscustomobject]@{
      Results = @()
      CombinedJsonPath = ''
    }
  }

  $filtered = @(Apply-ProcessingScope -Tickets $ticketList -Scope $ProcessingScope)
  if ($MaxTickets -gt 0 -and $filtered.Count -gt $MaxTickets) {
    $filtered = @($filtered | Select-Object -First $MaxTickets)
  }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Tickets queued: total=$($ticketList.Count), filtered=$($filtered.Count), scope=$ProcessingScope"

  $session = $null
  $results = New-Object System.Collections.Generic.List[object]
  try {
    $session = New-ServiceNowSession -Config $Config -RunContext $RunContext

    $i = 0
    foreach ($ticket in $filtered) {
      $i++
      $swTicket = Start-PerfStopwatch
      Write-RunLog -RunContext $RunContext -Level INFO -Message ("[{0}/{1}] Extracting {2}" -f $i, $filtered.Count, $ticket)

      $res = Get-ServiceNowTicket -Session $session -Ticket $ticket -PiSearchMode $PiSearchMode -SkipLegalNameFallback:$SkipLegalNameFallback

      # Legacy parity: for RITM in Force Update, ensure comments/journal path is also evaluated.
      # If the selected mode returned only CI evidence (or no PI), retry with CommentsAndCI and prefer it when available.
      if ($ticket -like 'RITM*') {
        $currentPi = ''
        $currentSource = ''
        try { if ($res.PSObject.Properties['detected_pi_machine']) { $currentPi = ("" + $res.detected_pi_machine).Trim() } } catch {}
        try { if ($res.PSObject.Properties['pi_source']) { $currentSource = ("" + $res.pi_source).Trim().ToLowerInvariant() } } catch {}

        $needsCommentsRetry = $false
        if ($PiSearchMode -eq 'ConfigurationItemOnly') { $needsCommentsRetry = $true }
        elseif ([string]::IsNullOrWhiteSpace($currentPi)) { $needsCommentsRetry = $true }
        elseif ($currentSource -eq 'ci') { $needsCommentsRetry = $true }

        if ($needsCommentsRetry) {
          try {
            $resComments = Get-ServiceNowTicket -Session $session -Ticket $ticket -PiSearchMode 'CommentsAndCI' -SkipLegalNameFallback:$SkipLegalNameFallback
            if ($resComments -and $resComments.ok -eq $true) {
              $retryPi = ''
              $retrySource = ''
              try { if ($resComments.PSObject.Properties['detected_pi_machine']) { $retryPi = ("" + $resComments.detected_pi_machine).Trim() } } catch {}
              try { if ($resComments.PSObject.Properties['pi_source']) { $retrySource = ("" + $resComments.pi_source).Trim().ToLowerInvariant() } } catch {}

              if (-not [string]::IsNullOrWhiteSpace($retryPi) -and $retrySource -ne 'ci') {
                $res = $resComments
                Write-RunLog -RunContext $RunContext -Level INFO -Message ("{0} PI upgraded from comments/journal source='{1}'." -f $ticket, $retrySource)
              }
              elseif ([string]::IsNullOrWhiteSpace($currentPi) -and -not [string]::IsNullOrWhiteSpace($retryPi)) {
                $res = $resComments
                Write-RunLog -RunContext $RunContext -Level INFO -Message ("{0} PI recovered via comments/journal source='{1}'." -f $ticket, $retrySource)
              }
            }
          }
          catch {
            Write-RunLog -RunContext $RunContext -Level WARN -Message ("{0} comments/journal retry failed: {1}" -f $ticket, $_.Exception.Message)
          }
        }
      }

      if ($res.ok -eq $false) {
        Write-RunLog -RunContext $RunContext -Level WARN -Message ("{0} extraction failed: {1}" -f $ticket, $res.reason)
      } else {
        $piValue = if ($res.PSObject.Properties['detected_pi_machine']) { ("" + $res.detected_pi_machine).Trim() } else { '' }
        $piSource = if ($res.PSObject.Properties['pi_source']) { ("" + $res.pi_source).Trim() } else { 'none' }
        if (-not $piSource) { $piSource = 'none' }
        Write-RunLog -RunContext $RunContext -Level INFO -Message ("{0} status='{1}' open_tasks={2} pi_source='{3}' pi='{4}'" -f $ticket, $res.status, $res.open_tasks, $piSource, $piValue)
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
    $writeSummary = Write-TicketResultsToExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
      -ResultByTicket $resultMap -NameHeader $NameHeader -PhoneHeader $PhoneHeader -ActionHeader $ActionHeader -SCTasksHeader $SCTasksHeader
    if ($writeSummary) {
      Write-RunLog -RunContext $RunContext -Level INFO -Message ("Writeback summary: matched={0}, updated={1}, skipped={2}" -f [int]$writeSummary.MatchedRows, [int]$writeSummary.UpdatedRows, [int]$writeSummary.SkippedRows)
    }
  }

  return [pscustomobject]@{
    Results = @($results.ToArray())
    CombinedJsonPath = $combinedJsonPath
  }
}

function Apply-ProcessingScope {
  param(
    [Parameter(Mandatory = $true)][string[]]$Tickets,
    [Parameter(Mandatory = $true)][ValidateSet('Auto','RitmOnly','IncOnly','IncAndRitm','All')][string]$Scope
  )

  switch ($Scope) {
    'RitmOnly' { return @($Tickets | Where-Object { $_ -like 'RITM*' }) }
    'IncOnly' { return @($Tickets | Where-Object { $_ -like 'INC*' }) }
    'IncAndRitm' { return @($Tickets | Where-Object { $_ -like 'RITM*' -or $_ -like 'INC*' }) }
    'All' { return @($Tickets) }
    default {
      $ritm = @($Tickets | Where-Object { $_ -like 'RITM*' })
      if ($ritm.Count -gt 0) { return $ritm }
      return @($Tickets)
    }
  }
}
