Set-StrictMode -Version Latest

function Get-SafeCount {
  param($Object)
  if ($null -eq $Object) { return 0 }
  $countProp = $Object.PSObject.Properties['Count']
  if ($countProp) {
    try { return [int]$countProp.Value } catch {}
  }
  $lengthProp = $Object.PSObject.Properties['Length']
  if ($lengthProp) {
    try { return [int]$lengthProp.Value } catch {}
  }
  if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
    $c = 0
    foreach ($nullItem in $Object) { $c++ }
    return $c
  }
  return 1
}
function Invoke-WithExcelWorkbook {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][bool]$ReadOnly,
    [Parameter(Mandatory = $true)][scriptblock]$Action
  )

  $excel = $null
  $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AskToUpdateLinks = $false } catch {}
    try { $excel.EnableEvents = $false } catch {}

    $wb = $excel.Workbooks.Open($ExcelPath, $null, $ReadOnly)
    & $Action $excel $wb
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Get-ExcelHeaderMap {
  param($Worksheet)
  $map = @{}
  $cols = [int]$Worksheet.UsedRange.Columns.Count
  for ($c = 1; $c -le $cols; $c++) {
    $h = ("" + $Worksheet.Cells.Item(1, $c).Text).Trim()
    if ($h -and -not $map.ContainsKey($h)) { $map[$h] = $c }
  }
  return $map
}

function Get-OrCreateHeaderColumn {
  param(
    $Worksheet,
    [hashtable]$HeaderMap,
    [string]$Header
  )

  if ($HeaderMap.ContainsKey($Header)) {
    return [int]$HeaderMap[$Header]
  }

  $newCol = [int]$Worksheet.UsedRange.Columns.Count + 1
  $Worksheet.Cells.Item(1, $newCol) = $Header
  $HeaderMap[$Header] = $newCol
  return $newCol
}

function Resolve-HeaderColumn {
  param(
    [hashtable]$HeaderMap,
    [string[]]$Names,
    [int]$Fallback = 0
  )

  foreach ($name in $Names) {
    foreach ($k in $HeaderMap.Keys) {
      if (("" + $k).Trim().ToLowerInvariant() -eq $name.Trim().ToLowerInvariant()) {
        return [int]$HeaderMap[$k]
      }
    }
  }

  if ($Fallback -gt 0) { return $Fallback }
  return $null
}

function Read-TicketsFromExcel {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][string]$TicketHeader,
    [int]$TicketColumn = 0,
    [int]$StopAfterEmptyRows = 50,
    [int]$MaxRowsAfterFirstTicket = 300
  )

  $tickets = New-Object System.Collections.Generic.List[string]

  Invoke-WithExcelWorkbook -ExcelPath $ExcelPath -ReadOnly $true -Action {
    param($excel, $wb)
    $ws = $null
    $range = $null
    try {
      $ws = $wb.Worksheets.Item($SheetName)
      $map = Get-ExcelHeaderMap -Worksheet $ws
      $ticketCol = Resolve-HeaderColumn -HeaderMap $map -Names @($TicketHeader) -Fallback $TicketColumn
      if (-not $ticketCol) {
        throw "Ticket column not found. Header '$TicketHeader' does not exist."
      }

      $xlUp = -4162
      $rows = 0
      try { $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ticketCol).End($xlUp).Row } catch {}
      if ($rows -le 0) {
        try { $rows = [int]($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1) } catch { $rows = 0 }
      }
      if ($rows -lt 2) { return }

      $range = $ws.Range($ws.Cells.Item(2, $ticketCol), $ws.Cells.Item($rows, $ticketCol))
      $vals = $range.Value2
      $countRows = if ($vals -is [System.Array]) { $vals.GetLength(0) } else { 1 }

      $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
      $emptyStreak = 0
      $firstFoundRow = $null
      for ($i = 1; $i -le $countRows; $i++) {
        $row = $i + 1
        $raw = if ($vals -is [System.Array]) { $vals[$i, 1] } else { $vals }
        $ticket = ("" + $raw).Trim().ToUpperInvariant()

        if ($ticket -match '^(INC|RITM|SCTASK)\d{6,8}$') {
          if ($seen.Add($ticket)) { [void]$tickets.Add($ticket) }
          if (-not $firstFoundRow) { $firstFoundRow = $row }
          $emptyStreak = 0
        } elseif ([string]::IsNullOrWhiteSpace($ticket)) {
          $emptyStreak++
          if ($tickets.Count -gt 0 -and $emptyStreak -ge $StopAfterEmptyRows) { break }
        } else {
          $emptyStreak = 0
        }

        if ($firstFoundRow -and (($row - $firstFoundRow) -ge $MaxRowsAfterFirstTicket) -and $emptyStreak -ge 10) {
          break
        }
      }
    }
    finally {
      try { if ($range) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) } } catch {}
      try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    }
  }

  return @($tickets)
}

function Write-TicketResultsToExcel {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][string]$TicketHeader,
    [int]$TicketColumn = 0,
    [Parameter(Mandatory = $true)]$ResultByTicket,
    [string]$NameHeader = 'Name',
    [string]$PhoneHeader = 'PI',
    [string]$ActionHeader = 'Estado de RITM',
    [string]$SCTasksHeader = 'SCTasks'
  )

  Invoke-WithExcelWorkbook -ExcelPath $ExcelPath -ReadOnly $false -Action {
    param($excel, $wb)

    $ws = $null
    try {
      $ws = $wb.Worksheets.Item($SheetName)
      $map = Get-ExcelHeaderMap -Worksheet $ws
      $ticketCol = Resolve-HeaderColumn -HeaderMap $map -Names @($TicketHeader) -Fallback $TicketColumn
      if (-not $ticketCol) { throw "Ticket column not found. Header '$TicketHeader' does not exist." }

      $nameCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header $NameHeader
      $phoneCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header $PhoneHeader
      $actionCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header $ActionHeader
      $tasksCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header $SCTasksHeader

      $xlUp = -4162
      $rows = 0
      try { $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ticketCol).End($xlUp).Row } catch {}
      if ($rows -le 0) {
        try { $rows = [int]($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1) } catch { $rows = 0 }
      }
      if ($rows -le 1) { return }
      for ($r = 2; $r -le $rows; $r++) {
        $ticket = ("" + $ws.Cells.Item($r, $ticketCol).Text).Trim().ToUpperInvariant()
        if (-not $ResultByTicket.ContainsKey($ticket)) { continue }

        $res = $ResultByTicket[$ticket]
        $affectedUser = if ($res.PSObject.Properties['affected_user']) { ("" + $res.affected_user).Trim() } else { '' }
        if ($affectedUser) {
          $current = ("" + $ws.Cells.Item($r, $nameCol).Text).Trim()
          if ([string]::IsNullOrWhiteSpace($current) -or $current -eq $ticket) {
            $ws.Cells.Item($r, $nameCol) = $affectedUser
          }
        }

        $detectedPi = if ($res.PSObject.Properties['detected_pi_machine']) { ("" + $res.detected_pi_machine).Trim() } else { '' }
        if ($detectedPi) {
          $current = ("" + $ws.Cells.Item($r, $phoneCol).Text).Trim()
          if ([string]::IsNullOrWhiteSpace($current) -or $current -eq $ticket) {
            $ws.Cells.Item($r, $phoneCol) = $detectedPi
          }
        }

        $completion = if ($res.PSObject.Properties['completion_status']) { ("" + $res.completion_status).Trim() } else { 'Pending' }
        $ws.Cells.Item($r, $actionCol) = $completion

        $openTaskNumbers = if ($res.PSObject.Properties['open_task_numbers']) { @($res.open_task_numbers) } else { @() }
        $ws.Cells.Item($r, $tasksCol) = if ((Get-SafeCount $openTaskNumbers) -gt 0) { ($openTaskNumbers -join ', ') } else { '' }
      }

      $wb.Save()
    }
    finally {
      try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    }
  }
}

function Search-DashboardRows {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [string]$SearchText = ''
  )

  $query = ("" + $SearchText).Trim()
  $rowsOut = New-Object System.Collections.Generic.List[object]

  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -Worksheet $ws
    $ritmCol = Resolve-HeaderColumn -HeaderMap $map -Names @('RITM', 'Request Item', 'Number')
    if (-not $ritmCol) { throw 'Dashboard requires a RITM/Number column.' }

    $statusCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Dashboard Status')
    $presentCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Present Time')
    $closedCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Closed Time')
    $taskCol = Resolve-HeaderColumn -HeaderMap $map -Names @('SCTasks', 'SCTask', 'SC Task')
    $nameCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Requested for', 'Name', 'User')

    $rows = [int]($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1)
    $cols = [int]$ws.UsedRange.Columns.Count
    for ($r = 2; $r -le $rows; $r++) {
      $blobParts = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $cols; $c++) {
        $v = ("" + $ws.Cells.Item($r, $c).Text).Trim()
        if ($v) { [void]$blobParts.Add($v) }
      }
      $blob = $blobParts -join ' '
      if ([string]::IsNullOrWhiteSpace($blob)) { continue }
      if ($query) {
        $blobNorm = $blob.ToLowerInvariant()
        $queryNorm = $query.ToLowerInvariant()
        if (-not $blobNorm.Contains($queryNorm)) { continue }
      }

      $rowsOut.Add([pscustomobject]@{
        Row = $r
        RITM = ("" + $ws.Cells.Item($r, $ritmCol).Text).Trim().ToUpperInvariant()
        RequestedFor = if ($nameCol) { ("" + $ws.Cells.Item($r, $nameCol).Text).Trim() } else { '' }
        DashboardStatus = if ($statusCol) { ("" + $ws.Cells.Item($r, $statusCol).Text).Trim() } else { '' }
        PresentTime = if ($presentCol) { ("" + $ws.Cells.Item($r, $presentCol).Text).Trim() } else { '' }
        ClosedTime = if ($closedCol) { ("" + $ws.Cells.Item($r, $closedCol).Text).Trim() } else { '' }
        SCTASK = if ($taskCol) { ("" + $ws.Cells.Item($r, $taskCol).Text).Trim() } else { '' }
      }) | Out-Null
    }
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  return @($rowsOut.ToArray())
}

function Update-DashboardRow {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][int]$Row,
    [Parameter(Mandatory = $true)][string]$Status,
    [string]$SCTaskNumber = ''
  )

  Invoke-WithExcelWorkbook -ExcelPath $ExcelPath -ReadOnly $false -Action {
    param($excel, $wb)
    $ws = $null
    try {
      $ws = $wb.Worksheets.Item($SheetName)
      $map = Get-ExcelHeaderMap -Worksheet $ws
      $statusCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header 'Dashboard Status'
      $presentCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header 'Present Time'
      $closedCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header 'Closed Time'
      $taskCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header 'SCTasks'

      $ws.Cells.Item($Row, $statusCol) = $Status
      if ($Status -eq 'Checked-In') {
        if ([string]::IsNullOrWhiteSpace(("" + $ws.Cells.Item($Row, $presentCol).Text).Trim())) {
          $ws.Cells.Item($Row, $presentCol) = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
      }
      if ($Status -eq 'Checked-Out') {
        if ([string]::IsNullOrWhiteSpace(("" + $ws.Cells.Item($Row, $closedCol).Text).Trim())) {
          $ws.Cells.Item($Row, $closedCol) = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
      }

      if ($SCTaskNumber -and [string]::IsNullOrWhiteSpace(("" + $ws.Cells.Item($Row, $taskCol).Text).Trim())) {
        $ws.Cells.Item($Row, $taskCol) = $SCTaskNumber
      }

      $wb.Save()
    }
    finally {
      try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    }
  }
}


