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

function Test-InvalidUserCellValue {
  param([string]$Value)

  $v = ("" + $Value).Trim()
  if (-not $v) { return $true }
  if ($v -match '^[0-9a-fA-F]{32}$') { return $true }
  if ($v -match '(?i)\bnew\b.*\bep\b.*\busers?\b') { return $true }
  if ($v -match '(?i)^new\b.*\busers?\b') { return $true }
  if ($v -match '(?i)^unknown$|^n/?a$|^null$') { return $true }
  return $false
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
    if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
      Register-SchumanOwnedComResource -Kind 'ExcelApplication' -Object $excel -Tag 'excel-workbook' | Out-Null
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.AskToUpdateLinks = $false } catch {}
    try { $excel.EnableEvents = $false } catch {}

    $wb = $excel.Workbooks.Open($ExcelPath, $null, $ReadOnly)
    if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
      Register-SchumanOwnedComResource -Kind 'Workbook' -Object $wb -Tag 'excel-workbook' | Out-Null
    }
    & $Action $excel $wb
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($wb -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $wb } } catch {}
    try { if ($excel -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $excel } } catch {}
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
    [string]$ActionHeader = 'RITM State',
    [string]$SCTasksHeader = 'SCTasks'
  )

  $summary = Invoke-WithExcelWorkbook -ExcelPath $ExcelPath -ReadOnly $false -Action {
    param($excel, $wb)

    $ws = $null
    $debugLogPath = Join-Path $env:TEMP 'Schuman\pi-transplant-debug.log'
    try { New-Item -ItemType Directory -Force -Path (Split-Path -Parent $debugLogPath) | Out-Null } catch {}
    $debugCount = 0
    $stats = [ordered]@{
      MatchedRows = 0
      UpdatedRows = 0
      SkippedRows = 0
    }
    try {
      $ws = $wb.Worksheets.Item($SheetName)
      $map = Get-ExcelHeaderMap -Worksheet $ws
      $ticketCol = Resolve-HeaderColumn -HeaderMap $map -Names @($TicketHeader) -Fallback $TicketColumn
      if (-not $ticketCol) { throw "Ticket column not found. Header '$TicketHeader' does not exist." }

      $namePrimaryCol = Resolve-HeaderColumn -HeaderMap $map -Names @($NameHeader, 'Requested for', 'Name', 'User')
      if (-not $namePrimaryCol) {
        $namePrimaryCol = Get-OrCreateHeaderColumn -Worksheet $ws -HeaderMap $map -Header $NameHeader
      }
      $nameCols = New-Object System.Collections.Generic.List[int]
      foreach ($candidate in @(
          $namePrimaryCol,
          (Resolve-HeaderColumn -HeaderMap $map -Names @('Requested for')),
          (Resolve-HeaderColumn -HeaderMap $map -Names @('Name')),
          (Resolve-HeaderColumn -HeaderMap $map -Names @('User'))
        )) {
        if ($candidate -and -not $nameCols.Contains([int]$candidate)) {
          [void]$nameCols.Add([int]$candidate)
        }
      }
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

        $stats.MatchedRows++
        $rowChanged = $false
        $res = $ResultByTicket[$ticket]
        $affectedUser = if ($res.PSObject.Properties['affected_user']) { ("" + $res.affected_user).Trim() } else { '' }
        $legalName = if ($res.PSObject.Properties['legal_name']) { ("" + $res.legal_name).Trim() } else { '' }
        if (($ticket -like 'RITM*') -and $legalName) {
          if (Test-InvalidUserCellValue -Value $affectedUser) {
            $affectedUser = $legalName
          }
        }
        if ($affectedUser) {
          foreach ($nameCol in $nameCols) {
            $current = ("" + $ws.Cells.Item($r, $nameCol).Text).Trim()
            $replaceCurrent = [string]::IsNullOrWhiteSpace($current) -or ($current -eq $ticket) -or (Test-InvalidUserCellValue -Value $current)
            if ($replaceCurrent) {
              if ($current -ne $affectedUser) {
                $ws.Cells.Item($r, $nameCol) = $affectedUser
                $rowChanged = $true
              }
            }
          }
        }

        $detectedPi = if ($res.PSObject.Properties['detected_pi_machine']) { ("" + $res.detected_pi_machine).Trim() } else { '' }
        $piSource = if ($res.PSObject.Properties['pi_source']) { ("" + $res.pi_source).Trim() } else { 'none' }
        if (-not $piSource) { $piSource = 'none' }
        if ($detectedPi) {
          $current = ("" + $ws.Cells.Item($r, $phoneCol).Text).Trim()
          if ([string]::IsNullOrWhiteSpace($current) -or $current -eq $ticket) {
            if ($current -ne $detectedPi) {
              $ws.Cells.Item($r, $phoneCol) = $detectedPi
              $rowChanged = $true
            }
          }
        }
        if ($debugCount -lt 10) {
          $debugCount++
          $line = "[{0}] ticket={1} row={2} col={3} pi='{4}' source='{5}'" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $ticket, $r, $phoneCol, $detectedPi, $piSource
          if (-not $detectedPi) {
            $line += " empty_reason='pi_source:" + $piSource + "'"
          }
          try { Add-Content -Path $debugLogPath -Value $line -Encoding UTF8 } catch {}
        }

        $completion = if ($res.PSObject.Properties['completion_status']) { ("" + $res.completion_status).Trim() } else { 'Pending' }
        $currentCompletion = ("" + $ws.Cells.Item($r, $actionCol).Text).Trim()
        if ($currentCompletion -ne $completion) {
          $ws.Cells.Item($r, $actionCol) = $completion
          $rowChanged = $true
        }

        $openTaskNumbers = if ($res.PSObject.Properties['open_task_numbers']) { @($res.open_task_numbers) } else { @() }
        $newTaskText = if ((Get-SafeCount $openTaskNumbers) -gt 0) { ($openTaskNumbers -join ', ') } else { '' }
        $currentTaskText = ("" + $ws.Cells.Item($r, $tasksCol).Text).Trim()
        if ($currentTaskText -ne $newTaskText) {
          $ws.Cells.Item($r, $tasksCol) = $newTaskText
          $rowChanged = $true
        }

        if ($rowChanged) { $stats.UpdatedRows++ } else { $stats.SkippedRows++ }
      }

      $wb.Save()
      return [pscustomobject]$stats
    }
    finally {
      try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    }
  }
  return $summary
}

function global:Search-DashboardRows {
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
    if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
      Register-SchumanOwnedComResource -Kind 'ExcelApplication' -Object $excel -Tag 'dashboard-search' | Out-Null
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
    if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
      Register-SchumanOwnedComResource -Kind 'Workbook' -Object $wb -Tag 'dashboard-search' | Out-Null
    }
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -Worksheet $ws
    $ritmCol = Resolve-HeaderColumn -HeaderMap $map -Names @('RITM', 'Request Item', 'Number')
    if (-not $ritmCol) { throw 'Dashboard requires a RITM/Number column.' }

    $statusCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Dashboard Status')
    $ritmStateCol = Resolve-HeaderColumn -HeaderMap $map -Names @(
      'RITM State',
      'RITM Status',
      'Estado de RITM',
      'Estado RITM',
      'Completion Status',
      'Action'
    )
    $taskStateCol = Resolve-HeaderColumn -HeaderMap $map -Names @(
      'SCTASK State',
      'SCTASK Status',
      'SC Task State',
      'Task State',
      'Estado de SCTASK',
      'Estado SCTASK'
    )
    $presentCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Present Time')
    $closedCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Closed Time')
    $taskCol = Resolve-HeaderColumn -HeaderMap $map -Names @('SCTasks', 'SCTask', 'SC Task')
    $nameCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Requested for', 'Name', 'User')
    $piCol = Resolve-HeaderColumn -HeaderMap $map -Names @('PI', 'Phone', 'Configuration Item', 'CI')

    $xlUp = -4162
    $rows = 0
    try {
      $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ritmCol).End($xlUp).Row
    }
    catch {
      $rows = [int]($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1)
    }
    if ($rows -lt 2) { return }

    $getColValues = {
      param([int]$ColumnIndex)
      $out = @{}
      if (-not $ColumnIndex -or $ColumnIndex -le 0) { return $out }

      # Some workbooks return unexpected COM array shapes for Value2, so read cell-by-cell for stability.
      for ($excelRow = 2; $excelRow -le $rows; $excelRow++) {
        $cellText = ''
        try {
          $cellObj = $ws.Cells.Item($excelRow, $ColumnIndex)
          $cellText = ("" + $cellObj.Text).Trim()
        }
        catch {
          $cellText = ''
        }
        $out[$excelRow] = $cellText
      }
      return $out
    }

    $ritmValues = [hashtable](& $getColValues $ritmCol)
    $nameValues = [hashtable](& $getColValues $nameCol)
    $statusValues = [hashtable](& $getColValues $statusCol)
    $ritmStateValues = [hashtable](& $getColValues $ritmStateCol)
    $taskStateValues = [hashtable](& $getColValues $taskStateCol)
    $presentValues = [hashtable](& $getColValues $presentCol)
    $closedValues = [hashtable](& $getColValues $closedCol)
    $taskValues = [hashtable](& $getColValues $taskCol)
    $piValues = [hashtable](& $getColValues $piCol)

    $normalizeStateText = {
      param([string]$Text)
      $raw = ("" + $Text).Trim()
      if (-not $raw) { return '' }
      $norm = $raw.ToLowerInvariant()
      if ($norm -match 'closed|complete|completed|resolved|cancel|cerrad|complet|resuelt|cancelad|chius|ferm|annul') {
        return 'Complete'
      }
      if ($norm -match 'work.?in.?progress|in.?progress|progress|wip|en progreso|progreso') {
        return 'Work in Progress'
      }
      if ($norm -match 'open|pending|new|abierto|pendiente|nuevo') {
        return 'Pending'
      }
      return $raw
    }

    $queryNorm = $query.ToLowerInvariant()
    for ($r = 2; $r -le $rows; $r++) {
      $ritm = if ($ritmValues.ContainsKey($r)) { ("" + $ritmValues[$r]).Trim().ToUpperInvariant() } else { '' }
      if (-not $ritm) { continue }

      $requestedFor = if ($nameValues.ContainsKey($r)) { ("" + $nameValues[$r]).Trim() } else { '' }
      $dashboardStatus = if ($statusValues.ContainsKey($r)) { ("" + $statusValues[$r]).Trim() } else { '' }
      $ritmStateRaw = if ($ritmStateValues.ContainsKey($r)) { ("" + $ritmStateValues[$r]).Trim() } else { '' }
      $taskStateRaw = if ($taskStateValues.ContainsKey($r)) { ("" + $taskStateValues[$r]).Trim() } else { '' }
      $ritmState = & $normalizeStateText $ritmStateRaw
      $taskState = & $normalizeStateText $taskStateRaw
      $presentTime = if ($presentValues.ContainsKey($r)) { ("" + $presentValues[$r]).Trim() } else { '' }
      $closedTime = if ($closedValues.ContainsKey($r)) { ("" + $closedValues[$r]).Trim() } else { '' }
      $sctask = if ($taskValues.ContainsKey($r)) { ("" + $taskValues[$r]).Trim() } else { '' }
      $pi = if ($piValues.ContainsKey($r)) { ("" + $piValues[$r]).Trim() } else { '' }

      if ($query) {
        $blobNorm = ("{0} {1} {2} {3} {4} {5} {6} {7} {8}" -f $requestedFor, $ritm, $sctask, $pi, $dashboardStatus, $ritmState, $taskState, $presentTime, $closedTime).ToLowerInvariant()
        if (-not $blobNorm.Contains($queryNorm)) { continue }
      }

      $rowsOut.Add([pscustomobject]@{
        Row = $r
        RITM = $ritm
        RequestedFor = $requestedFor
        PI = $pi
        DashboardStatus = $dashboardStatus
        RITMState = $ritmState
        SCTASKState = $taskState
        PresentTime = $presentTime
        ClosedTime = $closedTime
        SCTASK = $sctask
      }) | Out-Null
    }
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $ws } } catch {}
    try { if ($wb -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $wb } } catch {}
    try { if ($excel -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $excel } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  return @($rowsOut.ToArray())
}

function global:Update-DashboardRow {
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


