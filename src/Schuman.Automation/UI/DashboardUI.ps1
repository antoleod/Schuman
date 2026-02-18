Set-StrictMode -Version Latest

function New-DashboardUI {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [string]$SheetName = 'BRU',
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext
  )

  $fontName = Get-UiFontName
  $defaultCheckIn = 'Deliver all credentials to the new user'
  $defaultCheckOut = "Laptop has been delivered.`r`nFirst login made with the user.`r`nOutlook, Teams and Jabber successfully tested."

  $cBg = [System.Drawing.Color]::FromArgb(24, 24, 26)
  $cPanel = [System.Drawing.Color]::FromArgb(30, 30, 32)
  $cPanel2 = [System.Drawing.Color]::FromArgb(36, 36, 38)
  $cBorder = [System.Drawing.Color]::FromArgb(58, 58, 62)
  $cText = [System.Drawing.Color]::FromArgb(230, 230, 230)
  $cMuted = [System.Drawing.Color]::FromArgb(178, 178, 182)
  $cHint = [System.Drawing.Color]::FromArgb(120, 180, 255)
  $cAccent = [System.Drawing.Color]::FromArgb(0, 122, 255)

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Check-in / Check-out Dashboard'
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.Size = New-Object System.Drawing.Size(1120, 760)
  $form.MinimumSize = New-Object System.Drawing.Size(980, 680)
  $form.BackColor = $cBg
  $form.ForeColor = $cText
  $form.Font = New-Object System.Drawing.Font($fontName, 10)

  $btnStyle = {
    param($b, [bool]$accent = $false)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 1
    if ($accent) {
      $b.BackColor = $cAccent
      $b.ForeColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
      $b.FlatAppearance.BorderColor = $cAccent
      $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(20, 138, 255)
      $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0, 106, 226)
    }
    else {
      $b.BackColor = $cPanel2
      $b.ForeColor = $cText
      $b.FlatAppearance.BorderColor = $cBorder
      $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(44, 44, 48)
      $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(32, 32, 34)
    }
  }

  $lblSearch = New-Object System.Windows.Forms.Label
  $lblSearch.Text = 'Search Last Name or First Name:'
  $lblSearch.Location = New-Object System.Drawing.Point(16, 16)
  $lblSearch.AutoSize = $true
  $lblSearch.ForeColor = $cMuted

  $txtSearch = New-Object System.Windows.Forms.ComboBox
  $txtSearch.Location = New-Object System.Drawing.Point(16, 38)
  $txtSearch.Size = New-Object System.Drawing.Size(360, 24)
  $txtSearch.BackColor = [System.Drawing.Color]::FromArgb(34, 34, 36)
  $txtSearch.ForeColor = $cText
  $txtSearch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $txtSearch.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
  $txtSearch.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
  $txtSearch.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems
  $txtSearch.MaxDropDownItems = 14

  $btnRefresh = New-Object System.Windows.Forms.Button
  $btnRefresh.Text = 'Refresh'
  $btnRefresh.Location = New-Object System.Drawing.Point(390, 36)
  $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
  & $btnStyle $btnRefresh $false

  $btnClear = New-Object System.Windows.Forms.Button
  $btnClear.Text = 'Clear'
  $btnClear.Location = New-Object System.Drawing.Point(500, 36)
  $btnClear.Size = New-Object System.Drawing.Size(80, 28)
  & $btnStyle $btnClear $false

  $btnRecalc = New-Object System.Windows.Forms.Button
  $btnRecalc.Text = 'Recalculate from SNOW'
  $btnRecalc.Location = New-Object System.Drawing.Point(590, 36)
  $btnRecalc.Size = New-Object System.Drawing.Size(170, 28)
  & $btnStyle $btnRecalc $false

  $btnOpenSnow = New-Object System.Windows.Forms.Button
  $btnOpenSnow.Text = 'Open in ServiceNow'
  $btnOpenSnow.Location = New-Object System.Drawing.Point(770, 36)
  $btnOpenSnow.Size = New-Object System.Drawing.Size(160, 28)
  & $btnStyle $btnOpenSnow $false

  $chkOpenOnly = New-Object System.Windows.Forms.CheckBox
  $chkOpenOnly.Text = 'Solo RITM abiertos'
  $chkOpenOnly.Location = New-Object System.Drawing.Point(940, 40)
  $chkOpenOnly.Size = New-Object System.Drawing.Size(180, 24)
  $chkOpenOnly.ForeColor = $cMuted
  $chkOpenOnly.BackColor = $cBg

  $chkUltraFast = New-Object System.Windows.Forms.CheckBox
  $chkUltraFast.Text = 'Ultra fast mode'
  $chkUltraFast.Location = New-Object System.Drawing.Point(940, 16)
  $chkUltraFast.Size = New-Object System.Drawing.Size(160, 22)
  $chkUltraFast.ForeColor = $cMuted
  $chkUltraFast.BackColor = $cBg
  $chkUltraFast.Checked = $true

  $lblHint = New-Object System.Windows.Forms.Label
  $lblHint.Text = 'Live filter enabled. Start typing to load matching users and tasks.'
  $lblHint.Location = New-Object System.Drawing.Point(16, 68)
  $lblHint.AutoSize = $true
  $lblHint.ForeColor = $cHint

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Location = New-Object System.Drawing.Point(16, 96)
  $grid.Size = New-Object System.Drawing.Size(1070, 360)
  $grid.ReadOnly = $true
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AllowUserToResizeRows = $false
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EnableHeadersVisualStyles = $false
  $grid.BackgroundColor = $cBg
  $grid.GridColor = $cBorder
  $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $grid.ColumnHeadersDefaultCellStyle.BackColor = $cPanel
  $grid.ColumnHeadersDefaultCellStyle.ForeColor = $cMuted
  $grid.DefaultCellStyle.BackColor = $cBg
  $grid.DefaultCellStyle.ForeColor = $cText
  $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(42, 54, 72)
  $grid.DefaultCellStyle.SelectionForeColor = $cText
  $grid.AlternatingRowsDefaultCellStyle.BackColor = $cPanel
  $grid.RowHeadersVisible = $false
  $grid.RowTemplate.Height = 30
  $grid.AutoGenerateColumns = $false
  $grid.Columns.Clear()

  $colRow = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRow.Name = 'Row'
  $colRow.HeaderText = 'Row'
  $colRow.Visible = $false
  [void]$grid.Columns.Add($colRow)

  $colRequestedFor = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRequestedFor.Name = 'Requested For'
  $colRequestedFor.HeaderText = 'Requested For'
  [void]$grid.Columns.Add($colRequestedFor)

  $colRitm = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRitm.Name = 'RITM'
  $colRitm.HeaderText = 'RITM'
  [void]$grid.Columns.Add($colRitm)

  $colSctask = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colSctask.Name = 'SCTASK'
  $colSctask.HeaderText = 'SCTASK'
  [void]$grid.Columns.Add($colSctask)

  $lblComment = New-Object System.Windows.Forms.Label
  $lblComment.Text = 'Work Note (editable before submit):'
  $lblComment.Location = New-Object System.Drawing.Point(16, 470)
  $lblComment.AutoSize = $true
  $lblComment.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)

  $txtComment = New-Object System.Windows.Forms.TextBox
  $txtComment.Location = New-Object System.Drawing.Point(16, 492)
  $txtComment.Size = New-Object System.Drawing.Size(1070, 130)
  $txtComment.Multiline = $true
  $txtComment.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
  $txtComment.Text = $defaultCheckIn
  $txtComment.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
  $txtComment.ForeColor = $cText
  $txtComment.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

  $btnUseCheckInNote = New-Object System.Windows.Forms.Button
  $btnUseCheckInNote.Text = 'Use Check-In Note'
  $btnUseCheckInNote.Location = New-Object System.Drawing.Point(16, 626)
  $btnUseCheckInNote.Size = New-Object System.Drawing.Size(160, 28)
  & $btnStyle $btnUseCheckInNote $false

  $btnUseCheckOutNote = New-Object System.Windows.Forms.Button
  $btnUseCheckOutNote.Text = 'Use Check-Out Note'
  $btnUseCheckOutNote.Location = New-Object System.Drawing.Point(188, 626)
  $btnUseCheckOutNote.Size = New-Object System.Drawing.Size(160, 28)
  & $btnStyle $btnUseCheckOutNote $false

  $btnCheckIn = New-Object System.Windows.Forms.Button
  $btnCheckIn.Text = 'CHECK-IN'
  $btnCheckIn.Location = New-Object System.Drawing.Point(360, 626)
  $btnCheckIn.Size = New-Object System.Drawing.Size(160, 36)
  & $btnStyle $btnCheckIn $true

  $btnCheckOut = New-Object System.Windows.Forms.Button
  $btnCheckOut.Text = 'CHECK-OUT'
  $btnCheckOut.Location = New-Object System.Drawing.Point(532, 626)
  $btnCheckOut.Size = New-Object System.Drawing.Size(160, 36)
  & $btnStyle $btnCheckOut $false

  $lblStatus = New-Object System.Windows.Forms.Label
  $lblStatus.Text = 'Type to filter users. Nothing is loaded by default.'
  $lblStatus.Location = New-Object System.Drawing.Point(710, 634)
  $lblStatus.Size = New-Object System.Drawing.Size(700, 28)
  $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)

  $form.Controls.AddRange(@(
      $lblSearch, $txtSearch, $btnRefresh, $btnClear, $btnRecalc, $btnOpenSnow, $chkUltraFast, $chkOpenOnly, $lblHint,
      $grid, $lblComment, $txtComment, $btnUseCheckInNote, $btnUseCheckOutNote, $btnCheckIn, $btnCheckOut, $lblStatus
    ))

  $state = [pscustomobject]@{
    Config = $Config
    RunContext = $RunContext
    ExcelPath = $ExcelPath
    SheetName = $SheetName
    Session = $null
    Rows = @()
    AllRows = @()
    AllRowsUniverse = @()
    LastSearch = ''
    QueryCache = @{}
    ExcelStamp = 0L
    UserDirectory = @()
    UltraFast = $true
    DefaultCheckIn = $defaultCheckIn
    DefaultCheckOut = $defaultCheckOut
    Controls = @{
      Search = $txtSearch
      Grid = $grid
      OpenOnly = $chkOpenOnly
      UltraFast = $chkUltraFast
      Comment = $txtComment
      Status = $lblStatus
      SearchTimer = $null
    }
  }
  $form.Tag = $state

  $bindRowsToGrid = {
    param($rows)
    $grid.SuspendLayout()
    try {
      $grid.Rows.Clear()
      foreach ($x in @($rows)) {
        [void]$grid.Rows.Add(
          ("" + $x.Row),
          ("" + $x.RequestedFor),
          ("" + $x.RITM),
          ("" + $x.SCTASK)
        )
      }
      $grid.ClearSelection()
      if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
        $grid.CurrentCell = $grid.Rows[0].Cells['Requested For']
      }
    }
    finally {
      $grid.ResumeLayout()
    }
  }

  $getSelectedRow = {
    if ($grid.SelectedRows.Count -eq 0) { return $null }
    $selected = $grid.SelectedRows[0]
    if (-not $selected) { return $null }

    $rowText = ("" + $selected.Cells['Row'].Value).Trim()
    $rowNum = 0
    if (-not [int]::TryParse($rowText, [ref]$rowNum)) { return $null }

    foreach ($item in $state.Rows) {
      if ([int]$item.Row -eq $rowNum) { return $item }
    }
    return $null
  }

  $isRowOpenLocal = {
    param($rowItem)
    if (-not $rowItem) { return $false }
    $status = ("" + $rowItem.DashboardStatus).Trim().ToLowerInvariant()
    if (-not $status) { return $true }
    return ($status -match 'open|new|pending|in\s*progress|checked-in|work')
  }

  $updateActionButtons = {
    $sel = & $getSelectedRow
    $hasValidRitm = $false
    if ($sel) {
      $ritmTxt = ("" + $sel.RITM).Trim().ToUpperInvariant()
      $hasValidRitm = ($ritmTxt -match '^RITM\d{6,8}$')
    }
    $btnCheckIn.Enabled = $hasValidRitm
    $btnCheckOut.Enabled = $hasValidRitm
    $btnRecalc.Enabled = $hasValidRitm
    $btnOpenSnow.Enabled = $hasValidRitm
  }

  $getVisibleRows = {
    $rows = @($state.AllRows)
    if ($chkOpenOnly.Checked) {
      $rows = @($rows | Where-Object { & $isRowOpenLocal $_ })
    }
    return @($rows)
  }

  $getExcelStamp = {
    try {
      return [int64](Get-Item -LiteralPath $state.ExcelPath).LastWriteTimeUtc.Ticks
    }
    catch {
      return 0L
    }
  }

  $fetchRows = {
    param(
      [string]$QueryText = '',
      [switch]$ForceReload
    )

    $stamp = & $getExcelStamp
    if ($ForceReload -or ($stamp -ne [int64]$state.ExcelStamp)) {
      $state.QueryCache = @{}
      $state.ExcelStamp = [int64]$stamp
    }

    $key = ("" + $QueryText).Trim().ToLowerInvariant()

    if ($state.UltraFast) {
      if ($ForceReload -or $state.AllRowsUniverse.Count -eq 0) {
        $state.AllRowsUniverse = @(Search-DashboardRows -ExcelPath $state.ExcelPath -SheetName $state.SheetName -SearchText '')
        foreach ($r in $state.AllRowsUniverse) {
          $blob = "{0} {1} {2} {3} {4} {5}" -f ("" + $r.RequestedFor), ("" + $r.RITM), ("" + $r.SCTASK), ("" + $r.DashboardStatus), ("" + $r.PresentTime), ("" + $r.ClosedTime)
          $r | Add-Member -NotePropertyName __search -NotePropertyValue $blob.ToLowerInvariant() -Force
        }
      }

      if (-not $key) { return @($state.AllRowsUniverse) }
      if ($state.QueryCache.ContainsKey($key)) { return @($state.QueryCache[$key]) }
      $rowsFast = @($state.AllRowsUniverse | Where-Object { ("" + $_.__search).Contains($key) })
      $state.QueryCache[$key] = @($rowsFast)
      return @($rowsFast)
    }

    if ($state.QueryCache.ContainsKey($key)) { return @($state.QueryCache[$key]) }
    $rows = @(Search-DashboardRows -ExcelPath $state.ExcelPath -SheetName $state.SheetName -SearchText $QueryText)
    $state.QueryCache[$key] = @($rows)
    return @($rows)
  }

  $updateSearchUserSuggestions = ({
    $q = ("" + $state.Controls.Search.Text).Trim()
    $allUsers = @($state.UserDirectory)
    if ($allUsers.Count -eq 0) { return }

    $matches = @()
    if ([string]::IsNullOrWhiteSpace($q)) {
      $matches = @($allUsers | Select-Object -First 200)
    }
    else {
      $matches = @($allUsers | Where-Object { ("" + $_).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | Select-Object -First 200)
    }

    $caret = $state.Controls.Search.SelectionStart
    $state.Controls.Search.BeginUpdate()
    try {
      $state.Controls.Search.Items.Clear()
      foreach ($u in $matches) { [void]$state.Controls.Search.Items.Add($u) }
    }
    finally {
      $state.Controls.Search.EndUpdate()
    }
    $state.Controls.Search.SelectionStart = [Math]::Min($caret, $state.Controls.Search.Text.Length)
    $state.Controls.Search.SelectionLength = 0
  }).GetNewClosure()

  $ensureSession = {
    if ($state.Session) { return $true }
    try {
      $lblStatus.Text = 'Opening ServiceNow session...'
      $state.Session = New-ServiceNowSession -Config $state.Config -RunContext $state.RunContext
      return $true
    }
    catch {
      $lblStatus.Text = 'ServiceNow session failed.'
      [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'ServiceNow') | Out-Null
      return $false
    }
  }

  $openRowInServiceNow = {
    param($rowItem)
    if (-not $rowItem) { return }
    $ritm = ("" + $rowItem.RITM).Trim().ToUpperInvariant()
    if (-not ($ritm -match '^RITM\d{6,8}$')) { return }
    $url = "{0}/nav_to.do?uri=%2Fsc_req_item_list.do%3Fsysparm_query%3Dnumber%3D{1}" -f $state.Config.ServiceNow.BaseUrl, $ritm
    try { Start-Process $url | Out-Null } catch {}
  }

  $performSearch = ({
    param([switch]$ReloadFromExcel)
    try {
      $q = ("" + $state.Controls.Search.Text).Trim()
      $state.AllRows = @(& $fetchRows -QueryText $q -ForceReload:$ReloadFromExcel)

      $rows = & $getVisibleRows
      $state.Rows = @($rows)
      $state.LastSearch = $q
      & $bindRowsToGrid $rows
      & $updateActionButtons

      $filterNote = if ($chkOpenOnly.Checked) { ' (solo abiertos)' } else { '' }
      if ([string]::IsNullOrWhiteSpace($q)) {
        $lblStatus.Text = "Preloaded: $($rows.Count) rows$filterNote"
      }
      else {
        $lblStatus.Text = "Results: $($rows.Count) for '$q'$filterNote"
      }
    }
    catch {
      $err = $_.Exception.Message
      [System.Windows.Forms.MessageBox]::Show("Search failed: $err", 'Dashboard Error') | Out-Null
    }
  }).GetNewClosure()

  $applyAction = {
    param([string]$action)
    $row = & $getSelectedRow
    if (-not $row) {
      [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
      return
    }

    if (-not (& $ensureSession)) { return }

    $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
    $note = ("" + $state.Controls.Comment.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($note)) {
      $note = if ($action -eq 'checkin') { $state.DefaultCheckIn } else { $state.DefaultCheckOut }
      $state.Controls.Comment.Text = $note
    }

    $tasks = @(Get-ServiceNowTasksForRitm -Session $state.Session -RitmNumber $ritm)
    if ($tasks.Count -eq 0) {
      [System.Windows.Forms.MessageBox]::Show("No task found for $ritm.", 'Tasks') | Out-Null
      return
    }
    $task = $tasks[0]

    $targetLabel = if ($action -eq 'checkin') { 'Work in Progress' } else { 'Closed Complete' }
    $dashboardStatus = if ($action -eq 'checkin') { 'Checked-In' } else { 'Checked-Out' }

    $lblStatus.Text = "Updating $($task.number)..."
    $ok = Set-ServiceNowTaskState -Session $state.Session -TaskSysId ("" + $task.sys_id) -TargetStateLabel $targetLabel -WorkNote $note
    if (-not $ok) {
      [System.Windows.Forms.MessageBox]::Show("ServiceNow update failed for $($task.number).", 'Action failed') | Out-Null
      $lblStatus.Text = 'Action failed.'
      return
    }

    Update-DashboardRow -ExcelPath $state.ExcelPath -SheetName $state.SheetName -Row ([int]$row.Row) -Status $dashboardStatus -SCTaskNumber ("" + $task.number)
    $lblStatus.Text = ("{0}: {1} ({2})" -f $dashboardStatus, $row.RITM, $task.number)
    & $performSearch -ReloadFromExcel
  }

  $recalculateRow = {
    $row = & $getSelectedRow
    if (-not $row) {
      [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
      return
    }
    if (-not (& $ensureSession)) { return }

    $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
    $tasks = @(Get-ServiceNowTasksForRitm -Session $state.Session -RitmNumber $ritm)
    if ($tasks.Count -eq 0) {
      $lblStatus.Text = "Recalculated: no tasks for $ritm."
      return
    }

    $openCount = 0
    $wipCount = 0
    foreach ($t in $tasks) {
      $st = ("" + $t.state).Trim().ToLowerInvariant()
      $sv = ("" + $t.state_value).Trim().ToLowerInvariant()
      if ($st -match 'work|progress' -or $sv -eq '2') {
        $wipCount++
      }
      elseif ($st -match 'open|new|pending' -or $sv -eq '1') {
        $openCount++
      }
    }

    $newStatus = 'Checked-Out'
    if ($openCount -gt 0) { $newStatus = 'Open' }
    elseif ($wipCount -gt 0) { $newStatus = 'Checked-In' }

    Update-DashboardRow -ExcelPath $state.ExcelPath -SheetName $state.SheetName -Row ([int]$row.Row) -Status $newStatus -SCTaskNumber ("" + $tasks[0].number)
    $lblStatus.Text = "Recalculated: $ritm => $newStatus"
    & $performSearch -ReloadFromExcel
  }

  $searchTimer = New-Object System.Windows.Forms.Timer
  $searchTimer.Interval = 260
  $state.Controls.SearchTimer = $searchTimer
  $searchTimer.Add_Tick(({
    if ($state.Controls.SearchTimer) { $state.Controls.SearchTimer.Stop() }
    & $performSearch
  }).GetNewClosure())

  $scheduleSearch = ({
    if (-not $state.Controls.SearchTimer) { return }
    $state.Controls.SearchTimer.Stop()
    $state.Controls.SearchTimer.Start()
  }).GetNewClosure()

  $txtSearch.Add_KeyDown(({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
      $e.SuppressKeyPress = $true
      if ($state.Controls.SearchTimer) { $state.Controls.SearchTimer.Stop() }
      & $performSearch
    }
  }).GetNewClosure())
  $txtSearch.Add_TextUpdate(({
    & $updateSearchUserSuggestions
    & $scheduleSearch
  }).GetNewClosure())
  $txtSearch.Add_DropDown(({ & $updateSearchUserSuggestions }.GetNewClosure()))
  $txtSearch.Add_SelectedIndexChanged(({ & $scheduleSearch }.GetNewClosure()))

  $chkOpenOnly.Add_CheckedChanged(({
    if ([string]::IsNullOrWhiteSpace($state.LastSearch)) {
      $state.Rows = @()
      & $bindRowsToGrid @()
      & $updateActionButtons
      $lblStatus.Text = 'Type First/Last name to search.'
      return
    }
    $rows = & $getVisibleRows
    $state.Rows = @($rows)
    & $bindRowsToGrid $rows
    & $updateActionButtons
    $filterNote = if ($chkOpenOnly.Checked) { ' (solo abiertos)' } else { '' }
    $lblStatus.Text = "Results: $($rows.Count) for '$($state.LastSearch)'$filterNote"
  }).GetNewClosure())

  $chkUltraFast.Add_CheckedChanged(({
    $state.UltraFast = [bool]$chkUltraFast.Checked
    $state.QueryCache = @{}
    $state.AllRowsUniverse = @()
    & $performSearch -ReloadFromExcel
  }).GetNewClosure())

  $grid.Add_SelectionChanged(({ & $updateActionButtons }.GetNewClosure()))
  $grid.Add_CellDoubleClick(({ & $applyAction 'checkin' }.GetNewClosure()))

  $btnRefresh.Add_Click(({
    $state.Controls.Search.Text = $state.LastSearch
    & $performSearch -ReloadFromExcel
  }).GetNewClosure())

  $btnClear.Add_Click(({
    $state.Controls.Search.Text = ''
    & $performSearch
  }).GetNewClosure())

  $btnUseCheckInNote.Add_Click(({ $state.Controls.Comment.Text = $state.DefaultCheckIn }.GetNewClosure()))
  $btnUseCheckOutNote.Add_Click(({ $state.Controls.Comment.Text = $state.DefaultCheckOut }.GetNewClosure()))
  $btnOpenSnow.Add_Click(({
    $row = & $getSelectedRow
    if (-not $row) {
      [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
      return
    }
    & $openRowInServiceNow $row
  }).GetNewClosure())
  $btnRecalc.Add_Click(({ & $recalculateRow }.GetNewClosure()))
  $btnCheckIn.Add_Click(({ & $applyAction 'checkin' }.GetNewClosure()))
  $btnCheckOut.Add_Click(({ & $applyAction 'checkout' }.GetNewClosure()))

  $form.add_FormClosed(({
    param($sender, $eventArgs)
    try { if ($sender.Tag.Controls.SearchTimer) { $sender.Tag.Controls.SearchTimer.Stop(); $sender.Tag.Controls.SearchTimer.Dispose() } } catch {}
    try { if ($sender.Tag.Session) { Close-ServiceNowSession -Session $sender.Tag.Session } } catch {}
  }).GetNewClosure())

  try {
    $allRows = @(& $fetchRows -QueryText '' -ForceReload)
    $state.UserDirectory = @($allRows | ForEach-Object { ("" + $_.RequestedFor).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    & $updateSearchUserSuggestions
    $state.AllRows = @($allRows)
    $state.Rows = @(& $getVisibleRows)
    & $bindRowsToGrid $state.Rows
    $lblStatus.Text = "Ready. Users loaded: $($state.UserDirectory.Count). Preloaded rows: $($state.Rows.Count)."
  }
  catch {
    $lblStatus.Text = 'Ready.'
  }

  & $updateActionButtons
  return $form
}
