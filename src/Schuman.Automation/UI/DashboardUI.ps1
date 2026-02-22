Set-StrictMode -Version Latest

function global:New-DashboardUI {
  <#
  .SYNOPSIS
  Creates and returns the Check-in/Check-out Dashboard form.
  .DESCRIPTION
  Builds the Dashboard WinForms UI, wires event handlers through safe UI wrappers,
  and loads Excel-backed rows for ServiceNow-assisted operations.
  .PARAMETER ExcelPath
  Path to the Excel workbook used as data source/writeback target.
  .PARAMETER SheetName
  Worksheet name used by Dashboard operations.
  .PARAMETER Config
  Runtime configuration hashtable.
  .PARAMETER RunContext
  Run context used by logging/integrations.
  .PARAMETER InitialSession
  Optional pre-authenticated ServiceNow session.
  .OUTPUTS
  System.Windows.Forms.Form
  .NOTES
  UI updates are performed on the UI thread through safe wrappers.
  #>
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    $InitialSession = $null,
    [scriptblock]$OnRefreshExcel = $null
  )

  $fontName = 'Segoe UI'
  $uiFontCmd = Get-Command -Name Get-UiFontName -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($uiFontCmd) {
    try {
      $candidateFont = ("" + (& $uiFontCmd)).Trim()
      if ($candidateFont) { $fontName = $candidateFont }
    } catch {}
  }
  $defaultCheckIn = 'Deliver all credentials to the new user'
  $defaultCheckOut = "Laptop has been delivered.`r`nFirst login made with the user.`r`nOutlook, Teams and Jabber successfully tested."
  $defaultTemplates = @{
    checkIn = $defaultCheckIn
    checkOut = $defaultCheckOut
  }

  $templateStorePath = Join-Path (Join-Path $env:APPDATA 'Schuman') 'dashboard-templates.json'

  $cBg = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
  $cPanel = [System.Drawing.ColorTranslator]::FromHtml('#1E293B')
  $cPanel2 = [System.Drawing.ColorTranslator]::FromHtml('#172033')
  $cBorder = [System.Drawing.ColorTranslator]::FromHtml('#334155')
  $cText = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
  $cMuted = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
  $cHint = [System.Drawing.ColorTranslator]::FromHtml('#3B82F6')
  $cAccent = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Check-in / Check-out Dashboard'
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.Size = New-Object System.Drawing.Size(1120, 760)
  $form.MinimumSize = New-Object System.Drawing.Size(980, 680)
  $form.BackColor = $cBg
  $form.ForeColor = $cText
  $form.Font = New-Object System.Drawing.Font($fontName, 11)
  $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

  if (Get-Command -Name Initialize-SchumanRuntime -ErrorAction SilentlyContinue) {
    Initialize-SchumanRuntime -RevalidateOnly
  }
  $getFunctionCommand = ({
    param([string]$Name)
    return (Get-Command -Name $Name -CommandType Function -ErrorAction SilentlyContinue | Select-Object -First 1)
  }).GetNewClosure()
  $searchDashboardRowsCommand = & $getFunctionCommand 'Search-DashboardRows'
  if (-not $searchDashboardRowsCommand) {
    throw 'Search-DashboardRows is not available. Runtime initialization failed.'
  }
  $testClosedStateHandler = ${function:Test-ClosedState}
  if (-not $testClosedStateHandler) {
    $testClosedStateCommand = & $getFunctionCommand 'Test-ClosedState'
    if ($testClosedStateCommand -and $testClosedStateCommand.ScriptBlock) {
      $testClosedStateHandler = $testClosedStateCommand.ScriptBlock
    }
  }
  if (-not $testClosedStateHandler) {
    throw 'Test-ClosedState is not available. Runtime initialization failed.'
  }
  $newServiceNowSessionHandler = ${function:New-ServiceNowSession}
  if (-not $newServiceNowSessionHandler) {
    $cmd = & $getFunctionCommand 'New-ServiceNowSession'
    if ($cmd -and $cmd.ScriptBlock) { $newServiceNowSessionHandler = $cmd.ScriptBlock }
  }
  $getServiceNowTasksForRitmHandler = ${function:Get-ServiceNowTasksForRitm}
  if (-not $getServiceNowTasksForRitmHandler) {
    $cmd = & $getFunctionCommand 'Get-ServiceNowTasksForRitm'
    if ($cmd -and $cmd.ScriptBlock) { $getServiceNowTasksForRitmHandler = $cmd.ScriptBlock }
  }
  $setServiceNowTaskStateHandler = ${function:Set-ServiceNowTaskState}
  if (-not $setServiceNowTaskStateHandler) {
    $cmd = & $getFunctionCommand 'Set-ServiceNowTaskState'
    if ($cmd -and $cmd.ScriptBlock) { $setServiceNowTaskStateHandler = $cmd.ScriptBlock }
  }
  $updateDashboardRowHandler = ${function:Update-DashboardRow}
  if (-not $updateDashboardRowHandler) {
    $cmd = & $getFunctionCommand 'Update-DashboardRow'
    if ($cmd -and $cmd.ScriptBlock) { $updateDashboardRowHandler = $cmd.ScriptBlock }
  }

  $btnStyle = ({
    param($b, [bool]$accent = $false)
    $b.Font = New-Object System.Drawing.Font($fontName, 10.5, [System.Drawing.FontStyle]::Bold)
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
    $b.Add_EnabledChanged(({
          param($sender, $eventArgs)
          try {
            if ($sender.Enabled) {
              $sender.ForeColor = $cText
            }
            else {
              $sender.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
            }
          }
          catch {}
        }).GetNewClosure())
  }).GetNewClosure()

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

  $btnOpenSnow = New-Object System.Windows.Forms.Button
  $btnOpenSnow.Text = 'Open in ServiceNow'
  $btnOpenSnow.Location = New-Object System.Drawing.Point(690, 36)
  $btnOpenSnow.Size = New-Object System.Drawing.Size(160, 28)
  & $btnStyle $btnOpenSnow $false

  $btnTemplateSettings = New-Object System.Windows.Forms.Button
  $btnTemplateSettings.Text = [string][char]0x2699
  $btnTemplateSettings.Location = New-Object System.Drawing.Point(940, 36)
  $btnTemplateSettings.Size = New-Object System.Drawing.Size(38, 28)
  & $btnStyle $btnTemplateSettings $false

  $chkOpenOnly = New-Object System.Windows.Forms.CheckBox
  $chkOpenOnly.Text = 'Open RITM only'
  $chkOpenOnly.Location = New-Object System.Drawing.Point(984, 40)
  $chkOpenOnly.Size = New-Object System.Drawing.Size(180, 24)
  $chkOpenOnly.ForeColor = $cMuted
  $chkOpenOnly.BackColor = $cBg

  $lblHint = New-Object System.Windows.Forms.Label
  $lblHint.Text = 'Live filter enabled. Start typing to load matching users and tasks.'
  $lblHint.Location = New-Object System.Drawing.Point(16, 68)
  $lblHint.AutoSize = $true
  $lblHint.ForeColor = $cHint

  $loadingBar = New-Object System.Windows.Forms.ProgressBar
  $loadingBar.Location = New-Object System.Drawing.Point(16, 90)
  $loadingBar.Size = New-Object System.Drawing.Size(1070, 4)
  $loadingBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $loadingBar.MarqueeAnimationSpeed = 28
  $loadingBar.Visible = $false

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

  $colPi = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colPi.Name = 'PI'
  $colPi.HeaderText = 'PI'
  [void]$grid.Columns.Add($colPi)

  $colSctask = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colSctask.Name = 'SCTASK'
  $colSctask.HeaderText = 'SCTASK'
  [void]$grid.Columns.Add($colSctask)

  $colRitmState = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRitmState.Name = 'RITM State'
  $colRitmState.HeaderText = 'RITM State'
  [void]$grid.Columns.Add($colRitmState)

  $colTaskState = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colTaskState.Name = 'SCTASK State'
  $colTaskState.HeaderText = 'SCTASK State'
  [void]$grid.Columns.Add($colTaskState)

  $workNoteHost = New-Object System.Windows.Forms.TableLayoutPanel
  $workNoteHost.ColumnCount = 2
  $workNoteHost.RowCount = 1
  $workNoteHost.Margin = New-Object System.Windows.Forms.Padding(0)
  $workNoteHost.Padding = New-Object System.Windows.Forms.Padding(0)
  [void]$workNoteHost.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$workNoteHost.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 240)))

  $workNoteLeft = New-Object System.Windows.Forms.TableLayoutPanel
  $workNoteLeft.Dock = [System.Windows.Forms.DockStyle]::Fill
  $workNoteLeft.ColumnCount = 1
  $workNoteLeft.RowCount = 2
  [void]$workNoteLeft.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$workNoteLeft.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $workNoteLeft.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
  [void]$workNoteHost.Controls.Add($workNoteLeft, 0, 0)

  $lblComment = New-Object System.Windows.Forms.Label
  $lblComment.Text = 'Work Note (editable before submit):'
  $lblComment.AutoSize = $true
  $lblComment.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)
  $lblComment.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
  [void]$workNoteLeft.Controls.Add($lblComment, 0, 0)

  $txtComment = New-Object System.Windows.Forms.TextBox
  $txtComment.Dock = [System.Windows.Forms.DockStyle]::Fill
  $txtComment.Multiline = $true
  $txtComment.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
  $txtComment.Text = $defaultCheckIn
  $txtComment.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
  $txtComment.ForeColor = $cText
  $txtComment.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  [void]$workNoteLeft.Controls.Add($txtComment, 0, 1)

  $workNoteRight = New-Object System.Windows.Forms.TableLayoutPanel
  $workNoteRight.Dock = [System.Windows.Forms.DockStyle]::Fill
  $workNoteRight.ColumnCount = 1
  $workNoteRight.RowCount = 3
  [void]$workNoteRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$workNoteRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$workNoteRight.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $workNoteRight.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
  [void]$workNoteHost.Controls.Add($workNoteRight, 1, 0)

  $lblNoteTools = New-Object System.Windows.Forms.Label
  $lblNoteTools.Text = 'Note Templates:'
  $lblNoteTools.AutoSize = $true
  $lblNoteTools.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)
  $lblNoteTools.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
  [void]$workNoteRight.Controls.Add($lblNoteTools, 0, 0)

  $btnUseCheckInNote = New-Object System.Windows.Forms.Button
  $btnUseCheckInNote.Text = 'Use Check-In Note'
  $btnUseCheckInNote.Size = New-Object System.Drawing.Size(220, 28)
  $btnUseCheckInNote.Dock = [System.Windows.Forms.DockStyle]::Fill
  $btnUseCheckInNote.Visible = $true
  & $btnStyle $btnUseCheckInNote $false

  $btnUseCheckOutNote = New-Object System.Windows.Forms.Button
  $btnUseCheckOutNote.Text = 'Use Check-Out Note'
  $btnUseCheckOutNote.Size = New-Object System.Drawing.Size(220, 28)
  $btnUseCheckOutNote.Dock = [System.Windows.Forms.DockStyle]::Fill
  $btnUseCheckOutNote.Visible = $true
  & $btnStyle $btnUseCheckOutNote $false

  $noteButtonsPanel = New-Object System.Windows.Forms.TableLayoutPanel
  $noteButtonsPanel.Dock = [System.Windows.Forms.DockStyle]::Top
  $noteButtonsPanel.ColumnCount = 1
  $noteButtonsPanel.RowCount = 2
  [void]$noteButtonsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
  [void]$noteButtonsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
  $noteButtonsPanel.AutoSize = $false
  $noteButtonsPanel.Height = 68
  $noteButtonsPanel.Margin = New-Object System.Windows.Forms.Padding(0)
  $noteButtonsPanel.Visible = $true
  [void]$noteButtonsPanel.Controls.Add($btnUseCheckInNote, 0, 0)
  [void]$noteButtonsPanel.Controls.Add($btnUseCheckOutNote, 0, 1)
  [void]$workNoteRight.Controls.Add($noteButtonsPanel, 0, 1)

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

  $btnCloseCode = New-Object System.Windows.Forms.Button
  $btnCloseCode.Text = 'Close All'
  $btnCloseCode.Location = New-Object System.Drawing.Point(704, 626)
  $btnCloseCode.Size = New-Object System.Drawing.Size(210, 36)
  & $btnStyle $btnCloseCode $false
  $btnCloseCode.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#7F1D1D')
  $btnCloseCode.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
  $btnCloseCode.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')

  $lblStatus = New-Object System.Windows.Forms.Label
  $lblStatus.Text = 'Type to filter users. Nothing is loaded by default.'
  $lblStatus.Location = New-Object System.Drawing.Point(16, 668)
  $lblStatus.Size = New-Object System.Drawing.Size(1070, 28)
  $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170)

  $form.Controls.AddRange(@(
      $lblSearch, $txtSearch, $btnRefresh, $btnClear, $btnOpenSnow, $btnTemplateSettings, $chkOpenOnly, $lblHint, $loadingBar,
      $grid, $workNoteHost, $btnCheckIn, $btnCheckOut, $btnCloseCode, $lblStatus
    ))

  $layoutDashboard = ({
    $clientW = [Math]::Max(980, $form.ClientSize.Width)
    $clientH = [Math]::Max(680, $form.ClientSize.Height)
    $m = 16
    $g = 10

    $lblSearch.Location = New-Object System.Drawing.Point($m, 16)

    $rightColumnW = 196
    $rightColumnX = [Math]::Max(760, $clientW - $rightColumnW - $m)
    $chkOpenOnly.Location = New-Object System.Drawing.Point($rightColumnX, 40)

    $topButtonsY = 36
    $btnTemplateSettings.Size = New-Object System.Drawing.Size(38, 28)
    $btnTemplateSettings.Location = New-Object System.Drawing.Point(($rightColumnX - $g - $btnTemplateSettings.Width), $topButtonsY)

    $btnOpenSnow.Size = New-Object System.Drawing.Size(160, 28)
    $btnOpenSnow.Location = New-Object System.Drawing.Point(($btnTemplateSettings.Left - $g - $btnOpenSnow.Width), $topButtonsY)
    $btnClear.Size = New-Object System.Drawing.Size(80, 28)
    $btnClear.Location = New-Object System.Drawing.Point(($btnOpenSnow.Left - $g - $btnClear.Width), $topButtonsY)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $btnRefresh.Location = New-Object System.Drawing.Point(($btnClear.Left - $g - $btnRefresh.Width), $topButtonsY)

    $txtSearchX = $m
    $txtSearchW = [Math]::Max(240, $btnRefresh.Left - $g - $txtSearchX)
    $txtSearch.Location = New-Object System.Drawing.Point($txtSearchX, 38)
    $txtSearch.Size = New-Object System.Drawing.Size($txtSearchW, 24)

    $lblHint.Location = New-Object System.Drawing.Point($m, 68)
    $loadingBar.Location = New-Object System.Drawing.Point($m, 90)
    $loadingBar.Size = New-Object System.Drawing.Size(($clientW - ($m * 2)), 4)

    $gridTop = 96
    $bottomButtonsTop = $clientH - 54
    $commentTop = [Math]::Max(($gridTop + 180), ($bottomButtonsTop - 210))
    $commentHeight = [Math]::Max(120, $bottomButtonsTop - $commentTop - 16)
    $gridHeight = [Math]::Max(220, $commentTop - $gridTop - 14)

    $grid.Location = New-Object System.Drawing.Point($m, $gridTop)
    $grid.Size = New-Object System.Drawing.Size(($clientW - ($m * 2)), $gridHeight)

    $workNoteHost.Location = New-Object System.Drawing.Point($m, ($commentTop + 2))
    $workNoteHost.Size = New-Object System.Drawing.Size(($clientW - ($m * 2)), $commentHeight)

    $btnCloseCode.Location = New-Object System.Drawing.Point(($clientW - $m - $btnCloseCode.Width), $bottomButtonsTop)
    $btnCheckOut.Location = New-Object System.Drawing.Point(($btnCloseCode.Left - 10 - $btnCheckOut.Width), $bottomButtonsTop)
    $btnCheckIn.Location = New-Object System.Drawing.Point(($btnCheckOut.Left - 10 - $btnCheckIn.Width), $bottomButtonsTop)

    $lblStatus.Location = New-Object System.Drawing.Point($m, ($clientH - 24))
    $lblStatus.Size = New-Object System.Drawing.Size(($clientW - ($m * 2)), 20)
  }).GetNewClosure()

  $state = [pscustomobject]@{
    Config = $Config
    RunContext = $RunContext
    ExcelPath = $ExcelPath
    SheetName = $SheetName
    Session = $InitialSession
    OwnsSession = ($null -eq $InitialSession)
    Rows = @()
    AllRows = @()
    AllRowsUniverse = @()
    LastSearch = ''
    QueryCache = @{}
    ExcelStamp = 0L
    ExcelReady = $false
    IsLoading = $false
    UserDirectory = @()
    UltraFast = $true
    DefaultCheckIn = $defaultCheckIn
    DefaultCheckOut = $defaultCheckOut
    TemplateStorePath = $templateStorePath
    Templates = @{}
    Controls = @{
      Search = $txtSearch
      Grid = $grid
      OpenOnly = $chkOpenOnly
      Comment = $txtComment
      Status = $lblStatus
      TemplateSettings = $btnTemplateSettings
      LoadingBar = $loadingBar
      SearchTimer = $null
    }
  }
  $form.Tag = $state

  $setLoadingState = ({
    param(
      [bool]$IsLoading,
      [string]$Message = ''
    )
    $state.IsLoading = [bool]$IsLoading
    if ($state.Controls.LoadingBar) {
      $state.Controls.LoadingBar.Visible = $IsLoading
    }
    if ($IsLoading -and -not [string]::IsNullOrWhiteSpace($Message)) {
      $lblStatus.Text = $Message
      try { [System.Windows.Forms.Application]::DoEvents() } catch {}
    }
    try { & $updateActionButtons } catch {}
  }).GetNewClosure()

  $ensureTemplateStoreDirectory = ({
    $dir = Split-Path -Parent $state.TemplateStorePath
    if (-not $dir) { return }
    if (-not (Test-Path -LiteralPath $dir)) {
      New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
  }).GetNewClosure()

  $saveTemplates = ({
    param([hashtable]$TemplatesToSave)
    & $ensureTemplateStoreDirectory
    $payload = @{
      checkIn = ("" + $TemplatesToSave.checkIn)
      checkOut = ("" + $TemplatesToSave.checkOut)
    }
    ($payload | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $state.TemplateStorePath -Encoding UTF8
  }).GetNewClosure()

  $loadTemplates = ({
    $templates = @{
      checkIn = $defaultTemplates.checkIn
      checkOut = $defaultTemplates.checkOut
    }
    try {
      if (Test-Path -LiteralPath $state.TemplateStorePath) {
        $raw = Get-Content -LiteralPath $state.TemplateStorePath -Raw
        if ($raw) {
          $obj = $raw | ConvertFrom-Json
          if ($obj) {
            $ci = ("" + $obj.checkIn)
            $co = ("" + $obj.checkOut)
            if (-not [string]::IsNullOrWhiteSpace($ci)) { $templates.checkIn = $ci }
            if (-not [string]::IsNullOrWhiteSpace($co)) { $templates.checkOut = $co }
          }
        }
      }
    } catch {}
    return $templates
  }).GetNewClosure()

  $resolveRowLifecycle = ({
    param($rowItem)
    if (-not $rowItem) {
      return [pscustomobject]@{
        IsClosed = $false
        IsOpen = $true
      }
    }

    $dashboardStatus = ("" + $rowItem.DashboardStatus).Trim()
    $ritmState = ("" + $rowItem.RITMState).Trim()
    $taskState = ("" + $rowItem.SCTASKState).Trim()
    $closedByDashboard = ($dashboardStatus -match '(?i)checked-out|closed|complete|resolved|done')
    $closedByStates = (& $testClosedStateHandler -StateLabel $ritmState -StateValue $ritmState) -or (& $testClosedStateHandler -StateLabel $taskState -StateValue $taskState)
    $isClosed = $closedByDashboard -or $closedByStates
    return [pscustomobject]@{
      IsClosed = $isClosed
      IsOpen = (-not $isClosed)
    }
  }).GetNewClosure()

  $renderTemplateForRow = ({
    param(
      [string]$TemplateText,
      $rowItem
    )
    if ([string]::IsNullOrWhiteSpace($TemplateText)) { return '' }
    if (-not $rowItem) { return $TemplateText }
    $stateValue = ("" + $rowItem.RITMState).Trim()
    if (-not $stateValue) { $stateValue = ("" + $rowItem.DashboardStatus).Trim() }
    if (-not $stateValue) { $stateValue = ("" + $rowItem.SCTASKState).Trim() }

    $out = ("" + $TemplateText)
    $replacements = @{
      '{RequestedFor}' = ("" + $rowItem.RequestedFor).Trim()
      '{RITM}' = ("" + $rowItem.RITM).Trim()
      '{SCTASK}' = ("" + $rowItem.SCTASK).Trim()
      '{PI}' = ("" + $rowItem.PI).Trim()
      '{State}' = $stateValue
    }
    foreach ($k in $replacements.Keys) {
      $v = ("" + $replacements[$k])
      $out = $out.Replace($k, $v)
    }
    return $out
  }).GetNewClosure()

  $bindRowsToGrid = ({
    param($rows)
    $grid.SuspendLayout()
    try {
      $grid.Rows.Clear()
      foreach ($x in @($rows)) {
        $ritmStateText = ("" + $x.RITMState).Trim()
        if (-not $ritmStateText) { $ritmStateText = ("" + $x.DashboardStatus).Trim() }
        $taskStateText = ("" + $x.SCTASKState).Trim()
        if (-not $taskStateText) {
          if ($ritmStateText -match '(?i)checked-out|closed|complete|resolved|done') { $taskStateText = 'Closed' }
          elseif ($ritmStateText -match '(?i)checked-in|work|progress') { $taskStateText = 'Work in Progress' }
          else { $taskStateText = 'Open' }
        }
        [void]$grid.Rows.Add(
          ("" + $x.Row),
          ("" + $x.RequestedFor),
          ("" + $x.RITM),
          ("" + $x.PI),
          ("" + $x.SCTASK),
          $ritmStateText,
          $taskStateText
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
  }).GetNewClosure()

  $getSelectedRow = ({
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
  }).GetNewClosure()

  $isRowOpenLocal = ({
    param($rowItem)
    if (-not $rowItem) { return $false }
    $lifecycle = & $resolveRowLifecycle $rowItem
    return [bool]$lifecycle.IsOpen
  }).GetNewClosure()

  $ensureExcelReady = ({
    if ([bool]$state.ExcelReady) { return $true }
    $lblStatus.Text = 'Excel still loading...'
    return $false
  }).GetNewClosure()

  $updateActionButtons = ({
    $excelReady = [bool]$state.ExcelReady
    $sel = & $getSelectedRow
    $hasValidRitm = $false
    $isClosed = $false
    if ($sel) {
      $ritmTxt = ("" + $sel.RITM).Trim().ToUpperInvariant()
      $hasValidRitm = ($ritmTxt -match '^RITM\d{6,8}$')
      $life = & $resolveRowLifecycle $sel
      $isClosed = [bool]$life.IsClosed
    }
    $allowActions = ($excelReady -and -not $state.IsLoading)
    $btnRefresh.Enabled = $true
    $btnClear.Enabled = $allowActions
    $btnTemplateSettings.Enabled = $allowActions
    $btnCheckIn.Enabled = ($allowActions -and $hasValidRitm -and -not $isClosed)
    $btnCheckOut.Enabled = ($allowActions -and $hasValidRitm)
    # Recalculate button intentionally removed.
    $btnOpenSnow.Enabled = ($allowActions -and $hasValidRitm)
    $txtSearch.Enabled = $allowActions
    $chkOpenOnly.Enabled = $allowActions
  }).GetNewClosure()

  $getVisibleRows = ({
    $rows = @($state.AllRows)
    if ($chkOpenOnly.Checked) {
      $filtered = New-Object System.Collections.Generic.List[object]
      foreach ($candidate in $rows) {
        $isOpen = $false
        try { $isOpen = [bool](& $isRowOpenLocal $candidate) } catch { $isOpen = $false }
        if ($isOpen) { [void]$filtered.Add($candidate) }
      }
      $rows = @($filtered.ToArray())
    }
    return @($rows)
  }).GetNewClosure()

  $getExcelStamp = ({
    try {
      return [int64](Get-Item -LiteralPath $state.ExcelPath).LastWriteTimeUtc.Ticks
    }
    catch {
      return 0L
    }
  }).GetNewClosure()

  $fetchRows = ({
    param(
      [string]$QueryText = '',
      [switch]$ForceReload
    )

    if ([string]::IsNullOrWhiteSpace(("" + $state.ExcelPath).Trim()) -or -not (Test-Path -LiteralPath $state.ExcelPath)) {
      return @()
    }

    $stamp = & $getExcelStamp
    if ($ForceReload -or ($stamp -ne [int64]$state.ExcelStamp)) {
      $state.QueryCache = @{}
      $state.ExcelStamp = [int64]$stamp
    }

    $key = ("" + $QueryText).Trim().ToLowerInvariant()

    if ($state.UltraFast) {
      if (-not $ForceReload -and $state.AllRowsUniverse.Count -eq 0 -and $key) {
        if ($state.QueryCache.ContainsKey($key)) { return @($state.QueryCache[$key]) }
        $rowsQuick = @(Search-DashboardRows -ExcelPath $state.ExcelPath -SheetName $state.SheetName -SearchText $QueryText)
        $state.QueryCache[$key] = @($rowsQuick)
        return @($rowsQuick)
      }

      if ($ForceReload -or ($state.AllRowsUniverse.Count -eq 0 -and -not $key)) {
        $state.AllRowsUniverse = @(Search-DashboardRows -ExcelPath $state.ExcelPath -SheetName $state.SheetName -SearchText '')
        foreach ($r in $state.AllRowsUniverse) {
          $blob = "{0} {1} {2} {3} {4} {5} {6} {7} {8}" -f ("" + $r.RequestedFor), ("" + $r.RITM), ("" + $r.PI), ("" + $r.SCTASK), ("" + $r.DashboardStatus), ("" + $r.RITMState), ("" + $r.SCTASKState), ("" + $r.PresentTime), ("" + $r.ClosedTime)
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
  }).GetNewClosure()

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

  $ensureSession = ({
    if ($state.Session) { return $true }
    if (-not $newServiceNowSessionHandler) {
      [System.Windows.Forms.MessageBox]::Show('ServiceNow integration is not loaded. Please restart Schuman and try again.', 'ServiceNow') | Out-Null
      $lblStatus.Text = 'ServiceNow integration is unavailable.'
      return $false
    }
    try {
      $lblStatus.Text = 'Opening ServiceNow session...'
      $state.Session = & $newServiceNowSessionHandler -Config $state.Config -RunContext $state.RunContext
      return $true
    }
    catch {
      $lblStatus.Text = 'ServiceNow session failed.'
      $msg = "Cannot open ServiceNow session.`r`n`r`n$($_.Exception.Message)`r`n`r`nTry again. If it persists, check network/SSO."
      [System.Windows.Forms.MessageBox]::Show($msg, 'ServiceNow') | Out-Null
      return $false
    }
  }).GetNewClosure()

  $openRowInServiceNow = ({
    param($rowItem)
    if (-not $rowItem) { return }
    $ritm = ("" + $rowItem.RITM).Trim().ToUpperInvariant()
    if (-not ($ritm -match '^RITM\d{6,8}$')) { return }
    $url = "{0}/nav_to.do?uri=%2Fsc_req_item_list.do%3Fsysparm_query%3Dnumber%3D{1}" -f $state.Config.ServiceNow.BaseUrl, $ritm
    try { Start-Process $url | Out-Null } catch {}
  }).GetNewClosure()

  $refreshSelectionStatus = ({
    $row = & $getSelectedRow
    if (-not $row) { return }
    $ritm = ("" + $row.RITM).Trim()
    $sctask = ("" + $row.SCTASK).Trim()
    $ritmState = ("" + $row.RITMState).Trim()
    $taskState = ("" + $row.SCTASKState).Trim()
    $life = & $resolveRowLifecycle $row
    $openState = if ($life.IsClosed) { 'CLOSED' } else { 'OPEN' }
    $lblStatus.Text = ("Selected: {0} | Task: {1} | RITM State: {2} | SCTASK State: {3} | {4}" -f $ritm, $sctask, $ritmState, $taskState, $openState)
  }).GetNewClosure()

  $performSearch = ({
    param([switch]$ReloadFromExcel)
    try {
      & $setLoadingState -IsLoading $true -Message 'Loading Excel data...'
      $q = ("" + $state.Controls.Search.Text).Trim()
      if ([string]::IsNullOrWhiteSpace($q)) {
        if ($ReloadFromExcel -or $state.AllRows.Count -eq 0) {
          $state.AllRows = @(& $fetchRows -QueryText '' -ForceReload:$ReloadFromExcel)
        }
        $state.Rows = @(& $getVisibleRows)
        $state.ExcelReady = ($state.Rows.Count -gt 0)
        $state.LastSearch = ''
        & $bindRowsToGrid $state.Rows
        & $updateActionButtons
        if ($state.ExcelReady) {
          $lblStatus.Text = "Preloaded $($state.Rows.Count) rows from Excel."
          & $appendHistory "Preloaded $($state.Rows.Count) rows from Excel."
        } else {
          $lblStatus.Text = if ((Test-Path -LiteralPath $state.ExcelPath)) { 'Excel is empty' } else { 'Excel still loading...' }
        }
        return
      }

      $state.AllRows = @(& $fetchRows -QueryText $q -ForceReload:$ReloadFromExcel)
      $newUsers = @($state.AllRows | ForEach-Object { ("" + $_.RequestedFor).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
      if ($newUsers.Count -gt 0) {
        $state.UserDirectory = @($state.UserDirectory + $newUsers | Sort-Object -Unique)
      }

      $rows = & $getVisibleRows
      $state.Rows = @($rows)
      $state.ExcelReady = ($state.Rows.Count -gt 0)
      $state.LastSearch = $q
      & $bindRowsToGrid $rows
      & $updateActionButtons

      $filterNote = if ($chkOpenOnly.Checked) { ' (open only)' } else { '' }
      if ($state.ExcelReady) {
        $lblStatus.Text = "Results: $($rows.Count) for '$q'$filterNote"
        & $appendHistory "Search '$q' => $($rows.Count) row(s)$filterNote."
      } else {
        $lblStatus.Text = if ((Test-Path -LiteralPath $state.ExcelPath)) { 'Excel is empty' } else { 'Excel still loading...' }
      }
    }
    catch {
      $err = $_.Exception.Message
      $stack = ''
      $position = ''
      try { $stack = ("" + $_.ScriptStackTrace).Trim() } catch {}
      try { if ($_.InvocationInfo) { $position = ("" + $_.InvocationInfo.PositionMessage).Trim() } } catch {}
      $state.ExcelReady = $false
      & $updateActionButtons
      try {
        if ($stack -or $position) {
          Write-Log -Level ERROR -Message ("Dashboard search failed: {0} | {1} | {2}" -f $err, $position, $stack)
        }
        else {
          Write-Log -Level ERROR -Message ("Dashboard search failed: " + $err)
        }
      } catch {}
      $globalShowUiError = & $getFunctionCommand 'global:Show-UiError'
      if ($globalShowUiError -and $globalShowUiError.ScriptBlock) {
        & $globalShowUiError.ScriptBlock -Title 'Dashboard Error' -Message 'Search failed.' -Exception $_.Exception
      }
      else {
        [System.Windows.Forms.MessageBox]::Show("Search failed: $err", 'Dashboard Error') | Out-Null
      }
    }
    finally {
      & $setLoadingState -IsLoading $false
    }
  }).GetNewClosure()

  $applyAction = ({
    param([string]$action)
    try {
      if (-not (& $ensureExcelReady)) { return }
      $row = & $getSelectedRow
      if (-not $row) {
        [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
        return
      }

      if (-not (& $ensureSession)) { return }

      $lifecycle = & $resolveRowLifecycle $row
      if ($action -eq 'checkin' -and $lifecycle.IsClosed) {
        [System.Windows.Forms.MessageBox]::Show("RITM appears closed (`"$($row.RITM)`"). Check-in is blocked.", 'Ticket Closed') | Out-Null
        $lblStatus.Text = "Blocked: $($row.RITM) is closed."
        return
      }

      $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
      $note = ("" + $state.Controls.Comment.Text).Trim()
      if ([string]::IsNullOrWhiteSpace($note)) {
        $note = if ($action -eq 'checkin') { $state.DefaultCheckIn } else { $state.DefaultCheckOut }
        $state.Controls.Comment.Text = $note
      }

      if (-not $getServiceNowTasksForRitmHandler -or -not $setServiceNowTaskStateHandler -or -not $updateDashboardRowHandler) {
        [System.Windows.Forms.MessageBox]::Show('Dashboard dependencies are missing. Restart the app to reload integrations.', 'Dashboard Error') | Out-Null
        $lblStatus.Text = 'Dependencies missing. Action canceled.'
        return
      }

      $tasks = @(& $getServiceNowTasksForRitmHandler -Session $state.Session -RitmNumber $ritm)
      if ($tasks.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No task found for $ritm.", 'Tasks') | Out-Null
        return
      }
      $task = $tasks[0]

      $targetLabel = if ($action -eq 'checkin') { 'Work in Progress' } else { 'Closed Complete' }
      $dashboardStatus = if ($action -eq 'checkin') { 'Checked-In' } else { 'Checked-Out' }

      $lblStatus.Text = "Updating $($task.number)..."
      $ok = & $setServiceNowTaskStateHandler -Session $state.Session -TaskSysId ("" + $task.sys_id) -TargetStateLabel $targetLabel -WorkNote $note
      if (-not $ok) {
        [System.Windows.Forms.MessageBox]::Show("ServiceNow update failed for $($task.number). Please try again and verify connection/permissions.", 'Action failed') | Out-Null
        $lblStatus.Text = 'Action failed.'
        return
      }

      & $updateDashboardRowHandler -ExcelPath $state.ExcelPath -SheetName $state.SheetName -Row ([int]$row.Row) -Status $dashboardStatus -SCTaskNumber ("" + $task.number)
      $lblStatus.Text = ("{0}: {1} ({2})" -f $dashboardStatus, $row.RITM, $task.number)
      & $appendHistory ("{0}: {1} ({2})" -f $dashboardStatus, $row.RITM, $task.number)
      & $performSearch -ReloadFromExcel
    }
    catch {
      $err = $_.Exception.Message
      try { Write-Log -Level ERROR -Message ("Dashboard action failed: " + $err) } catch {}
      [System.Windows.Forms.MessageBox]::Show("Action failed.`r`n$err`r`n`r`nTry again. If it persists, check connection and login.", 'Dashboard Error') | Out-Null
    }
  }).GetNewClosure()

  $recalculateRow = ({
    try {
      if (-not (& $ensureExcelReady)) { return }
      & $setLoadingState -IsLoading $true -Message 'Recalculating row from ServiceNow...'
      $row = & $getSelectedRow
      if (-not $row) {
        [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
        return
      }
      if (-not (& $ensureSession)) { return }

      if (-not $getServiceNowTasksForRitmHandler -or -not $updateDashboardRowHandler) {
        [System.Windows.Forms.MessageBox]::Show('Cannot recalculate because ServiceNow helpers are missing. Restart Schuman and try again.', 'Dashboard Error') | Out-Null
        $lblStatus.Text = 'Recalculate failed: integration unavailable.'
        return
      }

      $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
      $tasks = @(& $getServiceNowTasksForRitmHandler -Session $state.Session -RitmNumber $ritm)
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

      & $updateDashboardRowHandler -ExcelPath $state.ExcelPath -SheetName $state.SheetName -Row ([int]$row.Row) -Status $newStatus -SCTaskNumber ("" + $tasks[0].number)
      $lblStatus.Text = "Recalculated: $ritm => $newStatus"
      & $appendHistory ("Recalculated {0} => {1}" -f $ritm, $newStatus)
      & $performSearch -ReloadFromExcel
    }
    catch {
      $err = $_.Exception.Message
      try { Write-Log -Level ERROR -Message ("Dashboard recalculate failed: " + $err) } catch {}
      [System.Windows.Forms.MessageBox]::Show("Recalculate failed.`r`n$err`r`n`r`nTry again. If it persists, check network/SSO and Excel file lock.", 'Dashboard Error') | Out-Null
    }
    finally {
      & $setLoadingState -IsLoading $false
    }
  }).GetNewClosure()

  $openTemplateManager = ({
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Templates'
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.Size = New-Object System.Drawing.Size(760, 560)
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.BackColor = $cBg
    $dlg.ForeColor = $cText
    $dlg.Font = New-Object System.Drawing.Font($fontName, 10)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tabs.Padding = New-Object System.Drawing.Point(14, 6)
    $dlg.Controls.Add($tabs)

    $tabCheckIn = New-Object System.Windows.Forms.TabPage
    $tabCheckIn.Text = 'Check-In'
    $tabCheckIn.BackColor = $cBg
    $tabCheckIn.ForeColor = $cText
    $tabs.TabPages.Add($tabCheckIn) | Out-Null

    $tabCheckOut = New-Object System.Windows.Forms.TabPage
    $tabCheckOut.Text = 'Check-Out'
    $tabCheckOut.BackColor = $cBg
    $tabCheckOut.ForeColor = $cText
    $tabs.TabPages.Add($tabCheckOut) | Out-Null

    $txtCheckInTemplate = New-Object System.Windows.Forms.TextBox
    $txtCheckInTemplate.Multiline = $true
    $txtCheckInTemplate.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtCheckInTemplate.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtCheckInTemplate.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $txtCheckInTemplate.ForeColor = $cText
    $txtCheckInTemplate.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtCheckInTemplate.Text = ("" + $state.Templates.checkIn)
    $tabCheckIn.Controls.Add($txtCheckInTemplate)

    $txtCheckOutTemplate = New-Object System.Windows.Forms.TextBox
    $txtCheckOutTemplate.Multiline = $true
    $txtCheckOutTemplate.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtCheckOutTemplate.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtCheckOutTemplate.BackColor = [System.Drawing.Color]::FromArgb(37, 37, 38)
    $txtCheckOutTemplate.ForeColor = $cText
    $txtCheckOutTemplate.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtCheckOutTemplate.Text = ("" + $state.Templates.checkOut)
    $tabCheckOut.Controls.Add($txtCheckOutTemplate)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = 'Placeholders: {RequestedFor} {RITM} {SCTASK} {PI} {State}'
    $hint.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $hint.Height = 26
    $hint.ForeColor = $cMuted
    $dlg.Controls.Add($hint)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $buttons.Height = 44
    $buttons.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $dlg.Controls.Add($buttons)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = 'Save'
    $btnSave.Width = 96
    & $btnStyle $btnSave $true
    $buttons.Controls.Add($btnSave) | Out-Null

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Width = 96
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    & $btnStyle $btnCancel $false
    $buttons.Controls.Add($btnCancel) | Out-Null

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = 'Reset defaults'
    $btnReset.Width = 120
    & $btnStyle $btnReset $false
    $buttons.Controls.Add($btnReset) | Out-Null

    $btnReset.Add_Click(({
      $txtCheckInTemplate.Text = ("" + $defaultTemplates.checkIn)
      $txtCheckOutTemplate.Text = ("" + $defaultTemplates.checkOut)
    }).GetNewClosure())

    $btnSave.Add_Click(({
      $newTemplates = @{
        checkIn = ("" + $txtCheckInTemplate.Text)
        checkOut = ("" + $txtCheckOutTemplate.Text)
      }
      & $saveTemplates $newTemplates
      $state.Templates = $newTemplates
      $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
      $dlg.Close()
    }).GetNewClosure())

    $dlg.CancelButton = $btnCancel
    $dlg.AcceptButton = $btnSave
    $applyThemeToControlTreeHandler = ${function:Apply-ThemeToControlTree}
    $themeVar = Get-Variable -Name CurrentMainTheme -Scope Global -ErrorAction SilentlyContinue
    $scaleVar = Get-Variable -Name CurrentMainFontScale -Scope Global -ErrorAction SilentlyContinue
    $scale = 1.0
    if ($scaleVar -and $scaleVar.Value) { try { $scale = [double]$scaleVar.Value } catch { $scale = 1.0 } }
    if ($applyThemeToControlTreeHandler -and $themeVar -and $themeVar.Value) {
      & $applyThemeToControlTreeHandler -RootControl $dlg -Theme $themeVar.Value -FontScale $scale
    }
    [void]$dlg.ShowDialog($form)
  }).GetNewClosure()

  $preloadRows = ({
    try {
      & $setLoadingState -IsLoading $true -Message 'Preloading Excel rows...'
      $state.AllRows = @(& $fetchRows -QueryText '' -ForceReload)
      $state.Rows = @(& $getVisibleRows)
      $state.ExcelReady = ($state.Rows.Count -gt 0)
      $state.LastSearch = ''
      $state.UserDirectory = @($state.AllRows | ForEach-Object { ("" + $_.RequestedFor).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
      & $bindRowsToGrid $state.Rows
      & $updateActionButtons
      if ($state.ExcelReady) {
        $lblStatus.Text = "Preloaded $($state.Rows.Count) rows from Excel."
      } else {
        $lblStatus.Text = if ((Test-Path -LiteralPath $state.ExcelPath)) { 'Excel is empty' } else { 'Excel still loading...' }
      }
      try { Write-Log -Level INFO -Message ("Preloaded {0} rows from Excel." -f $state.Rows.Count) } catch {}
    }
    catch {
      $state.ExcelReady = $false
      & $updateActionButtons
      $lblStatus.Text = 'Preload failed.'
      $globalShowUiError = & $getFunctionCommand 'global:Show-UiError'
      if ($globalShowUiError -and $globalShowUiError.ScriptBlock) {
        & $globalShowUiError.ScriptBlock -Title 'Dashboard' -Message 'Preload failed.' -Exception $_.Exception
      }
    }
    finally {
      & $setLoadingState -IsLoading $false
    }
  }).GetNewClosure()

  $searchTimer = New-Object System.Windows.Forms.Timer
  $searchTimer.Interval = 150
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

  $runSafeUi = ({
    param([string]$ctx, [scriptblock]$act)
    if (Get-Command -Name Invoke-SafeUiAction -CommandType Function -ErrorAction SilentlyContinue) {
      $null = Invoke-SafeUiAction -ActionName $ctx -Action $act
      return
    }
    try { & $act } catch {
      $globalShowUiError = & $getFunctionCommand 'global:Show-UiError'
      if ($globalShowUiError -and $globalShowUiError.ScriptBlock) {
        & $globalShowUiError.ScriptBlock -Title 'Dashboard' -Message ("{0} failed." -f $ctx) -Exception $_.Exception
      }
    }
  }).GetNewClosure()

  $appendHistory = ({
    param([string]$Line)
    $safeLine = ("" + $Line).Trim()
    if (-not $safeLine) { return }
    try { Write-Log -Level INFO -Message ("Dashboard: " + $safeLine) } catch {}
  }).GetNewClosure()

  $assignThemeRoles = ({
    $setRoleCmd = & $getFunctionCommand 'Set-UiControlRole'
    if (-not $setRoleCmd -or -not $setRoleCmd.ScriptBlock) { return }
    & $setRoleCmd.ScriptBlock -Control $btnRefresh -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnClear -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnOpenSnow -Role 'AccentButton'
    & $setRoleCmd.ScriptBlock -Control $btnTemplateSettings -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnUseCheckInNote -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnUseCheckOutNote -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnCheckIn -Role 'PrimaryButton'
    & $setRoleCmd.ScriptBlock -Control $btnCheckOut -Role 'AccentButton'
    & $setRoleCmd.ScriptBlock -Control $btnCloseCode -Role 'DangerButton'
    & $setRoleCmd.ScriptBlock -Control $lblSearch -Role 'MutedLabel'
    & $setRoleCmd.ScriptBlock -Control $lblHint -Role 'MutedLabel'
    & $setRoleCmd.ScriptBlock -Control $lblStatus -Role 'StatusLabel'
  }).GetNewClosure()

  $txtSearch.Add_KeyDown(({
    param($sender, $e)
    & $runSafeUi 'Dashboard Search KeyDown' {
      if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        if ($state.Controls.SearchTimer) { $state.Controls.SearchTimer.Stop() }
        & $performSearch
      }
    }
  }).GetNewClosure())
  $txtSearch.Add_TextUpdate(({
    & $runSafeUi 'Dashboard Search TextUpdate' {
      & $updateSearchUserSuggestions
      & $scheduleSearch
    }
  }).GetNewClosure())
  $txtSearch.Add_DropDown(({
        & $runSafeUi 'Dashboard Search DropDown' { & $updateSearchUserSuggestions }
      }).GetNewClosure())
  $txtSearch.Add_SelectedIndexChanged(({
        & $runSafeUi 'Dashboard Search SelectionChanged' { & $scheduleSearch }
      }).GetNewClosure())

  $chkOpenOnly.Add_CheckedChanged(({
    & $runSafeUi 'Dashboard OpenOnly Changed' {
      if ($state.AllRows.Count -eq 0) {
        & $performSearch -ReloadFromExcel
        return
      }
      $rows = & $getVisibleRows
      $state.Rows = @($rows)
      & $bindRowsToGrid $rows
      & $updateActionButtons
      if ([string]::IsNullOrWhiteSpace($state.LastSearch)) {
        $filterNote = if ($chkOpenOnly.Checked) { ' (open only)' } else { '' }
        $lblStatus.Text = "Preloaded $($rows.Count) rows from Excel$filterNote."
        return
      }
      $filterNote = if ($chkOpenOnly.Checked) { ' (open only)' } else { '' }
      $lblStatus.Text = "Results: $($rows.Count) for '$($state.LastSearch)'$filterNote"
    }
  }).GetNewClosure())

  $state.UltraFast = $true

  $grid.Add_SelectionChanged(({
        & $runSafeUi 'Dashboard Grid SelectionChanged Buttons' { & $updateActionButtons }
      }).GetNewClosure())
  $grid.Add_SelectionChanged(({
        & $runSafeUi 'Dashboard Grid SelectionChanged Status' { & $refreshSelectionStatus }
      }).GetNewClosure())
  $grid.Add_CellDoubleClick(({
        & $runSafeUi 'Dashboard Grid DoubleClick' { & $applyAction 'checkin' }
      }).GetNewClosure())

  $btnRefresh.Add_Click(({
    & $runSafeUi 'Dashboard Refresh' {
      if ([string]::IsNullOrWhiteSpace(("" + $state.ExcelPath).Trim()) -or -not (Test-Path -LiteralPath $state.ExcelPath)) {
        $lblStatus.Text = 'Excel not ready / Excel empty'
        return
      }
      if ($OnRefreshExcel) {
        $lblStatus.Text = 'Running Force Update Excel...'
        try {
          & $OnRefreshExcel
        }
        catch {
          $err = ("" + $_.Exception.Message).Trim()
          if (-not $err) { $err = 'Force Update failed.' }
          $lblStatus.Text = ("Force Update failed: {0}" -f $err)
          throw
        }
      }
      $state.QueryCache = @{}
      $state.AllRows = @()
      $state.AllRowsUniverse = @()
      $state.ExcelStamp = 0L
      $state.Controls.Search.Text = $state.LastSearch
      & $performSearch -ReloadFromExcel
      $lblStatus.Text = 'Done: dashboard refreshed from Excel.'
      & $appendHistory 'Dashboard refreshed from Excel.'
    }
  }).GetNewClosure())

  $btnClear.Add_Click(({
    & $runSafeUi 'Dashboard Clear' {
      if (-not (& $ensureExcelReady)) { return }
      $state.Controls.Search.Text = ''
      $state.LastSearch = ''
      & $performSearch
    }
  }).GetNewClosure())

  $btnUseCheckInNote.Add_Click(({
    & $runSafeUi 'Dashboard Use Check-In Note' {
      if (-not (& $ensureExcelReady)) { return }
      $row = & $getSelectedRow
      if (-not $row) {
        [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
        return
      }
      $state.Controls.Comment.Text = & $renderTemplateForRow -TemplateText $state.Templates.checkIn -rowItem $row
      & $appendHistory ("Applied Check-In template for {0}" -f ("" + $row.RITM))
    }
  }).GetNewClosure())
  $btnUseCheckOutNote.Add_Click(({
    & $runSafeUi 'Dashboard Use Check-Out Note' {
      if (-not (& $ensureExcelReady)) { return }
      $row = & $getSelectedRow
      if (-not $row) {
        [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
        return
      }
      $state.Controls.Comment.Text = & $renderTemplateForRow -TemplateText $state.Templates.checkOut -rowItem $row
      & $appendHistory ("Applied Check-Out template for {0}" -f ("" + $row.RITM))
    }
  }).GetNewClosure())
  $btnOpenSnow.Add_Click(({
    & $runSafeUi 'Dashboard Open ServiceNow' {
      if (-not (& $ensureExcelReady)) { return }
      $row = & $getSelectedRow
      if (-not $row) {
        [System.Windows.Forms.MessageBox]::Show('Select one row first.', 'Dashboard') | Out-Null
        return
      }
      & $openRowInServiceNow $row
      & $appendHistory ("Opened in ServiceNow: {0}" -f ("" + $row.RITM))
    }
  }).GetNewClosure())
  $btnTemplateSettings.Add_Click(({
        & $runSafeUi 'Dashboard Template Settings' { & $openTemplateManager }
      }).GetNewClosure())
  $btnCheckIn.Add_Click(({
        & $runSafeUi 'Dashboard CheckIn' { & $applyAction 'checkin' }
      }).GetNewClosure())
  $btnCheckOut.Add_Click(({
        & $runSafeUi 'Dashboard CheckOut' { & $applyAction 'checkout' }
      }).GetNewClosure())
  $btnCloseCode.Add_Click(({
    & $runSafeUi 'Dashboard Close All' {
      $shutdownCmd = & $getFunctionCommand 'Shutdown-SchumanApp'
      if ($shutdownCmd) {
        & $shutdownCmd.ScriptBlock -CurrentForm $form
        return
      }
      try { if (Get-Command -Name Close-SchumanUiWindows -ErrorAction SilentlyContinue) { Close-SchumanUiWindows -MainForm $null -GeneratorForm $null | Out-Null } } catch {}
      $r = $null
      if (Get-Command -Name Close-SchumanAllResources -ErrorAction SilentlyContinue) {
        $r = Close-SchumanAllResources -Mode 'All'
      }
      else {
        $fallback = Invoke-UiEmergencyClose -ActionLabel 'Close All' -ExecutableNames @('Code', 'Code - Insiders', 'Cursor', 'WINWORD', 'EXCEL') -Owner $form -Mode 'All'
        $r = @{
          ClosedProcesses = @()
          ClosedDocs      = 0
          Errors          = @()
          Skipped         = 0
        }
        if ($fallback) {
          $r.Errors = if ($fallback.FailedCount -gt 0) { @("Fallback close reported $($fallback.FailedCount) failures.") } else { @() }
          if ($fallback.KilledCount -gt 0) { $r.ClosedProcesses = @('fallback') }
        }
      }
      $closedProcCount = @($r.ClosedProcesses).Count
      $errorsCount = @($r.Errors).Count
      $summary = "Closed: $closedProcCount processes, $($r.ClosedDocs) documents. Skipped: $($r.Skipped). Errors: $errorsCount."
      $lblStatus.Text = $summary
      & $appendHistory $summary
      if ($errorsCount -gt 0) {
        try { Write-Log -Level WARN -Message ("Dashboard Close All errors: " + ($r.Errors -join ' | ')) } catch {}
      }
      try { if ($form -and -not $form.IsDisposed) { $form.Close() } } catch {}
    }
  }).GetNewClosure())
  # Close Documents button removed by design; merged into Close All.

  $form.add_FormClosed(({
    param($sender, $eventArgs)
    try { if ($sender.Tag.Controls.SearchTimer) { $sender.Tag.Controls.SearchTimer.Stop(); $sender.Tag.Controls.SearchTimer.Dispose() } } catch {}
    try {
      if ($sender.Tag.Session -and $sender.Tag.OwnsSession) {
        Close-ServiceNowSession -Session $sender.Tag.Session
      }
    } catch {}
  }).GetNewClosure())

  $state.UserDirectory = @()
  $state.Templates = & $loadTemplates
  $state.DefaultCheckIn = ("" + $state.Templates.checkIn)
  $state.DefaultCheckOut = ("" + $state.Templates.checkOut)
  $state.Controls.Comment.Text = $state.DefaultCheckIn
  $state.AllRows = @()
  $state.Rows = @()
  $state.ExcelReady = $false
  & $bindRowsToGrid @()
  $lblStatus.Text = 'Excel still loading...'
  [void]$form.Add_Shown(({
    & $runSafeUi 'Dashboard Shown' {
      & $layoutDashboard
      & $assignThemeRoles
      $applyThemeToControlTreeHandler = ${function:Apply-ThemeToControlTree}
      $themeVar = Get-Variable -Name CurrentMainTheme -Scope Global -ErrorAction SilentlyContinue
      $scaleVar = Get-Variable -Name CurrentMainFontScale -Scope Global -ErrorAction SilentlyContinue
      $scale = 1.0
      if ($scaleVar -and $scaleVar.Value) { try { $scale = [double]$scaleVar.Value } catch { $scale = 1.0 } }
      if ($applyThemeToControlTreeHandler -and $themeVar -and $themeVar.Value) {
        & $applyThemeToControlTreeHandler -RootControl $form -Theme $themeVar.Value -FontScale $scale
      }
      if (Get-Command -Name Apply-UiRoundedButtonsRecursive -ErrorAction SilentlyContinue) {
        Apply-UiRoundedButtonsRecursive -Root $form -Radius 10
      }
      if ($state.Controls.SearchTimer) { $state.Controls.SearchTimer.Stop() }
      & $preloadRows
      & $updateSearchUserSuggestions
    }
  }).GetNewClosure())
  $form.Add_SizeChanged(({
        & $runSafeUi 'Dashboard Resize' { & $layoutDashboard }
      }).GetNewClosure())

  & $layoutDashboard
  & $updateActionButtons
  return $form
}

