Set-StrictMode -Version Latest

function Write-Log {
  param(
    [string]$Message,
    [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
    [string]$LogPath = ''
  )

  try {
    $text = if ($null -eq $Message) { '' } else { [string]$Message }
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
      $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
      $logsDir = Join-Path $projectRoot 'system\logs'
      if (-not (Test-Path -LiteralPath $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
      }
      $LogPath = Join-Path $logsDir 'ui.log'
    }
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $text
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
  }
  catch {}
}

function Show-UiError {
  param(
    [string]$Context,
    [System.Management.Automation.ErrorRecord]$ErrorRecord
  )

  $ctx = if ([string]::IsNullOrWhiteSpace($Context)) { 'UI' } else { $Context }
  $msg = 'Unknown error.'

  try {
    if ($ErrorRecord -and $ErrorRecord.Exception) {
      $msg = ("" + $ErrorRecord.Exception.Message).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($msg)) { $msg = 'Unknown error.' }
  }
  catch {}

  Write-Log -Level ERROR -Message ("{0}: {1}" -f $ctx, $msg)
  try { [System.Windows.Forms.MessageBox]::Show("$ctx failed.`r`n`r`n$msg", 'Schuman', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
}

function New-UI {
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [string]$TemplatePath = '',
    [string]$OutputPath = '',
    [scriptblock]$OnOpenDashboard = $null,
    [scriptblock]$OnGenerate = $null
  )

  $UI = [hashtable]::Synchronized(@{})
  $UI.ExcelPath = ("" + $ExcelPath).Trim()
  $UI.SheetName = if ([string]::IsNullOrWhiteSpace($SheetName)) { 'BRU' } else { $SheetName }
  $UI.TemplatePath = ("" + $TemplatePath).Trim()
  $UI.OutputPath = ("" + $OutputPath).Trim()
  $UI.OnOpenDashboard = $OnOpenDashboard
  $UI.OnGenerate = $OnGenerate
  $UI.Theme = @{
    Light = @{
      Bg = [System.Drawing.Color]::FromArgb(245, 246, 248)
      Card = [System.Drawing.Color]::FromArgb(255, 255, 255)
      Text = [System.Drawing.Color]::FromArgb(24, 24, 26)
      Sub = [System.Drawing.Color]::FromArgb(110, 110, 115)
      Border = [System.Drawing.Color]::FromArgb(228, 228, 232)
      Accent = [System.Drawing.Color]::FromArgb(0, 122, 255)
      Success = [System.Drawing.Color]::FromArgb(220, 245, 231)
      Error = [System.Drawing.Color]::FromArgb(255, 230, 230)
      BadgeText = [System.Drawing.Color]::FromArgb(26, 26, 28)
      Shadow = [System.Drawing.Color]::FromArgb(236, 236, 240)
      StopBg = [System.Drawing.Color]::FromArgb(255, 235, 235)
      StopFg = [System.Drawing.Color]::FromArgb(170, 30, 30)
      StopBorder = [System.Drawing.Color]::FromArgb(210, 80, 80)
    }
    Dark = @{
      Bg = [System.Drawing.Color]::FromArgb(30, 30, 30)
      Card = [System.Drawing.Color]::FromArgb(37, 37, 38)
      Text = [System.Drawing.Color]::FromArgb(230, 230, 230)
      Sub = [System.Drawing.Color]::FromArgb(170, 170, 170)
      Border = [System.Drawing.Color]::FromArgb(60, 60, 60)
      Accent = [System.Drawing.Color]::FromArgb(10, 132, 255)
      Success = [System.Drawing.Color]::FromArgb(32, 60, 45)
      Error = [System.Drawing.Color]::FromArgb(70, 36, 36)
      BadgeText = [System.Drawing.Color]::FromArgb(235, 235, 235)
      Shadow = [System.Drawing.Color]::FromArgb(28, 28, 32)
      StopBg = [System.Drawing.Color]::FromArgb(88, 36, 36)
      StopFg = [System.Drawing.Color]::FromArgb(255, 210, 210)
      StopBorder = [System.Drawing.Color]::FromArgb(210, 80, 80)
    }
  }
  $UI.GridTable = New-Object System.Data.DataTable 'GenerateStatus'
  [void]$UI.GridTable.Columns.Add('Row', [string])
  [void]$UI.GridTable.Columns.Add('Ticket', [string])
  [void]$UI.GridTable.Columns.Add('User', [string])
  [void]$UI.GridTable.Columns.Add('File', [string])
  [void]$UI.GridTable.Columns.Add('Status', [string])
  [void]$UI.GridTable.Columns.Add('Message', [string])
  [void]$UI.GridTable.Columns.Add('Progress', [string])

  return $UI
}

function Initialize-Controls {
  param([hashtable]$UI)

  $fontName = Get-UiFontName
  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Schuman Word Generator'
  $form.Size = New-Object System.Drawing.Size(1120, 720)
  $form.MinimumSize = New-Object System.Drawing.Size(980, 640)
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.Font = New-Object System.Drawing.Font($fontName, 10)
  $form.Tag = $UI
  $UI.Form = $form

  try {
    $prop = $form.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($prop) { $prop.SetValue($form, $true, $null) }
  }
  catch {}

  $root = New-Object System.Windows.Forms.TableLayoutPanel
  $root.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Padding = New-Object System.Windows.Forms.Padding(16)
  $root.RowCount = 3
  $root.ColumnCount = 1
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$form.Controls.Add($root)
  $UI.Root = $root

  $headerCard = New-Object System.Windows.Forms.Panel
  $headerCard.Dock = [System.Windows.Forms.DockStyle]::Fill
  $headerCard.Padding = New-Object System.Windows.Forms.Padding(16, 16, 16, 12)
  $headerCard.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
  [void]$root.Controls.Add($headerCard, 0, 0)
  $UI.HeaderCard = $headerCard

  $headerGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $headerGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $headerGrid.ColumnCount = 2
  $headerGrid.RowCount = 2
  [void]$headerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
  [void]$headerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
  [void]$headerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$headerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$headerCard.Controls.Add($headerGrid)

  $lblTitle = New-Object System.Windows.Forms.Label
  $lblTitle.Text = 'Schuman Word Generator'
  $lblTitle.Font = New-Object System.Drawing.Font($fontName, 13, [System.Drawing.FontStyle]::Bold)
  $lblTitle.AutoSize = $true
  [void]$headerGrid.Controls.Add($lblTitle, 0, 0)
  $UI.LblTitle = $lblTitle

  $headerActions = New-Object System.Windows.Forms.FlowLayoutPanel
  $headerActions.AutoSize = $true
  $headerActions.WrapContents = $false
  $headerActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $headerActions.Anchor = 'Top,Right'
  [void]$headerGrid.Controls.Add($headerActions, 1, 0)

  $btnDashboard = New-Object System.Windows.Forms.Button
  $btnDashboard.Text = 'Open Check-in Dashboard'
  $btnDashboard.Width = 220
  $btnDashboard.Height = 32
  $btnDashboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnDashboard.Font = New-Object System.Drawing.Font($fontName, 9, [System.Drawing.FontStyle]::Bold)
  [void]$headerActions.Controls.Add($btnDashboard)
  $UI.BtnDashboard = $btnDashboard

  $statusPill = New-Object System.Windows.Forms.Panel
  $statusPill.AutoSize = $true
  $statusPill.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
  $statusPill.Margin = New-Object System.Windows.Forms.Padding(8, 4, 0, 0)
  [void]$headerActions.Controls.Add($statusPill)
  $UI.StatusPill = $statusPill

  $lblStatusPill = New-Object System.Windows.Forms.Label
  $lblStatusPill.Text = 'Idle'
  $lblStatusPill.AutoSize = $true
  [void]$statusPill.Controls.Add($lblStatusPill)
  $UI.LblStatusPill = $lblStatusPill

  $lblMetrics = New-Object System.Windows.Forms.Label
  $lblMetrics.Text = 'Total: 0 | Saved: 0 | Skipped: 0 | Errors: 0'
  $lblMetrics.AutoSize = $true
  $lblMetrics.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
  [void]$headerGrid.Controls.Add($lblMetrics, 0, 1)
  $UI.LblMetrics = $lblMetrics

  $lblStatusText = New-Object System.Windows.Forms.Label
  $lblStatusText.Text = 'Ready.'
  $lblStatusText.AutoSize = $true
  $lblStatusText.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
  $lblStatusText.Anchor = 'Top,Right'
  [void]$headerGrid.Controls.Add($lblStatusText, 1, 1)
  $UI.LblStatusText = $lblStatusText

  $centerGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $centerGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $centerGrid.ColumnCount = 1
  $centerGrid.RowCount = 2
  [void]$centerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20)))
  [void]$centerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.Controls.Add($centerGrid, 0, 1)

  $progressHost = New-Object System.Windows.Forms.Panel
  $progressHost.Height = 20
  $progressHost.Dock = [System.Windows.Forms.DockStyle]::Fill
  $progressHost.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
  [void]$centerGrid.Controls.Add($progressHost, 0, 0)
  $UI.ProgressHost = $progressHost

  $progressFill = New-Object System.Windows.Forms.Panel
  $progressFill.Height = 20
  $progressFill.Width = 0
  $progressFill.Dock = [System.Windows.Forms.DockStyle]::Left
  [void]$progressHost.Controls.Add($progressFill)
  $UI.ProgressFill = $progressFill

  $listCard = New-Object System.Windows.Forms.Panel
  $listCard.Dock = [System.Windows.Forms.DockStyle]::Fill
  $listCard.Padding = New-Object System.Windows.Forms.Padding(15)
  [void]$centerGrid.Controls.Add($listCard, 0, 1)
  $UI.ListCard = $listCard

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $grid.ReadOnly = $true
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AllowUserToResizeRows = $false
  $grid.RowHeadersVisible = $false
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.AutoGenerateColumns = $false
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EnableHeadersVisualStyles = $false
  $grid.ColumnHeadersHeight = 32
  $grid.RowTemplate.Height = 28
  try {
    $prop = $grid.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($prop) { $prop.SetValue($grid, $true, $null) }
  }
  catch {}
  [void]$listCard.Controls.Add($grid)
  $UI.Grid = $grid

  foreach ($columnName in @('Row', 'Ticket', 'User', 'File', 'Status', 'Message', 'Progress')) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $columnName
    $col.DataPropertyName = $columnName
    switch ($columnName) {
      'Row' { $col.HeaderText = 'Row #'; $col.FillWeight = 8 }
      'Ticket' { $col.HeaderText = 'Ticket/RITM'; $col.FillWeight = 16 }
      'User' { $col.HeaderText = 'User'; $col.FillWeight = 16 }
      'File' { $col.HeaderText = 'Output File'; $col.FillWeight = 24 }
      'Status' { $col.HeaderText = 'Status'; $col.FillWeight = 12 }
      'Message' { $col.HeaderText = 'Message'; $col.FillWeight = 24 }
      'Progress' {
        $col.HeaderText = 'Progress'
        $col.FillWeight = 10
        $col.DefaultCellStyle.Alignment = 'MiddleRight'
      }
    }
    [void]$grid.Columns.Add($col)
  }
  $grid.DataSource = $UI.GridTable

  $logPanel = New-Object System.Windows.Forms.Panel
  $logPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $logPanel.Padding = New-Object System.Windows.Forms.Padding(15)
  $logPanel.Height = 0
  $logPanel.Visible = $false
  [void]$listCard.Controls.Add($logPanel)
  $UI.LogPanel = $logPanel

  $logBox = New-Object System.Windows.Forms.RichTextBox
  $logBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $logBox.ReadOnly = $true
  $logBox.HideSelection = $false
  $logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  [void]$logPanel.Controls.Add($logBox)
  $UI.LogBox = $logBox

  $footer = New-Object System.Windows.Forms.Panel
  $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
  $footer.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
  [void]$root.Controls.Add($footer, 0, 2)
  $UI.Footer = $footer

  $footerGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $footerGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $footerGrid.ColumnCount = 2
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$footer.Controls.Add($footerGrid)

  $buttonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $buttonFlow.AutoSize = $true
  $buttonFlow.WrapContents = $false
  $buttonFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  [void]$footerGrid.Controls.Add($buttonFlow, 0, 0)

  $UI.BtnStart = New-Object System.Windows.Forms.Button
  $UI.BtnStart.Text = 'Generate Documents'
  $UI.BtnStart.Width = 170
  $UI.BtnStart.Height = 30
  $UI.BtnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnStart)

  $UI.BtnStop = New-Object System.Windows.Forms.Button
  $UI.BtnStop.Text = 'Stop'
  $UI.BtnStop.Width = 110
  $UI.BtnStop.Height = 30
  $UI.BtnStop.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $UI.BtnStop.Enabled = $false
  [void]$buttonFlow.Controls.Add($UI.BtnStop)

  $UI.BtnOpen = New-Object System.Windows.Forms.Button
  $UI.BtnOpen.Text = 'Open Output Folder'
  $UI.BtnOpen.Width = 170
  $UI.BtnOpen.Height = 30
  $UI.BtnOpen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $UI.BtnOpen.Enabled = $false
  [void]$buttonFlow.Controls.Add($UI.BtnOpen)

  $UI.BtnCloseCode = New-Object System.Windows.Forms.Button
  $UI.BtnCloseCode.Text = 'Cerrar codigo'
  $UI.BtnCloseCode.Width = 120
  $UI.BtnCloseCode.Height = 30
  $UI.BtnCloseCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnCloseCode)

  $UI.BtnCloseDocs = New-Object System.Windows.Forms.Button
  $UI.BtnCloseDocs.Text = 'Cerrar documentos'
  $UI.BtnCloseDocs.Width = 145
  $UI.BtnCloseDocs.Height = 30
  $UI.BtnCloseDocs.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnCloseDocs)

  $UI.BtnToggleLog = New-Object System.Windows.Forms.Button
  $UI.BtnToggleLog.Text = 'Show Log'
  $UI.BtnToggleLog.Width = 110
  $UI.BtnToggleLog.Height = 30
  $UI.BtnToggleLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnToggleLog)

  $optionsFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $optionsFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $optionsFlow.AutoSize = $true
  $optionsFlow.WrapContents = $false
  $optionsFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  [void]$footerGrid.Controls.Add($optionsFlow, 1, 0)

  $UI.ChkShowWord = New-Object System.Windows.Forms.CheckBox
  $UI.ChkShowWord.Text = 'Show Word after generation'
  $UI.ChkShowWord.AutoSize = $true
  [void]$optionsFlow.Controls.Add($UI.ChkShowWord)

  $UI.ChkSaveDocx = New-Object System.Windows.Forms.CheckBox
  $UI.ChkSaveDocx.Text = 'Save DOCX'
  $UI.ChkSaveDocx.AutoSize = $true
  $UI.ChkSaveDocx.Checked = $true
  [void]$optionsFlow.Controls.Add($UI.ChkSaveDocx)

  $UI.ChkSavePdf = New-Object System.Windows.Forms.CheckBox
  $UI.ChkSavePdf.Text = 'Save PDF'
  $UI.ChkSavePdf.AutoSize = $true
  $UI.ChkSavePdf.Checked = $true
  [void]$optionsFlow.Controls.Add($UI.ChkSavePdf)

  $UI.ChkDark = New-Object System.Windows.Forms.CheckBox
  $UI.ChkDark.Text = 'Dark theme'
  $UI.ChkDark.AutoSize = $true
  $UI.ChkDark.Checked = $true
  [void]$optionsFlow.Controls.Add($UI.ChkDark)

  return $UI
}

function Update-Grid {
  param(
    [hashtable]$UI,
    [object[]]$Rows
  )

  if (-not $UI.ContainsKey('Grid') -or -not $UI.Grid) { return }
  if (-not $UI.ContainsKey('GridTable') -or -not $UI.GridTable) { return }

  $rowsSafe = @($Rows)
  $UI.Grid.SuspendLayout()
  try {
    $UI.GridTable.BeginLoadData()
    try {
      $UI.GridTable.Rows.Clear()
      foreach ($item in $rowsSafe) {
        $dr = $UI.GridTable.NewRow()
        $dr['Row'] = if ($null -ne $item.Row) { "" + $item.Row } else { '' }
        $dr['Ticket'] = if ($null -ne $item.Ticket) { "" + $item.Ticket } else { '' }
        $dr['User'] = if ($null -ne $item.User) { "" + $item.User } else { '' }
        $dr['File'] = if ($null -ne $item.File) { "" + $item.File } else { '' }
        $dr['Status'] = if ($null -ne $item.Status) { "" + $item.Status } else { '' }
        $dr['Message'] = if ($null -ne $item.Message) { "" + $item.Message } else { '' }
        $dr['Progress'] = if ($null -ne $item.Progress) { "" + $item.Progress } else { '' }
        [void]$UI.GridTable.Rows.Add($dr)
      }
    }
    finally {
      $UI.GridTable.EndLoadData()
    }
    $UI.Grid.ClearSelection()
  }
  finally {
    $UI.Grid.ResumeLayout()
  }
}

function Load-Data {
  param([hashtable]$UI)

  if (-not $UI.ContainsKey('ExcelPath') -or [string]::IsNullOrWhiteSpace($UI.ExcelPath)) { return @() }
  if (-not (Test-Path -LiteralPath $UI.ExcelPath)) { return @() }
  if (-not (Get-Command -Name Search-DashboardRows -ErrorAction SilentlyContinue)) { return @() }

  $results = New-Object System.Collections.Generic.List[object]
  $rows = @(Search-DashboardRows -ExcelPath $UI.ExcelPath -SheetName $UI.SheetName -SearchText '')
  foreach ($r in $rows) {
    $ticket = ("" + $r.RITM).Trim()
    if (-not $ticket) { continue }
    $status = ("" + $r.DashboardStatus).Trim()
    if (-not $status) { $status = 'Ready' }
    $message = if ($status -eq 'Ready') { 'Preloaded from Excel' } else { 'Preloaded from Excel status' }
    $results.Add([pscustomobject]@{
        Row = ("" + $r.Row)
        Ticket = $ticket
        User = ("" + $r.RequestedFor)
        File = ''
        Status = $status
        Message = $message
        Progress = '0%'
      }) | Out-Null
  }
  return @($results.ToArray())
}

function Export-Excel {
  param([hashtable]$UI)
  return @(Load-Data -UI $UI)
}

function Generate-PDF {
  param([hashtable]$UI)

  if (-not $UI.ContainsKey('OnGenerate') -or -not $UI.OnGenerate) {
    throw 'Generate callback not configured.'
  }

  $argsObj = [pscustomobject]@{
    ExcelPath = $UI.ExcelPath
    TemplatePath = $UI.TemplatePath
    OutputPath = $UI.OutputPath
    ExportPdf = [bool]$UI.ChkSavePdf.Checked
    SaveDocx = [bool]$UI.ChkSaveDocx.Checked
    ShowWord = [bool]$UI.ChkShowWord.Checked
  }

  if (-not (Test-Path -LiteralPath $argsObj.ExcelPath)) { throw 'Excel file not found.' }
  if (-not (Test-Path -LiteralPath $argsObj.TemplatePath)) { throw 'Template file not found.' }
  if ([string]::IsNullOrWhiteSpace($argsObj.OutputPath)) { throw 'Output folder is required.' }

  return (& $UI.OnGenerate $argsObj)
}

function Set-GenerateUiTheme {
  param([hashtable]$UI)

  $palette = if ($UI.ChkDark.Checked) { $UI.Theme.Dark } else { $UI.Theme.Light }
  $UI.Form.BackColor = $palette.Bg
  $UI.Form.ForeColor = $palette.Text
  $UI.HeaderCard.BackColor = $palette.Card
  $UI.ListCard.BackColor = $palette.Card
  $UI.Footer.BackColor = $palette.Card
  $UI.LogPanel.BackColor = $palette.Card
  $UI.ProgressHost.BackColor = $palette.Border
  $UI.ProgressFill.BackColor = $palette.Accent
  $UI.LblTitle.ForeColor = $palette.Text
  $UI.LblMetrics.ForeColor = $palette.Sub
  $UI.LblStatusText.ForeColor = $palette.Sub
  $UI.LogBox.BackColor = $palette.Bg
  $UI.LogBox.ForeColor = $palette.Text

  foreach ($btn in @($UI.BtnDashboard, $UI.BtnStart, $UI.BtnStop, $UI.BtnOpen, $UI.BtnCloseCode, $UI.BtnCloseDocs, $UI.BtnToggleLog)) {
    $btn.BackColor = $palette.Card
    $btn.ForeColor = $palette.Text
    $btn.FlatAppearance.BorderColor = $palette.Border
    $btn.FlatAppearance.BorderSize = 1
  }
  $UI.BtnDashboard.BackColor = $palette.Bg
  $UI.BtnDashboard.ForeColor = $palette.Accent
  $UI.BtnDashboard.FlatAppearance.BorderColor = $palette.Accent
  $UI.BtnDashboard.FlatAppearance.BorderSize = 2
  $UI.BtnStart.BackColor = $palette.Accent
  $UI.BtnStart.ForeColor = $palette.BadgeText
  $UI.BtnStart.FlatAppearance.BorderColor = $palette.Accent
  $UI.BtnStop.BackColor = $palette.StopBg
  $UI.BtnStop.ForeColor = $palette.StopFg
  $UI.BtnStop.FlatAppearance.BorderColor = $palette.StopBorder

  foreach ($chk in @($UI.ChkDark, $UI.ChkSavePdf, $UI.ChkSaveDocx, $UI.ChkShowWord)) {
    $chk.BackColor = $palette.Card
    $chk.ForeColor = $palette.Sub
  }

  $UI.Grid.BackgroundColor = $palette.Bg
  $UI.Grid.GridColor = $palette.Border
  $UI.Grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UI.Grid.ColumnHeadersDefaultCellStyle.BackColor = $palette.Card
  $UI.Grid.ColumnHeadersDefaultCellStyle.ForeColor = $palette.Sub
  $UI.Grid.DefaultCellStyle.BackColor = $palette.Bg
  $UI.Grid.DefaultCellStyle.ForeColor = $palette.Text
  $UI.Grid.DefaultCellStyle.SelectionBackColor = $palette.Shadow
  $UI.Grid.DefaultCellStyle.SelectionForeColor = $palette.Text
  $UI.Grid.AlternatingRowsDefaultCellStyle.BackColor = $palette.Card
  $UI.Grid.AlternatingRowsDefaultCellStyle.ForeColor = $palette.Text
}

function Set-StatusPill {
  param(
    [hashtable]$UI,
    [string]$Text,
    [ValidateSet('idle', 'running', 'done', 'error')]$State = 'idle'
  )

  $palette = if ($UI.ChkDark.Checked) { $UI.Theme.Dark } else { $UI.Theme.Light }
  $UI.LblStatusPill.Text = $Text
  switch ($State) {
    'running' { $UI.StatusPill.BackColor = $palette.Accent }
    'done' { $UI.StatusPill.BackColor = $palette.Success }
    'error' { $UI.StatusPill.BackColor = $palette.Error }
    default { $UI.StatusPill.BackColor = $palette.Border }
  }
  $UI.LblStatusPill.ForeColor = $palette.BadgeText
}

function Append-GenerateLog {
  param([hashtable]$UI, [string]$Line)
  $text = if ($null -eq $Line) { '' } else { [string]$Line }
  $formatted = "[{0}] {1}{2}" -f (Get-Date -Format 'HH:mm:ss'), $text, [Environment]::NewLine
  try { $UI.LogBox.AppendText($formatted) } catch {}
  Write-Log -Level INFO -Message $text
}

function Update-OutputButton {
  param([hashtable]$UI)
  $out = ("" + $UI.OutputPath).Trim()
  $UI.BtnOpen.Enabled = (-not [string]::IsNullOrWhiteSpace($out)) -and (Test-Path -LiteralPath $out)
}

function Invoke-UiSafe {
  param(
    [hashtable]$UI,
    [string]$Context,
    [scriptblock]$Action
  )
  try { & $Action }
  catch { Show-UiError -Context $Context -ErrorRecord $_ }
}

function Register-GenerateHandlers {
  param([hashtable]$UI)

  $UI.BtnDashboard.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Open Dashboard' -Action {
      if ($UI.OnOpenDashboard) { & $UI.OnOpenDashboard; return }
      $UI.LblStatusText.Text = 'Dashboard callback not configured.'
      Set-StatusPill -UI $UI -Text 'Error' -State error
    }
  }).GetNewClosure())

  $UI.BtnOpen.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Open Output Folder' -Action {
      $out = ("" + $UI.OutputPath).Trim()
      if (-not $out) { return }
      if (-not (Test-Path -LiteralPath $out)) { return }
      Start-Process -FilePath $out | Out-Null
    }
  }).GetNewClosure())

  $UI.BtnCloseCode.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Cerrar codigo' -Action {
      $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar codigo' -ExecutableNames @('excel.exe', 'powershell.exe', 'pwsh.exe') -Owner $UI.Form -Mode 'Documents' -MainForm $UI.Form
      if (-not $r.Cancelled) {
        $UI.LblStatusText.Text = $r.Message
        Append-GenerateLog -UI $UI -Line $r.Message
      }
    }
  }).GetNewClosure())

  $UI.BtnCloseDocs.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Cerrar documentos' -Action {
      $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar documentos' -ExecutableNames @('winword.exe', 'excel.exe') -Owner $UI.Form
      if (-not $r.Cancelled) {
        $UI.LblStatusText.Text = $r.Message
        Append-GenerateLog -UI $UI -Line $r.Message
      }
    }
  }).GetNewClosure())

  $UI.BtnToggleLog.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Toggle Log' -Action {
      $UI.LogPanel.Visible = -not $UI.LogPanel.Visible
      if ($UI.LogPanel.Visible) {
        $UI.LogPanel.Height = 150
        $UI.BtnToggleLog.Text = 'Hide Log'
      }
      else {
        $UI.LogPanel.Height = 0
        $UI.BtnToggleLog.Text = 'Show Log'
      }
    }
  }).GetNewClosure())

  $UI.BtnStop.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Stop' -Action {
      [System.Windows.Forms.MessageBox]::Show('Stop is not available in this integrated mode.', 'Info') | Out-Null
    }
  }).GetNewClosure())

  $UI.BtnStart.Add_Click(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Generate Documents' -Action {
      $UI.BtnStart.Enabled = $false
      $UI.BtnDashboard.Enabled = $false
      try {
        $UI.LblStatusText.Text = 'Generating documents...'
        Set-StatusPill -UI $UI -Text 'Running' -State running
        Append-GenerateLog -UI $UI -Line 'Starting document generation.'

        $result = Generate-PDF -UI $UI
        if ($result -and $result.ok -eq $true) {
          $UI.LblStatusText.Text = 'Generation complete.'
          Set-StatusPill -UI $UI -Text 'Done' -State done
          $msg = if ($result.message) { "" + $result.message } else { 'Documents generated successfully.' }
          Append-GenerateLog -UI $UI -Line $msg
          if ($result.outputPath) { $UI.OutputPath = ("" + $result.outputPath).Trim() }
          Update-OutputButton -UI $UI
          Update-Grid -UI $UI -Rows @([pscustomobject]@{
              Row = ''
              Ticket = ''
              User = ''
              File = $UI.OutputPath
              Status = 'Done'
              Message = $msg
              Progress = '100%'
            })
          $UI.LblMetrics.Text = 'Total: 1 | Saved: 1 | Skipped: 0 | Errors: 0'
        }
        else {
          $msg = if ($result -and $result.message) { "" + $result.message } else { 'Unknown error.' }
          $UI.LblStatusText.Text = 'Generation failed.'
          Set-StatusPill -UI $UI -Text 'Error' -State error
          Append-GenerateLog -UI $UI -Line $msg
          $UI.LblMetrics.Text = 'Total: 1 | Saved: 0 | Skipped: 0 | Errors: 1'
        }
      }
      catch {
        $UI.LblStatusText.Text = 'Generation failed.'
        Set-StatusPill -UI $UI -Text 'Error' -State error
        Append-GenerateLog -UI $UI -Line ("ERROR: " + $_.Exception.Message)
        $UI.LblMetrics.Text = 'Total: 1 | Saved: 0 | Skipped: 0 | Errors: 1'
        Show-UiError -Context 'Generate-PDF' -ErrorRecord $_
      }
      finally {
        $UI.BtnStart.Enabled = $true
        $UI.BtnDashboard.Enabled = $true
      }
    }
  }).GetNewClosure())

  $UI.ChkDark.Add_CheckedChanged(({
    param($sender, $args)
    Invoke-UiSafe -UI $UI -Context 'Apply Theme' -Action {
      Set-GenerateUiTheme -UI $UI
      Set-StatusPill -UI $UI -Text $UI.LblStatusPill.Text -State idle
    }
  }).GetNewClosure())
}

function New-GeneratePdfUI {
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [string]$TemplatePath = '',
    [string]$OutputPath = '',
    [scriptblock]$OnOpenDashboard = $null,
    [scriptblock]$OnGenerate = $null
  )

  $UI = New-UI -ExcelPath $ExcelPath -SheetName $SheetName -TemplatePath $TemplatePath -OutputPath $OutputPath -OnOpenDashboard $OnOpenDashboard -OnGenerate $OnGenerate
  $null = Initialize-Controls -UI $UI
  Register-GenerateHandlers -UI $UI
  Set-GenerateUiTheme -UI $UI
  Set-StatusPill -UI $UI -Text 'Idle' -State idle
  Update-OutputButton -UI $UI

  try {
    $rows = @(Export-Excel -UI $UI)
    Update-Grid -UI $UI -Rows $rows
    $total = $rows.Count
    $UI.LblMetrics.Text = "Total: $total | Saved: 0 | Skipped: 0 | Errors: 0"
    if ($total -gt 0) {
      $UI.LblStatusText.Text = "Ready. Preloaded $total rows from Excel."
      Append-GenerateLog -UI $UI -Line "Preloaded $total rows from Excel."
    }
  }
  catch {
    Show-UiError -Context 'Load-Data' -ErrorRecord $_
  }

  return $UI.Form
}
