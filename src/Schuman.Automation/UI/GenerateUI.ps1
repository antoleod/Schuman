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
  $globalShowUiError = Get-Command -Name global:Show-UiError -CommandType Function -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($globalShowUiError -and $globalShowUiError.ScriptBlock) {
    try {
      if ($ErrorRecord -and $ErrorRecord.Exception) {
        & $globalShowUiError.ScriptBlock -Title 'Schuman' -Message ("{0} failed." -f $ctx) -Exception $ErrorRecord.Exception
      }
      else {
        & $globalShowUiError.ScriptBlock -Title 'Schuman' -Message ("{0} failed." -f $ctx)
      }
      return
    }
    catch {}
  }
  $msg = if ($ErrorRecord -and $ErrorRecord.Exception) { ("" + $ErrorRecord.Exception.Message).Trim() } else { "$ctx failed." }
  if (-not $msg) { $msg = "$ctx failed." }
  $stack = ''
  try { if ($ErrorRecord) { $stack = ("" + $ErrorRecord.ScriptStackTrace).Trim() } } catch {}
  if ($stack) {
    Write-Log -Level ERROR -Message ("{0}: {1} | {2}" -f $ctx, $msg, $stack)
  }
  else {
    Write-Log -Level ERROR -Message ("{0}: {1}" -f $ctx, $msg)
  }
}

function New-UI {
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [string]$TemplatePath = '',
    [string]$OutputPath = '',
    [scriptblock]$OnOpenDashboard = $null,
    [scriptblock]$OnGenerate = $null,
    [scriptblock]$OnCloseAll = $null
  )

  $UI = [hashtable]::Synchronized(@{})
  $UI.ExcelPath = ("" + $ExcelPath).Trim()
  $UI.SheetName = if ([string]::IsNullOrWhiteSpace($SheetName)) { 'BRU' } else { $SheetName }
  $UI.TemplatePath = ("" + $TemplatePath).Trim()
  $UI.OutputPath = ("" + $OutputPath).Trim()
  $UI.OnOpenDashboard = $OnOpenDashboard
  $UI.OnGenerate = $OnGenerate
  $UI.OnCloseAll = $OnCloseAll
  $UI.SelectAllSyncing = $false
  $UI.BulkSelectToggle = $false
  $UI.ExcelReady = $false
  $UI.Theme = @{
    Dark = @{
      Bg = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
      Card = [System.Drawing.ColorTranslator]::FromHtml('#1E293B')
      Input = [System.Drawing.ColorTranslator]::FromHtml('#0B1220')
      Text = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
      Sub = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
      Border = [System.Drawing.ColorTranslator]::FromHtml('#334155')
      Accent = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
      AccentHover = [System.Drawing.ColorTranslator]::FromHtml('#3B82F6')
      AccentPressed = [System.Drawing.ColorTranslator]::FromHtml('#1D4ED8')
      Success = [System.Drawing.ColorTranslator]::FromHtml('#22C55E')
      Error = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
      BadgeText = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
      GridAlt = [System.Drawing.ColorTranslator]::FromHtml('#172033')
      Selection = [System.Drawing.ColorTranslator]::FromHtml('#1E40AF')
      StopBg = [System.Drawing.ColorTranslator]::FromHtml('#7F1D1D')
      StopFg = [System.Drawing.ColorTranslator]::FromHtml('#FFFFFF')
      StopBorder = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
    }
  }
  $UI.GridTable = New-Object System.Data.DataTable 'GenerateStatus'
  [void]$UI.GridTable.Columns.Add('Generate', [bool])
  [void]$UI.GridTable.Columns.Add('Row', [string])
  [void]$UI.GridTable.Columns.Add('Ticket', [string])
  [void]$UI.GridTable.Columns.Add('User', [string])
  [void]$UI.GridTable.Columns.Add('PI', [string])
  [void]$UI.GridTable.Columns.Add('File', [string])
  [void]$UI.GridTable.Columns.Add('Status', [string])
  [void]$UI.GridTable.Columns.Add('Message', [string])
  [void]$UI.GridTable.Columns.Add('Progress', [string])

  return $UI
}

function Initialize-Controls {
  <#
  .SYNOPSIS
  Builds the Word Generator form controls and layout.
  .DESCRIPTION
  Creates a responsive WinForms layout with grid, progress, action buttons, and options.
  Control arrangement is optimized to avoid overlap on resize.
  .PARAMETER UI
  Shared synchronized UI state hashtable.
  .OUTPUTS
  hashtable
  #>
  param([hashtable]$UI)

  $fontName = Get-UiFontName
  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Schuman Word Generator'
  $form.Size = New-Object System.Drawing.Size(1120, 720)
  $form.MinimumSize = New-Object System.Drawing.Size(980, 640)
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.Font = New-Object System.Drawing.Font($fontName, 11)
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
  $grid.ReadOnly = $false
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AllowUserToResizeRows = $false
  $grid.RowHeadersVisible = $false
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.AutoGenerateColumns = $false
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EnableHeadersVisualStyles = $false
  $grid.ColumnHeadersVisible = $true
  $grid.ColumnHeadersHeight = 36
  $grid.RowTemplate.Height = 32
  try {
    $prop = $grid.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($prop) { $prop.SetValue($grid, $true, $null) }
  }
  catch {}
  [void]$listCard.Controls.Add($grid)
  $UI.Grid = $grid

  $colGenerate = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
  $colGenerate.Name = 'Generate'
  $colGenerate.DataPropertyName = 'Generate'
  $colGenerate.HeaderText = 'Generate'
  $colGenerate.FillWeight = 8
  [void]$grid.Columns.Add($colGenerate)

  foreach ($columnName in @('Row', 'Ticket', 'User', 'PI', 'File', 'Status', 'Message', 'Progress')) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $columnName
    $col.DataPropertyName = $columnName
    $col.ReadOnly = $true
    switch ($columnName) {
      'Row' { $col.HeaderText = 'Row #'; $col.FillWeight = 8 }
      'Ticket' { $col.HeaderText = 'Ticket/RITM'; $col.FillWeight = 16 }
      'User' { $col.HeaderText = 'User'; $col.FillWeight = 16 }
      'PI' { $col.HeaderText = 'PI'; $col.FillWeight = 14 }
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
  $footerGrid.ColumnCount = 3
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$footer.Controls.Add($footerGrid)

  $buttonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $buttonFlow.AutoSize = $true
  $buttonFlow.WrapContents = $true
  $buttonFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  [void]$footerGrid.Controls.Add($buttonFlow, 0, 0)

  $UI.BtnStart = New-Object System.Windows.Forms.Button
  $UI.BtnStart.Text = 'Generate Documents'
  $UI.BtnStart.Width = 170
  $UI.BtnStart.Height = 30
  $UI.BtnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnStart)

  $UI.ChkSelectAll = New-Object System.Windows.Forms.CheckBox
  $UI.ChkSelectAll.Text = 'Select all'
  $UI.ChkSelectAll.AutoSize = $true
  $UI.ChkSelectAll.ThreeState = $true
  $UI.ChkSelectAll.CheckState = [System.Windows.Forms.CheckState]::Checked
  $UI.ChkSelectAll.Margin = New-Object System.Windows.Forms.Padding(10, 6, 8, 0)
  [void]$buttonFlow.Controls.Add($UI.ChkSelectAll)

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

  $closeFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $closeFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $closeFlow.AutoSize = $true
  $closeFlow.WrapContents = $false
  $closeFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  [void]$footerGrid.Controls.Add($closeFlow, 1, 0)

  $UI.BtnCloseCode = New-Object System.Windows.Forms.Button
  $UI.BtnCloseCode.Text = 'Close All'
  $UI.BtnCloseCode.Width = 130
  $UI.BtnCloseCode.Height = 30
  $UI.BtnCloseCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$closeFlow.Controls.Add($UI.BtnCloseCode)

  $UI.BtnToggleLog = New-Object System.Windows.Forms.Button
  $UI.BtnToggleLog.Text = 'Show Log'
  $UI.BtnToggleLog.Width = 110
  $UI.BtnToggleLog.Height = 30
  $UI.BtnToggleLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($UI.BtnToggleLog)

  $optionsFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $optionsFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $optionsFlow.AutoSize = $true
  $optionsFlow.WrapContents = $true
  $optionsFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  [void]$footerGrid.Controls.Add($optionsFlow, 2, 0)

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

  return $UI
}

function Ensure-GenerateGridHeaders {
  param([hashtable]$UI)

  if (-not $UI -or -not $UI.Grid) { return }
  $UI.Grid.AutoGenerateColumns = $false
  $UI.Grid.ColumnHeadersVisible = $true
  $UI.Grid.EnableHeadersVisualStyles = $false
  $UI.Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::EnableResizing
  if ($UI.Grid.ColumnHeadersHeight -lt 30) { $UI.Grid.ColumnHeadersHeight = 36 }
  $UI.Grid.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
  $UI.Grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft

  $requiredHeaders = [ordered]@{
    Generate = 'Generate'
    Row = 'Row #'
    Ticket = 'Ticket/RITM'
    User = 'User'
    File = 'Output File'
    Status = 'Status'
    Message = 'Message'
    Progress = 'Progress'
  }
  $optionalHeaders = [ordered]@{
    PI = 'PI'
  }

  foreach ($columnName in $requiredHeaders.Keys) {
    try {
      $col = $UI.Grid.Columns[$columnName]
      if (-not $col) {
        if ($columnName -eq 'Generate') {
          $col = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
          $col.Name = 'Generate'
          $col.DataPropertyName = 'Generate'
        }
        else {
          $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
          $col.Name = $columnName
          $col.DataPropertyName = $columnName
          $col.ReadOnly = $true
        }
        [void]$UI.Grid.Columns.Add($col)
      }
      $col.HeaderText = $requiredHeaders[$columnName]
      $col.Visible = $true
    } catch {}
  }

  foreach ($columnName in $optionalHeaders.Keys) {
    try {
      $col = $UI.Grid.Columns[$columnName]
      if ($col) {
        $col.HeaderText = $optionalHeaders[$columnName]
        $col.Visible = $true
      }
    } catch {}
  }

  $displayOrder = @('Generate', 'Row', 'Ticket', 'User', 'PI', 'File', 'Status', 'Message', 'Progress')
  for ($i = 0; $i -lt $displayOrder.Count; $i++) {
    $name = $displayOrder[$i]
    try {
      $col = $UI.Grid.Columns[$name]
      if ($col -and $col.Visible) { $col.DisplayIndex = $i }
    } catch {}
  }
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
    Ensure-GenerateGridHeaders -UI $UI
    $UI.GridTable.BeginLoadData()
    try {
      $UI.GridTable.Rows.Clear()
      foreach ($item in $rowsSafe) {
        $dr = $UI.GridTable.NewRow()
        $dr['Generate'] = $true
        $dr['Row'] = if ($null -ne $item.Row) { "" + $item.Row } else { '' }
        $dr['Ticket'] = if ($null -ne $item.Ticket) { "" + $item.Ticket } else { '' }
        $dr['User'] = if ($null -ne $item.User) { "" + $item.User } else { '' }
        $dr['PI'] = if ($null -ne $item.PI) { "" + $item.PI } else { '' }
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
    try { Update-GenerateSelectAllState -UI $UI } catch {}
  }
  finally {
    $UI.Grid.ResumeLayout()
  }
}

function Get-GenerateSafeFileName {
  param([string]$Text)
  $name = ("" + $Text).Trim()
  if (-not $name) { $name = 'Unknown' }
  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($ch in $invalid) { $name = $name.Replace($ch, '_') }
  $name = [System.Text.RegularExpressions.Regex]::Replace($name, '\s+', ' ').Trim()
  if ($name.Length -gt 80) { $name = $name.Substring(0, 80).Trim() }
  if (-not $name) { $name = 'Unknown' }
  return $name
}

function Get-GenerateExpectedPaths {
  param(
    [hashtable]$UI,
    [string]$Ticket,
    [string]$User
  )
  $outDir = ("" + $UI.OutputPath).Trim()
  if (-not $outDir) { return [pscustomobject]@{ Docx=''; Pdf='' } }
  $ticketValue = ("" + $Ticket).Trim()
  if (-not $ticketValue) { $ticketValue = 'UNKNOWN_TICKET' }
  $nameSource = ("" + $User).Trim()
  if (-not $nameSource) { $nameSource = $ticketValue }
  $safeName = Get-GenerateSafeFileName -Text $nameSource
  $baseFile = "{0}_{1}" -f $ticketValue, $safeName
  return [pscustomobject]@{
    Docx = [System.IO.Path]::GetFullPath((Join-Path $outDir ("$baseFile.docx")))
    Pdf = [System.IO.Path]::GetFullPath((Join-Path $outDir ("$baseFile.pdf")))
  }
}

function Test-GenerateRowAlreadyExported {
  param(
    [hashtable]$UI,
    [string]$Ticket,
    [string]$User
  )
  $paths = Get-GenerateExpectedPaths -UI $UI -Ticket $Ticket -User $User
  $docxExists = $false
  $pdfExists = $false
  try { if ($paths.Docx) { $docxExists = Test-Path -LiteralPath $paths.Docx } } catch {}
  try { if ($paths.Pdf) { $pdfExists = Test-Path -LiteralPath $paths.Pdf } } catch {}
  $exists = ($docxExists -or $pdfExists)
  return [pscustomobject]@{
    Exists = [bool]$exists
    DocxExists = [bool]$docxExists
    PdfExists = [bool]$pdfExists
    Paths = $paths
  }
}

function Resolve-GenerateExcelPath {
  param([hashtable]$UI)

  $candidate = ''
  try { $candidate = ("" + $UI.ExcelPath).Trim() } catch { $candidate = '' }
  if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }

  $projectRoot = ''
  try { $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) } catch { $projectRoot = '' }
  if ($projectRoot) {
    $prefPath = Join-Path $projectRoot 'system\db\ui-preferences.json'
    if (Test-Path -LiteralPath $prefPath) {
      try {
        $json = Get-Content -LiteralPath $prefPath -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($json -and $json.PSObject.Properties['excelPath']) {
          $prefExcel = ("" + $json.excelPath).Trim()
          if ($prefExcel -and (Test-Path -LiteralPath $prefExcel)) {
            return $prefExcel
          }
        }
      }
      catch {}
    }

    $defaultExcel = Join-Path $projectRoot 'Schuman List.xlsx'
    if (Test-Path -LiteralPath $defaultExcel) { return $defaultExcel }
  }

  return ''
}

function Load-Data {
  param([hashtable]$UI)

  $resolvedExcel = Resolve-GenerateExcelPath -UI $UI
  if (-not $resolvedExcel) { return @() }
  $UI.ExcelPath = $resolvedExcel

  $results = New-Object System.Collections.Generic.List[object]
  $rows = @()
  $docReader = Get-Command -Name Read-DocumentRowsFromExcel -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($docReader -and $docReader.ScriptBlock) {
    try {
      $rows = @(Read-DocumentRowsFromExcel -ExcelPath $UI.ExcelPath -SheetName $UI.SheetName)
    }
    catch {
      Write-Log -Level ERROR -Message ("Load-Data/Read-DocumentRowsFromExcel failed: " + $_.Exception.Message)
      $rows = @()
    }
  }
  if (@($rows).Count -eq 0) {
    try {
      $rows = @(Search-DashboardRows -ExcelPath $UI.ExcelPath -SheetName $UI.SheetName -SearchText '')
    }
    catch {
      Write-Log -Level ERROR -Message ("Load-Data/Search-DashboardRows failed: " + $_.Exception.Message)
      throw
    }
  }
  foreach ($r in $rows) {
    $ticket = ("" + $r.RITM).Trim()
    if (-not $ticket) { $ticket = 'UNKNOWN_TICKET' }
    $user = ("" + $r.RequestedFor).Trim()
    $pi = ("" + $r.PI).Trim()
    if (-not $pi) { $pi = '-' }
    $status = ''
    $statusProp = $r.PSObject.Properties['DashboardStatus']
    if ($statusProp) { $status = ("" + $statusProp.Value).Trim() }
    if (-not $status) { $status = 'Ready' }
    $message = if ($status -eq 'Ready') { 'Preloaded from Excel' } else { 'Preloaded from Excel status' }
    $fileText = ("PI: {0}" -f $pi)
    $progressText = '0%'
    $checkDefault = $true

    $existing = Test-GenerateRowAlreadyExported -UI $UI -Ticket $ticket -User $user
    if ($existing.Exists) {
      $status = 'Done'
      $message = 'Already exported in WORD files'
      $progressText = '100%'
      $checkDefault = $false
      $parts = @()
      if ($existing.DocxExists) { $parts += 'DOCX' }
      if ($existing.PdfExists) { $parts += 'PDF' }
      $fileText = (($parts -join '+') + " | " + [System.IO.Path]::GetFileName($existing.Paths.Docx))
    }

    $results.Add([pscustomobject]@{
        Generate = $checkDefault
        Row = ("" + $r.Row)
        Ticket = $ticket
        User = $user
        PI = $pi
        File = $fileText
        Status = $status
        Message = $message
        Progress = $progressText
      }) | Out-Null
  }
  $excelText = ("" + $UI.ExcelPath).Trim()
  $resultCount = [int]$results.Count
  Write-Log -Level INFO -Message ("Load-Data completed. Excel='{0}', rows={1}" -f $excelText, $resultCount)
  return @($results.ToArray())
}

function Export-Excel {
  param([hashtable]$UI)
  return @(Load-Data -UI $UI)
}

function Generate-PDF {
  param(
    [hashtable]$UI,
    [int[]]$SelectedRowNumbers = @()
  )

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
    RowNumbers = @($SelectedRowNumbers)
  }

  if (-not (Test-Path -LiteralPath $argsObj.ExcelPath)) { throw 'Excel file not found.' }
  if (-not (Test-Path -LiteralPath $argsObj.TemplatePath)) { throw 'Template file not found.' }
  if ([string]::IsNullOrWhiteSpace($argsObj.OutputPath)) { throw 'Output folder is required.' }

  return (& $UI.OnGenerate $argsObj)
}

function Set-GenerateUiTheme {
  <#
  .SYNOPSIS
  Applies the Midnight theme to Generator controls.
  .DESCRIPTION
  Styles form, grid, buttons, and status controls with readable high-contrast colors.
  .PARAMETER UI
  Shared synchronized UI state hashtable.
  .NOTES
  Must run on UI thread.
  #>
  param([hashtable]$UI)

  $globalTheme = $null
  $globalScale = 1.0
  try {
    $themeVar = Get-Variable -Name CurrentMainTheme -Scope Global -ErrorAction SilentlyContinue
    if ($themeVar -and $themeVar.Value) { $globalTheme = $themeVar.Value }
    $scaleVar = Get-Variable -Name CurrentMainFontScale -Scope Global -ErrorAction SilentlyContinue
    if ($scaleVar -and $scaleVar.Value) { $globalScale = [double]$scaleVar.Value }
  } catch {}

  $palette = $UI.Theme.Dark
  $setRoleCmd = Get-Command -Name Set-UiControlRole -CommandType Function -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($setRoleCmd -and $setRoleCmd.ScriptBlock) {
    & $setRoleCmd.ScriptBlock -Control $UI.BtnStart -Role 'PrimaryButton'
    & $setRoleCmd.ScriptBlock -Control $UI.BtnDashboard -Role 'AccentButton'
    & $setRoleCmd.ScriptBlock -Control $UI.BtnOpen -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $UI.BtnStop -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $UI.BtnToggleLog -Role 'SecondaryButton'
    & $setRoleCmd.ScriptBlock -Control $UI.BtnCloseCode -Role 'DangerButton'
    & $setRoleCmd.ScriptBlock -Control $UI.LblMetrics -Role 'MutedLabel'
    & $setRoleCmd.ScriptBlock -Control $UI.LblStatusText -Role 'StatusLabel'
  }

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
  $UI.LogBox.BackColor = $palette.Input
  $UI.LogBox.ForeColor = $palette.Text

  foreach ($btn in @($UI.BtnDashboard, $UI.BtnStart, $UI.BtnStop, $UI.BtnOpen, $UI.BtnCloseCode, $UI.BtnToggleLog)) {
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
  $UI.BtnStart.FlatAppearance.MouseOverBackColor = $palette.AccentHover
  $UI.BtnStart.FlatAppearance.MouseDownBackColor = $palette.AccentPressed
  $UI.BtnStop.BackColor = $palette.StopBg
  $UI.BtnStop.ForeColor = $palette.StopFg
  $UI.BtnStop.FlatAppearance.BorderColor = $palette.StopBorder

  foreach ($chk in @($UI.ChkSelectAll, $UI.ChkSavePdf, $UI.ChkSaveDocx, $UI.ChkShowWord)) {
    $chk.BackColor = $palette.Card
    $chk.ForeColor = $palette.Text
  }

  $UI.Grid.BackgroundColor = $palette.Card
  $UI.Grid.GridColor = $palette.Border
  $UI.Grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $UI.Grid.EnableHeadersVisualStyles = $false
  $UI.Grid.RowHeadersVisible = $false
  $UI.Grid.ColumnHeadersDefaultCellStyle.BackColor = $palette.Card
  $UI.Grid.ColumnHeadersDefaultCellStyle.ForeColor = $palette.Sub
  $UI.Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $palette.Card
  $UI.Grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = $palette.Sub
  $UI.Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::EnableResizing
  if ($UI.Grid.ColumnHeadersHeight -lt 30) { $UI.Grid.ColumnHeadersHeight = 36 }
  $UI.Grid.DefaultCellStyle.BackColor = $palette.Input
  $UI.Grid.DefaultCellStyle.ForeColor = $palette.Text
  $UI.Grid.DefaultCellStyle.SelectionBackColor = $palette.Selection
  $UI.Grid.DefaultCellStyle.SelectionForeColor = $palette.BadgeText
  $UI.Grid.AlternatingRowsDefaultCellStyle.BackColor = $palette.GridAlt
  $UI.Grid.AlternatingRowsDefaultCellStyle.ForeColor = $palette.Text

  if ($globalTheme -and (Get-Command -Name Apply-ThemeToControlTree -ErrorAction SilentlyContinue)) {
    try { Apply-ThemeToControlTree -RootControl $UI.Form -Theme $globalTheme -FontScale $globalScale } catch {}
  }
  if (Get-Command -Name Apply-UiRoundedButtonsRecursive -ErrorAction SilentlyContinue) {
    try { Apply-UiRoundedButtonsRecursive -Root $UI.Form -Radius 10 } catch {}
  }
  Ensure-GenerateGridHeaders -UI $UI
}

function Update-GenerateDataState {
  param([hashtable]$UI,[bool]$ExcelReady,[string]$Reason = '')
  $UI.ExcelReady = [bool]$ExcelReady
  $UI.BtnStart.Enabled = [bool]$ExcelReady
  $UI.BtnDashboard.Enabled = $true
  if (-not $ExcelReady) {
    $msg = ("" + $Reason).Trim()
    if (-not $msg) { $msg = 'Excel still loading...' }
    $UI.LblStatusText.Text = $msg
    Set-StatusPill -UI $UI -Text 'Missing Excel' -State error
  }
}

function Set-StatusPill {
  param(
    [hashtable]$UI,
    [string]$Text,
    [ValidateSet('idle', 'running', 'done', 'error')]$State = 'idle'
  )

  $palette = $UI.Theme.Dark
  $UI.LblStatusPill.Text = $Text
  switch ($State) {
    'running' { $UI.StatusPill.BackColor = $palette.Accent }
    'done' { $UI.StatusPill.BackColor = $palette.Success }
    'error' { $UI.StatusPill.BackColor = $palette.Error }
    default { $UI.StatusPill.BackColor = $palette.Border }
  }
  $UI.LblStatusPill.ForeColor = $palette.BadgeText
}

function Set-GenerateProgress {
  param(
    [hashtable]$UI,
    [int]$Percent
  )
  if (-not $UI -or -not $UI.ProgressFill -or -not $UI.ProgressHost) { return }
  $p = [Math]::Max(0, [Math]::Min(100, [int]$Percent))
  $hostW = [Math]::Max(1, [int]$UI.ProgressHost.ClientSize.Width)
  $newW = [Math]::Max(0, [int][Math]::Floor($hostW * ($p / 100.0)))
  try { $UI.ProgressFill.Width = $newW } catch {}
}

function Start-GenerateProgressAnimation {
  param([hashtable]$UI)
  if (-not $UI) { return }
  try {
    if ($UI.ProgressTimer) {
      try { $UI.ProgressTimer.Stop() } catch {}
      try { $UI.ProgressTimer.Dispose() } catch {}
      $UI.ProgressTimer = $null
    }
  } catch {}
  $UI.ProgressAnimDirection = 1
  $UI.ProgressAnimPercent = 8
  Set-GenerateProgress -UI $UI -Percent 8
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 120
  $timer.Add_Tick(({
      try {
        $val = [int]$UI.ProgressAnimPercent
        $dir = [int]$UI.ProgressAnimDirection
        $val += (8 * $dir)
        if ($val -ge 92) { $val = 92; $dir = -1 }
        elseif ($val -le 12) { $val = 12; $dir = 1 }
        $UI.ProgressAnimPercent = $val
        $UI.ProgressAnimDirection = $dir
        Set-GenerateProgress -UI $UI -Percent $val
      } catch {}
    }).GetNewClosure())
  $UI.ProgressTimer = $timer
  try { $timer.Start() } catch {}
}

function Stop-GenerateProgressAnimation {
  param(
    [hashtable]$UI,
    [int]$FinalPercent = 0
  )
  if (-not $UI) { return }
  try {
    if ($UI.ProgressTimer) {
      try { $UI.ProgressTimer.Stop() } catch {}
      try { $UI.ProgressTimer.Dispose() } catch {}
      $UI.ProgressTimer = $null
    }
  } catch {}
  Set-GenerateProgress -UI $UI -Percent $FinalPercent
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

function Get-CheckedGenerateRows {
  param([hashtable]$UI)
  $picked = New-Object System.Collections.Generic.List[object]
  if (-not $UI) { return @() }
  try {
    if ($UI.Grid -and $UI.Grid.Rows) {
      for ($i = 0; $i -lt $UI.Grid.Rows.Count; $i++) {
        $gr = $UI.Grid.Rows[$i]
        if (-not $gr -or $gr.IsNewRow) { continue }
        $checked = $false
        try {
          $cell = $gr.Cells['Generate']
          if ($null -ne $cell -and $null -ne $cell.Value) {
            $checked = [bool]$cell.Value
          }
        } catch { $checked = $false }
        if (-not $checked) { continue }

        $rowNum = 0
        $rowText = ''
        try { $rowText = ("" + $gr.Cells['Row'].Value).Trim() } catch { $rowText = '' }
        if (-not [int]::TryParse($rowText, [ref]$rowNum)) { continue }
        $picked.Add([pscustomobject]@{
            Row = $rowNum
            Ticket = ("" + $gr.Cells['Ticket'].Value).Trim()
            User = ("" + $gr.Cells['User'].Value).Trim()
            PI = ("" + $gr.Cells['PI'].Value).Trim()
          }) | Out-Null
      }
      return @($picked.ToArray())
    }

    if (-not $UI.GridTable) { return @() }
    for ($i = 0; $i -lt $UI.GridTable.Rows.Count; $i++) {
      $dr = $UI.GridTable.Rows[$i]
      if (-not $dr) { continue }
      $checked = $false
      try {
        $cell = $dr['Generate']
        if ($null -ne $cell) { $checked = [bool]$cell }
      } catch { $checked = $false }
      if (-not $checked) { continue }

      $rowNum = 0
      $rowText = ("" + $dr['Row']).Trim()
      if (-not [int]::TryParse($rowText, [ref]$rowNum)) { continue }
      $picked.Add([pscustomobject]@{
          Row = $rowNum
          Ticket = ("" + $dr['Ticket']).Trim()
          User = ("" + $dr['User']).Trim()
          PI = ("" + $dr['PI']).Trim()
        }) | Out-Null
    }
  }
  catch {
    Write-Log -Level ERROR -Message ("Get-CheckedGenerateRows failed: " + $_.Exception.Message)
  }
  return @($picked.ToArray())
}

function Get-AllGenerateRows {
  param([hashtable]$UI)
  $picked = New-Object System.Collections.Generic.List[object]
  if (-not $UI) { return @() }
  try {
    if ($UI.Grid -and $UI.Grid.Rows) {
      for ($i = 0; $i -lt $UI.Grid.Rows.Count; $i++) {
        $gr = $UI.Grid.Rows[$i]
        if (-not $gr -or $gr.IsNewRow) { continue }
        $rowNum = 0
        $rowText = ''
        try { $rowText = ("" + $gr.Cells['Row'].Value).Trim() } catch { $rowText = '' }
        if (-not [int]::TryParse($rowText, [ref]$rowNum)) { continue }
        $picked.Add([pscustomobject]@{
            Row = $rowNum
            Ticket = ("" + $gr.Cells['Ticket'].Value).Trim()
            User = ("" + $gr.Cells['User'].Value).Trim()
            PI = ("" + $gr.Cells['PI'].Value).Trim()
          }) | Out-Null
      }
      return @($picked.ToArray())
    }
    if ($UI.GridTable) {
      for ($i = 0; $i -lt $UI.GridTable.Rows.Count; $i++) {
        $dr = $UI.GridTable.Rows[$i]
        if (-not $dr) { continue }
        $rowNum = 0
        $rowText = ("" + $dr['Row']).Trim()
        if (-not [int]::TryParse($rowText, [ref]$rowNum)) { continue }
        $picked.Add([pscustomobject]@{
            Row = $rowNum
            Ticket = ("" + $dr['Ticket']).Trim()
            User = ("" + $dr['User']).Trim()
            PI = ("" + $dr['PI']).Trim()
          }) | Out-Null
      }
    }
  }
  catch {
    Write-Log -Level ERROR -Message ("Get-AllGenerateRows failed: " + $_.Exception.Message)
  }
  return @($picked.ToArray())
}

function Update-GenerateSelectAllState {
  param([hashtable]$UI)
  if (-not $UI -or -not $UI.ChkSelectAll) { return }
  if ([bool]$UI.BulkSelectToggle) { return }
  $UI.SelectAllSyncing = $true
  try {
    $total = 0
    $checked = 0
    if ($UI.Grid -and $UI.Grid.Rows) {
      for ($i = 0; $i -lt $UI.Grid.Rows.Count; $i++) {
        $gr = $UI.Grid.Rows[$i]
        if (-not $gr -or $gr.IsNewRow) { continue }
        $total++
        $isChecked = $false
        try {
          $cell = $gr.Cells['Generate']
          if ($null -ne $cell -and $null -ne $cell.Value) {
            $isChecked = [bool]$cell.Value
          }
        } catch { $isChecked = $false }
        if ($isChecked) { $checked++ }
      }
    }
    elseif ($UI.GridTable) {
      $total = [int]$UI.GridTable.Rows.Count
      for ($i = 0; $i -lt $UI.GridTable.Rows.Count; $i++) {
        $dr = $UI.GridTable.Rows[$i]
        if (-not $dr) { continue }
        $isChecked = $false
        try {
          $cell = $dr['Generate']
          if ($null -ne $cell) { $isChecked = [bool]$cell }
        } catch { $isChecked = $false }
        if ($isChecked) { $checked++ }
      }
    }

    if ($total -le 0) {
      $UI.ChkSelectAll.CheckState = [System.Windows.Forms.CheckState]::Unchecked
      return
    }
    if ($checked -eq 0) {
      $UI.ChkSelectAll.CheckState = [System.Windows.Forms.CheckState]::Unchecked
    }
    elseif ($checked -ge $total) {
      $UI.ChkSelectAll.CheckState = [System.Windows.Forms.CheckState]::Checked
    }
    else {
      $UI.ChkSelectAll.CheckState = [System.Windows.Forms.CheckState]::Indeterminate
    }
  }
  finally {
    $UI.SelectAllSyncing = $false
  }
}

function Invoke-GenerateUiSafe {
  param(
    [hashtable]$UI,
    [string]$Context,
    [scriptblock]$Action
  )
  $ctx = ("" + $Context).Trim()
  if (-not $ctx) { $ctx = 'Generate UI Action' }
  if (-not $Action) { return }
  try {
    & $Action
  }
  catch {
    $errMsg = ''
    $errStack = ''
    try { $errMsg = ("" + $_.Exception.Message).Trim() } catch { $errMsg = '' }
    if (-not $errMsg) { $errMsg = 'Unhandled exception.' }
    try { $errStack = ("" + $_.ScriptStackTrace).Trim() } catch { $errStack = '' }
    try {
      if ($errStack) {
        Write-Log -Level ERROR -Message ("{0}: {1} | {2}" -f $ctx, $errMsg, $errStack)
      }
      else {
        Write-Log -Level ERROR -Message ("{0}: {1}" -f $ctx, $errMsg)
      }
    } catch {}
    Show-UiError -Context $ctx -ErrorRecord $_
  }
}

function Register-GenerateHandlers {
  <#
  .SYNOPSIS
  Registers Generator button and interaction handlers.
  .DESCRIPTION
  Wraps all UI actions with Invoke-GenerateUiSafe to avoid unhandled exceptions.
  .PARAMETER UI
  Shared synchronized UI state hashtable.
  .NOTES
  Event handlers execute on the WinForms UI thread.
  #>
  param([hashtable]$UI)

  $UI.Grid.Add_CurrentCellDirtyStateChanged(({
    param($sender, $args)
    try {
      if ($sender.IsCurrentCellDirty) { $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    } catch {}
  }).GetNewClosure())

  $UI.Grid.Add_CellValueChanged(({
    param($sender, $args)
    try {
      if ([bool]$UI.BulkSelectToggle) { return }
      if ($args -and $args.ColumnIndex -ge 0) {
        $col = $sender.Columns[$args.ColumnIndex]
        if ($col -and $col.Name -eq 'Generate') {
          Update-GenerateSelectAllState -UI $UI
        }
      }
    } catch {}
  }).GetNewClosure())

  $UI.Grid.Add_CellFormatting(({
    param($sender, $args)
    try {
      if ($args.RowIndex -lt 0 -or $args.ColumnIndex -lt 0) { return }
      $col = $sender.Columns[$args.ColumnIndex]
      if (-not $col -or $col.Name -ne 'Status') { return }
      $valueText = ("" + $args.Value).Trim().ToLowerInvariant()
      if ($valueText -eq 'done') {
        $args.CellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#14532D')
        $args.CellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#ECFDF5')
      }
      elseif ($valueText -eq 'error' -or $valueText -eq 'failed') {
        $args.CellStyle.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#7F1D1D')
        $args.CellStyle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#FEF2F2')
      }
      else {
        $args.CellStyle.BackColor = $sender.DefaultCellStyle.BackColor
        $args.CellStyle.ForeColor = $sender.DefaultCellStyle.ForeColor
      }
    } catch {}
  }).GetNewClosure())

  $UI.ChkSelectAll.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Select All' -Action {
      if ([bool]$UI.SelectAllSyncing) { return }
      $checkValue = [bool]$UI.ChkSelectAll.Checked
      $UI.BulkSelectToggle = $true
      try {
        try { if ($UI.Grid -and $UI.Grid.IsCurrentCellDirty) { $UI.Grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } } catch {}
        if ($UI.Grid -and $UI.Grid.Rows) {
          for ($i = 0; $i -lt $UI.Grid.Rows.Count; $i++) {
            $gr = $UI.Grid.Rows[$i]
            if (-not $gr -or $gr.IsNewRow) { continue }
            try { $gr.Cells['Generate'].Value = $checkValue } catch {}
          }
        }
        elseif ($UI.GridTable) {
          for ($i = 0; $i -lt $UI.GridTable.Rows.Count; $i++) {
            $dr = $UI.GridTable.Rows[$i]
            if (-not $dr) { continue }
            try { $dr['Generate'] = $checkValue } catch {}
          }
        }
      }
      finally {
        $UI.BulkSelectToggle = $false
      }
      Update-GenerateSelectAllState -UI $UI
    }
  }).GetNewClosure())

  $UI.BtnDashboard.Add_Click(({
    param($sender, $args)
    try { Write-Log -Level INFO -Message 'CLICK: Open Dashboard from Generator' } catch {}
    Invoke-GenerateUiSafe -UI $UI -Context 'Open Dashboard' -Action {
      if ($UI.OnOpenDashboard) { & $UI.OnOpenDashboard; return }
      $UI.LblStatusText.Text = 'Dashboard callback not configured.'
      Set-StatusPill -UI $UI -Text 'Error' -State error
    }
  }).GetNewClosure())

  $UI.BtnOpen.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Open Output Folder' -Action {
      $out = ("" + $UI.OutputPath).Trim()
      if (-not $out) { return }
      if (-not (Test-Path -LiteralPath $out)) { return }
      Start-Process -FilePath $out | Out-Null
    }
  }).GetNewClosure())

  $UI.BtnCloseCode.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Close All' -Action {
      if ($UI.OnCloseAll) {
        try { & $UI.OnCloseAll $UI } catch {}
        return
      }
      $r = $null
      if (Get-Command -Name Close-SchumanAllResources -ErrorAction SilentlyContinue) {
        $r = Close-SchumanAllResources -Mode 'All'
      }
      else {
        $fallback = Invoke-UiEmergencyClose -ActionLabel 'Close All' -ExecutableNames @('Code', 'Code - Insiders', 'Cursor', 'WINWORD', 'EXCEL') -Owner $UI.Form -Mode 'All'
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
      $UI.LblStatusText.Text = $summary
      Append-GenerateLog -UI $UI -Line $summary
      try { if ($UI.Form -and -not $UI.Form.IsDisposed) { $UI.Form.Close() } } catch {}
    }
  }).GetNewClosure())

  $UI.BtnToggleLog.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Toggle Log' -Action {
      $UI.LogPanel.Visible = -not $UI.LogPanel.Visible
      if ($UI.LogPanel.Visible) {
        $UI.LogPanel.Height = 150
        $UI.BtnToggleLog.Text = 'Hide Log'
        $logPath = ''
        try { if ($script:LogPath) { $logPath = ("" + $script:LogPath).Trim() } } catch {}
        if (-not $logPath) {
          try {
            $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
            $logPath = Join-Path $projectRoot 'system\logs\ui.log'
          } catch {}
        }
        if ($logPath -and (Test-Path -LiteralPath $logPath)) {
          try {
            $content = Get-Content -LiteralPath $logPath -Tail 250 -ErrorAction Stop
            $UI.LogBox.Text = (($content -join [Environment]::NewLine) + [Environment]::NewLine)
          }
          catch {
            $UI.LogBox.Text = ("Unable to read log: " + $_.Exception.Message)
          }
        }
        else {
          $UI.LogBox.Text = "No log yet.`r`nExpected path: $logPath"
        }
      }
      else {
        $UI.LogPanel.Height = 0
        $UI.BtnToggleLog.Text = 'Show Log'
      }
    }
  }).GetNewClosure())

  $UI.BtnStop.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Stop' -Action {
      [System.Windows.Forms.MessageBox]::Show('Stop is not available in this integrated mode.', 'Info') | Out-Null
    }
  }).GetNewClosure())

  $UI.BtnStart.Add_Click(({
    param($sender, $args)
    Invoke-GenerateUiSafe -UI $UI -Context 'Generate Documents' -Action {
      if (-not [bool]$UI.ExcelReady) {
        $UI.LblStatusText.Text = 'Excel still loading...'
        Set-StatusPill -UI $UI -Text 'Loading' -State running
        return
      }
      $selectedRows = @(Get-CheckedGenerateRows -UI $UI)
      if ($selectedRows.Count -eq 0) {
        $fallbackRows = @(Get-AllGenerateRows -UI $UI)
        if ($fallbackRows.Count -gt 0) {
          $selectedRows = @($fallbackRows)
          Append-GenerateLog -UI $UI -Line ("No rows checked; using all visible rows ({0}) like legacy flow." -f $selectedRows.Count)
        }
        else {
          $UI.LblStatusText.Text = 'Select at least one row to generate.'
          Set-StatusPill -UI $UI -Text 'Select Rows' -State error
          Append-GenerateLog -UI $UI -Line 'Select at least one row to generate.'
          return
        }
      }
      if ((-not [bool]$UI.ChkSaveDocx.Checked) -and (-not [bool]$UI.ChkSavePdf.Checked)) {
        $UI.LblStatusText.Text = 'Choose at least one format (DOCX or PDF).'
        Set-StatusPill -UI $UI -Text 'Choose Format' -State error
        Append-GenerateLog -UI $UI -Line 'Generation cancelled: both Save DOCX and Save PDF are unchecked.'
        return
      }
      $UI.BtnStart.Enabled = $false
      $UI.BtnDashboard.Enabled = $false
      Start-GenerateProgressAnimation -UI $UI
      try {
        $UI.LblStatusText.Text = ("Generating documents for {0} selected row(s)..." -f $selectedRows.Count)
        Set-StatusPill -UI $UI -Text 'Running' -State running
        Append-GenerateLog -UI $UI -Line ("Starting document generation for {0} selected row(s)." -f $selectedRows.Count)
        $alreadyDoneRows = New-Object System.Collections.Generic.List[int]
        $rowsToGenerate = New-Object System.Collections.Generic.List[int]
        foreach ($item in $selectedRows) {
          $rnum = 0
          if (-not [int]::TryParse(("" + $item.Row).Trim(), [ref]$rnum)) { continue }
          $existing = Test-GenerateRowAlreadyExported -UI $UI -Ticket $item.Ticket -User $item.User
          if ($existing.Exists) {
            $alreadyDoneRows.Add($rnum) | Out-Null
          }
          else {
            $rowsToGenerate.Add($rnum) | Out-Null
          }
        }

        foreach ($dr in @($UI.GridTable.Rows)) {
          if (-not $dr) { continue }
          $rowN = 0
          if (-not [int]::TryParse(("" + $dr['Row']).Trim(), [ref]$rowN)) { continue }
          if ($alreadyDoneRows.Contains($rowN)) {
            $dr['Generate'] = $false
            $dr['Status'] = 'Done'
            $dr['Message'] = 'Already exported in WORD files'
            $dr['Progress'] = '100%'
          }
        }

        if ($rowsToGenerate.Count -eq 0) {
          Stop-GenerateProgressAnimation -UI $UI -FinalPercent 100
          $UI.LblStatusText.Text = 'All selected files already exist.'
          Set-StatusPill -UI $UI -Text 'Done' -State done
          Append-GenerateLog -UI $UI -Line 'No new files to generate; all selected rows were already exported.'
          $UI.LblMetrics.Text = ("Total: {0} | Saved: {1} | Skipped: 0 | Errors: 0" -f $selectedRows.Count, $selectedRows.Count)
          Update-GenerateSelectAllState -UI $UI
          return
        }

        $selectedRowNumbers = @($rowsToGenerate.ToArray())
        $result = Generate-PDF -UI $UI -SelectedRowNumbers $selectedRowNumbers
        $generatedCount = 0
        try { if ($result -and $result.generatedCount -ne $null) { $generatedCount = [int]$result.generatedCount } } catch { $generatedCount = 0 }
        if ($result -and $result.ok -eq $true -and $generatedCount -gt 0) {
          Stop-GenerateProgressAnimation -UI $UI -FinalPercent 100
          $UI.LblStatusText.Text = 'Generation complete.'
          Set-StatusPill -UI $UI -Text 'Done' -State done
          $msg = if ($result.message) { "" + $result.message } else { 'Documents generated successfully.' }
          Append-GenerateLog -UI $UI -Line $msg
          if ($result.outputPath) { $UI.OutputPath = ("" + $result.outputPath).Trim() }
          Update-OutputButton -UI $UI
          $savedCount = 0
          foreach ($dr in @($UI.GridTable.Rows)) {
            if (-not $dr) { continue }
            $checked = $false
            try { $checked = [bool]$dr['Generate'] } catch { $checked = $false }
            if (-not $checked) { continue }
            $rnum = 0
            if (-not [int]::TryParse(("" + $dr['Row']).Trim(), [ref]$rnum)) { continue }
            if (-not $rowsToGenerate.Contains($rnum)) { continue }
            $pi = ("" + $dr['PI']).Trim()
            if (-not $pi) { $pi = '-' }
            $pathText = ("" + $UI.OutputPath).Trim()
            if ($pathText) { $dr['File'] = ("{0} | PI: {1}" -f $pathText, $pi) } else { $dr['File'] = ("PI: {0}" -f $pi) }
            $dr['Status'] = 'Done'
            $dr['Message'] = $msg
            $dr['Progress'] = '100%'
            $dr['Generate'] = $false
            $savedCount++
          }
          $savedTotal = $savedCount + $alreadyDoneRows.Count
          $UI.LblMetrics.Text = ("Total: {0} | Saved: {1} | Skipped: {2} | Errors: 0" -f $selectedRows.Count, $savedTotal, [Math]::Max(0, ($selectedRows.Count - $savedTotal)))
          Update-GenerateSelectAllState -UI $UI
        }
        else {
          Stop-GenerateProgressAnimation -UI $UI -FinalPercent 0
          $msg = if ($result -and $result.message) { "" + $result.message } else { 'Generation failed without details.' }
          $UI.LblStatusText.Text = 'Generation failed.'
          Set-StatusPill -UI $UI -Text 'Error' -State error
          Append-GenerateLog -UI $UI -Line $msg
          $UI.LblMetrics.Text = ("Total: {0} | Saved: 0 | Skipped: 0 | Errors: {0}" -f $selectedRows.Count)
        }
      }
      catch {
        Stop-GenerateProgressAnimation -UI $UI -FinalPercent 0
        $UI.LblStatusText.Text = 'Generation failed.'
        Set-StatusPill -UI $UI -Text 'Error' -State error
        $errMsg = ("" + $_.Exception.Message).Trim()
        if (-not $errMsg) { $errMsg = 'Unknown generation error.' }
        $errStack = ''
        try { $errStack = ("" + $_.ScriptStackTrace).Trim() } catch {}
        Append-GenerateLog -UI $UI -Line ("ERROR: " + $errMsg)
        try {
          if ($errStack) {
            Write-Log -Level ERROR -Message ("Generate Documents failed: {0} | {1}" -f $errMsg, $errStack)
          }
          else {
            Write-Log -Level ERROR -Message ("Generate Documents failed: " + $errMsg)
          }
        } catch {}
        $selCount = 1
        try { $selCount = @(Get-CheckedGenerateRows -UI $UI).Count } catch {}
        if ($selCount -lt 1) { $selCount = 1 }
        $UI.LblMetrics.Text = ("Total: {0} | Saved: 0 | Skipped: 0 | Errors: {0}" -f $selCount)
        Show-UiError -Context 'Generate-PDF' -ErrorRecord $_
      }
      finally {
        if ($UI.Grid) { try { $UI.Grid.Refresh() } catch {} }
        if ([bool]$UI.ExcelReady) {
          $UI.BtnStart.Enabled = $true
          $UI.BtnDashboard.Enabled = $true
        }
      }
    }
  }).GetNewClosure())

}

function global:New-GeneratePdfUI {
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [string]$TemplatePath = '',
    [string]$OutputPath = '',
    [scriptblock]$OnOpenDashboard = $null,
    [scriptblock]$OnGenerate = $null,
    [scriptblock]$OnCloseAll = $null
  )

  $UI = New-UI -ExcelPath $ExcelPath -SheetName $SheetName -TemplatePath $TemplatePath -OutputPath $OutputPath -OnOpenDashboard $OnOpenDashboard -OnGenerate $OnGenerate -OnCloseAll $OnCloseAll
  $null = Initialize-Controls -UI $UI
  Ensure-GenerateGridHeaders -UI $UI
  Register-GenerateHandlers -UI $UI
  Set-GenerateUiTheme -UI $UI
  Set-StatusPill -UI $UI -Text 'Idle' -State idle
  Update-OutputButton -UI $UI

  try {
    $rows = @(Export-Excel -UI $UI)
    Update-Grid -UI $UI -Rows $rows
    $total = $rows.Count
    $reason = if ($total -gt 0) { 'Ready' } elseif ((Test-Path -LiteralPath $UI.ExcelPath)) { 'Excel is empty' } else { 'Excel still loading...' }
    Update-GenerateDataState -UI $UI -ExcelReady ($total -gt 0) -Reason $reason
    $UI.LblMetrics.Text = "Total: $total | Saved: 0 | Skipped: 0 | Errors: 0"
    if ($total -gt 0) {
      $UI.LblStatusText.Text = "Ready. Preloaded $total rows from Excel."
      Append-GenerateLog -UI $UI -Line "Preloaded $total rows from Excel."
    }
  }
  catch {
    $msg = 'Excel still loading...'
    try {
      $msgEx = ("" + $_.Exception.Message).Trim()
      if ($msgEx) { $msg = $msgEx }
    } catch {}
    Update-GenerateDataState -UI $UI -ExcelReady $false -Reason $msg
    Show-UiError -Context 'Load-Data' -ErrorRecord $_
  }

  return $UI.Form
}
