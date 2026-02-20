Set-StrictMode -Version Latest

function New-GeneratePdfUI {
  param(
    [string]$ExcelPath = '',
    [string]$SheetName = 'BRU',
    [string]$TemplatePath = '',
    [string]$OutputPath = '',
    [scriptblock]$OnOpenDashboard = $null,
    [scriptblock]$OnGenerate = $null
  )

  $fontName = Get-UiFontName

  $theme = @{
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

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'Schuman Word Generator'
  $form.Size = New-Object System.Drawing.Size(1120, 720)
  $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
  $form.MinimumSize = New-Object System.Drawing.Size(980, 640)
  $form.BackColor = $theme.Dark.Bg
  $form.Font = New-Object System.Drawing.Font($fontName, 10)
  try {
    $prop = $form.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($prop) { $prop.SetValue($form, $true, $null) }
  } catch {}

  $root = New-Object System.Windows.Forms.TableLayoutPanel
  $root.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Padding = New-Object System.Windows.Forms.Padding(16)
  $root.RowCount = 3
  $root.ColumnCount = 1
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$form.Controls.Add($root)

  $headerCard = New-Object System.Windows.Forms.Panel
  $headerCard.Dock = [System.Windows.Forms.DockStyle]::Fill
  $headerCard.Padding = New-Object System.Windows.Forms.Padding(16, 16, 16, 12)
  $headerCard.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
  [void]$root.Controls.Add($headerCard, 0, 0)

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

  $headerActions = New-Object System.Windows.Forms.FlowLayoutPanel
  $headerActions.AutoSize = $true
  $headerActions.WrapContents = $false
  $headerActions.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $headerActions.Anchor = 'Top,Right'
  $headerActions.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
  [void]$headerGrid.Controls.Add($headerActions, 1, 0)

  $btnDashboard = New-Object System.Windows.Forms.Button
  $btnDashboard.Text = 'Open Check-in Dashboard'
  $btnDashboard.Width = 220
  $btnDashboard.Height = 32
  $btnDashboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnDashboard.Font = New-Object System.Drawing.Font($fontName, 9, [System.Drawing.FontStyle]::Bold)
  $btnDashboard.FlatAppearance.BorderSize = 2
  [void]$headerActions.Controls.Add($btnDashboard)

  $statusPill = New-Object System.Windows.Forms.Panel
  $statusPill.AutoSize = $true
  $statusPill.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
  $statusPill.Margin = New-Object System.Windows.Forms.Padding(8, 4, 0, 0)
  [void]$headerActions.Controls.Add($statusPill)

  $lblStatusPill = New-Object System.Windows.Forms.Label
  $lblStatusPill.Text = 'Idle'
  $lblStatusPill.AutoSize = $true
  [void]$statusPill.Controls.Add($lblStatusPill)

  $lblMetrics = New-Object System.Windows.Forms.Label
  $lblMetrics.Text = 'Total: 0 | Saved: 0 | Skipped: 0 | Errors: 0'
  $lblMetrics.AutoSize = $true
  $lblMetrics.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
  [void]$headerGrid.Controls.Add($lblMetrics, 0, 1)

  $lblStatusText = New-Object System.Windows.Forms.Label
  $lblStatusText.Text = 'Ready.'
  $lblStatusText.AutoSize = $true
  $lblStatusText.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
  $lblStatusText.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
  $lblStatusText.Anchor = 'Top,Right'
  [void]$headerGrid.Controls.Add($lblStatusText, 1, 1)

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

  $progressFill = New-Object System.Windows.Forms.Panel
  $progressFill.Height = 20
  $progressFill.Width = 0
  $progressFill.Dock = [System.Windows.Forms.DockStyle]::Left
  [void]$progressHost.Controls.Add($progressFill)

  $listCard = New-Object System.Windows.Forms.Panel
  $listCard.Dock = [System.Windows.Forms.DockStyle]::Fill
  $listCard.Padding = New-Object System.Windows.Forms.Padding(15)
  [void]$centerGrid.Controls.Add($listCard, 0, 1)

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $grid.ReadOnly = $true
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AllowUserToResizeRows = $false
  $grid.RowHeadersVisible = $false
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EnableHeadersVisualStyles = $false
  $grid.ColumnHeadersHeight = 32
  $grid.RowTemplate.Height = 28
  try {
    $prop = $grid.GetType().GetProperty('DoubleBuffered', 'NonPublic,Instance')
    if ($prop) { $prop.SetValue($grid, $true, $null) }
  } catch {}
  [void]$listCard.Controls.Add($grid)

  [void]$grid.Columns.Add('Row', 'Row #')
  [void]$grid.Columns.Add('Ticket', 'Ticket/RITM')
  [void]$grid.Columns.Add('User', 'User')
  [void]$grid.Columns.Add('File', 'Output File')
  [void]$grid.Columns.Add('Status', 'Status')
  [void]$grid.Columns.Add('Message', 'Message')
  [void]$grid.Columns.Add('Progress', 'Progress')
  $grid.Columns['Row'].FillWeight = 8
  $grid.Columns['Ticket'].FillWeight = 16
  $grid.Columns['User'].FillWeight = 16
  $grid.Columns['File'].FillWeight = 24
  $grid.Columns['Status'].FillWeight = 12
  $grid.Columns['Message'].FillWeight = 24
  $grid.Columns['Progress'].FillWeight = 10
  $grid.Columns['Progress'].DefaultCellStyle.Alignment = 'MiddleRight'

  $logPanel = New-Object System.Windows.Forms.Panel
  $logPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $logPanel.Padding = New-Object System.Windows.Forms.Padding(15)
  $logPanel.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 0)
  $logPanel.Height = 0
  $logPanel.Visible = $false
  [void]$listCard.Controls.Add($logPanel)

  $logBox = New-Object System.Windows.Forms.RichTextBox
  $logBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $logBox.ReadOnly = $true
  $logBox.HideSelection = $false
  $logBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  [void]$logPanel.Controls.Add($logBox)

  $footer = New-Object System.Windows.Forms.Panel
  $footer.Dock = [System.Windows.Forms.DockStyle]::Fill
  $footer.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
  [void]$root.Controls.Add($footer, 0, 2)

  $footerGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $footerGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
  $footerGrid.ColumnCount = 2
  $footerGrid.RowCount = 1
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$footerGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$footer.Controls.Add($footerGrid)

  $buttonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $buttonFlow.AutoSize = $true
  $buttonFlow.WrapContents = $false
  $buttonFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $buttonFlow.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
  [void]$footerGrid.Controls.Add($buttonFlow, 0, 0)

  $btnStart = New-Object System.Windows.Forms.Button
  $btnStart.Text = 'Generate Documents'
  $btnStart.Width = 170
  $btnStart.Height = 30
  $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($btnStart)

  $btnStop = New-Object System.Windows.Forms.Button
  $btnStop.Text = 'Stop'
  $btnStop.Width = 110
  $btnStop.Height = 30
  $btnStop.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnStop.Enabled = $false
  [void]$buttonFlow.Controls.Add($btnStop)

  $btnOpen = New-Object System.Windows.Forms.Button
  $btnOpen.Text = 'Open Output Folder'
  $btnOpen.Width = 170
  $btnOpen.Height = 30
  $btnOpen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnOpen.Enabled = $false
  [void]$buttonFlow.Controls.Add($btnOpen)

  $btnCloseCode = New-Object System.Windows.Forms.Button
  $btnCloseCode.Text = 'Cerrar codigo'
  $btnCloseCode.Width = 120
  $btnCloseCode.Height = 30
  $btnCloseCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($btnCloseCode)

  $btnCloseDocs = New-Object System.Windows.Forms.Button
  $btnCloseDocs.Text = 'Cerrar documentos'
  $btnCloseDocs.Width = 145
  $btnCloseDocs.Height = 30
  $btnCloseDocs.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($btnCloseDocs)

  $btnToggleLog = New-Object System.Windows.Forms.Button
  $btnToggleLog.Text = 'Show Log'
  $btnToggleLog.Width = 110
  $btnToggleLog.Height = 30
  $btnToggleLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  [void]$buttonFlow.Controls.Add($btnToggleLog)

  $optionsFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $optionsFlow.Dock = [System.Windows.Forms.DockStyle]::Fill
  $optionsFlow.AutoSize = $true
  $optionsFlow.WrapContents = $false
  $optionsFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $optionsFlow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
  [void]$footerGrid.Controls.Add($optionsFlow, 1, 0)

  $chkShowWord = New-Object System.Windows.Forms.CheckBox
  $chkShowWord.Text = 'Show Word after generation'
  $chkShowWord.AutoSize = $true
  $chkShowWord.Checked = $false
  [void]$optionsFlow.Controls.Add($chkShowWord)

  $chkSaveDocx = New-Object System.Windows.Forms.CheckBox
  $chkSaveDocx.Text = 'Save DOCX'
  $chkSaveDocx.AutoSize = $true
  $chkSaveDocx.Checked = $true
  [void]$optionsFlow.Controls.Add($chkSaveDocx)

  $chkSavePdf = New-Object System.Windows.Forms.CheckBox
  $chkSavePdf.Text = 'Save PDF'
  $chkSavePdf.AutoSize = $true
  $chkSavePdf.Checked = $true
  [void]$optionsFlow.Controls.Add($chkSavePdf)

  $chkDark = New-Object System.Windows.Forms.CheckBox
  $chkDark.Text = 'Dark theme'
  $chkDark.AutoSize = $true
  $chkDark.Checked = $true
  [void]$optionsFlow.Controls.Add($chkDark)

  $appendLog = {
    param([string]$line)
    $logBox.AppendText("[" + (Get-Date -Format 'HH:mm:ss') + "] $line" + [Environment]::NewLine)
  }

  $setStatusPill = {
    param([string]$Text, [string]$State)
    $p = if ($chkDark.Checked) { $theme.Dark } else { $theme.Light }
    $lblStatusPill.Text = $Text
    switch ($State) {
      'running' { $statusPill.BackColor = $p.Accent }
      'error' { $statusPill.BackColor = $p.Error }
      'done' { $statusPill.BackColor = $p.Success }
      default { $statusPill.BackColor = $p.Border }
    }
    $lblStatusPill.ForeColor = $p.BadgeText
  }

  $applyTheme = {
    $p = if ($chkDark.Checked) { $theme.Dark } else { $theme.Light }

    $form.BackColor = $p.Bg
    $form.ForeColor = $p.Text
    $headerCard.BackColor = $p.Card
    $listCard.BackColor = $p.Card
    $footer.BackColor = $p.Card
    $logPanel.BackColor = $p.Card
    $progressHost.BackColor = $p.Border
    $progressFill.BackColor = $p.Accent

    $lblTitle.ForeColor = $p.Text
    $lblMetrics.ForeColor = $p.Sub
    $lblStatusText.ForeColor = $p.Sub
    $logBox.BackColor = $p.Bg
    $logBox.ForeColor = $p.Text

    foreach ($btn in @($btnDashboard, $btnStart, $btnStop, $btnOpen, $btnCloseCode, $btnCloseDocs, $btnToggleLog)) {
      $btn.BackColor = $p.Card
      $btn.ForeColor = $p.Text
      $btn.FlatAppearance.BorderColor = $p.Border
      $btn.FlatAppearance.BorderSize = 1
    }

    $btnDashboard.BackColor = $p.Bg
    $btnDashboard.ForeColor = $p.Accent
    $btnDashboard.FlatAppearance.BorderColor = $p.Accent
    $btnDashboard.FlatAppearance.BorderSize = 2

    $btnStart.BackColor = $p.Accent
    $btnStart.ForeColor = $p.BadgeText
    $btnStart.FlatAppearance.BorderColor = $p.Accent

    $btnStop.BackColor = $p.StopBg
    $btnStop.ForeColor = $p.StopFg
    $btnStop.FlatAppearance.BorderColor = $p.StopBorder

    foreach ($chk in @($chkDark, $chkSavePdf, $chkSaveDocx, $chkShowWord)) {
      $chk.BackColor = $p.Card
      $chk.ForeColor = $p.Sub
    }

    $grid.BackgroundColor = $p.Bg
    $grid.GridColor = $p.Border
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $p.Card
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $p.Sub
    $grid.DefaultCellStyle.BackColor = $p.Bg
    $grid.DefaultCellStyle.ForeColor = $p.Text
    $grid.DefaultCellStyle.SelectionBackColor = $p.Shadow
    $grid.DefaultCellStyle.SelectionForeColor = $p.Text
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $p.Card
    $grid.AlternatingRowsDefaultCellStyle.ForeColor = $p.Text

    & $setStatusPill $lblStatusPill.Text 'idle'
  }

  $updateOutputButton = {
    $out = ("" + $OutputPath).Trim()
    $btnOpen.Enabled = (-not [string]::IsNullOrWhiteSpace($out)) -and (Test-Path -LiteralPath $out)
  }

  $btnDashboard.Add_Click({
    if ($OnOpenDashboard) {
      & $OnOpenDashboard
      return
    }
    $lblStatusText.Text = 'Dashboard callback not configured.'
    & $setStatusPill 'Error' 'error'
  })

  $btnOpen.Add_Click({
    $out = ("" + $OutputPath).Trim()
    if (-not $out) { return }
    try { Start-Process $out | Out-Null } catch {}
  })

  $btnCloseCode.Add_Click({
    $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar codigo' -ExecutableNames @('code.exe', 'code-insiders.exe', 'cursor.exe') -Owner $form
    if (-not $r.Cancelled) {
      $lblStatusText.Text = $r.Message
      & $appendLog $r.Message
    }
  })

  $btnCloseDocs.Add_Click({
    $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar documentos' -ExecutableNames @('winword.exe', 'excel.exe') -Owner $form
    if (-not $r.Cancelled) {
      $lblStatusText.Text = $r.Message
      & $appendLog $r.Message
    }
  })

  $btnToggleLog.Add_Click({
    $logPanel.Visible = -not $logPanel.Visible
    if ($logPanel.Visible) {
      $logPanel.Height = 150
      $btnToggleLog.Text = 'Hide Log'
    } else {
      $logPanel.Height = 0
      $btnToggleLog.Text = 'Show Log'
    }
  })

  $btnStop.Add_Click({
    [System.Windows.Forms.MessageBox]::Show('Stop is not available in this integrated mode.', 'Info') | Out-Null
  })

  $btnStart.Add_Click({
    if (-not $OnGenerate) {
      & $appendLog 'Generate callback not configured.'
      $lblStatusText.Text = 'Generate callback not configured.'
      & $setStatusPill 'Error' 'error'
      return
    }

    $argsObj = [pscustomobject]@{
      ExcelPath = ("" + $ExcelPath).Trim()
      TemplatePath = ("" + $TemplatePath).Trim()
      OutputPath = ("" + $OutputPath).Trim()
      ExportPdf = [bool]$chkSavePdf.Checked
      SaveDocx = [bool]$chkSaveDocx.Checked
      ShowWord = [bool]$chkShowWord.Checked
    }

    if (-not (Test-Path -LiteralPath $argsObj.ExcelPath)) {
      [System.Windows.Forms.MessageBox]::Show('Excel file not found.', 'Validation') | Out-Null
      return
    }
    if (-not (Test-Path -LiteralPath $argsObj.TemplatePath)) {
      [System.Windows.Forms.MessageBox]::Show('Template file not found.', 'Validation') | Out-Null
      return
    }
    if ([string]::IsNullOrWhiteSpace($argsObj.OutputPath)) {
      [System.Windows.Forms.MessageBox]::Show('Output folder is required.', 'Validation') | Out-Null
      return
    }

    $btnStart.Enabled = $false
    $btnDashboard.Enabled = $false
    $lblStatusText.Text = 'Generating documents...'
    & $setStatusPill 'Running' 'running'
    & $appendLog 'Starting document generation.'

    $result = $null
    try {
      $result = & $OnGenerate $argsObj
      if ($result -and $result.ok -eq $true) {
        $lblStatusText.Text = 'Generation complete.'
        & $setStatusPill 'Done' 'done'
        $msg = if ($result.message) { "" + $result.message } else { 'Documents generated successfully.' }
        & $appendLog $msg

        if ($result.outputPath) { $OutputPath = "" + $result.outputPath }
        & $updateOutputButton

        $null = $grid.Rows.Add(@(
          '' ,
          '' ,
          '' ,
          $OutputPath,
          'Done',
          $msg,
          '100%'
        ))
        $lblMetrics.Text = 'Total: 1 | Saved: 1 | Skipped: 0 | Errors: 0'
      }
      else {
        $msg = if ($result -and $result.message) { "" + $result.message } else { 'Unknown error.' }
        $lblStatusText.Text = 'Generation failed.'
        & $setStatusPill 'Error' 'error'
        & $appendLog $msg
        $lblMetrics.Text = 'Total: 1 | Saved: 0 | Skipped: 0 | Errors: 1'
      }
    }
    catch {
      $lblStatusText.Text = 'Generation failed.'
      & $setStatusPill 'Error' 'error'
      & $appendLog ("ERROR: " + $_.Exception.Message)
      [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Generate') | Out-Null
      $lblMetrics.Text = 'Total: 1 | Saved: 0 | Skipped: 0 | Errors: 1'
    }
    finally {
      $btnStart.Enabled = $true
      $btnDashboard.Enabled = $true
    }
  })

  $chkDark.Add_CheckedChanged({ & $applyTheme })

  $preloadGrid = {
    if ([string]::IsNullOrWhiteSpace($ExcelPath)) { return }
    if (-not (Test-Path -LiteralPath $ExcelPath)) { return }
    if (-not (Get-Command -Name Search-DashboardRows -ErrorAction SilentlyContinue)) { return }

    try {
      $rows = @(Search-DashboardRows -ExcelPath $ExcelPath -SheetName $SheetName -SearchText '')
      $grid.Rows.Clear()

      foreach ($r in $rows) {
        $ticket = ("" + $r.RITM).Trim()
        if (-not $ticket) { continue }
        $status = ("" + $r.DashboardStatus).Trim()
        if (-not $status) { $status = 'Ready' }
        $msg = if ($status -eq 'Ready') { 'Preloaded from Excel' } else { 'Preloaded from Excel status' }

        $null = $grid.Rows.Add(@(
            ("" + $r.Row),
            $ticket,
            ("" + $r.RequestedFor),
            '',
            $status,
            $msg,
            '0%'
          ))
      }

      $total = $grid.Rows.Count
      $lblMetrics.Text = "Total: $total | Saved: 0 | Skipped: 0 | Errors: 0"
      if ($total -gt 0) {
        $lblStatusText.Text = "Ready. Preloaded $total rows from Excel."
        & $appendLog "Preloaded $total rows from Excel."
      }
    }
    catch {
      & $appendLog ("Preload failed: " + $_.Exception.Message)
    }
  }

  & $applyTheme
  & $setStatusPill 'Idle' 'idle'
  & $updateOutputButton
  & $preloadGrid
  return $form
}
