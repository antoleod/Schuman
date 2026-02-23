#Requires -Version 5.1
<#
.SYNOPSIS
    Generador de documentos Word y PDF basado en una plantilla y una lista de Excel.
.DESCRIPTION
    Este script proporciona una interfaz gráfica (GUI) para procesar filas de un Excel.
    Utiliza un hilo de trabajo (Worker) en segundo plano para no congelar la interfaz durante la automatización de Word.
#>
param(
  # >>> CHANGE EXCEL DEFAULT NAME HERE if the planning file is renamed again <<<
  [string]$ExcelPath = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) "Schuman List.xlsx"),
  [string]$TemplatePath = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) "Reception_ITequipment.docx"),
  [string]$OutDir = (Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))) "WORD files"),
  [string]$LogPath = (Join-Path $PSScriptRoot "Generate-Schuman-Words.log"),
  [string]$PreferredSheet = "BRU",
  [string]$AutoExcelScript = (Join-Path $PSScriptRoot "auto-excel.ps1"),
  [string]$DashboardScript = (Join-Path $PSScriptRoot "dashboard-checkin-checkout.ps1")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

try {
  if (-not [System.IO.Path]::IsPathRooted($AutoExcelScript)) {
    $AutoExcelScript = Join-Path $PSScriptRoot $AutoExcelScript
  }
  if (-not [System.IO.Path]::IsPathRooted($DashboardScript)) {
    $DashboardScript = Join-Path $PSScriptRoot $DashboardScript
  }
}
catch {}

# ----------------------------
# Config
# ----------------------------
# Values are parameterized above.

# ----------------------------
# Logging (never crash)
# ----------------------------
function Write-Log([string]$Message) {
  try {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$ts] $Message" -Encoding UTF8
  } catch {
    # swallow - log must never kill the app
  }
}

function Resolve-ExcelPath {
  param(
    [string]$CurrentExcelPath
  )

  if ($CurrentExcelPath -and (Test-Path -LiteralPath $CurrentExcelPath)) {
    return $CurrentExcelPath
  }

  $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
  $defaultExcel = Join-Path $projectRoot "Schuman List.xlsx"
  if (Test-Path -LiteralPath $defaultExcel) {
    return $defaultExcel
  }
  throw "Excel file not found. Expected: $defaultExcel"
}

$script:AutoExcelReady = $false
$script:AutoExcelWarning = ''
try {
  $ExcelPath = Resolve-ExcelPath -CurrentExcelPath $ExcelPath
}
catch {
  $script:AutoExcelWarning = ("Excel path validation failed: " + $_.Exception.Message)
  Write-Log $script:AutoExcelWarning
}

# ----------------------------
# Preflight (UI will also show message)
# ----------------------------
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
if (-not (Test-Path -LiteralPath $ExcelPath)) { Write-Log "Missing Excel: $ExcelPath" }
if (-not (Test-Path -LiteralPath $TemplatePath)) { Write-Log "Missing Template: $TemplatePath" }

# ----------------------------
# Theme
# ----------------------------
$Theme = @{
  Light = @{
    Bg = [Drawing.Color]::FromArgb(245,246,248)
    Card = [Drawing.Color]::White
    Text = [Drawing.Color]::FromArgb(20,20,20)
    Sub  = [Drawing.Color]::FromArgb(100,100,100)
    Border = [Drawing.Color]::FromArgb(228,228,232)
    Accent = [Drawing.Color]::FromArgb(0,122,255)
    Warning = [Drawing.Color]::FromArgb(255,185,0)
    Success = [Drawing.Color]::FromArgb(220,245,231)
    Error = [Drawing.Color]::FromArgb(255,230,230)
    BadgeText = [Drawing.Color]::FromArgb(30,30,30)
    Shadow = [Drawing.Color]::FromArgb(235,235,240)
  }
  Dark = @{
    Bg = [Drawing.ColorTranslator]::FromHtml('#0F172A')
    Card = [Drawing.ColorTranslator]::FromHtml('#1E293B')
    Text = [Drawing.ColorTranslator]::FromHtml('#E5E7EB')
    Sub  = [Drawing.ColorTranslator]::FromHtml('#94A3B8')
    Border = [Drawing.ColorTranslator]::FromHtml('#334155')
    Accent = [Drawing.ColorTranslator]::FromHtml('#2563EB')
    Warning = [Drawing.ColorTranslator]::FromHtml('#F59E0B')
    Success = [Drawing.ColorTranslator]::FromHtml('#14532D')
    Error = [Drawing.ColorTranslator]::FromHtml('#7F1D1D')
    BadgeText = [Drawing.ColorTranslator]::FromHtml('#FFFFFF')
    Shadow = [Drawing.ColorTranslator]::FromHtml('#172033')
  }
}

# ----------------------------
# UI
# ----------------------------
# Clase para gestionar el cambio de tema (Claro/Oscuro) en todos los controles de WinForms.
# Recorre recursivamente los controles y aplica los colores definidos en la paleta.
class ThemeManager {
  [hashtable]$Palette
  ThemeManager([hashtable]$palette){ $this.Palette = $palette }
  [void]SetPalette([hashtable]$palette){ $this.Palette = $palette }
  [void]ApplyControl($c){
    $p = $this.Palette
    switch ($c.GetType().Name) {
      "Form" { $c.BackColor = $p.Bg; $c.ForeColor = $p.Text }
      "Panel" { $c.BackColor = $p.Card; $c.ForeColor = $p.Text }
      "TableLayoutPanel" { $c.BackColor = $p.Bg; $c.ForeColor = $p.Text }
      "Label" { $c.ForeColor = $p.Text }
      "CheckBox" { $c.ForeColor = $p.Sub }
      "Button" {
        $c.BackColor = $p.Card
        $c.ForeColor = $p.Text
        $c.FlatAppearance.BorderColor = $p.Border
        $c.FlatAppearance.BorderSize = 1
      }
      "RichTextBox" { $c.BackColor = $p.Bg; $c.ForeColor = $p.Text; $c.BorderStyle = "None" }
      "DataGridView" {
        $c.BackgroundColor = $p.Bg
        $c.GridColor = $p.Border
      }
      default { $c.ForeColor = $p.Text }
    }
    foreach($child in $c.Controls){ $this.ApplyControl($child) }
  }
  [void]ApplyCard($p){
    $p.BackColor = $this.Palette.Card
    $p.ForeColor = $this.Palette.Text
  }
}

$form = New-Object Windows.Forms.Form
$form.Text = "Schuman Word Generator"
$form.Size = New-Object Drawing.Size(1120, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object Drawing.Size(980, 640)
try {
  $prop = $form.GetType().GetProperty("DoubleBuffered", "NonPublic,Instance")
  if ($prop) { $prop.SetValue($form, $true, $null) }
} catch {}
$form.Font = New-Object Drawing.Font("Segoe UI", 10)

$themeMgr = [ThemeManager]::new($Theme.Dark)

$root = New-Object Windows.Forms.TableLayoutPanel
$root.Dock = "Fill"
$root.Padding = New-Object Windows.Forms.Padding(16,16,16,16)
$root.RowCount = 3
$root.ColumnCount = 1
$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100)))
$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($root)

# Header card
$headerCard = New-Object Windows.Forms.Panel
$headerCard.Dock = "Fill"
$headerCard.Padding = New-Object Windows.Forms.Padding(16,16,16,12)
$headerCard.Margin = New-Object Windows.Forms.Padding(0,0,0,12)
$root.Controls.Add($headerCard, 0, 0)

$headerGrid = New-Object Windows.Forms.TableLayoutPanel
$headerGrid.Dock = "Fill"
$headerGrid.ColumnCount = 2
$headerGrid.RowCount = 2
$headerGrid.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 70)))
$headerGrid.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 30)))
$headerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$headerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$headerCard.Controls.Add($headerGrid)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = "Schuman Word Generator"
$lblTitle.Font = New-Object Drawing.Font("Segoe UI Semibold", 13)
$lblTitle.AutoSize = $true
$headerGrid.Controls.Add($lblTitle, 0, 0)

$btnDashboard = New-Object Windows.Forms.Button
$btnDashboard.Text = "Open Check-in Dashboard"
$btnDashboard.Width = 220
$btnDashboard.Height = 32
$btnDashboard.FlatStyle = "Flat"
$btnDashboard.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)
$btnDashboard.FlatAppearance.BorderSize = 2

$btnForceExcel = New-Object Windows.Forms.Button
$btnForceExcel.Text = "Force Excel Update"
$btnForceExcel.Width = 170
$btnForceExcel.Height = 32
$btnForceExcel.FlatStyle = "Flat"
$btnForceExcel.Font = New-Object Drawing.Font("Segoe UI Semibold", 9)

$headerActions = New-Object Windows.Forms.FlowLayoutPanel
$headerActions.AutoSize = $true
$headerActions.WrapContents = $false
$headerActions.FlowDirection = "LeftToRight"
$headerActions.Anchor = "Top,Right"
$headerActions.Margin = New-Object Windows.Forms.Padding(0,2,0,0)
$headerGrid.Controls.Add($headerActions, 1, 0)

$statusPill = New-Object Windows.Forms.Panel
$statusPill.AutoSize = $true
$statusPill.Padding = New-Object Windows.Forms.Padding(10,4,10,4)
$statusPill.Margin = New-Object Windows.Forms.Padding(8,4,0,0)
$statusPill.BorderStyle = "None"
$headerActions.Controls.Add($btnDashboard)
$headerActions.Controls.Add($btnForceExcel)
$headerActions.Controls.Add($statusPill)

$lblStatusPill = New-Object Windows.Forms.Label
$lblStatusPill.Text = "Idle"
$lblStatusPill.AutoSize = $true
$lblStatusPill.TextAlign = "MiddleCenter"
$statusPill.Controls.Add($lblStatusPill)

$lblStatusText = New-Object Windows.Forms.Label
$lblStatusText.Text = "Ready."
$lblStatusText.AutoSize = $true
$lblStatusText.Margin = New-Object Windows.Forms.Padding(0,2,0,0)
$lblStatusText.TextAlign = "MiddleRight"
$lblStatusText.Anchor = "Top,Right"
$headerGrid.Controls.Add($lblStatusText, 1, 1)

$lblMetrics = New-Object Windows.Forms.Label
$lblMetrics.Text = "Total: 0 | Saved: 0 | Skipped: 0 | Errors: 0"
$lblMetrics.AutoSize = $true
$lblMetrics.Margin = New-Object Windows.Forms.Padding(0,6,0,0)
$headerGrid.Controls.Add($lblMetrics, 0, 1)

# Main content (progress + dynamic area)
$centerGrid = New-Object Windows.Forms.TableLayoutPanel
$centerGrid.Dock = "Fill"
$centerGrid.ColumnCount = 1
$centerGrid.RowCount = 2
$centerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Absolute, 20)))
$centerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100)))
$root.Controls.Add($centerGrid, 0, 1)

$progressHost = New-Object Windows.Forms.Panel
$progressHost.Height = 20
$progressHost.Dock = "Fill"
$progressHost.Margin = New-Object Windows.Forms.Padding(0,0,0,12)
$centerGrid.Controls.Add($progressHost, 0, 0)

$progressFill = New-Object Windows.Forms.Panel
$progressFill.Height = 20
$progressFill.Width = 0
$progressFill.Dock = "Left"
$progressHost.Controls.Add($progressFill)

$listCard = New-Object Windows.Forms.Panel
$listCard.Dock = "Fill"
$listCard.Padding = New-Object Windows.Forms.Padding(15)
$centerGrid.Controls.Add($listCard, 0, 1)

$grid = New-Object Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersHeight = 32
$grid.RowTemplate.Height = 28
try {
  $prop = $grid.GetType().GetProperty("DoubleBuffered", "NonPublic,Instance")
  if ($prop) { $prop.SetValue($grid, $true, $null) }
} catch {}
$listCard.Controls.Add($grid)

[void]$grid.Columns.Add("Row", "Row #")
[void]$grid.Columns.Add("Ticket", "Ticket/RITM")
[void]$grid.Columns.Add("User", "User")
[void]$grid.Columns.Add("File", "Output File")
[void]$grid.Columns.Add("Status", "Status")
[void]$grid.Columns.Add("Message", "Message")
[void]$grid.Columns.Add("Progress", "Progress")
$grid.Columns["Row"].FillWeight = 8
$grid.Columns["Ticket"].FillWeight = 16
$grid.Columns["User"].FillWeight = 16
$grid.Columns["File"].FillWeight = 24
$grid.Columns["Status"].FillWeight = 12
$grid.Columns["Message"].FillWeight = 24
$grid.Columns["Progress"].FillWeight = 10
$grid.Columns["Status"].DefaultCellStyle.Alignment = "MiddleLeft"
$grid.Columns["Progress"].DefaultCellStyle.Alignment = "MiddleRight"

$logPanel = New-Object Windows.Forms.Panel
$logPanel.Dock = "Bottom"
$logPanel.Padding = New-Object Windows.Forms.Padding(15)
$logPanel.Margin = New-Object Windows.Forms.Padding(0,12,0,0)
$logPanel.Height = 0
$logPanel.Visible = $false
$listCard.Controls.Add($logPanel)

$logBox = New-Object Windows.Forms.RichTextBox
$logBox.Dock = "Fill"
$logBox.ReadOnly = $true
$logBox.HideSelection = $false
$logBox.BorderStyle = "None"
$logPanel.Controls.Add($logBox)

# Footer
$footer = New-Object Windows.Forms.Panel
$footer.Dock = "Fill"
$footer.Padding = New-Object Windows.Forms.Padding(16,12,16,12)
$root.Controls.Add($footer, 0, 2)

$footerGrid = New-Object Windows.Forms.TableLayoutPanel
$footerGrid.Dock = "Fill"
$footerGrid.ColumnCount = 1
$footerGrid.RowCount = 2
$footerGrid.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 100)))
$footerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$footerGrid.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
$footer.Controls.Add($footerGrid)

$buttonFlow = New-Object Windows.Forms.FlowLayoutPanel
$buttonFlow.Dock = "Fill"
$buttonFlow.AutoSize = $true
$buttonFlow.WrapContents = $true
$buttonFlow.FlowDirection = "LeftToRight"
$buttonFlow.Padding = New-Object Windows.Forms.Padding(0,2,0,0)
$footerGrid.Controls.Add($buttonFlow, 0, 0)

$btnStart = New-Object Windows.Forms.Button
$btnStart.Text = "Generate Documents"
$btnStart.Width = 170
$btnStart.Height = 30
$btnStart.FlatStyle = "Flat"
$buttonFlow.Controls.Add($btnStart)

$btnStop = New-Object Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Width = 110
$btnStop.Height = 30
$btnStop.FlatStyle = "Flat"
$btnStop.Enabled = $false
$buttonFlow.Controls.Add($btnStop)

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = "Open Output Folder"
$btnOpen.Width = 170
$btnOpen.Height = 30
$btnOpen.FlatStyle = "Flat"
$btnOpen.Enabled = (Test-Path -LiteralPath $OutDir)
$buttonFlow.Controls.Add($btnOpen)

$btnToggleLog = New-Object Windows.Forms.Button
$btnToggleLog.Text = "Show Log"
$btnToggleLog.Width = 110
$btnToggleLog.Height = 30
$btnToggleLog.FlatStyle = "Flat"
$buttonFlow.Controls.Add($btnToggleLog)

$btnCloseAll = New-Object Windows.Forms.Button
$btnCloseAll.Text = "Close All"
$btnCloseAll.Width = 120
$btnCloseAll.Height = 30
$btnCloseAll.FlatStyle = "Flat"
$buttonFlow.Controls.Add($btnCloseAll)

$optionsFlow = New-Object Windows.Forms.FlowLayoutPanel
$optionsFlow.Dock = "Fill"
$optionsFlow.AutoSize = $true
$optionsFlow.WrapContents = $true
$optionsFlow.FlowDirection = "LeftToRight"
$optionsFlow.Padding = New-Object Windows.Forms.Padding(0,6,0,0)
$footerGrid.Controls.Add($optionsFlow, 0, 1)

$chkShowWord = New-Object Windows.Forms.CheckBox
$chkShowWord.Text = "Show Word after generation"
$chkShowWord.AutoSize = $true
$chkShowWord.Checked = $false

$chkDark = New-Object Windows.Forms.CheckBox
$chkDark.Text = "Dark theme"
$chkDark.AutoSize = $true
$chkDark.Checked = $true
$optionsFlow.Controls.Add($chkDark)

$chkSavePdf = New-Object Windows.Forms.CheckBox
$chkSavePdf.Text = "Save PDF"
$chkSavePdf.AutoSize = $true
$chkSavePdf.Checked = $true
$optionsFlow.Controls.Add($chkSavePdf)

$chkSaveDocx = New-Object Windows.Forms.CheckBox
$chkSaveDocx.Text = "Save DOCX"
$chkSaveDocx.AutoSize = $true
$chkSaveDocx.Checked = $true
$optionsFlow.Controls.Add($chkSaveDocx)

$optionsFlow.Controls.Add($chkShowWord)

# Flags de rendimiento:
# FastMode: Reduce el logging detallado para ganar velocidad.
# TurboMode: Usa una plantilla en memoria en lugar de abrir el archivo disco por cada fila (mucho más rápido).
# Performance flags for the worker thread
$script:UseFastMode = $true  # Skips some verbose logging in the worker
$script:UseTurboMode = $true # Uses an in-memory template for faster document creation

$btnOpen.Add_Click({ if (Test-Path -LiteralPath $OutDir) { Start-Process explorer.exe $OutDir } })
$btnForceExcel.Add_Click({
  if (Invoke-AutoExcelRefresh) {
    $lblStatusText.Text = "Excel refreshed."
    Set-StatusPill -text "Ready" -state "idle"
    Append-Log "Force Excel Update completed."
  }
})

$btnDashboard.Add_Click({
  if (-not (Test-Path -LiteralPath $DashboardScript)) {
    $lblStatusText.Text = "Dashboard script not found: $DashboardScript"
    Set-StatusPill -text "Error" -state "error"
    Append-Log "Dashboard script not found: $DashboardScript"
    return
  }
  try {
    Start-Process powershell.exe -ArgumentList @(
      "-NoProfile",
      "-ExecutionPolicy", "Bypass",
      "-File", $DashboardScript,
      "-ExcelPath", $ExcelPath,
      "-SheetName", $PreferredSheet
    ) -WindowStyle Hidden | Out-Null
    Append-Log "Opened dashboard UI."
  }
  catch {
    $lblStatusText.Text = "Failed to open dashboard: $($_.Exception.Message)"
    Set-StatusPill -text "Error" -state "error"
    Append-Log "Failed to open dashboard: $($_.Exception.Message)"
  }
})

function Start-Confetti {
  return
}

$script:StatusState = "idle"
function Set-StatusPill([string]$text, [string]$state){
  $script:StatusState = $state
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  $lblStatusPill.Text = $text
  switch($state){
    "running" { $statusPill.BackColor = $t.Accent }
    "error" { $statusPill.BackColor = $t.Error }
    "done" { $statusPill.BackColor = $t.Success }
    default { $statusPill.BackColor = $t.Border }
  }
  $lblStatusPill.ForeColor = $t.BadgeText
}

function Apply-Theme {
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  $themeMgr.SetPalette($t)
  $themeMgr.ApplyControl($form)

  $headerCard.BackColor = $t.Card
  $listCard.BackColor = $t.Card
  $footer.BackColor = $t.Card
  $logPanel.BackColor = $t.Card

  $progressHost.BackColor = $t.Border
  $progressFill.BackColor = $t.Accent
  $lblStatusText.ForeColor = $t.Sub
  $lblMetrics.ForeColor = $t.Sub

  $grid.ColumnHeadersDefaultCellStyle.BackColor = $t.Card
  $grid.ColumnHeadersDefaultCellStyle.ForeColor = $t.Sub
  $grid.DefaultCellStyle.BackColor = $t.Bg
  $grid.DefaultCellStyle.ForeColor = $t.Text
  $grid.DefaultCellStyle.SelectionBackColor = $t.Shadow
  $grid.DefaultCellStyle.SelectionForeColor = $t.Text
  $grid.AlternatingRowsDefaultCellStyle.BackColor = $t.Card
  $grid.AlternatingRowsDefaultCellStyle.ForeColor = $t.Text

  # Primary/secondary button emphasis for faster scanning.
  $btnStart.BackColor = $t.Accent
  $btnStart.ForeColor = $t.BadgeText
  $btnStart.FlatAppearance.BorderColor = $t.Accent
  $btnStart.FlatAppearance.BorderSize = 1

  # Stop is intentionally red-highlighted for visibility.
  if ($chkDark.Checked) {
    $btnStop.BackColor = [Drawing.Color]::FromArgb(88,36,36)
    $btnStop.ForeColor = [Drawing.Color]::FromArgb(255,210,210)
    $btnStop.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(210,80,80)
    $btnStop.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(110,44,44)
    $btnStop.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(76,30,30)
  } else {
    $btnStop.BackColor = [Drawing.Color]::FromArgb(255,235,235)
    $btnStop.ForeColor = [Drawing.Color]::FromArgb(170,30,30)
    $btnStop.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(210,80,80)
    $btnStop.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(255,220,220)
    $btnStop.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(255,205,205)
  }
  $btnStop.FlatAppearance.BorderSize = 2

  # Close All: danger style
  if ($chkDark.Checked) {
    $btnCloseAll.BackColor = [Drawing.Color]::FromArgb(88,36,36)
    $btnCloseAll.ForeColor = [Drawing.Color]::FromArgb(255,210,210)
    $btnCloseAll.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(210,80,80)
    $btnCloseAll.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(110,44,44)
    $btnCloseAll.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(76,30,30)
  } else {
    $btnCloseAll.BackColor = [Drawing.Color]::FromArgb(255,235,235)
    $btnCloseAll.ForeColor = [Drawing.Color]::FromArgb(170,30,30)
    $btnCloseAll.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(210,80,80)
    $btnCloseAll.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(255,220,220)
    $btnCloseAll.FlatAppearance.MouseDownBackColor = [Drawing.Color]::FromArgb(255,205,205)
  }
  $btnCloseAll.FlatAppearance.BorderSize = 2

  $btnDashboard.BackColor = $t.Card
  $btnDashboard.ForeColor = $t.Accent
  $btnDashboard.FlatAppearance.BorderColor = $t.Accent
  $btnDashboard.FlatAppearance.BorderSize = 2
  $btnDashboard.FlatAppearance.MouseOverBackColor = if ($chkDark.Checked) { [Drawing.Color]::FromArgb(34,44,58) } else { [Drawing.Color]::FromArgb(235,245,255) }
  $btnDashboard.FlatAppearance.MouseDownBackColor = if ($chkDark.Checked) { [Drawing.Color]::FromArgb(28,38,52) } else { [Drawing.Color]::FromArgb(220,238,255) }

  $btnForceExcel.BackColor = $t.Card
  $btnForceExcel.ForeColor = $t.Accent
  $btnForceExcel.FlatAppearance.BorderColor = $t.Accent
  $btnForceExcel.FlatAppearance.BorderSize = 2
  $btnForceExcel.FlatAppearance.MouseOverBackColor = if ($chkDark.Checked) { [Drawing.Color]::FromArgb(34,44,58) } else { [Drawing.Color]::FromArgb(235,245,255) }
  $btnForceExcel.FlatAppearance.MouseDownBackColor = if ($chkDark.Checked) { [Drawing.Color]::FromArgb(28,38,52) } else { [Drawing.Color]::FromArgb(220,238,255) }

  Set-StatusPill -text $lblStatusPill.Text -state $script:StatusState
}

$chkDark.Add_CheckedChanged({ Apply-Theme })
Apply-Theme

$script:RowMap = @{}
$script:ActiveRow = $null

$script:ProgressTarget = 0
$script:ProgressCurrent = 0
$progressTimer = New-Object Windows.Forms.Timer
$progressTimer.Interval = 16
$progressTimer.Add_Tick({
  if($progressHost.Width -le 0){ return }
  $delta = $script:ProgressTarget - $script:ProgressCurrent
  if([Math]::Abs($delta) -lt 0.5){
    $script:ProgressCurrent = $script:ProgressTarget
  } else {
    $script:ProgressCurrent += ($delta * 0.2)
  }
  $pct = [Math]::Max(0, [Math]::Min(1, $script:ProgressCurrent))
  $progressFill.Width = [int]($progressHost.Width * $pct)
})

function Invoke-Ui([scriptblock]$action){
  if($form.InvokeRequired){
    [void]$form.BeginInvoke($action)
  } else {
    & $action
  }
}

function Append-Log([string]$line){
  Invoke-Ui {
    $ts = Get-Date -Format "HH:mm:ss"
    $logBox.AppendText("[$ts] $line`r`n")
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
  }
}

function Invoke-AutoExcelRefresh {
  param([switch]$Quiet)
  $projectRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
  try { $ExcelPath = Resolve-ExcelPath -CurrentExcelPath $ExcelPath } catch {
    $msg = ("Excel validation failed: " + $_.Exception.Message)
    if (-not $Quiet) {
      $lblStatusText.Text = $msg
      Set-StatusPill -text "Error" -state "error"
      Append-Log $msg
    }
    Write-Log $msg
    return $false
  }

  if (-not (Test-Path -LiteralPath $AutoExcelScript)) {
    $msg = "auto-excel.ps1 was not found: $AutoExcelScript"
    if (-not $Quiet) {
      $lblStatusText.Text = $msg
      Set-StatusPill -text "Error" -state "error"
      Append-Log $msg
    }
    Write-Log $msg
    return $false
  }

  if (-not $Quiet) {
    $lblStatusText.Text = "Updating Excel..."
    Set-StatusPill -text "Updating" -state "running"
    Append-Log "Updating Excel before generation..."
  }
  Write-Log "Running auto-excel.ps1 before PDF generation."

  $procArgsLine = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', ('"{0}"' -f $AutoExcelScript),
    '-ExcelPath', ('"{0}"' -f $ExcelPath),
    '-DefaultStartDir', ('"{0}"' -f $projectRoot),
    '-DefaultExcelName', '"Schuman List.xlsx"',
    '-SheetName', ('"{0}"' -f $PreferredSheet),
    '-TicketHeader', '"Number"',
    '-TicketColumn', '"4"',
    '-NameHeader', '"Name"',
    '-PhoneHeader', '"PI"',
    '-ActionHeader', '"Estado de RITM"',
    '-NoPopups'
  ) -join ' '

  $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $procArgsLine -PassThru -WindowStyle Hidden -WorkingDirectory $projectRoot
  $tick = 0
  while (-not $proc.HasExited) {
    Start-Sleep -Milliseconds 200
    [System.Windows.Forms.Application]::DoEvents()
    if (-not $Quiet) {
      $dots = '.' * (($tick % 3) + 1)
      $lblStatusText.Text = ("Updating Excel{0}" -f $dots)
      $tick++
    }
  }

  if ($proc.ExitCode -ne 0) {
    $msg = ("Excel update failed (exit code {0})." -f $proc.ExitCode)
    if (-not $Quiet) {
      $lblStatusText.Text = $msg
      Set-StatusPill -text "Error" -state "error"
      Append-Log $msg
    }
    Write-Log $msg
    return $false
  }

  $script:AutoExcelReady = $true
  $script:AutoExcelWarning = ''
  if (-not $Quiet) {
    $lblStatusText.Text = "Excel ready."
    Set-StatusPill -text "Ready" -state "idle"
    Append-Log "Excel update completed."
    [void](Load-ExcelPreview)
  }
  Write-Log "Excel is ready. Choose DOCX/PDF and click Generate Documents."
  return $true
}

if ($script:AutoExcelWarning) {
  $lblStatusText.Text = "Excel pending update."
  Set-StatusPill -text "Warning" -state "error"
  Append-Log $script:AutoExcelWarning
} else {
  $lblStatusText.Text = "Ready. Click Generate Documents."
  Set-StatusPill -text "Idle" -state "idle"
  Append-Log "Generator loaded. Excel will update on Generate."
}

function Get-IntPropOrDefault($obj, [string]$name, [int]$defaultValue = 0){
  if ($null -eq $obj) { return $defaultValue }
  if ($obj.PSObject -and $obj.PSObject.Properties[$name]) {
    try { return [int]$obj.$name } catch {}
  }
  return $defaultValue
}

function Sanitize-OutputName([string]$s){
  if([string]::IsNullOrWhiteSpace($s)){ return "" }
  $s = $s -replace '[\\/:*?"<>|'']',''
  $s = $s -replace '\s+',' '
  return $s.Trim()
}

function Get-ExpectedOutputPaths([string]$ticket, [string]$user){
  $safeTicket = Sanitize-OutputName $ticket
  $safeUser = Sanitize-OutputName $user
  if([string]::IsNullOrWhiteSpace($safeTicket)){ $safeTicket = "UNKNOWN_TICKET" }
  if([string]::IsNullOrWhiteSpace($safeUser)){ $safeUser = "UNKNOWN_NAME" }
  $docx = Join-Path $OutDir ("{0}_{1}.docx" -f $safeTicket, $safeUser)
  $pdf = [System.IO.Path]::ChangeExtension($docx, ".pdf")
  return [pscustomobject]@{
    Docx = $docx
    Pdf = $pdf
    BaseName = [System.IO.Path]::GetFileName($docx)
  }
}

function Ensure-GridRow([int]$rowId, [string]$fileName, [string]$ticket, [string]$user){
  if($script:RowMap.ContainsKey($rowId)){ return $script:RowMap[$rowId] }
  $r = $grid.Rows.Add()
  $row = $grid.Rows[$r]
  $row.Cells["Row"].Value = $rowId
  $row.Cells["Ticket"].Value = $ticket
  $row.Cells["User"].Value = $user
  $row.Cells["File"].Value = $fileName
  $row.Cells["Status"].Value = "Pending"
  $row.Cells["Message"].Value = "Queued"
  $row.Cells["Progress"].Value = "-"
  $script:RowMap[$rowId] = $row
  return $row
}

function Load-ExcelPreview {
  param([switch]$Silent)

  try { $ExcelPath = Resolve-ExcelPath -CurrentExcelPath $ExcelPath } catch {
    $msg = ("Excel preview failed: " + $_.Exception.Message)
    if (-not $Silent) {
      $lblStatusText.Text = $msg
      Set-StatusPill -text "Error" -state "error"
      Append-Log $msg
    }
    Write-Log $msg
    return $false
  }

  $excel = $null
  $wb = $null
  $ws = $null
  $used = $null
  $previewRows = @()
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $false, $true)
    try { $ws = $wb.Worksheets.Item($PreferredSheet) } catch { $ws = $wb.Worksheets.Item(1) }
    $used = $ws.UsedRange
    $lastRow = [int]($used.Row + $used.Rows.Count - 1)
    if ($lastRow -lt 2) { $lastRow = 1 }

    $ticketCol = 4
    $nameCol = 0
    $piCol = 0
    for ($c = 1; $c -le 30; $c++) {
      $h = ("" + $ws.Cells.Item(1, $c).Text).Trim().ToLowerInvariant()
      if (-not $h) { continue }
      if ($h -eq 'number' -or $h -eq 'ticket' -or $h -eq 'ritm') { $ticketCol = $c }
      if ($nameCol -eq 0 -and ($h -eq 'name' -or $h -eq 'requested for' -or $h -eq 'user')) { $nameCol = $c }
      if ($piCol -eq 0 -and $h -eq 'pi') { $piCol = $c }
    }
    if ($nameCol -eq 0) { $nameCol = 2 }

    for ($r = 2; $r -le $lastRow; $r++) {
      $ticket = ("" + $ws.Cells.Item($r, $ticketCol).Text).Trim()
      if (-not $ticket) { continue }
      $user = ("" + $ws.Cells.Item($r, $nameCol).Text).Trim()
      if (-not $user) { $user = '-' }
      $pi = if ($piCol -gt 0) { ("" + $ws.Cells.Item($r, $piCol).Text).Trim() } else { '' }
      $previewRows += [pscustomobject]@{
        Row    = $r
        Ticket = $ticket
        User   = $user
        PI     = $pi
      }
    }
  }
  catch {
    $msg = ("Excel preview read failed: " + $_.Exception.Message)
    if (-not $Silent) {
      $lblStatusText.Text = $msg
      Set-StatusPill -text "Error" -state "error"
      Append-Log $msg
    }
    Write-Log $msg
    return $false
  }
  finally {
    try { if ($used) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($used) } } catch {}
    try { if ($ws) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($wb) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }

  $grid.Rows.Clear()
  $script:RowMap.Clear()
  $existingDone = 0
  foreach ($item in $previewRows) {
    $piText = if ($item.PI) { $item.PI } else { '-' }
    $paths = Get-ExpectedOutputPaths -ticket $item.Ticket -user $item.User
    $hasPdf = Test-Path -LiteralPath $paths.Pdf
    $hasDocx = Test-Path -LiteralPath $paths.Docx
    $outputLabel = "PI: " + $piText
    if ($hasPdf) {
      $outputLabel = [System.IO.Path]::GetFileName($paths.Pdf)
    } elseif ($hasDocx) {
      $outputLabel = [System.IO.Path]::GetFileName($paths.Docx)
    }
    $row = Ensure-GridRow -rowId ([int]$item.Row) -fileName $outputLabel -ticket $item.Ticket -user $item.User
    if ($hasPdf -or $hasDocx) {
      $existingDone++
      $msg = if ($hasPdf -and $hasDocx) { "PDF/DOCX already generated" } elseif ($hasPdf) { "PDF already generated" } else { "DOCX already generated" }
      Set-RowStatus -row $row -status "Done" -message $msg -state "success"
      $row.Cells["Progress"].Value = "100%"
    } else {
      Set-RowStatus -row $row -status "Ready" -message "Loaded from Excel" -state "normal"
      $row.Cells["Progress"].Value = "-"
    }
  }
  $total = @($previewRows).Count
  $lblMetrics.Text = "Total: $total | Saved: $existingDone | Skipped: 0 | Errors: 0"
  $btnOpen.Enabled = (Test-Path -LiteralPath $OutDir)
  if (-not $Silent) {
    $lblStatusText.Text = "Ready. Preloaded $total rows from Excel."
    Set-StatusPill -text "Ready" -state "idle"
    Append-Log "Preloaded $total rows from Excel. Existing outputs: $existingDone."
  }
  Write-Log ("Load-Data completed. Excel='{0}', rows={1}" -f $ExcelPath, $total)
  return $true
}

# Actualiza el estado visual de una fila en la tabla (Iconos y Colores)
function Set-RowStatus($row, [string]$status, [string]$message, [string]$state){
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  $icon = switch($state){
    "active"  { ">" }
    "success" { "OK" }
    "error"   { "ERROR" }
    "hold"    { "HOLD" }
    "skipped" { "SKIPPED" }
    default   { "" }
  }
  $row.Cells["Status"].Value = "$icon $status"
  $row.Cells["Message"].Value = $message
  switch($state){
    "active" {
      $row.DefaultCellStyle.BackColor = $t.Shadow
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = $t.Accent
    }
    "success" {
      $row.DefaultCellStyle.BackColor = $t.Success
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = $t.Accent
    }
    "error" {
      $row.DefaultCellStyle.BackColor = $t.Error
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = [Drawing.Color]::FromArgb(220,80,80)
    }
    "hold" {
      $row.DefaultCellStyle.BackColor = $t.Card
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = $t.Warning
    }
    "skipped" {
      $row.DefaultCellStyle.BackColor = $t.Border
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = $t.Sub
    }
    default {
      $row.DefaultCellStyle.BackColor = $t.Bg
      $row.DefaultCellStyle.ForeColor = $t.Text
      $row.Cells["Status"].Style.ForeColor = $t.Sub
    }
  }
}

# Additional visual status mapping for the Status column text itself.
$grid.Add_CellFormatting({
  param($sender, $e)
  try {
    if ($e.RowIndex -lt 0) { return }
    if ($e.ColumnIndex -ne $grid.Columns["Status"].Index) { return }
    $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
    $txt = ("" + $grid.Rows[$e.RowIndex].Cells["Status"].Value).ToLowerInvariant()
    if ($txt -match 'hold') {
      $e.CellStyle.ForeColor = $t.Warning
    }
    elseif ($txt -match 'error') {
      $e.CellStyle.ForeColor = [Drawing.Color]::FromArgb(220,80,80)
    }
    elseif ($txt -match 'ok|done|complete') {
      $e.CellStyle.ForeColor = $t.Accent
    }
    elseif ($txt -match 'pending|processing|running') {
      $e.CellStyle.ForeColor = $t.Accent
    }
  } catch {}
})

function Show-Toast([string]$title, [string]$body){
  return
}

$btnToggleLog.Add_Click({
  if($logPanel.Height -eq 0){
    $logPanel.Height = 160
    $logPanel.Visible = $true
    $btnToggleLog.Text = "Hide Log"
  } else {
    $logPanel.Height = 0
    $logPanel.Visible = $false
    $btnToggleLog.Text = "Show Log"
  }
})

# ----------------------------
# Worker communication
# ----------------------------
$script:SyncHash = [hashtable]::Synchronized(@{
  Running = $false
  Cancel  = $false
  Status  = "Ready"
  UiEvents = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
  Result = $null
  Error  = $null
})

$script:PSInstance = $null
$script:PSAsync = $null
$script:RunStarted = $null
$script:LastCounters = @{
  Total = 0
  Saved = 0
  Skipped = 0
  Errors = 0
}

# ----------------------------
# Worker logic (STA runspace) - SAFE PowerShell execution
# ----------------------------
# Este bloque de script se ejecuta en un hilo separado (Runspace).
# Es necesario usar STA (Single Threaded Apartment) para que la automatización COM de Office funcione correctamente.
# La comunicación con la GUI se hace a través de la variable sincronizada $SyncHash.
$script:WorkerLogic = {
  param($SyncHash, $Config)

  # --- Constants ---
  $wdFormatDOCX = 16
  $wdFormatPDF = 17
  $xlCalculationManual = -4135

  function WriteLog($Path, $Msg) {
    try {
      $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      Add-Content -LiteralPath $Path -Value "[$ts] $Msg" -Encoding UTF8
    } catch {}
  }
  function Release-Com($o){ if($o){ try{[Runtime.InteropServices.Marshal]::ReleaseComObject($o)|Out-Null}catch{} } }
  function Normalize([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $n = $s.Trim().ToLowerInvariant()
    $n = $n -replace '^field',''
    $n = $n -replace '[\s_\-]',''
    return $n
  }
  function Sanitize([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $s = $s -replace '[\\/:*?"<>|'']',''
    $s = $s -replace '\s+',' '
    return $s.Trim()
  }
  function Get-UniquePath([string]$Dir, [string]$BaseName){
    $base = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)
    $ext  = [System.IO.Path]::GetExtension($BaseName)
    if([string]::IsNullOrWhiteSpace($ext)){ $ext = ".docx" }
    $candidate = Join-Path $Dir ($base + $ext)
    $i = 2
    while(Test-Path -LiteralPath $candidate){
      $candidate = Join-Path $Dir ("$base ($i)$ext")
      $i++
    }
    return $candidate
  }

  # Inspecciona la plantilla de Word para listar qué marcadores o campos existen.
  # Útil para depuración si los datos no se están rellenando donde se espera.
  function Inspect-Template($WordApp, [string]$TemplatePath, [string]$LogPath){
    $doc = $null
    try {
      WriteLog $LogPath "Template path: $TemplatePath"
      $doc = $WordApp.Documents.Open($TemplatePath, $false, $true, $false)
      WriteLog $LogPath "Template inspection: opened"

      try {
        $ffNames = @()
        foreach($ff in $doc.FormFields){ $ffNames += $ff.Name }
        WriteLog $LogPath ("Template FormFields (" + $ffNames.Count + "): " + ($ffNames -join ", "))
      } catch { WriteLog $LogPath "Template FormFields: (error reading)" }

      try {
        $ccNames = @()
        foreach($cc in $doc.ContentControls){
          $ccNames += ("Title='" + $cc.Title + "' Tag='" + $cc.Tag + "'")
        }
        WriteLog $LogPath ("Template ContentControls (" + $ccNames.Count + "): " + ($ccNames -join ", "))
      } catch { WriteLog $LogPath "Template ContentControls: (error reading)" }

      try {
        $bmNames = @()
        foreach($bm in $doc.Bookmarks){ $bmNames += $bm.Name }
        WriteLog $LogPath ("Template Bookmarks (" + $bmNames.Count + "): " + ($bmNames -join ", "))
      } catch { WriteLog $LogPath "Template Bookmarks: (error reading)" }

      try {
        $text = [string]$doc.Content.Text
        $matches = [regex]::Matches($text, '\bField[A-Za-z0-9_]+\b')
        $uniq = @($matches | ForEach-Object { $_.Value } | Sort-Object -Unique)
        if($uniq.Count -gt 0){
          WriteLog $LogPath ("Template Text Placeholders (" + $uniq.Count + "): " + ($uniq -join ", "))
        } else {
          WriteLog $LogPath "Template Text Placeholders: (none)"
        }
        $matches2 = [regex]::Matches($text, '(\{\{|\[|<<)\s*Field[A-Za-z0-9_]+\s*(\}\}|\]|>>)')
        $uniq2 = @($matches2 | ForEach-Object { $_.Value } | Sort-Object -Unique)
        if($uniq2.Count -gt 0){
          WriteLog $LogPath ("Template Token Placeholders (" + $uniq2.Count + "): " + ($uniq2 -join ", "))
        }
      } catch { WriteLog $LogPath "Template Text Placeholders: (error scanning)" }
    }
    catch {
      WriteLog $LogPath ("Template inspection failed: " + $_.Exception.Message)
    }
    finally {
      try { if($doc){ $doc.Close($false) | Out-Null } } catch {}
      Release-Com $doc
    }
  }

  function Log-DocPlaceholders($Doc, [string]$LogPath, [string]$Prefix){
    try {
      $text = [string]$Doc.Content.Text
      $matches = [regex]::Matches($text, '\bField[A-Za-z0-9_]+\b')
      $uniq = @($matches | ForEach-Object { $_.Value } | Sort-Object -Unique)
      if($uniq.Count -gt 0){
        WriteLog $LogPath ("${Prefix}: Doc Text Placeholders (" + $uniq.Count + "): " + ($uniq -join ", "))
      } else {
        WriteLog $LogPath "${Prefix}: Doc Text Placeholders: (none)"
      }
    } catch {
      WriteLog $LogPath "${Prefix}: Doc Text Placeholders: (error scanning)"
    }
  }

  # Función principal de reemplazo. Intenta inyectar el valor usando varias estrategias en orden:
  # 1. ContentControls (Cajas de texto modernas en Word)
  # 2. Bookmarks (Marcadores clásicos)
  # 3. FormFields (Campos de formulario heredados)
  # 4. Búsqueda y Reemplazo de texto plano (ej: {{FieldTicketNumber}})
  function Set-WordPlaceholderValue($Doc, [string]$Key, [string]$Value, [string]$LogPath, [string]$LogPrefix, [bool]$FastMode){
    $changed = $false
    $method = "NotFound"
    $replaceCount = 0

    function Count-Occurrences([string]$Text, [string]$Pattern, [bool]$WholeWord){
      try {
        if([string]::IsNullOrEmpty($Text)){ return 0 }
        $escaped = [regex]::Escape($Pattern)
        $regex = if($WholeWord){ "(?<!\\w)$escaped(?!\\w)" } else { $escaped }
        return ([regex]::Matches($Text, $regex)).Count
      } catch { return 0 }
    }
    function Replace-InRangeFast($Range, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        $text = [string]$Range.Text
        $c = Count-Occurrences -Text $text -Pattern $Pattern -WholeWord $WholeWord
        if($c -gt 0){
          $Count.Value += $c
          $find = $Range.Find
          $find.ClearFormatting()
          $find.Replacement.ClearFormatting()
          $find.Text = $Pattern
          $find.Replacement.Text = $ReplaceValue
          $find.MatchCase = $false
          $find.MatchWholeWord = $WholeWord
          $find.MatchWildcards = $false
          $find.Wrap = 0
          $find.Forward = $true
          [void]$find.Execute($Pattern,$false,$false,$false,$false,$false,$true,1,$false,$ReplaceValue,2)
        }
      } catch {}
    }
    function Replace-InHeadersFooters($DocRef, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        foreach($sec in $DocRef.Sections){
          foreach($hf in @($sec.Headers, $sec.Footers)){
            foreach($item in @($hf.Item(1), $hf.Item(2), $hf.Item(3))){
              try {
                Replace-InRangeFast -Range $item.Range -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
              } catch {}
            }
          }
        }
      } catch {}
    }

    # ContentControls by exact Title/Tag
    try {
      $hits = 0
      foreach($cc in $Doc.ContentControls){
        if($cc.Title -eq $Key -or $cc.Tag -eq $Key){
          $cc.Range.Text = $Value
          $hits++
        }
      }
      if($hits -gt 0){
        $changed = $true
        $method = "ContentControl"
        $replaceCount = $hits
      }
    } catch {}

    # Bookmarks by exact name
    if(-not $changed){
      try {
        if($Doc.Bookmarks.Exists($Key)){
          $bm = $Doc.Bookmarks.Item($Key)
          $bm.Range.Text = $Value
          $Doc.Bookmarks.Add($Key, $bm.Range) | Out-Null
          $changed = $true
          $method = "Bookmark"
          $replaceCount = 1
        }
      } catch {}
    }

    # Legacy FormFields by exact name
    if(-not $changed){
      try {
        $ff = $Doc.FormFields.Item($Key)
        if($ff){
          $ff.Result = $Value
          $changed = $true
          $method = "FormField"
          $replaceCount = 1
        }
      } catch {}
    }

    # Literal Find/Replace for plain text placeholder (fast path)
    if(-not $changed){
      try {
        $tokens = @(
          $Key,
          "{{${Key}}}",
          "{${Key}}",
          "<<${Key}>>",
          "[${Key}]",
          "[[${Key}]]"
        )
        foreach($t in $tokens){
          Replace-InRangeFast -Range $Doc.Content -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
          Replace-InHeadersFooters -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
        }
        if($replaceCount -eq 0){
          foreach($t in $tokens){
            Replace-InRangeFast -Range $Doc.Content -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
            Replace-InHeadersFooters -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
          }
        }
        if($replaceCount -gt 0){
          $changed = $true
          $method = "FindReplace"
        }
      } catch {}
    }

    if(-not $FastMode){
      WriteLog $LogPath ("${LogPrefix}: Set $Key -> $changed via $method (replacements=$replaceCount)")
    }
    return [pscustomobject]@{
      Success = [bool]$changed
      MethodUsed = $method
      ReplaceCount = [int]$replaceCount
    }
  }
  function GetCol($Sheet, [string]$Header){
    $used = $Sheet.UsedRange
    $cols = $used.Columns.Count
    for($c=1;$c -le $cols;$c++){
      $h = ([string]$Sheet.Cells.Item(1,$c).Text).Trim()
      if($h -eq $Header){ return $c }
    }
    return $null
  }
  function Normalize-Header([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $n = $s.ToLowerInvariant().Trim()
    $n = [regex]::Replace($n, '[^a-z0-9\s]', '')
    $n = [regex]::Replace($n, '\s+', ' ')
    return $n
  }
  function Find-ColumnByKeywords($Sheet, [string[]]$Keywords){
    $used = $Sheet.UsedRange
    $cols = $used.Columns.Count
    $bestScore = 0
    $bestCol = $null
    for($c=1;$c -le $cols;$c++){
      $h = Normalize-Header ([string]$Sheet.Cells.Item(1,$c).Text)
      if([string]::IsNullOrWhiteSpace($h)){ continue }
      $score = 0
      foreach($k in $Keywords){
        if($h -like "*$k*"){ $score++ }
      }
      if($score -gt $bestScore){
        $bestScore = $score
        $bestCol = $c
      }
    }
    return [pscustomobject]@{ Col=$bestCol; Score=$bestScore }
  }
  function Get-OrCreateColumn($Sheet, [string]$StandardHeader, [string[]]$Keywords){
    $res = Find-ColumnByKeywords $Sheet $Keywords
    if($res.Col -and $res.Score -ge 2){
      return [pscustomobject]@{ Col=$res.Col; Created=$false; Header=$StandardHeader; Score=$res.Score }
    }
    $used = $Sheet.UsedRange
    $newCol = $used.Columns.Count + 1
    $Sheet.Cells.Item(1,$newCol).Value2 = $StandardHeader
    return [pscustomobject]@{ Col=$newCol; Created=$true; Header=$StandardHeader; Score=0 }
  }
  function Get-ProtectionTypeName([int]$pt){
    switch($pt){
      0 { "wdNoProtection" }
      1 { "wdAllowOnlyRevisions" }
      2 { "wdAllowOnlyComments" }
      3 { "wdAllowOnlyFormFields" }
      4 { "wdAllowOnlyReading" }
      5 { "wdAllowOnlyFormFields" }
      default { "Unknown" }
    }
  }

  $excel=$null; $wb=$null; $sheet=$null; $word=$null; $doc=$null; $templateDoc=$null
  $excelCalc = $null
  $saved=0; $skipped=0; $errors=0; $total=0
  $templateInspected = $false

  try {
    if (-not (Test-Path -LiteralPath $Config.ExcelPath)) { throw "Missing Excel: $($Config.ExcelPath)" }
    if (-not (Test-Path -LiteralPath $Config.TemplatePath)) { throw "Missing template: $($Config.TemplatePath)" }
    New-Item -ItemType Directory -Force -Path $Config.OutDir | Out-Null

    # Unblock files if they came from the internet zone
    try { Unblock-File -LiteralPath $Config.ExcelPath -ErrorAction SilentlyContinue } catch {}
    try { Unblock-File -LiteralPath $Config.TemplatePath -ErrorAction SilentlyContinue } catch {}

    WriteLog $Config.LogPath "=== RUN START ==="

    $fast = [bool]($Config.FastMode -or $Config.TurboMode)

    $SyncHash.Status = "Opening Excel"
    WriteLog $Config.LogPath "Opening Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      # Desactivar actualizaciones de pantalla y cálculo automático acelera enormemente la lectura/escritura
      $excel.ScreenUpdating = $false
      $excel.EnableEvents = $false
      $excelCalc = $excel.Calculation
      $excel.Calculation = $xlCalculationManual
    } catch {}
    # Open writable so we can update status/PDF path
    $wb = $excel.Workbooks.Open($Config.ExcelPath, $null, $false)
    WriteLog $Config.LogPath "Excel opened"

    try { $sheet = $wb.Worksheets.Item($Config.PreferredSheet) } catch { $sheet = $wb.Worksheets.Item(1) }
    WriteLog $Config.LogPath ("Using sheet: " + $sheet.Name)

    $nameCol   = GetCol $sheet "Name"
    $ticketCol = GetCol $sheet "Ticket"
    if(-not $ticketCol){ $ticketCol = GetCol $sheet "Number" }
    if(-not $ticketCol){
      $ticketCol = 4
      WriteLog $Config.LogPath "Header 'Ticket/Number' not found. Falling back to column 4."
    }
    $piCol     = GetCol $sheet "PI"
    $equipCol  = GetCol $sheet "Receive ID Equipment" # optional

    $statusInfo = Get-OrCreateColumn $sheet "Export Status" @("status","export","result","done","ok")
    $docxInfo   = Get-OrCreateColumn $sheet "DOCX File"     @("docx","word","generated","output","file")
    $pdfInfo    = Get-OrCreateColumn $sheet "PDF File"      @("pdf","export","generated","output","file")

    $statusCol = $statusInfo.Col
    $docxCol   = $docxInfo.Col
    $pdfCol    = $pdfInfo.Col

    WriteLog $Config.LogPath ("Excel column status: Col=" + $statusCol + " Created=" + $statusInfo.Created + " Score=" + $statusInfo.Score)
    WriteLog $Config.LogPath ("Excel column docx:   Col=" + $docxCol + " Created=" + $docxInfo.Created + " Score=" + $docxInfo.Score)
    WriteLog $Config.LogPath ("Excel column pdf:    Col=" + $pdfCol + " Created=" + $pdfInfo.Created + " Score=" + $pdfInfo.Score)

    if(-not $nameCol){ WriteLog $Config.LogPath "Header 'Name' not found. FieldDisplayName will be left blank." }
    if(-not $piCol){ WriteLog $Config.LogPath "Header 'PI' not found. FieldPINumber will be left blank." }

    $used = $sheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $total = [Math]::Max(0, $lastRow - 1)
    WriteLog $Config.LogPath ("Rows detected: " + $total)

    $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Init"; Total=$total })
    $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=0; Skipped=0; Errors=0 })

    $SyncHash.Status = "Opening Word"
    if(-not $fast){ WriteLog $Config.LogPath "Opening Word" }
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Config.ShowWord
    $word.DisplayAlerts = 0
    try {
      $word.ScreenUpdating = $false
      $word.Options.ConfirmConversions = $false
      $word.Options.SaveNormalPrompt = $false
      $word.Options.BackgroundSave = $false
      $word.Options.AllowFastSave = $false
      $word.Options.UpdateLinksAtOpen = $false
      $word.Options.CheckSpellingAsYouType = $false
      $word.Options.CheckGrammarAsYouType = $false
    } catch {}
    try {
      # 3 = msoAutomationSecurityForceDisable (avoid macro prompts)
      $word.AutomationSecurity = 3
    } catch {}
    if(-not $fast){ WriteLog $Config.LogPath "Word opened" }

    # En modo Turbo, abrimos la plantilla una sola vez y la mantenemos en memoria ($templateDoc).
    # Para cada fila, copiamos el contenido de $templateDoc a un documento nuevo en blanco.
    if($Config.TurboMode){
      try {
        $templateDoc = $word.Documents.Open($Config.TemplatePath, $false, $true, $false)
        if(-not $fast){ WriteLog $Config.LogPath "Template opened in memory (Turbo)" }
      } catch {
        WriteLog $Config.LogPath ("Turbo template open failed: " + $_.Exception.Message)
        $templateDoc = $null
      }
    }

    if(-not $templateInspected -and -not $fast){
      Inspect-Template -WordApp $word -TemplatePath $Config.TemplatePath -LogPath $Config.LogPath
      $templateInspected = $true
    }

    for($r=2;$r -le $lastRow;$r++){
      if($SyncHash.Cancel){ break }

      $name   = if($nameCol){ [string]$sheet.Cells.Item($r,$nameCol).Text } else { "" }
      $ticket = if($ticketCol){ [string]$sheet.Cells.Item($r,$ticketCol).Text } else { "" }
      $pi     = if($piCol){ [string]$sheet.Cells.Item($r,$piCol).Text } else { "" }

      if([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($ticket) -and [string]::IsNullOrWhiteSpace($pi)){
        continue
      }
      $equipment = "Laptop"
      if($equipCol){
        $tmp = [string]$sheet.Cells.Item($r,$equipCol).Text
        if(-not [string]::IsNullOrWhiteSpace($tmp)){ $equipment = $tmp.Trim() }
      }

      $safeTicket = Sanitize $ticket
      $safeName   = Sanitize $name
      if([string]::IsNullOrWhiteSpace($safeTicket)){ $safeTicket="UNKNOWN_TICKET" }
      if([string]::IsNullOrWhiteSpace($safeName)){ $safeName="UNKNOWN_NAME" }

      $fileName = "${safeTicket}_${safeName}.docx"
      $filePath = Join-Path $Config.OutDir $fileName
      $pdfPath = [System.IO.Path]::ChangeExtension([string]$filePath, ".pdf")

      $hasDocx = Test-Path -LiteralPath $filePath
      $hasPdf = Test-Path -LiteralPath $pdfPath
      $docxReady = (-not [bool]$Config.ExportDocx) -or $hasDocx
      $pdfReady = (-not [bool]$Config.ExportPdf) -or $hasPdf
      if ($docxReady -and $pdfReady) {
        $skipped++
        try {
          $sheet.Cells.Item($r,$statusCol).Value2 = "OK (existing)"
          if ($docxCol) { $sheet.Cells.Item($r,$docxCol).Value2 = if ($hasDocx) { [string]$filePath } else { "" } }
          if ($pdfCol) { $sheet.Cells.Item($r,$pdfCol).Value2 = if ($hasPdf) { [string]$pdfPath } else { "" } }
        } catch {}
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{
          Type="RowSkip"; Row=$r; File=$fileName; Ticket=$ticket; User=$name; Message="Already generated"
        })
        WriteLog $Config.LogPath "Row ${r} skipped: existing output(s) detected."
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors })
        continue
      }

      # UI row start
      $rowStart = Get-Date
      $SyncHash.Status = "Saving $fileName"
      $SyncHash.UiEvents.Enqueue([pscustomobject]@{
        Type="RowStart"; Row=$r; File=$fileName; Ticket=$ticket; User=$name
      })
      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Creating file..." })
      WriteLog $Config.LogPath "Row ${r} start: $fileName"
      if(-not $fast){
        WriteLog $Config.LogPath "Row ${r} values: Name='$name' Ticket='$ticket' PI='$pi' Equipment='$equipment'"
      }

      try {
        if($Config.TurboMode){
          $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Creating from template..." })
          # Crear documento en blanco e inyectar contenido formateado (más rápido que IO de disco)
          $doc = $word.Documents.Add()
          if($templateDoc){
            try {
              $doc.Range.FormattedText = $templateDoc.Range.FormattedText
            } catch {
              $doc = $word.Documents.Add($Config.TemplatePath)
            }
          } else {
            $doc = $word.Documents.Add($Config.TemplatePath)
          }
          if(-not $fast){ WriteLog $Config.LogPath "Row ${r}: Doc created in memory" }
        } else {
          # Modo estándar: Copiar archivo físico y abrir la copia. Más seguro para formatos complejos.
          # overwrite target by copying template first (avoids SaveAs/SaveAs2 COM issues)
          if(Test-Path -LiteralPath $filePath){ Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue }
          Copy-Item -LiteralPath $Config.TemplatePath -Destination $filePath -Force

          WriteLog $Config.LogPath "Row ${r}: Opening copied doc"
          $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Opening file..." })
          # Open existing file (not template) so we can Save() directly
          $doc = $word.Documents.Open($filePath, $false, $false, $false)
          WriteLog $Config.LogPath "Row ${r}: Doc opened"
        }
        if($SyncHash.Cancel){
          try { $doc.Close($false) | Out-Null } catch {}
          Release-Com $doc; $doc=$null
          break
        }
        if(-not $fast){
          Log-DocPlaceholders -Doc $doc -LogPath $Config.LogPath -Prefix "Row ${r}"
        }

        if(-not $Config.UltraMode){
          try {
            $prot = $doc.ProtectionType
            $protName = Get-ProtectionTypeName $prot
            WriteLog $Config.LogPath "Row ${r}: ProtectionType=$prot ($protName)"
            if($prot -ne 0){
              WriteLog $Config.LogPath "Row ${r}: Document protected, attempting unprotect"
              try {
                $doc.Unprotect() | Out-Null
              } catch {
                WriteLog $Config.LogPath ("Row ${r}: Unprotect failed: " + $_.Exception.Message)
              }
              try {
                $prot2 = $doc.ProtectionType
                $prot2Name = Get-ProtectionTypeName $prot2
                WriteLog $Config.LogPath "Row ${r}: Unprotect attempted, ProtectionType now=$prot2 ($prot2Name)"
              } catch {
                WriteLog $Config.LogPath "Row ${r}: Unprotect attempted, re-check failed"
              }
            }
          } catch { WriteLog $Config.LogPath "Row ${r}: Protection check failed" }
        }

        # Mapeo de datos: Nombre de campo en Word => Variable de PowerShell
        # Forced mapping (extend this hashtable to support additional placeholders)
        $forced = @{
          "FieldDisplayName"  = $name
          "FieldTicketNumber" = $ticket
          "FieldPINumber"     = $pi
          "FieldITEquipment"  = $equipment
        }

        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Filling fields..." })
        foreach($key in $forced.Keys){
          if($SyncHash.Cancel){ break }
          [void](Set-WordPlaceholderValue -Doc $doc -Key $key -Value $forced[$key] -LogPath $Config.LogPath -LogPrefix "Row ${r}" -FastMode $fast)
        }
        if($SyncHash.Cancel){
          try { $doc.Close($false) | Out-Null } catch {}
          Release-Com $doc; $doc=$null
          break
        }

        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Saving..." })
        if($Config.ExportDocx){
          if($Config.TurboMode){
            try { if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue } } catch {}
            $doc.SaveAs2([string]$filePath, $wdFormatDOCX)
          } else {
            WriteLog $Config.LogPath "Row ${r}: Saving"
            $doc.Save()
            WriteLog $Config.LogPath "Row ${r}: Saved"
          }
          $sheet.Cells.Item($r,$docxCol).Value2 = [string]$filePath
        } else {
          $sheet.Cells.Item($r,$docxCol).Value2 = ""
        }

        if($Config.ExportPdf){
          # Export PDF
          try {
            $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Exporting PDF..." })
            $doc.ExportAsFixedFormat($pdfPath, $wdFormatPDF)
            WriteLog $Config.LogPath "Row ${r}: PDF saved -> $pdfPath"
            $sheet.Cells.Item($r,$statusCol).Value2 = "OK"
            $sheet.Cells.Item($r,$pdfCol).Value2 = $pdfPath
          } catch {
            WriteLog $Config.LogPath "Row ${r}: PDF export failed -> $($_.Exception.Message)"
            $sheet.Cells.Item($r,$statusCol).Value2 = "FAILED: PDF export"
            $sheet.Cells.Item($r,$pdfCol).Value2 = ""
          }
        } else {
          if($sheet.Cells.Item($r,$statusCol).Value2 -ne "FAILED: PDF export"){
            $sheet.Cells.Item($r,$statusCol).Value2 = "OK"
          }
          $sheet.Cells.Item($r,$pdfCol).Value2 = ""
        }

        $doc.Close($false)
        Release-Com $doc; $doc=$null

        $saved++
        try {
          $ms = [int]((Get-Date) - $rowStart).TotalMilliseconds
        if(-not $fast){
          WriteLog $Config.LogPath ("Row ${r}: Done in ${ms} ms")
        }
        } catch {}
        WriteLog $Config.LogPath "Saved: $filePath"
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowDone"; Row=$r; File=$fileName; Ok=$true; Message="Saved" })
      }
      catch {
        $errors++
        WriteLog $Config.LogPath "Row $r ERROR: $($_.Exception.Message)"
        try {
          if($statusCol){ $sheet.Cells.Item($r,$statusCol).Value2 = ("FAILED: " + $_.Exception.Message) }
          if($docxCol){ $sheet.Cells.Item($r,$docxCol).Value2 = [string]$filePath }
          if($pdfCol){ $sheet.Cells.Item($r,$pdfCol).Value2 = "" }
        } catch {}
        try { if($doc){ $doc.Close($false) | Out-Null } } catch {}
        Release-Com $doc; $doc=$null
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowDone"; Row=$r; File=$fileName; Ok=$false; Error=$_.Exception.Message; Message="Error" })
      }

      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors })
    }

    try {
      $wb.Save()
      WriteLog $Config.LogPath "Excel saved"
    } catch {
      WriteLog $Config.LogPath ("Excel save failed: " + $_.Exception.Message)
    }

    $SyncHash.Result = [pscustomobject]@{ Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors }
    WriteLog $Config.LogPath "=== RUN END === Total=$total Saved=$saved Skipped=$skipped Errors=$errors"
  }
  catch {
    $SyncHash.Error = $_.Exception
    WriteLog $Config.LogPath ("FATAL: " + $_.Exception.ToString())
  }
  finally {
    try { if($wb){ $wb.Close($false) | Out-Null } } catch {}
    try { if($excel -and $excelCalc -ne $null){ $excel.Calculation = $excelCalc } } catch {}
    try { if($excel){ $excel.Quit() | Out-Null } } catch {}
    try { if($word){ $word.Quit() | Out-Null } } catch {}

    Release-Com $templateDoc
    Release-Com $sheet
    Release-Com $wb
    Release-Com $excel
    Release-Com $word

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $SyncHash.Running = $false
  }
}

# ----------------------------
# UI Timer consumes events
# ----------------------------
$uiTimer = New-Object Windows.Forms.Timer
$uiTimer.Interval = 80
$uiTimer.Add_Tick({
  # Watchdog: if runspace already ended but SyncHash stayed Running=true, unstick UI.
  if ($script:PSInstance -and $script:SyncHash.Running) {
    try {
      $state = $script:PSInstance.InvocationStateInfo.State
      if ($state -in @('Completed', 'Failed', 'Stopped')) {
        if ($state -eq 'Failed' -and -not $script:SyncHash.Error) {
          $script:SyncHash.Error = [System.Exception]::new("Generation worker ended with state: $state")
        }
        $script:SyncHash.Running = $false
      }
    } catch {}
  }

  while($script:SyncHash.UiEvents.Count -gt 0){
    $p = $script:SyncHash.UiEvents.Dequeue()
    switch($p.Type){
      "Init" {
        $grid.Rows.Clear()
        $script:RowMap.Clear()
        $lblMetrics.Text = "Total: $($p.Total) | Saved: 0 | Skipped: 0 | Errors: 0"
        $script:ProgressTarget = 0
        $script:ProgressCurrent = 0
        Set-StatusPill -text "Running" -state "running"
        Append-Log "Run initialized. Total rows: $($p.Total)"
      }
      "Counters" {
        $totalC = Get-IntPropOrDefault $p "Total" 0
        $savedC = Get-IntPropOrDefault $p "Saved" 0
        $skippedC = Get-IntPropOrDefault $p "Skipped" 0
        $errorsC = Get-IntPropOrDefault $p "Errors" 0
        $done = $savedC + $skippedC + $errorsC
        $lblMetrics.Text = "Total: $totalC | Saved: $savedC | Skipped: $skippedC | Errors: $errorsC"
        $script:ProgressTarget = if($totalC -gt 0){ [double]$done / [double]$totalC } else { 0 }
        $script:LastCounters.Total = $totalC
        $script:LastCounters.Saved = $savedC
        $script:LastCounters.Skipped = $skippedC
        $script:LastCounters.Errors = $errorsC
      }
      "RowStart" {
        if($script:ActiveRow -and $script:RowMap.ContainsKey($script:ActiveRow)){
          $prev = $script:RowMap[$script:ActiveRow]
          if($prev.Cells["Status"].Value -eq "Processing"){
            Set-RowStatus -row $prev -status "Pending" -message "Queued" -state "normal"
          }
        }
        $row = Ensure-GridRow -rowId $p.Row -fileName $p.File -ticket $p.Ticket -user $p.User
        Set-RowStatus -row $row -status "Processing" -message "Creating..." -state "active"
        $row.Cells["Progress"].Value = "0%"
        $script:ActiveRow = $p.Row
        Append-Log "Row $($p.Row) started: $($p.File)"
      }
      "RowStage" {
        if($script:RowMap.ContainsKey($p.Row)){
          $row = $script:RowMap[$p.Row]
          $row.Cells["Message"].Value = $p.Stage
          $row.Cells["Progress"].Value = "Working"
        }
      }
      "RowDone" {
        if($script:RowMap.ContainsKey($p.Row)){
          $row = $script:RowMap[$p.Row]
          if($p.Ok){
            Set-RowStatus -row $row -status "Done" -message "Saved" -state "success"
            $row.Cells["Progress"].Value = "100%"
          } else {
            Set-RowStatus -row $row -status "Error" -message $p.Error -state "error"
            $row.Cells["Progress"].Value = "Failed"
          }
          Append-Log "Row $($p.Row): $($row.Cells["Status"].Value) - $($row.Cells["Message"].Value)"
        }
      }
      "RowSkip" {
        $row = Ensure-GridRow -rowId $p.Row -fileName $p.File -ticket $p.Ticket -user $p.User
        if (("" + $p.Message).ToLowerInvariant().Contains("already")) {
          Set-RowStatus -row $row -status "Done" -message "Already generated" -state "success"
          $row.Cells["Progress"].Value = "100%"
          Append-Log "Row $($p.Row) already generated."
        } else {
          Set-RowStatus -row $row -status "Skipped" -message $p.Message -state "skipped"
          $row.Cells["Progress"].Value = "-"
          Append-Log "Row $($p.Row) skipped: $($p.Message)"
        }
      }
    }
  }

  if(-not $script:SyncHash.Running -and $script:PSInstance){
    $uiTimer.Stop()
    $progressTimer.Stop()
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnOpen.Enabled = $true

    $script:PSInstance.Dispose()
    $script:PSInstance = $null
    $script:PSAsync = $null

    if($script:SyncHash.Error){
      $lblStatusText.Text = "FAILED: " + $script:SyncHash.Error.Message
      Set-StatusPill -text "Error" -state "error"
      Show-Toast -title "Generation failed" -body "See log for details."
      Append-Log "Run failed: $($script:SyncHash.Error.Message)"
    } else {
      $res = $script:SyncHash.Result
      $savedR = Get-IntPropOrDefault $res "Saved" 0
      $skippedR = Get-IntPropOrDefault $res "Skipped" 0
      $errorsR = Get-IntPropOrDefault $res "Errors" 0
      $lblStatusText.Text = "Completed. Saved=$savedR Skipped=$skippedR Errors=$errorsR"
      Set-StatusPill -text "Completed" -state "done"
      Show-Toast -title "Generation completed" -body "Saved: $savedR, Skipped: $skippedR, Errors: $errorsR"
      if($errorsR -eq 0){ Start-Confetti }
      Append-Log "Run completed: Saved=$savedR Skipped=$skippedR Errors=$errorsR"
    }
  }
})

# ----------------------------
# Buttons
# ----------------------------
$btnStart.Add_Click({
  # Generate should run immediately; Excel refresh is handled manually via "Force Excel Update".
  if(-not (Test-Path -LiteralPath $ExcelPath)){
    $lblStatusText.Text = "Excel not found: $ExcelPath"
    Set-StatusPill -text "Error" -state "error"
    Append-Log "Excel not found: $ExcelPath"
    return
  }
  if(-not (Test-Path -LiteralPath $TemplatePath)){
    $lblStatusText.Text = "Template not found: $TemplatePath"
    Set-StatusPill -text "Error" -state "error"
    Append-Log "Template not found: $TemplatePath"
    return
  }
  if(-not $chkSaveDocx.Checked -and -not $chkSavePdf.Checked){
    $lblStatusText.Text = "Select at least one output: DOCX or PDF."
    Set-StatusPill -text "Error" -state "error"
    Append-Log "Validation failed: Select at least one output: DOCX or PDF."
    return
  }

  $btnStart.Enabled = $false
  $btnStop.Enabled = $true
  $btnOpen.Enabled = $false

  $grid.Rows.Clear()
  $script:RowMap.Clear()
  $logBox.Clear()

  $script:SyncHash.Cancel = $false
  $script:SyncHash.Error = $null
  $script:SyncHash.Result = $null
  $script:SyncHash.Running = $true
  $script:SyncHash.Status = "Starting"
  $script:RunStarted = Get-Date
  $lblStatusText.Text = "Starting..."
  Set-StatusPill -text "Running" -state "running"

  Write-Log "User clicked Start."

  # STA runspace for Office COM
  $rs = [RunspaceFactory]::CreateRunspace()
  $rs.ApartmentState = "STA"
  $rs.ThreadOptions = "ReuseThread"
  $rs.Open()

  $script:PSInstance = [PowerShell]::Create()
  $script:PSInstance.Runspace = $rs
  $script:PSInstance.AddScript($script:WorkerLogic) | Out-Null
  $script:PSInstance.AddArgument($script:SyncHash) | Out-Null
  $script:PSInstance.AddArgument(@{
    ExcelPath = $ExcelPath
    TemplatePath = $TemplatePath
    OutDir = $OutDir
    LogPath = $LogPath
    PreferredSheet = $PreferredSheet
    ShowWord = $chkShowWord.Checked
    ExportPdf = $chkSavePdf.Checked
    ExportDocx = $chkSaveDocx.Checked
    FastMode = $script:UseFastMode
    TurboMode = $script:UseTurboMode
  }) | Out-Null

  $script:PSAsync = $script:PSInstance.BeginInvoke()

  $uiTimer.Start()
  $progressTimer.Start()
})

$btnStop.Add_Click({
  $script:SyncHash.Cancel = $true
  $script:SyncHash.Status = "Stopping"
  $lblStatusText.Text = "Stopping..."
  Set-StatusPill -text "Stopping" -state "running"
  $btnStop.Enabled = $false
  if($script:PSInstance){
    try { $script:PSInstance.Stop() | Out-Null } catch {}
  }
  Write-Log "User clicked Stop."
})

$btnCloseAll.Add_Click({
  try {
    $script:SyncHash.Cancel = $true
    if($script:PSInstance){
      try { $script:PSInstance.Stop() | Out-Null } catch {}
    }
  } catch {}

  try {
    if (Get-Command -Name Close-SchumanAllResources -ErrorAction SilentlyContinue) {
      Close-SchumanAllResources -Mode 'All' | Out-Null
    }
  } catch {}

  try { $form.Close() } catch {}
})

# Start timers idle-safe
$progressTimer.Start()
$progressTimer.Stop()
Apply-Theme
[void](Load-ExcelPreview -Silent)

# Start app
[Windows.Forms.Application]::EnableVisualStyles()
[Windows.Forms.Application]::Run($form)
