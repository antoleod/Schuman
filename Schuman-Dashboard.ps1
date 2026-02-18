#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$script:EntryScript = Join-Path $PSScriptRoot 'Invoke-Schuman.ps1'
if (-not (Test-Path -LiteralPath $script:EntryScript)) {
  throw "Invoke-Schuman.ps1 not found: $script:EntryScript"
}

$script:State = [pscustomobject]@{
  IsBusy = $false
  Job = $null
  ExportReady = $false
}

function Invoke-Ui {
  param(
    [Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Control,
    [Parameter(Mandatory = $true)][scriptblock]$Action
  )

  try {
    if ($Control.IsDisposed) { return }
    if ($Control.InvokeRequired) { [void]$Control.BeginInvoke($Action) }
    else { & $Action }
  } catch {}
}

function Write-UiLog {
  param(
    [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$Target,
    [Parameter(Mandatory = $true)][ValidateSet('INFO','WARN','ERROR')][string]$Level,
    [Parameter(Mandatory = $true)][string]$Message
  )

  $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
  Invoke-Ui -Control $Target -Action { $Target.AppendText($line + [Environment]::NewLine) }
}

function Test-NoiseLogLine {
  param([string]$Line)
  if ([string]::IsNullOrWhiteSpace($Line)) { return $true }

  $trim = $Line.Trim()
  if ($trim -match '(?i)^Windows PowerShell$') { return $true }
  if ($trim -match '(?i)^Copyright \(C\) Microsoft Corporation') { return $true }
  if ($trim -match '(?i)Install the latest PowerShell for new features and improvements') { return $true }
  if ($trim -match '(?i)https://aka\.ms/PSWindows') { return $true }

  return $false
}

function Test-ExportAlreadyAvailable {
  $runsDir = Join-Path $PSScriptRoot 'system\runs'
  if (-not (Test-Path -LiteralPath $runsDir)) { return $false }

  try {
    $latest = Get-ChildItem -LiteralPath $runsDir -Directory -Filter 'export_*' -ErrorAction Stop |
      Sort-Object LastWriteTime -Descending |
      Select-Object -First 1
    if (-not $latest) { return $false }

    $json = Join-Path $latest.FullName 'tickets_export.json'
    return (Test-Path -LiteralPath $json)
  }
  catch {
    return $false
  }
}

function Set-MainBusy {
  param([bool]$Busy, [string]$Caption = 'Idle')

  $script:State.IsBusy = $Busy
  $btnExport.Enabled = -not $Busy
  $btnDashboard.Enabled = -not $Busy
  $btnDocs.Enabled = -not $Busy
  $btnBrowse.Enabled = -not $Busy
  $txtExcel.Enabled = -not $Busy
  $txtSheet.Enabled = -not $Busy
  $cmbScope.Enabled = -not $Busy
  $numMax.Enabled = -not $Busy

  if ($Busy) {
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $statusLabel.Text = $Caption
  }
  else {
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
    $progress.Value = 0
    $statusLabel.Text = 'Idle'
  }
}

function Get-CommonParamsOrNull {
  $excel = ("" + $txtExcel.Text).Trim()
  $sheet = ("" + $txtSheet.Text).Trim()

  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    [System.Windows.Forms.MessageBox]::Show('Excel file not found. Select a valid .xlsx file.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    return $null
  }
  if ([string]::IsNullOrWhiteSpace($sheet)) {
    [System.Windows.Forms.MessageBox]::Show('Sheet name is required.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    return $null
  }

  return @{
    ExcelPath = $excel
    SheetName = $sheet
    NoPopups = $true
  }
}

function Start-AsyncInvokeSchuman {
  param(
    [Parameter(Mandatory = $true)][string]$ActionName,
    [Parameter(Mandatory = $true)][hashtable]$ParamMap,
    [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$LogTarget,
    [scriptblock]$OnDone
  )

  if ($script:State.IsBusy) {
    Write-UiLog -Target $LogTarget -Level WARN -Message 'Another operation is already running.'
    return
  }

  Set-MainBusy -Busy $true -Caption ("Running: " + $ActionName)
  Write-UiLog -Target $LogTarget -Level INFO -Message ("Starting " + $ActionName)

  $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
  $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
  $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
  $runspace.Open()

  $ps = [PowerShell]::Create()
  $ps.Runspace = $runspace
  $scriptBlock = {
    param($Entry, $Params)
    & $Entry @Params 2>&1 | ForEach-Object { $_ }
  }

  [void]$ps.AddScript($scriptBlock)
  [void]$ps.AddArgument($script:EntryScript)
  [void]$ps.AddArgument($ParamMap)

  $input = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
  $async = $ps.BeginInvoke($input)

  $script:State.Job = [pscustomobject]@{
    Name = $ActionName
    PS = $ps
    Runspace = $runspace
    Async = $async
    LogTarget = $LogTarget
    OnDone = $OnDone
  }

  if (-not $pollTimer.Enabled) { $pollTimer.Start() }
}

function Complete-CurrentJob {
  param([bool]$Success)

  $job = $script:State.Job
  if (-not $job) { return }

  $callback = $job.OnDone
  $name = $job.Name

  try { if ($job.PS) { $job.PS.Dispose() } } catch {}
  try { if ($job.Runspace) { $job.Runspace.Close() } } catch {}
  try { if ($job.Runspace) { $job.Runspace.Dispose() } } catch {}

  $script:State.Job = $null
  Set-MainBusy -Busy $false

  if ($Success) {
    Write-UiLog -Target $job.LogTarget -Level INFO -Message ("Completed " + $name)
  }
  else {
    Write-UiLog -Target $job.LogTarget -Level ERROR -Message ("Failed " + $name)
  }

  if ($callback) {
    try { & $callback $Success } catch {}
  }
}

function Ensure-ExportPrerequisite {
  param(
    [Parameter(Mandatory = $true)][string]$FeatureName,
    [Parameter(Mandatory = $true)][System.Windows.Forms.TextBox]$LogTarget,
    [Parameter(Mandatory = $true)][scriptblock]$OnContinue
  )

  if ($script:State.ExportReady) {
    & $OnContinue
    return
  }

  $msg = "$FeatureName works best after 'Export Tickets'. Do you want to run Export Tickets first?"
  $choice = [System.Windows.Forms.MessageBox]::Show($msg,'Export Recommended',[System.Windows.Forms.MessageBoxButtons]::YesNoCancel,[System.Windows.Forms.MessageBoxIcon]::Question)
  if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

  if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
    $common = Get-CommonParamsOrNull
    if (-not $common) { return }

    $params = @{}
    foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
    $params.Operation = 'Export'
    $params.ProcessingScope = $cmbScope.SelectedItem.ToString()
    $params.MaxTickets = [int]$numMax.Value

    Start-AsyncInvokeSchuman -ActionName 'Export Tickets (Prerequisite)' -ParamMap $params -LogTarget $LogTarget -OnDone {
      param($ok)
      if ($ok) {
        $script:State.ExportReady = $true
        & $OnContinue
      }
    }
    return
  }

  & $OnContinue
}

function Show-DashboardWindow {
  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Dashboard'
  $dlg.StartPosition = 'CenterParent'
  $dlg.Size = New-Object System.Drawing.Size(980, 560)
  $dlg.MinimumSize = New-Object System.Drawing.Size(900, 520)
  $dlg.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
  $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 9)

  $lblTop = New-Object System.Windows.Forms.Label
  $lblTop.Text = 'Search and perform Check-In / Check-Out operations'
  $lblTop.Location = New-Object System.Drawing.Point(16, 14)
  $lblTop.AutoSize = $true
  $dlg.Controls.Add($lblTop)

  $lblSearch = New-Object System.Windows.Forms.Label
  $lblSearch.Text = 'Search'
  $lblSearch.Location = New-Object System.Drawing.Point(16, 46)
  $lblSearch.AutoSize = $true
  $dlg.Controls.Add($lblSearch)

  $txtSearchLocal = New-Object System.Windows.Forms.TextBox
  $txtSearchLocal.Location = New-Object System.Drawing.Point(72, 42)
  $txtSearchLocal.Size = New-Object System.Drawing.Size(360, 24)
  $dlg.Controls.Add($txtSearchLocal)

  $btnSearchLocal = New-Object System.Windows.Forms.Button
  $btnSearchLocal.Text = 'Search'
  $btnSearchLocal.Location = New-Object System.Drawing.Point(442, 40)
  $btnSearchLocal.Size = New-Object System.Drawing.Size(100, 28)
  $dlg.Controls.Add($btnSearchLocal)

  $lblRow = New-Object System.Windows.Forms.Label
  $lblRow.Text = 'Row'
  $lblRow.Location = New-Object System.Drawing.Point(560, 46)
  $lblRow.AutoSize = $true
  $dlg.Controls.Add($lblRow)

  $numRowLocal = New-Object System.Windows.Forms.NumericUpDown
  $numRowLocal.Location = New-Object System.Drawing.Point(596, 42)
  $numRowLocal.Size = New-Object System.Drawing.Size(90, 24)
  $numRowLocal.Minimum = 0
  $numRowLocal.Maximum = 100000
  $dlg.Controls.Add($numRowLocal)

  $btnCheckInLocal = New-Object System.Windows.Forms.Button
  $btnCheckInLocal.Text = 'Check-In'
  $btnCheckInLocal.Location = New-Object System.Drawing.Point(704, 40)
  $btnCheckInLocal.Size = New-Object System.Drawing.Size(110, 28)
  $dlg.Controls.Add($btnCheckInLocal)

  $btnCheckOutLocal = New-Object System.Windows.Forms.Button
  $btnCheckOutLocal.Text = 'Check-Out'
  $btnCheckOutLocal.Location = New-Object System.Drawing.Point(824, 40)
  $btnCheckOutLocal.Size = New-Object System.Drawing.Size(120, 28)
  $dlg.Controls.Add($btnCheckOutLocal)

  $txtLogLocal = New-Object System.Windows.Forms.TextBox
  $txtLogLocal.Location = New-Object System.Drawing.Point(16, 80)
  $txtLogLocal.Size = New-Object System.Drawing.Size(928, 420)
  $txtLogLocal.Multiline = $true
  $txtLogLocal.ScrollBars = 'Vertical'
  $txtLogLocal.ReadOnly = $true
  $txtLogLocal.BackColor = [System.Drawing.Color]::FromArgb(14, 20, 28)
  $txtLogLocal.ForeColor = [System.Drawing.Color]::FromArgb(224, 233, 244)
  $txtLogLocal.Font = New-Object System.Drawing.Font('Consolas', 9)
  $dlg.Controls.Add($txtLogLocal)

  $btnSearchLocal.Add_Click({
    $common = Get-CommonParamsOrNull
    if (-not $common) { return }

    $params = @{}
    foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
    $params.Operation = 'DashboardSearch'
    $params.SearchText = $txtSearchLocal.Text

    Start-AsyncInvokeSchuman -ActionName 'Dashboard Search' -ParamMap $params -LogTarget $txtLogLocal
  })

  $btnCheckInLocal.Add_Click({
    if ($numRowLocal.Value -le 0) {
      [System.Windows.Forms.MessageBox]::Show('Provide Row > 0.','Validation') | Out-Null
      return
    }

    $common = Get-CommonParamsOrNull
    if (-not $common) { return }

    $params = @{}
    foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
    $params.Operation = 'DashboardCheckIn'
    $params.Row = [int]$numRowLocal.Value

    Start-AsyncInvokeSchuman -ActionName 'Dashboard Check-In' -ParamMap $params -LogTarget $txtLogLocal
  })

  $btnCheckOutLocal.Add_Click({
    if ($numRowLocal.Value -le 0) {
      [System.Windows.Forms.MessageBox]::Show('Provide Row > 0.','Validation') | Out-Null
      return
    }

    $common = Get-CommonParamsOrNull
    if (-not $common) { return }

    $params = @{}
    foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
    $params.Operation = 'DashboardCheckOut'
    $params.Row = [int]$numRowLocal.Value

    Start-AsyncInvokeSchuman -ActionName 'Dashboard Check-Out' -ParamMap $params -LogTarget $txtLogLocal
  })

  Write-UiLog -Target $txtLogLocal -Level INFO -Message 'Dashboard ready.'
  [void]$dlg.ShowDialog($form)
}

# Main form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Schuman Automation Hub'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(980, 620)
$form.MinimumSize = New-Object System.Drawing.Size(920, 560)
$form.BackColor = [System.Drawing.Color]::FromArgb(241, 244, 249)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Schuman Automation Hub'
$title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
$title.Location = New-Object System.Drawing.Point(18, 14)
$title.AutoSize = $true
$form.Controls.Add($title)

$sub = New-Object System.Windows.Forms.Label
$sub.Text = 'Corporate-safe launcher (PowerShell 5.1, no installs, no admin).'
$sub.Location = New-Object System.Drawing.Point(20, 46)
$sub.AutoSize = $true
$sub.ForeColor = [System.Drawing.Color]::FromArgb(92, 106, 128)
$form.Controls.Add($sub)

$lblExcel = New-Object System.Windows.Forms.Label
$lblExcel.Text = 'Excel'
$lblExcel.Location = New-Object System.Drawing.Point(20, 82)
$lblExcel.AutoSize = $true
$form.Controls.Add($lblExcel)

$txtExcel = New-Object System.Windows.Forms.TextBox
$txtExcel.Location = New-Object System.Drawing.Point(66, 78)
$txtExcel.Size = New-Object System.Drawing.Size(760, 24)
$txtExcel.Text = $ExcelPath
$form.Controls.Add($txtExcel)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse'
$btnBrowse.Location = New-Object System.Drawing.Point(836, 76)
$btnBrowse.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnBrowse)

$lblSheet = New-Object System.Windows.Forms.Label
$lblSheet.Text = 'Sheet'
$lblSheet.Location = New-Object System.Drawing.Point(20, 118)
$lblSheet.AutoSize = $true
$form.Controls.Add($lblSheet)

$txtSheet = New-Object System.Windows.Forms.TextBox
$txtSheet.Location = New-Object System.Drawing.Point(66, 114)
$txtSheet.Size = New-Object System.Drawing.Size(96, 24)
$txtSheet.Text = $SheetName
$form.Controls.Add($txtSheet)

$lblScope = New-Object System.Windows.Forms.Label
$lblScope.Text = 'Scope'
$lblScope.Location = New-Object System.Drawing.Point(176, 118)
$lblScope.AutoSize = $true
$form.Controls.Add($lblScope)

$cmbScope = New-Object System.Windows.Forms.ComboBox
$cmbScope.Location = New-Object System.Drawing.Point(222, 114)
$cmbScope.Size = New-Object System.Drawing.Size(150, 24)
$cmbScope.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbScope.Items.AddRange(@('Auto','RitmOnly','IncAndRitm','All'))
$cmbScope.SelectedItem = 'RitmOnly'
$form.Controls.Add($cmbScope)

$lblMax = New-Object System.Windows.Forms.Label
$lblMax.Text = 'Max Tickets'
$lblMax.Location = New-Object System.Drawing.Point(390, 118)
$lblMax.AutoSize = $true
$form.Controls.Add($lblMax)

$numMax = New-Object System.Windows.Forms.NumericUpDown
$numMax.Location = New-Object System.Drawing.Point(462, 114)
$numMax.Size = New-Object System.Drawing.Size(90, 24)
$numMax.Minimum = 0
$numMax.Maximum = 10000
$numMax.Value = 0
$form.Controls.Add($numMax)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export Tickets'
$btnExport.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$btnExport.Location = New-Object System.Drawing.Point(20, 162)
$btnExport.Size = New-Object System.Drawing.Size(300, 44)
$form.Controls.Add($btnExport)

$btnDashboard = New-Object System.Windows.Forms.Button
$btnDashboard.Text = 'Dashboard'
$btnDashboard.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$btnDashboard.Location = New-Object System.Drawing.Point(334, 162)
$btnDashboard.Size = New-Object System.Drawing.Size(300, 44)
$form.Controls.Add($btnDashboard)

$btnDocs = New-Object System.Windows.Forms.Button
$btnDocs.Text = 'Generate Docs'
$btnDocs.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$btnDocs.Location = New-Object System.Drawing.Point(648, 162)
$btnDocs.Size = New-Object System.Drawing.Size(298, 44)
$form.Controls.Add($btnDocs)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Idle'
$statusLabel.Location = New-Object System.Drawing.Point(20, 216)
$statusLabel.AutoSize = $true
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(18, 92, 203)
$form.Controls.Add($statusLabel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(66, 214)
$progress.Size = New-Object System.Drawing.Size(880, 16)
$progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$form.Controls.Add($progress)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 244)
$txtLog.Size = New-Object System.Drawing.Size(926, 328)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(14, 20, 28)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(224, 233, 244)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

$pollTimer = New-Object System.Windows.Forms.Timer
$pollTimer.Interval = 120
$pollTimer.Add_Tick({
  $job = $script:State.Job
  if (-not $job) {
    $pollTimer.Stop()
    return
  }

  if ($job.Async.IsCompleted) {
    $ok = $true
    $results = @()
    try { $results = @($job.PS.EndInvoke($job.Async)) }
    catch {
      $ok = $false
      Write-UiLog -Target $job.LogTarget -Level ERROR -Message $_.Exception.Message
      if ($_.Exception.InnerException) {
        Write-UiLog -Target $job.LogTarget -Level ERROR -Message ("Inner: " + $_.Exception.InnerException.Message)
      }
    }
    if ($ok) {
      foreach ($r in $results) {
        $line = ("" + $r).TrimEnd()
        if (-not (Test-NoiseLogLine -Line $line)) {
          Write-UiLog -Target $job.LogTarget -Level INFO -Message $line
        }
      }
    }
    Complete-CurrentJob -Success $ok
  }
})

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx'
  $dlg.Title = 'Select Excel file'
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $txtExcel.Text = $dlg.FileName
  }
})

$btnExport.Add_Click({
  $common = Get-CommonParamsOrNull
  if (-not $common) { return }

  $params = @{}
  foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
  $params.Operation = 'Export'
  $params.ProcessingScope = $cmbScope.SelectedItem.ToString()
  $params.MaxTickets = [int]$numMax.Value

  Start-AsyncInvokeSchuman -ActionName 'Export Tickets' -ParamMap $params -LogTarget $txtLog -OnDone {
    param($ok)
    if ($ok) { $script:State.ExportReady = $true }
  }
})

$btnDashboard.Add_Click({
  Ensure-ExportPrerequisite -FeatureName 'Dashboard' -LogTarget $txtLog -OnContinue {
    Show-DashboardWindow
  }
})

$btnDocs.Add_Click({
  Ensure-ExportPrerequisite -FeatureName 'Generate Docs' -LogTarget $txtLog -OnContinue {
    $common = Get-CommonParamsOrNull
    if (-not $common) { return }

    $params = @{}
    foreach ($k in $common.Keys) { $params[$k] = $common[$k] }
    $params.Operation = 'DocsGenerate'

    Start-AsyncInvokeSchuman -ActionName 'Generate Docs' -ParamMap $params -LogTarget $txtLog
  }
})

$script:State.ExportReady = Test-ExportAlreadyAvailable
if ($script:State.ExportReady) {
  Write-UiLog -Target $txtLog -Level INFO -Message 'Previous export data detected. Dashboard/Docs can run directly.'
}
else {
  Write-UiLog -Target $txtLog -Level INFO -Message 'Run Export Tickets first for best results.'
}

Write-UiLog -Target $txtLog -Level INFO -Message 'Launcher ready.'

[void]$form.ShowDialog()
