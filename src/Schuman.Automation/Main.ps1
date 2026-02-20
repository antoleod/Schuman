#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot '..\..\Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$invokeScript = Join-Path $projectRoot 'Invoke-Schuman.ps1'
$importModulesPath = Join-Path $PSScriptRoot 'Import-SchumanModules.ps1'
$themePath = Join-Path $PSScriptRoot 'UI\Theme.ps1'
$dashboardUiPath = Join-Path $PSScriptRoot 'UI\DashboardUI.ps1'
$generateUiPath = Join-Path $PSScriptRoot 'UI\GenerateUI.ps1'

foreach ($p in @($invokeScript, $importModulesPath, $themePath, $dashboardUiPath, $generateUiPath)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required file not found: $p" }
}

. $importModulesPath
. $themePath
. $dashboardUiPath
. $generateUiPath

$globalConfig = Initialize-SchumanEnvironment -ProjectRoot $projectRoot
$uiRunContext = New-RunContext -Config $globalConfig -RunName 'mainui'

function Convert-ToArgumentString {
  param([string[]]$Tokens)
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($token in $Tokens) {
    $text = if ($null -eq $token) { '' } else { [string]$token }
    if ($text -match '^[A-Za-z0-9_:\\\.\-]+$') {
      [void]$parts.Add($text)
    } else {
      [void]$parts.Add('"' + ($text -replace '"','\"') + '"')
    }
  }
  return ($parts -join ' ')
}

function Get-RunningSchumanOperationProcesses {
  param(
    [string[]]$Operations = @('Export','DocsGenerate')
  )

  $ops = @($Operations | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToLowerInvariant() })
  if ($ops.Count -eq 0) { return @() }

  $matches = New-Object System.Collections.Generic.List[object]
  try {
    $procs = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'")
    foreach ($p in $procs) {
      $cmd = ("" + $p.CommandLine)
      if (-not $cmd) { continue }
      if ($p.ProcessId -eq $PID) { continue }
      if ($cmd -notmatch '(?i)Invoke-Schuman\.ps1') { continue }
      foreach ($op in $ops) {
        if ($cmd -match ("(?i)-Operation\s+{0}(\s|$)" -f [Regex]::Escape($op))) {
          $matches.Add([pscustomobject]@{
              ProcessId = [int]$p.ProcessId
              Operation = $op
              CommandLine = $cmd
            }) | Out-Null
          break
        }
      }
    }
  }
  catch {
    return @()
  }

  return @($matches.ToArray())
}

function Get-UiMetricsPath {
  param([hashtable]$Config)
  return (Join-Path (Join-Path $Config.Output.SystemRoot $Config.Output.DbSubdir) 'ui-metrics.json')
}

function Get-UiMetrics {
  param([hashtable]$Config)
  $path = Get-UiMetricsPath -Config $Config
  if (-not (Test-Path -LiteralPath $path)) { return @{} }
  try {
    $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($json -is [hashtable]) { return $json }
    if ($json -is [pscustomobject]) { return (ConvertTo-Hashtable $json) }
  }
  catch {}
  return @{}
}

function Save-UiMetrics {
  param(
    [hashtable]$Config,
    [hashtable]$Metrics
  )
  $path = Get-UiMetricsPath -Config $Config
  Ensure-Directory -Path (Split-Path -Parent $path) | Out-Null
  ($Metrics | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $path -Encoding UTF8
}

function Get-UiEstimatedSeconds {
  param(
    [hashtable]$Config,
    [string]$OperationKey,
    [int]$DefaultSeconds = 180
  )
  $m = Get-UiMetrics -Config $Config
  if ($m.ContainsKey($OperationKey)) {
    try {
      $item = $m[$OperationKey]
      $avg = [double]$item.avg_seconds
      if ($avg -gt 1) { return [int][Math]::Round($avg) }
    }
    catch {}
  }
  return $DefaultSeconds
}

function Update-UiMetricDuration {
  param(
    [hashtable]$Config,
    [string]$OperationKey,
    [int]$DurationSeconds
  )
  if ($DurationSeconds -le 0) { return }

  $m = Get-UiMetrics -Config $Config
  if (-not $m.ContainsKey($OperationKey)) {
    $m[$OperationKey] = @{
      avg_seconds = [double]$DurationSeconds
      samples = 1
      last_seconds = [int]$DurationSeconds
    }
  }
  else {
    $item = $m[$OperationKey]
    $samples = 0
    $avg = [double]$DurationSeconds
    try { $samples = [int]$item.samples } catch { $samples = 0 }
    try { $avg = [double]$item.avg_seconds } catch { $avg = [double]$DurationSeconds }

    $newSamples = [Math]::Min($samples + 1, 50)
    $newAvg = (($avg * [Math]::Max($samples, 1)) + [double]$DurationSeconds) / [double]([Math]::Max($samples, 1) + 1)
    $m[$OperationKey] = @{
      avg_seconds = [double]$newAvg
      samples = [int]$newSamples
      last_seconds = [int]$DurationSeconds
    }
  }

  Save-UiMetrics -Config $Config -Metrics $m
}

function Show-ForceUpdateOptionsDialog {
  param(
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Force Update Options'
  $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.Size = New-Object System.Drawing.Size(640, 420)
  $dlg.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 10)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 5
  $layout.Padding = New-Object System.Windows.Forms.Padding(16)
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$dlg.Controls.Add($layout)

  $lblTitle = New-Object System.Windows.Forms.Label
  $lblTitle.Text = 'Force Update Excel'
  $lblTitle.AutoSize = $true
  $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
  [void]$layout.Controls.Add($lblTitle, 0, 0)

  $body = New-Object System.Windows.Forms.TableLayoutPanel
  $body.Dock = [System.Windows.Forms.DockStyle]::Fill
  $body.ColumnCount = 2
  $body.RowCount = 1
  [void]$body.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
  [void]$body.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
  $body.Margin = New-Object System.Windows.Forms.Padding(0, 12, 0, 8)
  [void]$layout.Controls.Add($body, 0, 1)

  $newGroupPanel = {
    param([string]$Title)
    $p = New-Object System.Windows.Forms.Panel
    $p.Dock = [System.Windows.Forms.DockStyle]::Fill
    $p.Padding = New-Object System.Windows.Forms.Padding(12)
    $p.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $p.BackColor = [System.Drawing.Color]::FromArgb(32,32,34)

    $t = New-Object System.Windows.Forms.Label
    $t.Text = $Title
    $t.AutoSize = $true
    $t.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $t.ForeColor = [System.Drawing.Color]::FromArgb(170,170,175)
    $p.Controls.Add($t)

    $f = New-Object System.Windows.Forms.FlowLayoutPanel
    $f.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $f.AutoSize = $true
    $f.WrapContents = $false
    $f.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $f.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 0)
    $p.Controls.Add($f)

    return [pscustomobject]@{ Panel = $p; Flow = $f }
  }

  $scopeGroup = & $newGroupPanel 'Ticket Scope'
  $scopeGroup.Panel.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
  [void]$body.Controls.Add($scopeGroup.Panel, 0, 0)

  $piGroup = & $newGroupPanel 'PI Search Mode'
  $piGroup.Panel.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
  [void]$body.Controls.Add($piGroup.Panel, 1, 0)

  $newChoice = {
    param(
      [string]$Text,
      [string]$TagValue,
      [bool]$Checked = $false
    )
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = $Text
    $chk.Tag = $TagValue
    $chk.AutoSize = $true
    $chk.Checked = $Checked
    $chk.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $chk.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    return $chk
  }

  $scopeChoices = @(
    (& $newChoice 'RITM only (fastest)' 'RitmOnly' $true),
    (& $newChoice 'INC + RITM' 'IncAndRitm' $false),
    (& $newChoice 'All tickets' 'All' $false)
  )
  foreach ($c in $scopeChoices) { [void]$scopeGroup.Flow.Controls.Add($c) }

  $piChoices = @(
    (& $newChoice 'Configuration Item only (fastest)' 'ConfigurationItemOnly' $true),
    (& $newChoice 'Comments only' 'CommentsOnly' $false),
    (& $newChoice 'Comments + Configuration Item' 'CommentsAndCI' $false),
    (& $newChoice 'Auto' 'Auto' $false)
  )
  foreach ($c in $piChoices) { [void]$piGroup.Flow.Controls.Add($c) }

  $bindSingleSelect = {
    param([System.Windows.Forms.CheckBox[]]$Boxes)
    foreach ($box in $Boxes) {
      $box.Add_CheckedChanged(({
        param($sender, $eventArgs)
        if (-not $sender.Checked) { return }
        foreach ($other in $Boxes) {
          if ($other -ne $sender) { $other.Checked = $false }
        }
      }).GetNewClosure())
    }
  }
  & $bindSingleSelect $scopeChoices
  & $bindSingleSelect $piChoices

  $hint = New-Object System.Windows.Forms.Label
  $hint.Text = 'Tip: RITM only + Configuration Item only is usually the fastest combination.'
  $hint.AutoSize = $true
  $hint.ForeColor = [System.Drawing.Color]::FromArgb(120,120,126)
  $hint.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
  [void]$layout.Controls.Add($hint, 0, 2)

  $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttons.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $buttons.AutoSize = $true
  [void]$layout.Controls.Add($buttons, 0, 4)

  $btnOk = New-Object System.Windows.Forms.Button
  $btnOk.Text = 'Start'
  $btnOk.Width = 100
  $btnOk.Height = 32
  $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
  $btnOk.ForeColor = [System.Drawing.Color]::White
  $btnOk.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0,122,255)
  $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
  [void]$buttons.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = 'Cancel'
  $btnCancel.Width = 100
  $btnCancel.Height = 32
  $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(36,36,38)
  $btnCancel.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(58,58,62)
  $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  [void]$buttons.Controls.Add($btnCancel)

  $dlg.AcceptButton = $btnOk
  $dlg.CancelButton = $btnCancel

  $res = if ($Owner) { $dlg.ShowDialog($Owner) } else { $dlg.ShowDialog() }
  if ($res -ne [System.Windows.Forms.DialogResult]::OK) {
    return [pscustomobject]@{ ok = $false }
  }

  $scope = @($scopeChoices | Where-Object { $_.Checked } | Select-Object -First 1)
  $scopeValue = if ($scope.Count -gt 0) { "" + $scope[0].Tag } else { 'RitmOnly' }

  $pi = @($piChoices | Where-Object { $_.Checked } | Select-Object -First 1)
  $piValue = if ($pi.Count -gt 0) { "" + $pi[0].Tag } else { 'ConfigurationItemOnly' }

  return [pscustomobject]@{
    ok = $true
    processingScope = $scopeValue
    piSearchMode = $piValue
  }
}

function Invoke-PreloadExport {
  param(
    [string]$Excel,
    [string]$Sheet
  )

  $active = @(Get-RunningSchumanOperationProcesses -Operations @('Export'))
  if ($active.Count -gt 0) {
    $ids = (($active | Select-Object -ExpandProperty ProcessId) -join ', ')
    [System.Windows.Forms.MessageBox]::Show(
      "There is already an Export process running (PID: $ids). Wait for it to finish before starting another one.",
      'Schuman'
    ) | Out-Null
    return $false
  }

  $loading = New-Object System.Windows.Forms.Form
  $loading.Text = 'Schuman — Loading'
  $loading.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $loading.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $loading.ControlBox = $false
  $loading.Size = New-Object System.Drawing.Size(520, 180)
  $loading.TopMost = $true
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 3
  $layout.Padding = New-Object System.Windows.Forms.Padding(20)
  $loading.Controls.Add($layout)

  $title = New-Object System.Windows.Forms.Label
  $title.Text = 'Preparing data before opening module'
  $title.AutoSize = $true
  $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
  $layout.Controls.Add($title, 0, 0)

  $status = New-Object System.Windows.Forms.Label
  $status.Text = 'Running export...'
  $status.AutoSize = $true
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 12)
  $layout.Controls.Add($status, 0, 1)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220,220,225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 2)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(90, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
  $barTrack.Controls.Add($barFill)

  $animTimer = New-Object System.Windows.Forms.Timer
  $animTimer.Interval = 180

  $args = @(
    '-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
    '-File',$invokeScript,
    '-Operation','Export',
    '-ExcelPath',$Excel,
    '-SheetName',$Sheet,
    '-ProcessingScope','RitmOnly',
    '-MaxTickets','40',
    '-NoWriteBack',
    '-NoPopups'
  )

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = (Join-Path $PSHOME 'powershell.exe')
  $psi.Arguments = Convert-ToArgumentString -Tokens $args
  $psi.WorkingDirectory = $projectRoot
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi

  $state = @{
    Tick = 0
    Ok = $false
    Proc = $proc
    StatusLabel = $status
  }
  $loading.Tag = $state

  $animTimer.Add_Tick({
    param($sender, $eventArgs)
    $st = $loading.Tag
    if (-not $st) { return }
    $st.Tick = [int]$st.Tick + 1
    $dots = '.' * (([int]$st.Tick % 3) + 1)
    $st.StatusLabel.Text = 'Running export' + $dots
    $maxX = [Math]::Max(1, $barTrack.ClientSize.Width - $barFill.Width)
    $barFill.Left = ([int]$st.Tick * 14) % $maxX
  })

  $watchTimer = New-Object System.Windows.Forms.Timer
  $watchTimer.Interval = 250
  $watchTimer.Add_Tick({
    param($sender, $eventArgs)
    $st = $loading.Tag
    if (-not $st -or -not $st.Proc) { return }
    if ($st.Proc.HasExited) {
      $st.Ok = ($st.Proc.ExitCode -eq 0)
      $loading.Close()
    }
  })

  $loading.add_Shown({
    try {
      $st = $loading.Tag
      [void]$st.Proc.Start()
      $animTimer.Start()
      $watchTimer.Start()
    }
    catch {
      $loading.Tag.Ok = $false
      $loading.Close()
    }
  })

  $ok = $false
  try {
    [void]$loading.ShowDialog()
    $st = $loading.Tag
    if ($st) { $ok = [bool]$st.Ok }
  }
  finally {
    try { $animTimer.Stop() } catch {}
    try { $watchTimer.Stop() } catch {}
    try { $animTimer.Dispose() } catch {}
    try { $watchTimer.Dispose() } catch {}
    try { $loading.Close() } catch {}
    try { $loading.Dispose() } catch {}
    try { $proc.Dispose() } catch {}
  }

  return $ok
}

function Invoke-DocsGenerate {
  param(
    [string]$Excel,
    [string]$Sheet,
    [string]$Template,
    [string]$Output,
    [bool]$ExportPdf = $true
  )

  $active = @(Get-RunningSchumanOperationProcesses -Operations @('DocsGenerate'))
  if ($active.Count -gt 0) {
    $ids = (($active | Select-Object -ExpandProperty ProcessId) -join ', ')
    [System.Windows.Forms.MessageBox]::Show(
      "There is already a DocsGenerate process running (PID: $ids). Wait for it to finish before starting another one.",
      'Schuman'
    ) | Out-Null
    return $false
  }

  $loading = New-Object System.Windows.Forms.Form
  $loading.Text = 'Schuman — Generating'
  $loading.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $loading.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $loading.ControlBox = $false
  $loading.Size = New-Object System.Drawing.Size(560, 190)
  $loading.TopMost = $true
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 3
  $layout.Padding = New-Object System.Windows.Forms.Padding(20)
  $loading.Controls.Add($layout)

  $title = New-Object System.Windows.Forms.Label
  $title.Text = 'Generating Word/PDF documents'
  $title.AutoSize = $true
  $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
  $layout.Controls.Add($title, 0, 0)

  $status = New-Object System.Windows.Forms.Label
  $status.Text = 'Working...'
  $status.AutoSize = $true
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 12)
  $layout.Controls.Add($status, 0, 1)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220,220,225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 2)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(90, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
  $barTrack.Controls.Add($barFill)

  $args = @(
    '-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
    '-File',$invokeScript,
    '-Operation','DocsGenerate',
    '-ExcelPath',$Excel,
    '-SheetName',$Sheet,
    '-TemplatePath',$Template,
    '-OutputDirectory',$Output,
    '-NoPopups'
  )
  if ($ExportPdf) { $args += '-ExportPdf' }

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = (Join-Path $PSHOME 'powershell.exe')
  $psi.Arguments = Convert-ToArgumentString -Tokens $args
  $psi.WorkingDirectory = $projectRoot
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi

  $animTimer = New-Object System.Windows.Forms.Timer
  $animTimer.Interval = 180
  $watchTimer = New-Object System.Windows.Forms.Timer
  $watchTimer.Interval = 250

  $state = @{
    Tick = 0
    Ok = $false
    Proc = $proc
    StatusLabel = $status
  }
  $loading.Tag = $state

  $animTimer.Add_Tick({
    $st = $loading.Tag
    if (-not $st) { return }
    $st.Tick = [int]$st.Tick + 1
    $st.StatusLabel.Text = 'Working' + ('.' * (([int]$st.Tick % 3) + 1))
    $maxX = [Math]::Max(1, $barTrack.ClientSize.Width - $barFill.Width)
    $barFill.Left = ([int]$st.Tick * 14) % $maxX
  })
  $watchTimer.Add_Tick({
    $st = $loading.Tag
    if (-not $st -or -not $st.Proc) { return }
    if ($st.Proc.HasExited) {
      $st.Ok = ($st.Proc.ExitCode -eq 0)
      $loading.Close()
    }
  })

  $loading.add_Shown({
    try {
      [void]$loading.Tag.Proc.Start()
      $animTimer.Start()
      $watchTimer.Start()
    } catch {
      $loading.Tag.Ok = $false
      $loading.Close()
    }
  })

  $ok = $false
  try {
    [void]$loading.ShowDialog()
    if ($loading.Tag) { $ok = [bool]$loading.Tag.Ok }
  }
  finally {
    try { $animTimer.Stop(); $animTimer.Dispose() } catch {}
    try { $watchTimer.Stop(); $watchTimer.Dispose() } catch {}
    try { $loading.Dispose() } catch {}
    try { $proc.Dispose() } catch {}
  }

  return $ok
}

function Invoke-ForceExcelUpdate {
  param(
    [string]$Excel,
    [string]$Sheet,
    [ValidateSet('Auto','RitmOnly','IncAndRitm','All')][string]$ProcessingScope = 'All',
    [ValidateSet('Auto','ConfigurationItemOnly','CommentsOnly','CommentsAndCI')][string]$PiSearchMode = 'Auto'
  )

  $active = @(Get-RunningSchumanOperationProcesses -Operations @('Export'))
  if ($active.Count -gt 0) {
    $ids = (($active | Select-Object -ExpandProperty ProcessId) -join ', ')
    [System.Windows.Forms.MessageBox]::Show(
      "There is already an Export process running (PID: $ids). Stop/wait for that process before running Force Update.",
      'Schuman'
    ) | Out-Null
    return [pscustomobject]@{
      ok = $false
      durationSec = 0
      ticketCount = 0
      estimatedSec = 0
    }
  }

  $ticketCount = 0
  try {
    $ticketCount = @(
      Read-TicketsFromExcel -ExcelPath $Excel -SheetName $Sheet -TicketHeader 'Number' -TicketColumn 4 `
        -StopAfterEmptyRows $globalConfig.Excel.StopScanAfterEmptyRows -MaxRowsAfterFirstTicket $globalConfig.Excel.MaxRowsAfterFirstTicket
    ).Count
  }
  catch {}

  $historyEstimate = Get-UiEstimatedSeconds -Config $globalConfig -OperationKey 'force_update_excel' -DefaultSeconds 180
  $heuristicEstimate = if ($ticketCount -gt 0) { [int][Math]::Max(45, [Math]::Min(1800, $ticketCount * 2.2)) } else { $historyEstimate }
  $estimatedSeconds = [int][Math]::Round((0.65 * $historyEstimate) + (0.35 * $heuristicEstimate))
  if ($estimatedSeconds -lt 20) { $estimatedSeconds = 20 }

  $loading = New-Object System.Windows.Forms.Form
  $loading.Text = 'Schuman — Force Update Excel'
  $loading.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $loading.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $loading.ControlBox = $false
  $loading.Size = New-Object System.Drawing.Size(620, 240)
  $loading.TopMost = $true
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 5
  $layout.Padding = New-Object System.Windows.Forms.Padding(20)
  $loading.Controls.Add($layout)

  $title = New-Object System.Windows.Forms.Label
  $title.Text = 'Force updating Excel with latest ServiceNow data'
  $title.AutoSize = $true
  $title.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
  $layout.Controls.Add($title, 0, 0)

  $status = New-Object System.Windows.Forms.Label
  $status.Text = "Scanning tickets$([string]::Concat(''))..."
  $status.AutoSize = $true
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 6)
  $layout.Controls.Add($status, 0, 1)

  $timing = New-Object System.Windows.Forms.Label
  $timing.Text = ("Tickets: {0} | ETA: ~{1:mm\:ss} | Elapsed: 00:00" -f $ticketCount, [TimeSpan]::FromSeconds($estimatedSeconds))
  $timing.AutoSize = $true
  $timing.ForeColor = [System.Drawing.Color]::FromArgb(90,90,96)
  $timing.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
  $layout.Controls.Add($timing, 0, 2)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220,220,225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 3)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(100, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
  $barTrack.Controls.Add($barFill)

  $note = New-Object System.Windows.Forms.Label
  $note.Text = 'This operation rewrites Name/PI/Status/SCTasks in Excel (full refresh).'
  $note.AutoSize = $true
  $note.ForeColor = [System.Drawing.Color]::FromArgb(120,120,126)
  $note.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
  $layout.Controls.Add($note, 0, 4)

  $args = @(
    '-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
    '-File',$invokeScript,
    '-Operation','Export',
    '-ExcelPath',$Excel,
    '-SheetName',$Sheet,
    '-ProcessingScope',$ProcessingScope,
    '-PiSearchMode',$PiSearchMode,
    '-NoPopups'
  )

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = (Join-Path $PSHOME 'powershell.exe')
  $psi.Arguments = Convert-ToArgumentString -Tokens $args
  $psi.WorkingDirectory = $projectRoot
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $runsRoot = Join-Path $globalConfig.Output.SystemRoot $globalConfig.Output.RunsSubdir
  $knownRunDirs = @()
  if (Test-Path -LiteralPath $runsRoot) {
    $knownRunDirs = @(Get-ChildItem -LiteralPath $runsRoot -Directory -Filter 'export_*' | ForEach-Object { $_.FullName })
  }

  $animTimer = New-Object System.Windows.Forms.Timer
  $animTimer.Interval = 180
  $watchTimer = New-Object System.Windows.Forms.Timer
  $watchTimer.Interval = 250

  $state = @{
    Tick = 0
    Ok = $false
    Proc = $proc
    DurationSec = 0
    ProgressDone = 0
    ProgressTotal = 0
    ProgressPct = 0
    RunLogPath = ''
    LastStatusLine = ''
  }
  $loading.Tag = $state

  $resolveRunLog = {
    param($st)
    if ($st.RunLogPath -and (Test-Path -LiteralPath $st.RunLogPath)) { return }
    if (-not (Test-Path -LiteralPath $runsRoot)) { return }

    $dirs = @(Get-ChildItem -LiteralPath $runsRoot -Directory -Filter 'export_*' | Sort-Object LastWriteTime -Descending)
    foreach ($d in $dirs) {
      if ($knownRunDirs -contains $d.FullName) { continue }
      $candidate = Join-Path $d.FullName 'run.log.txt'
      if (Test-Path -LiteralPath $candidate) {
        $st.RunLogPath = $candidate
        return
      }
    }
  }

  $updateProgressFromLog = {
    param($st)
    if (-not $st.RunLogPath -or -not (Test-Path -LiteralPath $st.RunLogPath)) { return }
    try {
      $tail = @(Get-Content -LiteralPath $st.RunLogPath -Tail 120)
      foreach ($line in $tail) {
        if ($line -match '\[(\d+)\/(\d+)\]\s+Extracting\s+([A-Z]+\d+)') {
          $d = [int]$matches[1]
          $t = [int]$matches[2]
          if ($t -gt 0) {
            $st.ProgressDone = $d
            $st.ProgressTotal = $t
            $st.ProgressPct = [int][Math]::Min(99, [Math]::Round(($d * 100.0) / $t))
            $st.LastStatusLine = "Processing $($matches[3]) ($d/$t)"
          }
        }
      }
      if ($tail.Count -gt 0) {
        $last = ("" + $tail[$tail.Count - 1]).Trim()
        if ($last) { $st.LastStatusLine = $last }
      }
    }
    catch {}
  }

  $animTimer.Add_Tick({
    $st = $loading.Tag
    if (-not $st) { return }
    $st.Tick = [int]$st.Tick + 1
    & $resolveRunLog $st
    & $updateProgressFromLog $st

    $dots = '.' * (([int]$st.Tick % 3) + 1)
    if ($st.ProgressTotal -gt 0) {
      $status.Text = "Updating from ServiceNow$dots  [$($st.ProgressDone)/$($st.ProgressTotal)]"
    } else {
      $status.Text = 'Updating from ServiceNow' + $dots
    }

    $elapsed = [int][Math]::Floor($sw.Elapsed.TotalSeconds)
    $etaSec = [Math]::Max(0, $estimatedSeconds - $elapsed)
    if ($st.ProgressTotal -gt 0 -and $st.ProgressDone -gt 0) {
      $rate = $elapsed / [double]$st.ProgressDone
      $etaFromReal = [int][Math]::Round(($st.ProgressTotal - $st.ProgressDone) * $rate)
      if ($etaFromReal -ge 0) { $etaSec = $etaFromReal }
    }
    if ($elapsed -gt $estimatedSeconds) {
      $etaSec = [int][Math]::Max(5, [Math]::Round($estimatedSeconds * 0.2))
    }
    $pctText = if ($st.ProgressTotal -gt 0) { "$($st.ProgressPct)%" } else { '--%' }
    $timing.Text = ("Tickets: {0} | Progress: {1} | ETA: ~{2:mm\:ss} | Elapsed: {3:mm\:ss}" -f $ticketCount, $pctText, [TimeSpan]::FromSeconds($etaSec), [TimeSpan]::FromSeconds($elapsed))

    if ($st.ProgressTotal -gt 0) {
      $trackW = [Math]::Max(1, $barTrack.ClientSize.Width)
      $targetW = [int][Math]::Max(14, [Math]::Round(($trackW * $st.ProgressPct) / 100.0))
      $barFill.Width = [Math]::Min($trackW, $targetW)
      $barFill.Left = 0
    } else {
      $barFill.Width = 100
      $maxX = [Math]::Max(1, $barTrack.ClientSize.Width - $barFill.Width)
      $barFill.Left = ([int]$st.Tick * 16) % $maxX
    }
  })

  $watchTimer.Add_Tick({
    $st = $loading.Tag
    if (-not $st -or -not $st.Proc) { return }
    if ($st.Proc.HasExited) {
      $st.Ok = ($st.Proc.ExitCode -eq 0)
      if ($st.Ok) {
        $st.ProgressPct = 100
        $st.ProgressDone = [Math]::Max($st.ProgressDone, $st.ProgressTotal)
      }
      $st.DurationSec = [int][Math]::Max(1, [Math]::Round($sw.Elapsed.TotalSeconds))
      $loading.Close()
    }
  })

  $loading.add_Shown({
    try {
      [void]$loading.Tag.Proc.Start()
      $animTimer.Start()
      $watchTimer.Start()
    } catch {
      $loading.Tag.Ok = $false
      $loading.Close()
    }
  })

  $ok = $false
  $durationSec = 0
  try {
    [void]$loading.ShowDialog()
    if ($loading.Tag) {
      $ok = [bool]$loading.Tag.Ok
      $durationSec = [int]$loading.Tag.DurationSec
    }
  }
  finally {
    try { $animTimer.Stop(); $animTimer.Dispose() } catch {}
    try { $watchTimer.Stop(); $watchTimer.Dispose() } catch {}
    try { $loading.Dispose() } catch {}
    try { $proc.Dispose() } catch {}
    try { $sw.Stop() } catch {}
  }

  if ($ok -and $durationSec -gt 0) {
    Update-UiMetricDuration -Config $globalConfig -OperationKey 'force_update_excel' -DurationSeconds $durationSec
  }

  return [pscustomobject]@{
    ok = $ok
    durationSec = $durationSec
    ticketCount = $ticketCount
    estimatedSec = $estimatedSeconds
  }
}

function Resolve-UiForm {
  param(
    [Parameter(Mandatory = $true)]$UiResult,
    [Parameter(Mandatory = $true)][string]$UiName
  )

  if ($UiResult -is [System.Windows.Forms.Form]) { return $UiResult }

  $items = @($UiResult)
  if ($items.Count -gt 0) {
    $form = @($items | Where-Object { $_ -is [System.Windows.Forms.Form] } | Select-Object -Last 1)
    if ($form.Count -gt 0) { return $form[0] }
  }

  throw "UI factory '$UiName' did not return a WinForms Form."
}

function Start-StartupSsoSession {
  param(
    [hashtable]$Config,
    [hashtable]$RunContext
  )

  try {
    return (New-ServiceNowSession -Config $Config -RunContext $RunContext)
  }
  catch {
    [System.Windows.Forms.MessageBox]::Show(
      "ServiceNow SSO login is mandatory at startup.`r`n`r`n$($_.Exception.Message)",
      'SSO Verification'
    ) | Out-Null
    return $null
  }
}

$startupSession = Start-StartupSsoSession -Config $globalConfig -RunContext $uiRunContext
if (-not $startupSession) {
  return
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Schuman — Main'
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.MinimumSize = New-Object System.Drawing.Size(880, 520)
$form.Size = New-Object System.Drawing.Size(920, 560)
$form.BackColor = [System.Drawing.Color]::FromArgb(245,245,247)
$form.Font = New-Object System.Drawing.Font((Get-UiFontName), 10)

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = [System.Windows.Forms.DockStyle]::Fill
$root.ColumnCount = 3
$root.RowCount = 3
$root.Padding = New-Object System.Windows.Forms.Padding(24)
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 80)))
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($root)

$hdr = New-Object System.Windows.Forms.Label
$hdr.Text = 'Schuman Main'
$hdr.Font = New-Object System.Drawing.Font((Get-UiFontName), 20, [System.Drawing.FontStyle]::Bold)
$hdr.ForeColor = [System.Drawing.Color]::FromArgb(28,28,30)
$hdr.AutoSize = $true
$hdr.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$root.Controls.Add($hdr, 1, 0)

$card = New-CardContainer -Title 'Modules'
$card.Border.Dock = [System.Windows.Forms.DockStyle]::Fill
$root.Controls.Add($card.Border, 1, 1)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = [System.Windows.Forms.DockStyle]::Top
$layout.AutoSize = $true
$layout.ColumnCount = 1
$layout.RowCount = 8
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 8)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 8)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$card.Content.Controls.Add($layout)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = 'Excel file'
$lbl.AutoSize = $true
$lbl.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
$layout.Controls.Add($lbl, 0, 0)

$txtExcel = New-Object System.Windows.Forms.TextBox
$txtExcel.Dock = [System.Windows.Forms.DockStyle]::Top
$txtExcel.Text = $ExcelPath
$txtExcel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$layout.Controls.Add($txtExcel, 0, 1)

$btns = New-Object System.Windows.Forms.TableLayoutPanel
$btns.Dock = [System.Windows.Forms.DockStyle]::Top
$btns.AutoSize = $true
$btns.ColumnCount = 2
$btns.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$btns.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$layout.Controls.Add($btns, 0, 3)

$btnDashboard = New-Object System.Windows.Forms.Button
$btnDashboard.Text = 'Dashboard'
$btnDashboard.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnDashboard.Height = 46
$btnDashboard.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnDashboard.FlatAppearance.BorderSize = 0
$btnDashboard.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
$btnDashboard.ForeColor = [System.Drawing.Color]::White
$btnDashboard.Font = New-Object System.Drawing.Font((Get-UiFontName), 10, [System.Drawing.FontStyle]::Bold)
$btnDashboard.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btns.Controls.Add($btnDashboard, 0, 0)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = 'Generate'
$btnGenerate.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnGenerate.Height = 46
$btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenerate.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(230,230,235)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(255,255,255)
$btnGenerate.ForeColor = [System.Drawing.Color]::FromArgb(28,28,30)
$btnGenerate.Font = New-Object System.Drawing.Font((Get-UiFontName), 10, [System.Drawing.FontStyle]::Bold)
$btnGenerate.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btns.Controls.Add($btnGenerate, 1, 0)

$btnForce = New-Object System.Windows.Forms.Button
$btnForce.Text = 'Force Update Excel'
$btnForce.Dock = [System.Windows.Forms.DockStyle]::Top
$btnForce.Height = 40
$btnForce.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnForce.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(222, 152, 0)
$btnForce.FlatAppearance.BorderSize = 1
$btnForce.BackColor = [System.Drawing.Color]::FromArgb(255, 248, 230)
$btnForce.ForeColor = [System.Drawing.Color]::FromArgb(146, 88, 0)
$btnForce.Font = New-Object System.Drawing.Font((Get-UiFontName), 10, [System.Drawing.FontStyle]::Bold)
$btnForce.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$layout.Controls.Add($btnForce, 0, 5)

$btnEmergency = New-Object System.Windows.Forms.TableLayoutPanel
$btnEmergency.Dock = [System.Windows.Forms.DockStyle]::Top
$btnEmergency.AutoSize = $true
$btnEmergency.ColumnCount = 2
$btnEmergency.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$btnEmergency.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$layout.Controls.Add($btnEmergency, 0, 7)

$btnCloseCode = New-Object System.Windows.Forms.Button
$btnCloseCode.Text = 'Cerrar codigo'
$btnCloseCode.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnCloseCode.Height = 36
$btnCloseCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCloseCode.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210, 80, 80)
$btnCloseCode.FlatAppearance.BorderSize = 1
$btnCloseCode.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 238)
$btnCloseCode.ForeColor = [System.Drawing.Color]::FromArgb(160, 32, 32)
$btnCloseCode.Font = New-Object System.Drawing.Font((Get-UiFontName), 10, [System.Drawing.FontStyle]::Bold)
$btnCloseCode.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btnEmergency.Controls.Add($btnCloseCode, 0, 0)

$btnCloseDocs = New-Object System.Windows.Forms.Button
$btnCloseDocs.Text = 'Cerrar documentos'
$btnCloseDocs.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnCloseDocs.Height = 36
$btnCloseDocs.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCloseDocs.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(210, 80, 80)
$btnCloseDocs.FlatAppearance.BorderSize = 1
$btnCloseDocs.BackColor = [System.Drawing.Color]::FromArgb(255, 238, 238)
$btnCloseDocs.ForeColor = [System.Drawing.Color]::FromArgb(160, 32, 32)
$btnCloseDocs.Font = New-Object System.Drawing.Font((Get-UiFontName), 10, [System.Drawing.FontStyle]::Bold)
$btnCloseDocs.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btnEmergency.Controls.Add($btnCloseDocs, 1, 0)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Status: Ready (SSO connected)'
$status.AutoSize = $true
$status.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
$root.Controls.Add($status, 1, 2)

$openModule = {
  param([string]$module)

  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    [System.Windows.Forms.MessageBox]::Show('Select a valid Excel file first.','Validation') | Out-Null
    return
  }

  $status.Text = 'Status: Opening module...'
  if ($module -eq 'Dashboard') {
    $frm = Resolve-UiForm -UiResult (New-DashboardUI -ExcelPath $excel -SheetName $SheetName -Config $globalConfig -RunContext $uiRunContext -InitialSession $startupSession) -UiName 'New-DashboardUI'
    [void]$frm.ShowDialog($form)
  }
  else {
    $defaultTemplate = Join-Path $projectRoot $globalConfig.Documents.TemplateFile
    $defaultOutput = Join-Path $projectRoot $globalConfig.Documents.OutputFolder

    $frm = Resolve-UiForm -UiResult (New-GeneratePdfUI -ExcelPath $excel -SheetName $SheetName -TemplatePath $defaultTemplate -OutputPath $defaultOutput -OnOpenDashboard {
      $d = Resolve-UiForm -UiResult (New-DashboardUI -ExcelPath $excel -SheetName $SheetName -Config $globalConfig -RunContext $uiRunContext -InitialSession $startupSession) -UiName 'New-DashboardUI'
      [void]$d.ShowDialog($form)
    } -OnGenerate {
      param($argsObj)
      $okRun = Invoke-DocsGenerate -Excel $argsObj.ExcelPath -Sheet $SheetName -Template $argsObj.TemplatePath -Output $argsObj.OutputPath -ExportPdf:[bool]$argsObj.ExportPdf
      if ($okRun) {
        return [pscustomobject]@{
          ok = $true
          message = 'Documents generated successfully.'
          outputPath = $argsObj.OutputPath
        }
      }
      return [pscustomobject]@{
        ok = $false
        message = 'Document generation failed. Check logs under system/runs.'
      }
    }) -UiName 'New-GeneratePdfUI'
    [void]$frm.ShowDialog($form)
  }
  $status.Text = 'Status: Ready'
}

$btnDashboard.Add_Click({ & $openModule 'Dashboard' })
$btnGenerate.Add_Click({ & $openModule 'Generate' })
$btnCloseCode.Add_Click({
  $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar codigo' -ExecutableNames @('code.exe', 'code-insiders.exe', 'cursor.exe') -Owner $form
  if (-not $r.Cancelled) {
    $status.Text = "Status: $($r.Message)"
  }
})
$btnCloseDocs.Add_Click({
  $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar documentos' -ExecutableNames @('winword.exe', 'excel.exe') -Owner $form
  if (-not $r.Cancelled) {
    $status.Text = "Status: $($r.Message)"
  }
})
$btnForce.Add_Click({
  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    [System.Windows.Forms.MessageBox]::Show('Select a valid Excel file first.','Validation') | Out-Null
    return
  }

  $opts = Show-ForceUpdateOptionsDialog -Owner $form
  if (-not $opts -or -not $opts.ok) {
    $status.Text = 'Status: Force update canceled'
    return
  }

  $btnDashboard.Enabled = $false
  $btnGenerate.Enabled = $false
  $btnForce.Enabled = $false
  $status.Text = ("Status: Running force update ({0}, {1})..." -f $opts.processingScope, $opts.piSearchMode)
  try {
    $res = Invoke-ForceExcelUpdate -Excel $excel -Sheet $SheetName -ProcessingScope $opts.processingScope -PiSearchMode $opts.piSearchMode
    if ($res.ok) {
      $status.Text = ("Status: Force update complete ({0}s, tickets={1})" -f $res.durationSec, $res.ticketCount)
      [System.Windows.Forms.MessageBox]::Show(
        "Force update completed successfully.`r`nDuration: $($res.durationSec)s`r`nTickets: $($res.ticketCount)`r`nScope: $($opts.processingScope)`r`nPI mode: $($opts.piSearchMode)",
        'Force Update'
      ) | Out-Null
    } else {
      $status.Text = 'Status: Force update failed'
      [System.Windows.Forms.MessageBox]::Show('Force update failed. Check logs under system/runs.','Error') | Out-Null
    }
  }
  finally {
    $btnDashboard.Enabled = $true
    $btnGenerate.Enabled = $true
    $btnForce.Enabled = $true
  }
})

[void]$form.add_FormClosed(({
  try {
    if ($startupSession) { Close-ServiceNowSession -Session $startupSession }
  } catch {}
}).GetNewClosure())

[void]$form.ShowDialog()
