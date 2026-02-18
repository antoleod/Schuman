#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot '..\..\Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

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

function Invoke-PreloadExport {
  param(
    [string]$Excel,
    [string]$Sheet
  )

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

  $bar = New-Object System.Windows.Forms.ProgressBar
  $bar.Dock = [System.Windows.Forms.DockStyle]::Top
  $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $bar.MarqueeAnimationSpeed = 24
  $layout.Controls.Add($bar, 0, 2)

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

  $bar = New-Object System.Windows.Forms.ProgressBar
  $bar.Dock = [System.Windows.Forms.DockStyle]::Top
  $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
  $bar.MarqueeAnimationSpeed = 24
  $layout.Controls.Add($bar, 0, 2)

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
$layout.RowCount = 4
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 10)))
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

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Status: Ready'
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

  $ok = $true
  if ($module -eq 'Generate') {
    $btnDashboard.Enabled = $false
    $btnGenerate.Enabled = $false
    $status.Text = 'Status: Running pre-load export...'
    try { $ok = Invoke-PreloadExport -Excel $excel -Sheet $SheetName }
    finally {
      $btnDashboard.Enabled = $true
      $btnGenerate.Enabled = $true
    }
    if (-not $ok) {
      $status.Text = 'Status: Pre-load failed'
      [System.Windows.Forms.MessageBox]::Show('Pre-load export failed. Check SSO/config and retry.','Error') | Out-Null
      return
    }
  }

  $status.Text = 'Status: Opening module...'
  if ($module -eq 'Dashboard') {
    $frm = Resolve-UiForm -UiResult (New-DashboardUI -ExcelPath $excel -SheetName $SheetName -Config $globalConfig -RunContext $uiRunContext) -UiName 'New-DashboardUI'
    [void]$frm.ShowDialog($form)
  }
  else {
    $defaultTemplate = Join-Path $projectRoot $globalConfig.Documents.TemplateFile
    $defaultOutput = Join-Path $projectRoot $globalConfig.Documents.OutputFolder

    $frm = Resolve-UiForm -UiResult (New-GeneratePdfUI -ExcelPath $excel -TemplatePath $defaultTemplate -OutputPath $defaultOutput -OnOpenDashboard {
      $d = Resolve-UiForm -UiResult (New-DashboardUI -ExcelPath $excel -SheetName $SheetName -Config $globalConfig -RunContext $uiRunContext) -UiName 'New-DashboardUI'
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

[void]$form.ShowDialog()
