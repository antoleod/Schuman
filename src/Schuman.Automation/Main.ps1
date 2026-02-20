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

Initialize-SchumanRuntime -RevalidateOnly
Assert-SchumanRuntime -RequiredCommands @(
  'Get-UiFontName',
  'New-CardContainer',
  'Invoke-UiEmergencyClose',
  'Search-DashboardRows'
)

$script:GetUiFontNameHandler = ${function:Get-UiFontName}
function Get-UiFontNameSafe {
  if ($script:GetUiFontNameHandler) {
    try {
      $name = ("" + (& $script:GetUiFontNameHandler)).Trim()
      if ($name) { return $name }
    }
    catch {}
  }
  return 'Segoe UI'
}
$script:GetUiFontNameSafeHandler = ${function:Get-UiFontNameSafe}

$script:ExternalInvokeUiEmergencyClose = $null
try {
  $existingEmergencyClose = Get-Command -Name Invoke-UiEmergencyClose -CommandType Function -ErrorAction SilentlyContinue
  if ($existingEmergencyClose -and $existingEmergencyClose.ScriptBlock) {
    $script:ExternalInvokeUiEmergencyClose = $existingEmergencyClose.ScriptBlock
  }
}
catch {}

function Initialize-UiEmergencyCloseDependency {
  $scriptDir = Split-Path -Parent $PSCommandPath
  $candidates = @(
    (Join-Path $scriptDir 'ui_helpers.ps1'),
    (Join-Path $scriptDir 'helpers.ps1'),
    (Join-Path $scriptDir 'common.ps1'),
    (Join-Path $scriptDir 'Schuman.UI.Helpers.ps1'),
    (Join-Path $scriptDir 'UI\ui_helpers.ps1'),
    (Join-Path $scriptDir 'UI\helpers.ps1'),
    (Join-Path $scriptDir 'UI\common.ps1'),
    (Join-Path $scriptDir 'UI\Schuman.UI.Helpers.ps1')
  )

  foreach ($file in $candidates) {
    try {
      if (-not (Test-Path -LiteralPath $file)) { continue }
      . $file
      $cmd = Get-Command -Name Invoke-UiEmergencyClose -CommandType Function -ErrorAction SilentlyContinue
      if ($cmd -and $cmd.ScriptBlock) {
        $script:ExternalInvokeUiEmergencyClose = $cmd.ScriptBlock
        return
      }
    }
    catch {}
  }
}

function global:Release-ComObjectSafe {
  param($obj)
  if ($null -eq $obj) { return }
  try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch {}
}

if (-not (Get-Variable -Name SchumanOwnedComResources -Scope Script -ErrorAction SilentlyContinue)) {
  $script:SchumanOwnedComResources = New-Object System.Collections.ArrayList
}
if (-not (Get-Variable -Name SchumanOwnedProcesses -Scope Script -ErrorAction SilentlyContinue)) {
  $script:SchumanOwnedProcesses = @{}
}

function global:Register-SchumanOwnedComResource {
  param(
    [Parameter(Mandatory = $true)][string]$Kind,
    [Parameter(Mandatory = $true)]$Object,
    [string]$Tag = ''
  )
  if ($null -eq $Object) { return $null }
  $entry = [pscustomobject]@{
    Kind = ("" + $Kind).Trim()
    Object = $Object
    Tag = ("" + $Tag).Trim()
    AddedAt = [DateTime]::UtcNow
    Id = [Guid]::NewGuid().ToString('N')
  }
  [void]$script:SchumanOwnedComResources.Add($entry)
  return $entry
}

function global:Unregister-SchumanOwnedComResource {
  param($Object)
  if ($null -eq $Object) { return }
  for ($i = $script:SchumanOwnedComResources.Count - 1; $i -ge 0; $i--) {
    $item = $script:SchumanOwnedComResources[$i]
    if ($item -and $item.Object -eq $Object) {
      [void]$script:SchumanOwnedComResources.RemoveAt($i)
    }
  }
}

function global:Register-SchumanOwnedProcess {
  param(
    [Parameter(Mandatory = $true)][System.Diagnostics.Process]$Process,
    [string]$Tag = ''
  )
  if (-not $Process) { return }
  try {
    $script:SchumanOwnedProcesses[[string]$Process.Id] = [pscustomobject]@{
      Id = [int]$Process.Id
      Name = ("" + $Process.ProcessName).Trim()
      Process = $Process
      Tag = ("" + $Tag).Trim()
      AddedAt = [DateTime]::UtcNow
    }
  } catch {}
}

function global:Unregister-SchumanOwnedProcess {
  param([int]$ProcessId)
  $key = [string]$ProcessId
  if ($script:SchumanOwnedProcesses.ContainsKey($key)) {
    $script:SchumanOwnedProcesses.Remove($key)
  }
}

function global:Stop-SchumanOwnedResources {
  param(
    [ValidateSet('Code','Documents','All')][string]$Mode = 'All'
  )

  $comClosed = 0
  $procClosed = 0
  $procFailed = 0

  $closeComEntry = {
    param($entry)
    if (-not $entry) { return }
    $kind = ("" + $entry.Kind).Trim().ToLowerInvariant()
    $obj = $entry.Object
    if ($null -eq $obj) { return }
    try {
      if ($kind -like '*doc*' -or $kind -eq 'workbook') {
        try { $obj.Close($false) | Out-Null } catch {}
      }
      elseif ($kind -like '*word*' -or $kind -like '*excel*') {
        try { $obj.Quit() | Out-Null } catch {}
      }
    } catch {}
    try { Release-ComObjectSafe $obj } catch {}
    $script:SchumanOwnedComResources.Remove($entry) | Out-Null
    $comClosed++
  }

  if ($Mode -eq 'Documents' -or $Mode -eq 'All') {
    $snapshot = @($script:SchumanOwnedComResources)
    for ($i = $snapshot.Count - 1; $i -ge 0; $i--) {
      $entry = $snapshot[$i]
      if (-not $entry) { continue }
      $kind = ("" + $entry.Kind).Trim().ToLowerInvariant()
      if ($kind -like '*excel*' -or $kind -like '*word*' -or $kind -like '*doc*' -or $kind -eq 'workbook') {
        & $closeComEntry $entry
      }
    }
  }

  $ownedSnapshot = @($script:SchumanOwnedProcesses.Values)
  foreach ($p in $ownedSnapshot) {
    if (-not $p) { continue }
    $tag = ("" + $p.Tag).Trim().ToLowerInvariant()
    $matchesMode = ($Mode -eq 'All') -or (($Mode -eq 'Documents') -and ($tag -match 'doc|export|excel|word|force')) -or (($Mode -eq 'Code') -and ($tag -match 'code|main|ui'))
    if (-not $matchesMode) { continue }
    try {
      if ($p.Process -and (-not $p.Process.HasExited)) {
        try { $p.Process.CloseMainWindow() | Out-Null } catch {}
        Start-Sleep -Milliseconds 600
        if (-not $p.Process.HasExited) {
          try { Stop-Process -Id $p.Id -Force -ErrorAction Stop; $procClosed++ } catch { $procFailed++ }
        }
      }
      else { $procClosed++ }
    } catch { $procFailed++ }
    Unregister-SchumanOwnedProcess -ProcessId $p.Id
  }

  return [pscustomobject]@{
    ComClosedCount = [int]$comClosed
    ProcessClosedCount = [int]$procClosed
    ProcessFailedCount = [int]$procFailed
  }
}

function global:Invoke-UiEmergencyClose {
  param(
    [string]$ActionLabel = '',
    [string[]]$ExecutableNames = @(),
    [System.Windows.Forms.IWin32Window]$Owner = $null,
    [ValidateSet('Code','Documents','All')][string]$Mode = 'All',
    [System.Windows.Forms.Form]$MainForm = $null,
    [string]$BaseDir = ''
  )

  if (-not $script:ExternalInvokeUiEmergencyClose) {
    Initialize-UiEmergencyCloseDependency
  }

  $externalResult = $null
  if ($script:ExternalInvokeUiEmergencyClose) {
    try { $externalResult = (& $script:ExternalInvokeUiEmergencyClose -ActionLabel $ActionLabel -ExecutableNames $ExecutableNames -Owner $Owner) } catch {}
  }

  $labelText = ("" + $ActionLabel).Trim().ToLowerInvariant()
  $modeResolved = $Mode
  $modeWasExplicit = ($PSBoundParameters.ContainsKey('Mode') -and $Mode -ne 'All')
  if (-not $modeWasExplicit -and $labelText -match 'codigo') { $modeResolved = 'Code' }
  elseif (-not $modeWasExplicit -and $labelText -match 'document') { $modeResolved = 'Documents' }
  elseif (-not $modeWasExplicit -and $ExecutableNames -and @($ExecutableNames).Count -gt 0) {
    $joined = ("" + ($ExecutableNames -join ' ')).ToLowerInvariant()
    if ($joined -match 'code|cursor') { $modeResolved = 'Code' }
    elseif ($joined -match 'word|excel') { $modeResolved = 'Documents' }
  }

  $cleanup = Stop-SchumanOwnedResources -Mode $modeResolved
  if ($externalResult -and $externalResult.PSObject.Properties['KilledCount']) {
    try { $cleanup.ProcessClosedCount = [int]$cleanup.ProcessClosedCount + [int]$externalResult.KilledCount } catch {}
  }
  if ($externalResult -and $externalResult.PSObject.Properties['FailedCount']) {
    try { $cleanup.ProcessFailedCount = [int]$cleanup.ProcessFailedCount + [int]$externalResult.FailedCount } catch {}
  }

  if ($modeResolved -eq 'Code' -or $modeResolved -eq 'All') {
    try { Close-SchumanOpenForms } catch {}
    if ($MainForm) {
      try { $MainForm.Close() } catch {}
    }
  }

  return [pscustomobject]@{
    Cancelled = $false
    KilledCount = [int]$cleanup.ProcessClosedCount
    FailedCount = [int]$cleanup.ProcessFailedCount
    Message = ("Cleanup done (COM={0}, ProcClosed={1}, ProcFailed={2})" -f $cleanup.ComClosedCount, $cleanup.ProcessClosedCount, $cleanup.ProcessFailedCount)
  }
}

function global:Show-UiError {
  param(
    [string]$Title = 'Schuman',
    [string]$Message = '',
    [System.Exception]$Exception = $null,
    [string]$Context = '',
    $ErrorRecord = $null
  )

  $safeTitle = if ([string]::IsNullOrWhiteSpace($Title)) { 'Schuman' } else { $Title }
  $safeMessage = ("" + $Message).Trim()
  if ([string]::IsNullOrWhiteSpace($safeMessage) -and -not [string]::IsNullOrWhiteSpace($Context)) {
    $safeMessage = "$Context failed."
  }
  if ([string]::IsNullOrWhiteSpace($safeMessage)) {
    $safeMessage = 'An unexpected error occurred.'
  }

  $detail = ''
  try {
    if ($Exception) {
      $detail = ("" + $Exception.Message).Trim()
    }
    elseif ($ErrorRecord -and $ErrorRecord.Exception) {
      $detail = ("" + $ErrorRecord.Exception.Message).Trim()
    }
  }
  catch {}

  $full = if ($detail) { "$safeMessage`r`n`r`n$detail" } else { $safeMessage }
  try { [System.Windows.Forms.MessageBox]::Show($full, $safeTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
  try {
    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
      Write-Log -Level ERROR -Message $full
    }
  }
  catch {}
}

function global:Show-UiInfo {
  param(
    [string]$Title = 'Schuman',
    [string]$Message = 'Done.'
  )
  $safeTitle = if ([string]::IsNullOrWhiteSpace($Title)) { 'Schuman' } else { $Title }
  $safeMessage = if ([string]::IsNullOrWhiteSpace($Message)) { 'Done.' } else { $Message }
  try { [System.Windows.Forms.MessageBox]::Show($safeMessage, $safeTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null } catch {}
}

function global:Invoke-SafeUiAction {
  param(
    [string]$ActionName = 'UI Action',
    [scriptblock]$Action,
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  try {
    if (-not $Action) { return $null }
    return (& $Action)
  }
  catch {
    $errText = ("" + $_.Exception.Message).Trim()
    if (-not $errText) { $errText = 'Unknown UI error.' }
    try {
      if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
        Write-Log -Level ERROR -Message ("{0}: {1}" -f $ActionName, $errText)
      }
    }
    catch {}
    Show-UiError -Title 'Schuman' -Message ("{0} failed." -f $ActionName) -Exception $_.Exception
    return $null
  }
}

if (-not (Get-Variable -Name WinFormsExceptionHandlingRegistered -Scope Script -ErrorAction SilentlyContinue)) {
  [bool]$script:WinFormsExceptionHandlingRegistered = $false
}
$script:ThreadExceptionHandler = $null
$script:DomainExceptionHandler = $null

function Register-WinFormsGlobalExceptionHandling {
  if ($script:WinFormsExceptionHandlingRegistered) { return }

  try {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
  }
  catch {
    Write-Log -Level WARN -Message ("SetUnhandledExceptionMode failed: " + $_.Exception.Message)
  }

  $script:ThreadExceptionHandler = [System.Threading.ThreadExceptionEventHandler]{
    param($sender, $eventArgs)
    try {
      $ex = $null
      if ($eventArgs) { $ex = $eventArgs.Exception }
      $err = if ($ex) { [System.Management.Automation.ErrorRecord]::new($ex, 'ThreadException', [System.Management.Automation.ErrorCategory]::NotSpecified, $null) } else { $null }
      Show-UiError -Context 'ThreadException' -ErrorRecord $err
    }
    catch {}
  }
  [System.Windows.Forms.Application]::add_ThreadException($script:ThreadExceptionHandler)

  $script:DomainExceptionHandler = [System.UnhandledExceptionEventHandler]{
    param($sender, $eventArgs)
    try {
      $ex = $eventArgs.ExceptionObject -as [System.Exception]
      $err = if ($ex) { [System.Management.Automation.ErrorRecord]::new($ex, 'UnhandledException', [System.Management.Automation.ErrorCategory]::NotSpecified, $null) } else { $null }
      Show-UiError -Context 'AppDomain.UnhandledException' -ErrorRecord $err
    }
    catch {}
  }
  [AppDomain]::CurrentDomain.add_UnhandledException($script:DomainExceptionHandler)

  $script:WinFormsExceptionHandlingRegistered = $true
}

function global:Invoke-UiHandler {
  param(
    [string]$Context,
    [scriptblock]$Action
  )
  Invoke-SafeUiAction -ActionName $Context -Action $Action | Out-Null
}

Register-WinFormsGlobalExceptionHandling

function global:Close-SchumanOpenForms {
  param(
    [System.Windows.Forms.Form]$Except = $null
  )
  try {
    $forms = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ })
    foreach ($openForm in $forms) {
      if (-not $openForm -or $openForm.IsDisposed) { continue }
      if ($Except -and ($openForm -eq $Except)) { continue }
      try { $openForm.Close() } catch {}
    }
  }
  catch {}
}

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

function Get-MainUiPrefsPath {
  param([hashtable]$Config)
  return (Join-Path (Join-Path $Config.Output.SystemRoot $Config.Output.DbSubdir) 'ui-preferences.json')
}

function Get-MainUiPrefs {
  param([hashtable]$Config)

  $defaults = @{
    theme = 'Dark'
    accent = 'Blue'
    fontScale = 100
    compact = $false
  }

  $path = Get-MainUiPrefsPath -Config $Config
  if (-not (Test-Path -LiteralPath $path)) { return $defaults }
  try {
    $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($json) {
      $theme = ("" + $json.theme).Trim()
      $accent = ("" + $json.accent).Trim()
      $fontScale = 100
      try { $fontScale = [int]$json.fontScale } catch {}
      $compact = $false
      try { $compact = [bool]$json.compact } catch {}
      return @{
        theme = if ($theme) { $theme } else { $defaults.theme }
        accent = if ($accent) { $accent } else { $defaults.accent }
        fontScale = [Math]::Min(130, [Math]::Max(90, $fontScale))
        compact = $compact
      }
    }
  }
  catch {}

  return $defaults
}

function Save-MainUiPrefs {
  param(
    [hashtable]$Config,
    [hashtable]$Prefs
  )
  $path = Get-MainUiPrefsPath -Config $Config
  Ensure-Directory -Path (Split-Path -Parent $path) | Out-Null
  ($Prefs | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8
}

function global:Show-ForceUpdateOptionsDialog {
  param(
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Force Update Options'
  $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.Size = New-Object System.Drawing.Size(700, 520)
  $dlg.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 10)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 6
  $layout.Padding = New-Object System.Windows.Forms.Padding(16)
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
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

  $perfPanel = New-Object System.Windows.Forms.Panel
  $perfPanel.Dock = [System.Windows.Forms.DockStyle]::Top
  $perfPanel.Height = 96
  $perfPanel.BackColor = [System.Drawing.Color]::FromArgb(32,32,34)
  $perfPanel.Padding = New-Object System.Windows.Forms.Padding(12)
  $perfPanel.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
  [void]$layout.Controls.Add($perfPanel, 0, 2)

  $lblPerf = New-Object System.Windows.Forms.Label
  $lblPerf.Text = 'Performance'
  $lblPerf.AutoSize = $true
  $lblPerf.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
  $lblPerf.ForeColor = [System.Drawing.Color]::FromArgb(170,170,175)
  $perfPanel.Controls.Add($lblPerf)

  $chkFastMode = New-Object System.Windows.Forms.CheckBox
  $chkFastMode.Text = 'Fast mode (skip deep legal name fallback)'
  $chkFastMode.AutoSize = $true
  $chkFastMode.Checked = $true
  $chkFastMode.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $chkFastMode.Location = New-Object System.Drawing.Point(0, 30)
  $perfPanel.Controls.Add($chkFastMode)

  $lblMax = New-Object System.Windows.Forms.Label
  $lblMax.Text = 'Max tickets (0 = all)'
  $lblMax.AutoSize = $true
  $lblMax.ForeColor = [System.Drawing.Color]::FromArgb(180,180,184)
  $lblMax.Location = New-Object System.Drawing.Point(0, 58)
  $perfPanel.Controls.Add($lblMax)

  $numMaxTickets = New-Object System.Windows.Forms.NumericUpDown
  $numMaxTickets.Minimum = 0
  $numMaxTickets.Maximum = 10000
  $numMaxTickets.Value = 0
  $numMaxTickets.Width = 100
  $numMaxTickets.Location = New-Object System.Drawing.Point(170, 56)
  $numMaxTickets.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $numMaxTickets.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $perfPanel.Controls.Add($numMaxTickets)

  $hint = New-Object System.Windows.Forms.Label
  $hint.Text = 'Tip: RITM only + Configuration Item only is usually the fastest combination.'
  $hint.AutoSize = $true
  $hint.ForeColor = [System.Drawing.Color]::FromArgb(120,120,126)
  $hint.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
  [void]$layout.Controls.Add($hint, 0, 3)

  $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttons.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $buttons.AutoSize = $true
  [void]$layout.Controls.Add($buttons, 0, 5)

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
    fastMode = [bool]$chkFastMode.Checked
    maxTickets = [int]$numMaxTickets.Value
  }
}

function global:New-FallbackForceUpdateOptionsDialog {
  param(
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Force Update Options'
  $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dlg.Size = New-Object System.Drawing.Size(460, 270)
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

  $root = New-Object System.Windows.Forms.TableLayoutPanel
  $root.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Padding = New-Object System.Windows.Forms.Padding(12)
  $root.ColumnCount = 2
  $root.RowCount = 4
  [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $dlg.Controls.Add($root)

  $lblScope = New-Object System.Windows.Forms.Label
  $lblScope.Text = 'Scope'
  $lblScope.AutoSize = $true
  $root.Controls.Add($lblScope, 0, 0)

  $cmbScope = New-Object System.Windows.Forms.ComboBox
  $cmbScope.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$cmbScope.Items.AddRange(@('RitmOnly','IncAndRitm','All'))
  $cmbScope.SelectedItem = 'RitmOnly'
  $cmbScope.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Controls.Add($cmbScope, 1, 0)

  $lblPi = New-Object System.Windows.Forms.Label
  $lblPi.Text = 'PI mode'
  $lblPi.AutoSize = $true
  $root.Controls.Add($lblPi, 0, 1)

  $cmbPi = New-Object System.Windows.Forms.ComboBox
  $cmbPi.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$cmbPi.Items.AddRange(@('ConfigurationItemOnly','CommentsOnly','CommentsAndCI','Auto'))
  $cmbPi.SelectedItem = 'ConfigurationItemOnly'
  $cmbPi.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Controls.Add($cmbPi, 1, 1)

  $chkFast = New-Object System.Windows.Forms.CheckBox
  $chkFast.Text = 'Fast mode'
  $chkFast.Checked = $true
  $chkFast.AutoSize = $true
  $root.Controls.Add($chkFast, 1, 2)

  $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $buttons.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $buttons.WrapContents = $false
  $buttons.AutoSize = $true
  $root.Controls.Add($buttons, 0, 3)
  $root.SetColumnSpan($buttons, 2)

  $btnRun = New-Object System.Windows.Forms.Button
  $btnRun.Text = 'Run'
  $btnRun.Width = 90
  $btnRun.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $buttons.Controls.Add($btnRun) | Out-Null

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = 'Cancel'
  $btnCancel.Width = 90
  $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $buttons.Controls.Add($btnCancel) | Out-Null

  $dlg.AcceptButton = $btnRun
  $dlg.CancelButton = $btnCancel
  $res = if ($Owner) { $dlg.ShowDialog($Owner) } else { $dlg.ShowDialog() }
  if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

  return @{
    ok = $true
    processingScope = ("" + $cmbScope.SelectedItem)
    piSearchMode = ("" + $cmbPi.SelectedItem)
    fastMode = [bool]$chkFast.Checked
    maxTickets = 0
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
      Register-SchumanOwnedProcess -Process $st.Proc -Tag 'export'
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
    try { if ($proc) { Unregister-SchumanOwnedProcess -ProcessId $proc.Id } } catch {}
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
      Register-SchumanOwnedProcess -Process $loading.Tag.Proc -Tag 'docsgenerate'
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
    try { if ($proc) { Unregister-SchumanOwnedProcess -ProcessId $proc.Id } } catch {}
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
    [ValidateSet('Auto','ConfigurationItemOnly','CommentsOnly','CommentsAndCI')][string]$PiSearchMode = 'Auto',
    [int]$MaxTickets = 0,
    [bool]$FastMode = $true
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
      errorMessage = 'Another export process is already running.'
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
    '-MaxTickets',[string]$MaxTickets,
    '-NoPopups'
  )
  if ($FastMode) { $args += '-SkipLegalNameFallback' }

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
    LastLogRefreshMs = -10000
    ExitCode = -1
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
      $tail = @(Get-Content -LiteralPath $st.RunLogPath -Tail 40)
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
    $elapsedMs = [int][Math]::Floor($sw.Elapsed.TotalMilliseconds)
    if (($elapsedMs - [int]$st.LastLogRefreshMs) -ge 900) {
      & $updateProgressFromLog $st
      $st.LastLogRefreshMs = $elapsedMs
    }

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
      $st.ExitCode = [int]$st.Proc.ExitCode
      $st.Ok = ($st.ExitCode -eq 0)
      if ($st.Ok) {
        $st.ProgressPct = 100
        $st.ProgressDone = [Math]::Max($st.ProgressDone, $st.ProgressTotal)
      } else {
        & $resolveRunLog $st
        & $updateProgressFromLog $st
      }
      $st.DurationSec = [int][Math]::Max(1, [Math]::Round($sw.Elapsed.TotalSeconds))
      $loading.Close()
    }
  })

  $loading.add_Shown({
    try {
      [void]$loading.Tag.Proc.Start()
      Register-SchumanOwnedProcess -Process $loading.Tag.Proc -Tag 'forceupdate'
      $animTimer.Start()
      $watchTimer.Start()
    } catch {
      $loading.Tag.Ok = $false
      $loading.Close()
    }
  })

  $ok = $false
  $durationSec = 0
  $errorMessage = ''
  try {
    [void]$loading.ShowDialog()
    if ($loading.Tag) {
      $ok = [bool]$loading.Tag.Ok
      $durationSec = [int]$loading.Tag.DurationSec
      if (-not $ok) {
        $line = ("" + $loading.Tag.LastStatusLine).Trim()
        if ($line) {
          $errorMessage = "Export failed (exit=$($loading.Tag.ExitCode)). Last log: $line"
        }
        else {
          $errorMessage = "Export failed (exit=$($loading.Tag.ExitCode))."
        }
      }
    }
  }
  finally {
    try { if ($proc) { Unregister-SchumanOwnedProcess -ProcessId $proc.Id } } catch {}
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
    errorMessage = $errorMessage
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
$form.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10)

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
$hdr.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 20, [System.Drawing.FontStyle]::Bold)
$hdr.ForeColor = [System.Drawing.Color]::FromArgb(28,28,30)
$hdr.AutoSize = $true
$hdr.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$root.Controls.Add($hdr, 1, 0)

$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Text = '⚙'
$btnSettings.Width = 38
$btnSettings.Height = 34
$btnSettings.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSettings.FlatAppearance.BorderSize = 1
$btnSettings.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$root.Controls.Add($btnSettings, 2, 0)

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
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
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
$txtExcel.Add_TextChanged(({
  $ready = Test-MainExcelReady -ExcelPathValue $txtExcel.Text -SheetNameValue $SheetName
  if ($ready) {
    Update-MainExcelDependentControls -ExcelReady $true
  } else {
    Update-MainExcelDependentControls -ExcelReady $false -Reason 'Please load Excel first.'
  }
}).GetNewClosure())

$prefsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$prefsPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$prefsPanel.AutoSize = $true
$prefsPanel.WrapContents = $true
$prefsPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$layout.Controls.Add($prefsPanel, 0, 2)

$lblThemeOpt = New-Object System.Windows.Forms.Label
$lblThemeOpt.Text = 'Theme'
$lblThemeOpt.AutoSize = $true
$lblThemeOpt.Margin = New-Object System.Windows.Forms.Padding(0, 7, 8, 0)
$prefsPanel.Controls.Add($lblThemeOpt)

$cmbTheme = New-Object System.Windows.Forms.ComboBox
$cmbTheme.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbTheme.Width = 92
[void]$cmbTheme.Items.AddRange(@('Dark', 'Light', 'Matrix', 'Hot', 'Parliament'))
$cmbTheme.Margin = New-Object System.Windows.Forms.Padding(0, 2, 14, 0)
$prefsPanel.Controls.Add($cmbTheme)

$lblAccentOpt = New-Object System.Windows.Forms.Label
$lblAccentOpt.Text = 'Accent'
$lblAccentOpt.AutoSize = $true
$lblAccentOpt.Margin = New-Object System.Windows.Forms.Padding(0, 7, 8, 0)
$prefsPanel.Controls.Add($lblAccentOpt)

$cmbAccent = New-Object System.Windows.Forms.ComboBox
$cmbAccent.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbAccent.Width = 108
[void]$cmbAccent.Items.AddRange(@('Blue', 'Teal', 'Orange', 'Green', 'Red'))
$cmbAccent.Margin = New-Object System.Windows.Forms.Padding(0, 2, 14, 0)
$prefsPanel.Controls.Add($cmbAccent)

$lblScaleOpt = New-Object System.Windows.Forms.Label
$lblScaleOpt.Text = 'Font'
$lblScaleOpt.AutoSize = $true
$lblScaleOpt.Margin = New-Object System.Windows.Forms.Padding(0, 7, 8, 0)
$prefsPanel.Controls.Add($lblScaleOpt)

$cmbScale = New-Object System.Windows.Forms.ComboBox
$cmbScale.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbScale.Width = 88
[void]$cmbScale.Items.AddRange(@('90%', '100%', '110%', '120%'))
$cmbScale.Margin = New-Object System.Windows.Forms.Padding(0, 2, 14, 0)
$prefsPanel.Controls.Add($cmbScale)

$chkCompact = New-Object System.Windows.Forms.CheckBox
$chkCompact.Text = 'Compact'
$chkCompact.AutoSize = $true
$chkCompact.Margin = New-Object System.Windows.Forms.Padding(0, 6, 14, 0)
$prefsPanel.Controls.Add($chkCompact)
$prefsPanel.Visible = $false
$prefsPanel.Height = 0
$prefsPanel.Margin = New-Object System.Windows.Forms.Padding(0)

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
$btnDashboard.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
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
$btnGenerate.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
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
$btnForce.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
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
$btnCloseCode.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
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
$btnCloseDocs.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnCloseDocs.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btnEmergency.Controls.Add($btnCloseDocs, 1, 0)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Status: Ready (SSO connected)'
$status.AutoSize = $true
$status.ForeColor = [System.Drawing.Color]::FromArgb(110,110,115)
$statusHost = New-Object System.Windows.Forms.TableLayoutPanel
$statusHost.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusHost.ColumnCount = 1
$statusHost.RowCount = 2
[void]$statusHost.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$statusHost.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.Controls.Add($statusHost, 1, 2)
$statusHost.Controls.Add($status, 0, 0)

$loadingBar = New-Object System.Windows.Forms.ProgressBar
$loadingBar.Dock = [System.Windows.Forms.DockStyle]::Top
$loadingBar.Height = 10
$loadingBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$loadingBar.MarqueeAnimationSpeed = 28
$loadingBar.Visible = $false
$loadingBar.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$statusHost.Controls.Add($loadingBar, 0, 1)

$script:MainBusyTimer = New-Object System.Windows.Forms.Timer
$script:MainBusyTimer.Interval = 240
$script:MainBusyTick = 0
$script:MainBusyBaseText = 'Loading'
$script:MainBusyActive = $false

$script:MainBusyTimer.Add_Tick({
  try {
    if (-not $script:MainBusyActive) { return }
    if (-not $status -or $status.IsDisposed) { return }
    $script:MainBusyTick = [int]$script:MainBusyTick + 1
    $dots = '.' * (($script:MainBusyTick % 3) + 1)
    $status.Text = ("Status: {0}{1}" -f $script:MainBusyBaseText, $dots)
  } catch {}
})

function global:Set-MainBusyState {
  param(
    [bool]$IsBusy,
    [string]$Text = 'Loading'
  )

  if (-not $status -or $status.IsDisposed) { return }
  $script:MainBusyActive = $IsBusy

  if ($IsBusy) {
    $script:MainBusyBaseText = if ([string]::IsNullOrWhiteSpace($Text)) { 'Loading' } else { $Text.Trim() }
    $script:MainBusyTick = 0
    if ($loadingBar -and -not $loadingBar.IsDisposed) { $loadingBar.Visible = $true }
    try { $script:MainBusyTimer.Start() } catch {}
    $status.Text = ("Status: {0}." -f $script:MainBusyBaseText)
    try {
      if ($btnDashboard) { $btnDashboard.Enabled = $false }
      if ($btnGenerate) { $btnGenerate.Enabled = $false }
      if ($btnForce) { $btnForce.Enabled = $false }
    } catch {}
  }
  else {
    try { $script:MainBusyTimer.Stop() } catch {}
    if ($loadingBar -and -not $loadingBar.IsDisposed) { $loadingBar.Visible = $false }
    $status.Text = 'Status: Ready'
    try {
      Update-MainExcelDependentControls -ExcelReady $script:MainExcelReady
    } catch {}
  }
}

$script:UiRoleByControlId = @{}
$script:UiHoverBoundByControlId = @{}
$script:CurrentMainTheme = @{}
$script:CurrentMainFontScale = 1.0
$global:CurrentMainTheme = @{}
$global:CurrentMainFontScale = 1.0

function Convert-HexToUiColor {
  param(
    [string]$Hex,
    [System.Drawing.Color]$Fallback = [System.Drawing.Color]::FromArgb(0, 122, 255)
  )
  try {
    $h = ("" + $Hex).Trim()
    if (-not $h) { return $Fallback }
    if ($h -notmatch '^#') { $h = "#$h" }
    if ($h -notmatch '^#([0-9a-fA-F]{6})$') { return $Fallback }
    return [System.Drawing.ColorTranslator]::FromHtml($h)
  } catch {
    return $Fallback
  }
}

function Get-ShiftedUiColor {
  param([System.Drawing.Color]$Color,[int]$Delta = 0)
  $r = [Math]::Min(255, [Math]::Max(0, $Color.R + $Delta))
  $g = [Math]::Min(255, [Math]::Max(0, $Color.G + $Delta))
  $b = [Math]::Min(255, [Math]::Max(0, $Color.B + $Delta))
  return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

function Get-ReadableTextColor {
  param([System.Drawing.Color]$Background)
  if (-not $Background) { return [System.Drawing.Color]::FromArgb(245, 248, 252) }
  $luma = (0.299 * $Background.R) + (0.587 * $Background.G) + (0.114 * $Background.B)
  if ($luma -ge 150) { return [System.Drawing.Color]::FromArgb(18, 24, 38) }
  return [System.Drawing.Color]::FromArgb(245, 248, 252)
}

function New-UiThemeCatalog {
  $themes = [ordered]@{}
  $themes['Dark'] = @{ FormBG=Convert-HexToUiColor '#111827'; HeaderBG=Convert-HexToUiColor '#0F172A'; CardBG=Convert-HexToUiColor '#1E293B'; Border=Convert-HexToUiColor '#334155'; Text=Convert-HexToUiColor '#E2E8F0'; MutedText=Convert-HexToUiColor '#94A3B8'; Primary=Convert-HexToUiColor '#2563EB'; PrimaryHover=Convert-HexToUiColor '#3B82F6'; Secondary=Convert-HexToUiColor '#334155'; SecondaryHover=Convert-HexToUiColor '#475569'; Accent=Convert-HexToUiColor '#38BDF8'; AccentHover=Convert-HexToUiColor '#0EA5E9'; Danger=Convert-HexToUiColor '#7F1D1D'; DangerHover=Convert-HexToUiColor '#991B1B'; InputBG=Convert-HexToUiColor '#0F172A'; InputBorder=Convert-HexToUiColor '#334155'; FocusBorder=Convert-HexToUiColor '#38BDF8'; GridAltRow=Convert-HexToUiColor '#18243A' }
  $themes['Light'] = @{ FormBG=Convert-HexToUiColor '#F3F6FB'; HeaderBG=Convert-HexToUiColor '#E7EEF9'; CardBG=Convert-HexToUiColor '#FAFCFF'; Border=Convert-HexToUiColor '#CBD5E1'; Text=Convert-HexToUiColor '#0F172A'; MutedText=Convert-HexToUiColor '#475569'; Primary=Convert-HexToUiColor '#003399'; PrimaryHover=Convert-HexToUiColor '#1D4ED8'; Secondary=Convert-HexToUiColor '#E2E8F0'; SecondaryHover=Convert-HexToUiColor '#CBD5E1'; Accent=Convert-HexToUiColor '#FFCC00'; AccentHover=Convert-HexToUiColor '#EAB308'; Danger=Convert-HexToUiColor '#B91C1C'; DangerHover=Convert-HexToUiColor '#991B1B'; InputBG=Convert-HexToUiColor '#FFFFFF'; InputBorder=Convert-HexToUiColor '#CBD5E1'; FocusBorder=Convert-HexToUiColor '#003399'; GridAltRow=Convert-HexToUiColor '#F1F5F9' }
  $themes['Matrix'] = @{ FormBG=Convert-HexToUiColor '#06110C'; HeaderBG=Convert-HexToUiColor '#081810'; CardBG=Convert-HexToUiColor '#0D1F14'; Border=Convert-HexToUiColor '#185C28'; Text=Convert-HexToUiColor '#B0FFC8'; MutedText=Convert-HexToUiColor '#5ECB7E'; Primary=Convert-HexToUiColor '#22C55E'; PrimaryHover=Convert-HexToUiColor '#4ADE80'; Secondary=Convert-HexToUiColor '#163622'; SecondaryHover=Convert-HexToUiColor '#1E4A2D'; Accent=Convert-HexToUiColor '#30FF80'; AccentHover=Convert-HexToUiColor '#22C55E'; Danger=Convert-HexToUiColor '#5E1D1D'; DangerHover=Convert-HexToUiColor '#7F1D1D'; InputBG=Convert-HexToUiColor '#0A1C12'; InputBorder=Convert-HexToUiColor '#1F7A3A'; FocusBorder=Convert-HexToUiColor '#30FF80'; GridAltRow=Convert-HexToUiColor '#0A1911' }
  $themes['Hot'] = @{ FormBG=Convert-HexToUiColor '#180A10'; HeaderBG=Convert-HexToUiColor '#240C16'; CardBG=Convert-HexToUiColor '#2D0E1A'; Border=Convert-HexToUiColor '#5C1E2E'; Text=Convert-HexToUiColor '#FFE4E6'; MutedText=Convert-HexToUiColor '#F8B4BD'; Primary=Convert-HexToUiColor '#BE185D'; PrimaryHover=Convert-HexToUiColor '#DB2777'; Secondary=Convert-HexToUiColor '#581C2D'; SecondaryHover=Convert-HexToUiColor '#7F1D1D'; Accent=Convert-HexToUiColor '#F43F5E'; AccentHover=Convert-HexToUiColor '#E11D48'; Danger=Convert-HexToUiColor '#7F1D1D'; DangerHover=Convert-HexToUiColor '#991B1B'; InputBG=Convert-HexToUiColor '#260B18'; InputBorder=Convert-HexToUiColor '#7F1D3A'; FocusBorder=Convert-HexToUiColor '#F43F5E'; GridAltRow=Convert-HexToUiColor '#210B15' }
  $themes['Parliament'] = @{ FormBG=Convert-HexToUiColor '#0F172A'; HeaderBG=Convert-HexToUiColor '#003399'; CardBG=Convert-HexToUiColor '#1E293B'; Border=Convert-HexToUiColor '#334155'; Text=Convert-HexToUiColor '#E2E8F0'; MutedText=Convert-HexToUiColor '#94A3B8'; Primary=Convert-HexToUiColor '#003399'; PrimaryHover=Convert-HexToUiColor '#1D4ED8'; Secondary=Convert-HexToUiColor '#1E40AF'; SecondaryHover=Convert-HexToUiColor '#1D4ED8'; Accent=Convert-HexToUiColor '#FFCC00'; AccentHover=Convert-HexToUiColor '#EAB308'; Danger=Convert-HexToUiColor '#7F1D1D'; DangerHover=Convert-HexToUiColor '#991B1B'; InputBG=Convert-HexToUiColor '#111C33'; InputBorder=Convert-HexToUiColor '#334155'; FocusBorder=Convert-HexToUiColor '#FFCC00'; GridAltRow=Convert-HexToUiColor '#1A2436' }
  return $themes
}

function Resolve-AccentColor {
  param([string]$AccentValue,[System.Drawing.Color]$Fallback)
  $key = ("" + $AccentValue).Trim()
  switch ($key.ToLowerInvariant()) {
    'blue' { return Convert-HexToUiColor '#3B82F6' }
    'teal' { return Convert-HexToUiColor '#14B8A6' }
    'orange' { return Convert-HexToUiColor '#F59E0B' }
    'green' { return Convert-HexToUiColor '#22C55E' }
    'red' { return Convert-HexToUiColor '#EF4444' }
    default {
      if ($key -match '^#?[0-9a-fA-F]{6}$') { return Convert-HexToUiColor $key -Fallback $Fallback }
      return $Fallback
    }
  }
}

function Set-UiControlRole {
  param([System.Windows.Forms.Control]$Control,[string]$Role)
  if (-not $Control) { return }
  $script:UiRoleByControlId[[string]$Control.GetHashCode()] = ("" + $Role)
}

function Get-UiControlRole {
  param([System.Windows.Forms.Control]$Control)
  if (-not $Control) { return '' }
  $id = [string]$Control.GetHashCode()
  if ($script:UiRoleByControlId.ContainsKey($id)) { return ("" + $script:UiRoleByControlId[$id]) }
  return ''
}

function Update-UiButtonVisual {
  param([System.Windows.Forms.Button]$Button,[hashtable]$Theme,[bool]$Hover = $false)
  if (-not $Button -or -not $Theme -or -not $Theme.ContainsKey('Secondary')) { return }
  $roleHandler = ${function:Get-UiControlRole}
  $role = if ($roleHandler) { & $roleHandler -Control $Button } else { '' }
  $bg = $Theme.Secondary
  $bgHover = $Theme.SecondaryHover
  switch ($role) {
    'PrimaryButton' { $bg = $Theme.Primary; $bgHover = $Theme.PrimaryHover }
    'DangerButton' { $bg = $Theme.Danger; $bgHover = $Theme.DangerHover }
    'AccentButton' { $bg = $Theme.Accent; $bgHover = $Theme.AccentHover }
  }
  $fill = if ($Hover) { $bgHover } else { $bg }
  $Button.BackColor = $fill
  $Button.ForeColor = Get-ReadableTextColor -Background $fill
  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 1
  $Button.FlatAppearance.BorderColor = if ($Hover) { $Theme.FocusBorder } else { $Theme.Border }
  $Button.FlatAppearance.MouseOverBackColor = $bgHover
  $Button.FlatAppearance.MouseDownBackColor = Get-ShiftedUiColor -Color $bgHover -Delta -14
}

function Ensure-UiButtonHoverBinding {
  param([System.Windows.Forms.Button]$Button)
  if (-not $Button) { return }
  $id = [string]$Button.GetHashCode()
  if ($script:UiHoverBoundByControlId.ContainsKey($id)) { return }
  $updateHandler = ${function:Update-UiButtonVisual}
  $Button.Add_MouseEnter(({ param($sender,$eventArgs) try { if ($updateHandler) { & $updateHandler -Button $sender -Theme $script:CurrentMainTheme -Hover:$true } } catch {} }).GetNewClosure())
  $Button.Add_MouseLeave(({ param($sender,$eventArgs) try { if ($updateHandler) { & $updateHandler -Button $sender -Theme $script:CurrentMainTheme -Hover:$false } } catch {} }).GetNewClosure())
  $script:UiHoverBoundByControlId[$id] = $true
}

function Apply-ThemeToControlTree {
  param([System.Windows.Forms.Control]$RootControl,[hashtable]$Theme,[double]$FontScale = 1.0)
  if (-not $RootControl -or -not $Theme) { return }
  $fontName = Get-UiFontNameSafe
  $regular = New-Object System.Drawing.Font($fontName, [single](11 * $FontScale), [System.Drawing.FontStyle]::Regular)
  $bold = New-Object System.Drawing.Font($fontName, [single](11 * $FontScale), [System.Drawing.FontStyle]::Bold)
  $title = New-Object System.Drawing.Font($fontName, [single](18 * $FontScale), [System.Drawing.FontStyle]::Bold)

  $walk = $null
  $walk = {
    param([System.Windows.Forms.Control]$Ctrl)
    if (-not $Ctrl) { return }
    $role = Get-UiControlRole -Control $Ctrl
    if ($Ctrl -is [System.Windows.Forms.Form]) { $Ctrl.BackColor = $Theme.FormBG; $Ctrl.ForeColor = $Theme.Text; $Ctrl.Font = $regular }
    elseif ($Ctrl -is [System.Windows.Forms.Panel]) {
      switch ($role) {
        'CardBorder' { $Ctrl.BackColor = $Theme.Border }
        'CardSurface' { $Ctrl.BackColor = $Theme.CardBG }
        default { $Ctrl.BackColor = if ($Ctrl.Parent) { $Ctrl.Parent.BackColor } else { $Theme.FormBG } }
      }
    }
    elseif ($Ctrl -is [System.Windows.Forms.TableLayoutPanel] -or $Ctrl -is [System.Windows.Forms.FlowLayoutPanel]) { $Ctrl.BackColor = if ($Ctrl.Parent) { $Ctrl.Parent.BackColor } else { $Theme.FormBG } }
    elseif ($Ctrl -is [System.Windows.Forms.StatusStrip] -or $Ctrl -is [System.Windows.Forms.ToolStrip]) {
      $Ctrl.BackColor = $Theme.HeaderBG
      $Ctrl.ForeColor = $Theme.Text
      $Ctrl.Font = $regular
      try {
        foreach ($item in $Ctrl.Items) {
          $item.ForeColor = $Theme.Text
          $item.BackColor = $Theme.HeaderBG
        }
      } catch {}
    }
    elseif ($Ctrl -is [System.Windows.Forms.ProgressBar]) {
      $Ctrl.BackColor = $Theme.CardBG
      $Ctrl.ForeColor = $Theme.Primary
    }
    elseif ($Ctrl -is [System.Windows.Forms.Label]) { $Ctrl.ForeColor = if ($role -eq 'MutedLabel' -or $role -eq 'StatusLabel') { $Theme.MutedText } else { $Theme.Text }; $Ctrl.Font = if ($role -eq 'HeaderTitle') { $title } else { $regular } }
    elseif ($Ctrl -is [System.Windows.Forms.TextBox] -or $Ctrl -is [System.Windows.Forms.RichTextBox] -or $Ctrl -is [System.Windows.Forms.ComboBox]) { $Ctrl.BackColor = $Theme.InputBG; $Ctrl.ForeColor = $Theme.Text; $Ctrl.Font = $regular }
    elseif ($Ctrl -is [System.Windows.Forms.DataGridView]) {
      $Ctrl.BackgroundColor = $Theme.CardBG; $Ctrl.GridColor = $Theme.Border; $Ctrl.BorderStyle = [System.Windows.Forms.BorderStyle]::None
      $Ctrl.DefaultCellStyle.BackColor = $Theme.CardBG; $Ctrl.DefaultCellStyle.ForeColor = $Theme.Text
      $Ctrl.DefaultCellStyle.SelectionBackColor = $Theme.Primary; $Ctrl.DefaultCellStyle.SelectionForeColor = Get-ReadableTextColor -Background $Theme.Primary
      $Ctrl.AlternatingRowsDefaultCellStyle.BackColor = $Theme.GridAltRow; $Ctrl.AlternatingRowsDefaultCellStyle.ForeColor = $Theme.Text
      $Ctrl.ColumnHeadersDefaultCellStyle.BackColor = $Theme.HeaderBG; $Ctrl.ColumnHeadersDefaultCellStyle.ForeColor = $Theme.Text
      $Ctrl.EnableHeadersVisualStyles = $false
    }
    elseif ($Ctrl -is [System.Windows.Forms.Button]) {
      $Ctrl.Font = $bold
      Ensure-UiButtonHoverBinding -Button $Ctrl
      Update-UiButtonVisual -Button $Ctrl -Theme $Theme -Hover:$false
      if (Get-Command -Name Set-UiRoundedButton -ErrorAction SilentlyContinue) {
        Set-UiRoundedButton -Button $Ctrl -Radius 10
      }
    }
    foreach ($child in $Ctrl.Controls) { & $walk $child }
  }
  & $walk $RootControl
  if (Get-Command -Name Apply-UiRoundedButtonsRecursive -ErrorAction SilentlyContinue) {
    Apply-UiRoundedButtonsRecursive -Root $RootControl -Radius 10
  }
}

$script:SetUiControlRoleHandler = ${function:Set-UiControlRole}
$script:EnsureUiButtonHoverBindingHandler = ${function:Ensure-UiButtonHoverBinding}
$script:UpdateUiButtonVisualHandler = ${function:Update-UiButtonVisual}
$script:ApplyThemeToControlTreeHandler = ${function:Apply-ThemeToControlTree}
foreach ($handler in @(
    @{ Name = 'Set-UiControlRole'; Value = $script:SetUiControlRoleHandler },
    @{ Name = 'Ensure-UiButtonHoverBinding'; Value = $script:EnsureUiButtonHoverBindingHandler },
    @{ Name = 'Update-UiButtonVisual'; Value = $script:UpdateUiButtonVisualHandler },
    @{ Name = 'Apply-ThemeToControlTree'; Value = $script:ApplyThemeToControlTreeHandler }
  )) {
  if (-not $handler.Value) {
    throw ("Main UI runtime is incomplete. Missing private handler: {0}" -f $handler.Name)
  }
}

$script:Themes = New-UiThemeCatalog
if ($script:SetUiControlRoleHandler) {
  & $script:SetUiControlRoleHandler -Control $btnDashboard -Role 'PrimaryButton'
  & $script:SetUiControlRoleHandler -Control $btnGenerate -Role 'SecondaryButton'
  & $script:SetUiControlRoleHandler -Control $btnForce -Role 'AccentButton'
  & $script:SetUiControlRoleHandler -Control $btnCloseCode -Role 'DangerButton'
  & $script:SetUiControlRoleHandler -Control $btnCloseDocs -Role 'DangerButton'
  & $script:SetUiControlRoleHandler -Control $btnSettings -Role 'SecondaryButton'
  & $script:SetUiControlRoleHandler -Control $card.Border -Role 'CardBorder'
  & $script:SetUiControlRoleHandler -Control $card.Content.Parent -Role 'CardSurface'
  & $script:SetUiControlRoleHandler -Control $hdr -Role 'HeaderTitle'
  & $script:SetUiControlRoleHandler -Control $status -Role 'StatusLabel'
  & $script:SetUiControlRoleHandler -Control $lbl -Role 'MutedLabel'
}

function Apply-MainUiTheme {
  param(
    [string]$ThemeName,
    [string]$AccentName,
    [int]$FontScale = 100,
    [bool]$Compact = $false
  )

  $safeTheme = ("" + $ThemeName).Trim()
  if (-not $script:Themes.Contains($safeTheme)) { $safeTheme = 'Dark' }
  $safeAccent = ("" + $AccentName).Trim()
  if (-not $safeAccent) { $safeAccent = 'Blue' }
  $safeScale = [Math]::Min(130, [Math]::Max(90, $FontScale))

  $baseTheme = $script:Themes[$safeTheme]
  if (-not $baseTheme -or -not ($baseTheme -is [hashtable])) {
    $safeTheme = 'Dark'
    $baseTheme = $script:Themes['Dark']
  }
  $resolved = @{}
  foreach ($k in $baseTheme.Keys) { $resolved[$k] = $baseTheme[$k] }

  $accent = Resolve-AccentColor -AccentValue $safeAccent -Fallback $resolved.Accent
  $resolved.Primary = $accent
  $resolved.PrimaryHover = Get-ShiftedUiColor -Color $accent -Delta 18
  $resolved.Accent = $accent
  $resolved.AccentHover = Get-ShiftedUiColor -Color $accent -Delta 10
  $resolved.FocusBorder = $accent
  $script:CurrentMainTheme = $resolved
  $script:CurrentMainFontScale = ($safeScale / 100.0)
  $global:CurrentMainTheme = $resolved
  $global:CurrentMainFontScale = ($safeScale / 100.0)

  if ($script:ApplyThemeToControlTreeHandler) {
    if ($form -and -not $form.IsDisposed) {
      & $script:ApplyThemeToControlTreeHandler -RootControl $form -Theme $resolved -FontScale ($safeScale / 100.0)
      try { $form.Invalidate() } catch {}
      try { $form.Refresh() } catch {}
    }
    $formsToRefresh = @()
    try { $formsToRefresh = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ }) } catch {}
    try { Write-Log -Level INFO -Message ("Theme apply '{0}/{1}' on {2} open form(s)." -f $safeTheme, $safeAccent, $formsToRefresh.Count) } catch {}
    foreach ($openForm in $formsToRefresh) {
      if (-not $openForm -or $openForm.IsDisposed) { continue }
      if ($form -and ($openForm -eq $form)) { continue }
      try {
        & $script:ApplyThemeToControlTreeHandler -RootControl $openForm -Theme $resolved -FontScale ($safeScale / 100.0)
        $openForm.Invalidate()
        $openForm.Refresh()
        try { Write-Log -Level INFO -Message ("Theme applied to window: {0}" -f ("" + $openForm.Text)) } catch {}
      }
      catch {
        try { Write-Log -Level WARN -Message ("Theme apply skipped for a window: " + $_.Exception.Message) } catch {}
      }
    }
  }

  $moduleHeight = if ($Compact) { 38 } else { 46 }
  $dangerHeight = if ($Compact) { 32 } else { 36 }
  $forceHeight = if ($Compact) { 34 } else { 40 }
  if ($btnDashboard -and -not $btnDashboard.IsDisposed) { $btnDashboard.Height = $moduleHeight }
  if ($btnGenerate -and -not $btnGenerate.IsDisposed) { $btnGenerate.Height = $moduleHeight }
  if ($btnForce -and -not $btnForce.IsDisposed) { $btnForce.Height = $forceHeight }
  if ($btnCloseCode -and -not $btnCloseCode.IsDisposed) { $btnCloseCode.Height = $dangerHeight }
  if ($btnCloseDocs -and -not $btnCloseDocs.IsDisposed) { $btnCloseDocs.Height = $dangerHeight }

  $mainUiPrefs = @{
    theme = $safeTheme
    accent = $safeAccent
    fontScale = $safeScale
    compact = [bool]$Compact
  }
  Save-MainUiPrefs -Config $globalConfig -Prefs $mainUiPrefs
}

$mainUiPrefs = Get-MainUiPrefs -Config $globalConfig
$themePref = ("" + $mainUiPrefs.theme).Trim()
if (-not $script:Themes.Contains($themePref)) { $themePref = 'Dark' }
$accentPref = ("" + $mainUiPrefs.accent).Trim()
if (-not $accentPref) { $accentPref = 'Blue' }
$scalePref = [Math]::Min(130, [Math]::Max(90, [int]$mainUiPrefs.fontScale))
$compactPref = [bool]$mainUiPrefs.compact

$cmbTheme.SelectedItem = $themePref
if (-not $cmbTheme.SelectedItem) { $cmbTheme.SelectedItem = 'Dark' }
$cmbAccent.SelectedItem = $accentPref
if (-not $cmbAccent.SelectedItem) { $cmbAccent.SelectedItem = 'Blue' }
$scaleText = ('{0}%' -f $scalePref)
if ($cmbScale.Items -contains $scaleText) { $cmbScale.SelectedItem = $scaleText } else { $cmbScale.SelectedItem = '100%' }
$chkCompact.Checked = $compactPref

try { Apply-MainUiTheme -ThemeName $themePref -AccentName $accentPref -FontScale $scalePref -Compact:$compactPref }
catch { Show-UiError -Title 'Theme' -Message 'Could not apply initial theme.' -Exception $_.Exception }

$script:ApplyMainUiThemeHandler = ${function:Apply-MainUiTheme}
$script:MainExcelReady = $false

function Test-MainExcelReady {
  param([string]$ExcelPathValue,[string]$SheetNameValue)
  $path = ("" + $ExcelPathValue).Trim()
  if ([string]::IsNullOrWhiteSpace($path)) { return $false }
  if (-not (Test-Path -LiteralPath $path)) { return $false }
  try {
    $rows = @(Search-DashboardRows -ExcelPath $path -SheetName $SheetNameValue -SearchText '')
    return ($rows.Count -gt 0)
  } catch {
    return $false
  }
}

function Update-MainExcelDependentControls {
  param([bool]$ExcelReady,[string]$Reason = '')
  $script:MainExcelReady = [bool]$ExcelReady
  $btnDashboard.Enabled = [bool]$ExcelReady
  $btnGenerate.Enabled = [bool]$ExcelReady
  $btnForce.Enabled = [bool]$ExcelReady
  if (-not $ExcelReady -and $status -and -not $status.IsDisposed) {
    $msg = ("" + $Reason).Trim()
    if (-not $msg) { $msg = 'Please load Excel first.' }
    $status.Text = ("Status: {0}" -f $msg)
  }
  elseif ($ExcelReady -and $status -and -not $status.IsDisposed -and -not $script:MainBusyActive) {
    $status.Text = 'Status: Ready'
  }
}

$showSettingsDialog = ({
  if (-not $form -or $form.IsDisposed) {
    Show-UiError -Title 'Theme' -Message 'Owner form is not available.'
    return
  }
  if (-not $script:SetUiControlRoleHandler) {
    $cmd = Get-Command -Name Set-UiControlRole -CommandType Function -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.ScriptBlock) { $script:SetUiControlRoleHandler = $cmd.ScriptBlock }
  }
  if (-not $script:ApplyThemeToControlTreeHandler) {
    $cmd = Get-Command -Name Apply-ThemeToControlTree -CommandType Function -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.ScriptBlock) { $script:ApplyThemeToControlTreeHandler = $cmd.ScriptBlock }
  }
  if (-not $script:ApplyMainUiThemeHandler) {
    $cmd = Get-Command -Name Apply-MainUiTheme -CommandType Function -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.ScriptBlock) { $script:ApplyMainUiThemeHandler = $cmd.ScriptBlock }
  }

  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Personalization'
  $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.MinimumSize = New-Object System.Drawing.Size(460, 320)
  $dlg.MaximumSize = New-Object System.Drawing.Size(520, 360)
  $dlg.Size = New-Object System.Drawing.Size(470, 330)
  $dlg.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
  $dlg.BackColor = if ($form) { $form.BackColor } else { [System.Drawing.Color]::FromArgb(30,30,32) }
  $dlg.ForeColor = if ($hdr) { $hdr.ForeColor } else { [System.Drawing.Color]::FromArgb(230,230,230) }
  $dialogFontName = 'Segoe UI'
  if ($script:GetUiFontNameSafeHandler) {
    try {
      $candidateFont = ("" + (& $script:GetUiFontNameSafeHandler)).Trim()
      if ($candidateFont) { $dialogFontName = $candidateFont }
    } catch {}
  }
  $dlg.Font = New-Object System.Drawing.Font($dialogFontName, 10, [System.Drawing.FontStyle]::Regular)

  $root = New-Object System.Windows.Forms.TableLayoutPanel
  $root.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.ColumnCount = 1
  $root.RowCount = 2
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 60)))
  $dlg.Controls.Add($root)

  $gridPrefs = New-Object System.Windows.Forms.TableLayoutPanel
  $gridPrefs.Dock = [System.Windows.Forms.DockStyle]::Fill
  $gridPrefs.Padding = New-Object System.Windows.Forms.Padding(16)
  $gridPrefs.ColumnCount = 2
  $gridPrefs.RowCount = 6
  [void]$gridPrefs.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38)))
  [void]$gridPrefs.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 62)))
  [void]$root.Controls.Add($gridPrefs, 0, 0)

  $newRowLabel = {
    param([string]$text, [int]$row)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(0, 8, 8, 8)
    $gridPrefs.Controls.Add($l, 0, $row)
    return $l
  }

  $null = & $newRowLabel 'Theme' 0
  $dlgTheme = New-Object System.Windows.Forms.ComboBox
  $dlgTheme.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$dlgTheme.Items.AddRange(@('Dark', 'Light', 'Matrix', 'Hot', 'Parliament'))
  $dlgTheme.SelectedItem = if ($cmbTheme) { $cmbTheme.SelectedItem } else { 'Dark' }
  $dlgTheme.Dock = [System.Windows.Forms.DockStyle]::Top
  $gridPrefs.Controls.Add($dlgTheme, 1, 0)

  $null = & $newRowLabel 'Accent' 1
  $dlgAccent = New-Object System.Windows.Forms.ComboBox
  $dlgAccent.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$dlgAccent.Items.AddRange(@('Blue', 'Teal', 'Orange', 'Green', 'Red'))
  $dlgAccent.SelectedItem = if ($cmbAccent) { $cmbAccent.SelectedItem } else { 'Blue' }
  $dlgAccent.Dock = [System.Windows.Forms.DockStyle]::Top
  $gridPrefs.Controls.Add($dlgAccent, 1, 1)

  $null = & $newRowLabel 'Font scale' 2
  $dlgScale = New-Object System.Windows.Forms.ComboBox
  $dlgScale.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$dlgScale.Items.AddRange(@('90%', '100%', '110%', '120%'))
  $dlgScale.SelectedItem = if ($cmbScale) { $cmbScale.SelectedItem } else { '100%' }
  $dlgScale.Dock = [System.Windows.Forms.DockStyle]::Top
  $gridPrefs.Controls.Add($dlgScale, 1, 2)

  $null = & $newRowLabel 'Compact mode' 3
  $dlgCompact = New-Object System.Windows.Forms.CheckBox
  $dlgCompact.Checked = if ($chkCompact) { [bool]$chkCompact.Checked } else { $false }
  $dlgCompact.AutoSize = $true
  $gridPrefs.Controls.Add($dlgCompact, 1, 3)

  $note = New-Object System.Windows.Forms.Label
  $note.Text = 'Parliament=EP blue+gold | Matrix=green/black | Hot=burgundy/red'
  $note.AutoSize = $true
  $note.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
  $gridPrefs.Controls.Add($note, 0, 4)
  $gridPrefs.SetColumnSpan($note, 2)

  $miniStatus = New-Object System.Windows.Forms.Label
  $miniStatus.Text = ''
  $miniStatus.AutoSize = $true
  $miniStatus.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
  $gridPrefs.Controls.Add($miniStatus, 0, 5)
  $gridPrefs.SetColumnSpan($miniStatus, 2)

  $buttonBar = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonBar.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $buttonBar.Dock = [System.Windows.Forms.DockStyle]::Fill
  $buttonBar.Padding = New-Object System.Windows.Forms.Padding(12)
  $buttonBar.WrapContents = $false
  [void]$root.Controls.Add($buttonBar, 0, 1)

  $btnApply = New-Object System.Windows.Forms.Button
  $btnApply.Text = 'Apply'
  $btnApply.Size = New-Object System.Drawing.Size(110, 34)
  $btnApply.Visible = $true
  $buttonBar.Controls.Add($btnApply) | Out-Null

  $btnClose = New-Object System.Windows.Forms.Button
  $btnClose.Text = 'Close'
  $btnClose.Size = New-Object System.Drawing.Size(110, 34)
  $btnClose.Visible = $true
  $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $buttonBar.Controls.Add($btnClose) | Out-Null

  $btnApply.BringToFront()
  $btnClose.BringToFront()
  $dlg.AcceptButton = $btnApply
  $dlg.CancelButton = $btnClose

  if ($script:CurrentMainTheme -and $script:CurrentMainTheme.Count -gt 0) {
    if ($script:SetUiControlRoleHandler) {
      & $script:SetUiControlRoleHandler -Control $btnApply -Role 'PrimaryButton'
      & $script:SetUiControlRoleHandler -Control $btnClose -Role 'SecondaryButton'
    }
    if ($script:EnsureUiButtonHoverBindingHandler) {
      & $script:EnsureUiButtonHoverBindingHandler -Button $btnApply
      & $script:EnsureUiButtonHoverBindingHandler -Button $btnClose
    }
    if ($script:UpdateUiButtonVisualHandler) {
      & $script:UpdateUiButtonVisualHandler -Button $btnApply -Theme $script:CurrentMainTheme -Hover:$false
      & $script:UpdateUiButtonVisualHandler -Button $btnClose -Theme $script:CurrentMainTheme -Hover:$false
    }
    if ($script:ApplyThemeToControlTreeHandler) {
      & $script:ApplyThemeToControlTreeHandler -RootControl $dlg -Theme $script:CurrentMainTheme -FontScale 1.0
    }
  }

  $previewThemeAction = {
      if (-not $form -or $form.IsDisposed) { return }
      if (-not $script:Themes -or $script:Themes.Count -eq 0) { return }
      if (-not $dlgTheme -or -not $dlgAccent -or -not $dlgScale -or -not $dlgCompact) { return }
      $selectedTheme = ("" + $dlgTheme.SelectedItem).Trim()
      if (-not $selectedTheme -or -not $script:Themes.Contains($selectedTheme)) { $selectedTheme = 'Dark' }
      $selectedAccent = ("" + $dlgAccent.SelectedItem).Trim()
      if (-not $selectedAccent) { $selectedAccent = 'Blue' }
      $selectedScale = 100
      try { $selectedScale = [int]((("" + $dlgScale.SelectedItem).Trim()).TrimEnd('%')) } catch { $selectedScale = 100 }
      $selectedScale = [Math]::Min(130, [Math]::Max(90, $selectedScale))
      $selectedCompact = [bool]$dlgCompact.Checked

      if ($cmbTheme) { $cmbTheme.SelectedItem = $selectedTheme }
      if ($cmbAccent) { $cmbAccent.SelectedItem = $selectedAccent }
      if ($cmbScale) { $cmbScale.SelectedItem = ("{0}%" -f $selectedScale) }
      if ($chkCompact) { $chkCompact.Checked = $selectedCompact }

      try {
        Write-Log -Level INFO -Message ("Theme preview apply: theme={0}, accent={1}, scale={2}, compact={3}" -f $selectedTheme, $selectedAccent, $selectedScale, $selectedCompact)
      } catch {}
      if (-not $script:ApplyMainUiThemeHandler) {
        throw 'Apply-MainUiTheme handler is not available.'
      }
      & $script:ApplyMainUiThemeHandler -ThemeName $selectedTheme -AccentName $selectedAccent -FontScale $selectedScale -Compact:$selectedCompact
      if ($script:CurrentMainTheme -and $script:CurrentMainTheme.Count -gt 0 -and $dlg -and -not $dlg.IsDisposed -and $script:ApplyThemeToControlTreeHandler) {
        & $script:ApplyThemeToControlTreeHandler -RootControl $dlg -Theme $script:CurrentMainTheme -FontScale ($selectedScale / 100.0)
        try { $dlg.Invalidate() } catch {}
        try { $dlg.Refresh() } catch {}
      }
      if ($form -and -not $form.IsDisposed) {
        try { $form.Invalidate() } catch {}
        try { $form.Refresh() } catch {}
      }
      if ($status) { $status.Text = "Status: Theme applied ($selectedTheme/$selectedAccent)" }
      $miniStatus.Text = 'Applied'
  }.GetNewClosure()

  $applyPersonalizationAction = {
      if (-not $form -or $form.IsDisposed) { return }
      if (-not $script:Themes -or $script:Themes.Count -eq 0) {
        Show-UiError -Title 'Theme' -Message 'Theme catalog is not initialized.'
        return
      }
      if (-not $dlgTheme -or -not $dlgAccent -or -not $dlgScale -or -not $dlgCompact) {
        Show-UiError -Title 'Theme' -Message 'Personalization controls are not initialized.'
        return
      }

      $selectedTheme = ("" + $dlgTheme.SelectedItem).Trim()
      if (-not $selectedTheme -or -not $script:Themes.Contains($selectedTheme)) { $selectedTheme = 'Dark' }
      $selectedAccent = ("" + $dlgAccent.SelectedItem).Trim()
      if (-not $selectedAccent) { $selectedAccent = 'Blue' }
      $selectedScale = 100
      try { $selectedScale = [int]((("" + $dlgScale.SelectedItem).Trim()).TrimEnd('%')) } catch { $selectedScale = 100 }
      $selectedScale = [Math]::Min(130, [Math]::Max(90, $selectedScale))
      $selectedCompact = [bool]$dlgCompact.Checked

      if ($cmbTheme) { $cmbTheme.SelectedItem = $selectedTheme }
      if ($cmbAccent) { $cmbAccent.SelectedItem = $selectedAccent }
      if ($cmbScale) { $cmbScale.SelectedItem = ("{0}%" -f $selectedScale) }
      if ($chkCompact) { $chkCompact.Checked = $selectedCompact }

      & $previewThemeAction
  }.GetNewClosure()

  $dlgTheme.Add_SelectedIndexChanged(({
    Invoke-SafeUiAction -ActionName 'Preview Theme' -Owner $dlg -Action $previewThemeAction | Out-Null
  }).GetNewClosure())
  $dlgAccent.Add_SelectedIndexChanged(({
    Invoke-SafeUiAction -ActionName 'Preview Accent' -Owner $dlg -Action $previewThemeAction | Out-Null
  }).GetNewClosure())
  $dlgScale.Add_SelectedIndexChanged(({
    Invoke-SafeUiAction -ActionName 'Preview Font Scale' -Owner $dlg -Action $previewThemeAction | Out-Null
  }).GetNewClosure())
  $dlgCompact.Add_CheckedChanged(({
    Invoke-SafeUiAction -ActionName 'Preview Compact Mode' -Owner $dlg -Action $previewThemeAction | Out-Null
  }).GetNewClosure())

  $btnApply.Add_Click(({
    Invoke-SafeUiAction -ActionName 'Apply Personalization' -Owner $dlg -Action $applyPersonalizationAction | Out-Null
  }).GetNewClosure())

  [void]$dlg.ShowDialog($form)
}).GetNewClosure()

$openModule = {
  param([string]$module)

  if (-not $script:MainExcelReady) {
    [System.Windows.Forms.MessageBox]::Show('Please load Excel first.', 'Validation') | Out-Null
    return
  }
  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    [System.Windows.Forms.MessageBox]::Show('Please load Excel first.','Validation') | Out-Null
    Update-MainExcelDependentControls -ExcelReady $false -Reason 'Please load Excel first.'
    return
  }

  Set-MainBusyState -IsBusy $true -Text ("Opening {0}" -f $module)
  try {
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
  }
  finally {
    Set-MainBusyState -IsBusy $false
  }
}

$openDashboardAction = { & $openModule 'Dashboard' }.GetNewClosure()
$openGenerateAction = { & $openModule 'Generate' }.GetNewClosure()
$invokeForceExcelUpdateHandler = ${function:Invoke-ForceExcelUpdate}
$showForceUpdateOptionsDialogHandler = ${function:Show-ForceUpdateOptionsDialog}
$newFallbackForceUpdateOptionsDialogHandler = ${function:New-FallbackForceUpdateOptionsDialog}
$closeCodeAction = {
  $cmd = Get-Command -Name Invoke-UiEmergencyClose -CommandType Function -ErrorAction SilentlyContinue
  if (-not $cmd) {
    Show-UiError -Title 'Schuman' -Message 'Close helper not available.'
    return
  }
  try {
    $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar codigo' -ExecutableNames @('excel.exe', 'powershell.exe', 'pwsh.exe') -Owner $form -Mode 'All' -MainForm $form -BaseDir $projectRoot
    if ($r -and -not $r.Cancelled) {
      $status.Text = "Status: $($r.Message)"
      try { Close-SchumanOpenForms } catch {}
    }
  }
  catch {
    Show-UiError -Title 'Schuman' -Message 'Could not complete "Cerrar codigo".' -Exception $_.Exception
  }
}.GetNewClosure()
$closeDocsAction = {
  $cmd = Get-Command -Name Invoke-UiEmergencyClose -CommandType Function -ErrorAction SilentlyContinue
  if (-not $cmd) {
    Show-UiError -Title 'Schuman' -Message 'Close helper not available for documents.'
    return
  }
  try {
    $r = Invoke-UiEmergencyClose -ActionLabel 'Cerrar documentos' -ExecutableNames @('winword.exe', 'excel.exe') -Owner $form -Mode 'Documents' -MainForm $form -BaseDir $projectRoot
    if ($r -and -not $r.Cancelled) {
      $status.Text = "Status: $($r.Message)"
    }
  }
  catch {
    Show-UiError -Title 'Schuman' -Message 'Could not complete "Cerrar documentos".' -Exception $_.Exception
  }
}.GetNewClosure()
$forceUpdateAction = {
  if (-not $script:MainExcelReady) {
    [System.Windows.Forms.MessageBox]::Show('Please load Excel first.', 'Validation') | Out-Null
    return
  }
  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    [System.Windows.Forms.MessageBox]::Show('Please load Excel first.','Validation') | Out-Null
    Update-MainExcelDependentControls -ExcelReady $false -Reason 'Please load Excel first.'
    return
  }

  Set-MainBusyState -IsBusy $true -Text 'Preparing force update'
  $opts = $null
  try {
    if ($showForceUpdateOptionsDialogHandler) {
      $opts = & $showForceUpdateOptionsDialogHandler -Owner $form
    }
    else {
      if (-not $newFallbackForceUpdateOptionsDialogHandler) {
        throw 'Force update options dialog handler is not available.'
      }
      $opts = & $newFallbackForceUpdateOptionsDialogHandler -Owner $form
    }
  } finally {
    Set-MainBusyState -IsBusy $false
  }

  if (-not $opts -or -not $opts.ok) {
    $status.Text = 'Status: Force update canceled'
    return
  }

  $btnDashboard.Enabled = $false
  $btnGenerate.Enabled = $false
  $btnForce.Enabled = $false
  $status.Text = ("Status: Running force update ({0}, {1}, fast={2}, max={3})..." -f $opts.processingScope, $opts.piSearchMode, $opts.fastMode, $opts.maxTickets)
  try {
    if (-not $invokeForceExcelUpdateHandler) {
      throw 'Invoke-ForceExcelUpdate handler is not available.'
    }
    $res = & $invokeForceExcelUpdateHandler -Excel $excel -Sheet $SheetName -ProcessingScope $opts.processingScope -PiSearchMode $opts.piSearchMode -MaxTickets ([int]$opts.maxTickets) -FastMode ([bool]$opts.fastMode)
    if ($res.ok) {
      $status.Text = ("Status: Force update complete ({0}s, tickets={1})" -f $res.durationSec, $res.ticketCount)
      [System.Windows.Forms.MessageBox]::Show(
        "Force update completed successfully.`r`nDuration: $($res.durationSec)s`r`nTickets: $($res.ticketCount)`r`nScope: $($opts.processingScope)`r`nPI mode: $($opts.piSearchMode)`r`nFast mode: $($opts.fastMode)`r`nMax tickets: $($opts.maxTickets)",
        'Force Update'
      ) | Out-Null
    } else {
      $status.Text = 'Status: Force update failed'
      $err = ("" + $res.errorMessage).Trim()
      if (-not $err) { $err = 'Force update failed. Check logs under system/runs.' }
      [System.Windows.Forms.MessageBox]::Show($err,'Error') | Out-Null
    }
  }
  finally {
    $btnDashboard.Enabled = $true
    $btnGenerate.Enabled = $true
    $btnForce.Enabled = $true
  }
}.GetNewClosure()

$initialExcelReady = Test-MainExcelReady -ExcelPathValue $txtExcel.Text -SheetNameValue $SheetName
if ($initialExcelReady) {
  Update-MainExcelDependentControls -ExcelReady $true
} else {
  Update-MainExcelDependentControls -ExcelReady $false -Reason 'Please load Excel first.'
}

$btnDashboard.Add_Click(({
  Invoke-UiHandler -Context 'Open Dashboard' -Action $openDashboardAction
}).GetNewClosure())
$btnGenerate.Add_Click(({
  Invoke-UiHandler -Context 'Open Generate' -Action $openGenerateAction
}).GetNewClosure())
$btnSettings.Add_Click(({
  try {
    Set-MainBusyState -IsBusy $true -Text 'Opening settings'
    & $showSettingsDialog
  }
  catch { Show-UiError -Context 'Open Settings' -ErrorRecord $_ }
  finally { Set-MainBusyState -IsBusy $false }
}).GetNewClosure())
$btnCloseCode.Add_Click(({
  Invoke-UiHandler -Context 'Cerrar codigo' -Action $closeCodeAction
}).GetNewClosure())
$btnCloseDocs.Add_Click(({
  Invoke-UiHandler -Context 'Cerrar documentos' -Action $closeDocsAction
}).GetNewClosure())
$btnForce.Add_Click(({
  Invoke-UiHandler -Context 'Force Update' -Action $forceUpdateAction
}).GetNewClosure())

[void]$form.add_FormClosed(({
  try {
    if ($startupSession) { Close-ServiceNowSession -Session $startupSession }
  } catch {}
  try { Stop-SchumanOwnedResources -Mode 'All' | Out-Null } catch {}
  try {
    if ($script:MainBusyTimer) {
      $script:MainBusyTimer.Stop()
      $script:MainBusyTimer.Dispose()
    }
  } catch {}
}).GetNewClosure())

[void]$form.add_FormClosing(({
  param($sender, $eventArgs)
  try {
    try { Stop-SchumanOwnedResources -Mode 'All' | Out-Null } catch {}
    try { Close-SchumanOpenForms -Except $sender } catch {}
  }
  catch {}
}).GetNewClosure())

[void]$form.ShowDialog()
