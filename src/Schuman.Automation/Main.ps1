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
$uiHelpersPath = Join-Path $PSScriptRoot 'UI\UiHelpers.ps1'
$themePath = Join-Path $PSScriptRoot 'UI\Theme.ps1'
$dashboardUiPath = Join-Path $PSScriptRoot 'UI\DashboardUI.ps1'
$generateUiPath = Join-Path $PSScriptRoot 'UI\GenerateUI.ps1'

foreach ($p in @($invokeScript, $importModulesPath, $uiHelpersPath, $themePath, $dashboardUiPath, $generateUiPath)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required file not found: $p" }
}

. $importModulesPath
. $uiHelpersPath
. $themePath
. $dashboardUiPath
. $generateUiPath

$script:UiLogPath = Join-Path (Join-Path $env:TEMP 'Schuman') 'schuman-ui.log'
$script:LogPath = $script:UiLogPath
try {
  $uiLogDir = Split-Path -Parent $script:UiLogPath
  if ($uiLogDir -and -not (Test-Path -LiteralPath $uiLogDir)) {
    New-Item -ItemType Directory -Path $uiLogDir -Force | Out-Null
  }
}
catch {}

function Write-UiTrace {
  param(
    [string]$Level = 'INFO',
    [string]$Message = ''
  )
  $lvl = ("" + $Level).Trim().ToUpperInvariant()
  $msg = ("" + $Message).Trim()
  if (-not $msg) { return }
  $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $lvl, $msg
  try { Add-Content -LiteralPath $script:UiLogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
  try {
    if (Get-Command -Name Write-Log -ErrorAction SilentlyContinue) {
      Write-Log -Level $lvl -Message $msg
    }
  }
  catch {}
}

function global:Write-Log {
  param(
    [string]$Message,
    [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO',
    [string]$LogPath = ''
  )
  $msg = ("" + $Message).Trim()
  if (-not $msg) { return }
  $path = ("" + $LogPath).Trim()
  if (-not $path) {
    $path = Join-Path (Join-Path $env:TEMP 'Schuman') 'schuman-ui.log'
  }
  try {
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
  }
  catch {}
  $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $msg
  try { Add-Content -LiteralPath $path -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

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
    Kind    = ("" + $Kind).Trim()
    Object  = $Object
    Tag     = ("" + $Tag).Trim()
    AddedAt = [DateTime]::UtcNow
    Id      = [Guid]::NewGuid().ToString('N')
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
      Id      = [int]$Process.Id
      Name    = ("" + $Process.ProcessName).Trim()
      Process = $Process
      Tag     = ("" + $Tag).Trim()
      AddedAt = [DateTime]::UtcNow
    }
  }
  catch {}
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
    [ValidateSet('Code', 'Documents', 'All')][string]$Mode = 'All'
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
    }
    catch {}
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
    }
    catch { $procFailed++ }
    Unregister-SchumanOwnedProcess -ProcessId $p.Id
  }

  return [pscustomobject]@{
    ComClosedCount     = [int]$comClosed
    ProcessClosedCount = [int]$procClosed
    ProcessFailedCount = [int]$procFailed
  }
}

function global:Close-SchumanAllResources {
  param(
    [ValidateSet('Code', 'Documents', 'All')][string]$Mode = 'All'
  )

  $result = @{
    ClosedProcesses = New-Object System.Collections.Generic.List[string]
    ClosedDocs      = 0
    Errors          = New-Object System.Collections.Generic.List[string]
    Skipped         = 0
  }

  try {
    $owned = Stop-SchumanOwnedResources -Mode $Mode
    if ($owned) {
      try { $result.ClosedDocs = [int]$owned.ComClosedCount } catch {}
      if ($owned.ProcessFailedCount -gt 0) {
        $result.Errors.Add(("Owned process cleanup had {0} failure(s)." -f [int]$owned.ProcessFailedCount))
      }
    }
  }
  catch {
    $result.Errors.Add(("Owned resource cleanup failed: {0}" -f $_.Exception.Message))
  }

  $result.ClosedProcesses = @($result.ClosedProcesses | Sort-Object -Unique)
  $result.Errors = @($result.Errors)
  return $result
}

function global:Invoke-UiEmergencyClose {
  param(
    [string]$ActionLabel = '',
    [string[]]$ExecutableNames = @(),
    [System.Windows.Forms.IWin32Window]$Owner = $null,
    [ValidateSet('Code', 'Documents', 'All')][string]$Mode = 'All',
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
  if (-not $modeWasExplicit -and $labelText -match 'code') { $modeResolved = 'Code' }
  elseif (-not $modeWasExplicit -and $labelText -match 'document') { $modeResolved = 'Documents' }
  elseif (-not $modeWasExplicit -and $ExecutableNames -and @($ExecutableNames).Count -gt 0) {
    $joined = ("" + ($ExecutableNames -join ' ')).ToLowerInvariant()
    if ($joined -match 'code|cursor') { $modeResolved = 'Code' }
    elseif ($joined -match 'word|excel') { $modeResolved = 'Documents' }
  }

  $cleanup = Close-SchumanAllResources -Mode $modeResolved
  $killedCount = @($cleanup.ClosedProcesses).Count
  $failedCount = @($cleanup.Errors).Count
  if ($externalResult -and $externalResult.PSObject.Properties['KilledCount']) {
    try { $killedCount += [int]$externalResult.KilledCount } catch {}
  }
  if ($externalResult -and $externalResult.PSObject.Properties['FailedCount']) {
    try { $failedCount += [int]$externalResult.FailedCount } catch {}
  }

  return [pscustomobject]@{
    Cancelled   = $false
    KilledCount = [int]$killedCount
    FailedCount = [int]$failedCount
    Message     = ("Closed: {0} processes, {1} documents. Skipped: {2}. Errors: {3}." -f [int]$killedCount, [int]$cleanup.ClosedDocs, [int]$cleanup.Skipped, [int]$failedCount)
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
    $safeMessage = 'An unexpected UI error occurred.'
  }

  $ex = $null
  if ($Exception) { $ex = $Exception }
  elseif ($ErrorRecord -and $ErrorRecord.Exception) { $ex = $ErrorRecord.Exception }

  $detail = @()
  try {
    if ($ex) {
      $detail += ("Type: " + $ex.GetType().FullName)
      $detail += ("Message: " + ("" + $ex.Message).Trim())
      if ($ex.InnerException) {
        $detail += ("Inner Type: " + $ex.InnerException.GetType().FullName)
        $detail += ("Inner Message: " + ("" + $ex.InnerException.Message).Trim())
      }
      $stack = ("" + $ex.StackTrace).Trim()
      if ($stack) {
        $detail += 'StackTrace:'
        $detail += $stack
      }
    }
  }
  catch {}

  $detailText = if ($detail.Count -gt 0) { ($detail -join "`r`n") } else { 'No exception details were captured.' }
  $exceptionLine = ''
  if ($ex) {
    $exceptionLine = ("{0}: {1}" -f $ex.GetType().FullName, ("" + $ex.Message).Trim())
  }
  $summary = if ($exceptionLine) {
    "$safeMessage`r`n`r`n$exceptionLine`r`n`r`nFull details were written to:`r`n$script:LogPath"
  }
  else {
    "$safeMessage`r`n`r`nFull details were written to:`r`n$script:LogPath"
  }
  Write-UiTrace -Level 'ERROR' -Message ("{0} | {1} | {2}" -f $safeTitle, $safeMessage, $detailText)

  $dlg = $null
  try {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $safeTitle
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Size = New-Object System.Drawing.Size(840, 520)
    $dlg.MinimumSize = New-Object System.Drawing.Size(760, 420)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $dlg.MaximizeBox = $true
    $dlg.MinimizeBox = $false
    $dlg.TopMost = $true

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $layout.Padding = New-Object System.Windows.Forms.Padding(12)
    $dlg.Controls.Add($layout)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $true
    $lbl.Text = $summary
    $layout.Controls.Add($lbl, 0, 0)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txt.ReadOnly = $true
    $txt.WordWrap = $false
    $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txt.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txt.Text = $detailText
    $layout.Controls.Add($txt, 0, 1)

    $btnBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $btnBar.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $btnBar.AutoSize = $true
    $btnBar.Dock = [System.Windows.Forms.DockStyle]::Fill
    $layout.Controls.Add($btnBar, 0, 2)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Width = 90
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    [void]$btnBar.Controls.Add($btnClose)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy Details'
    $btnCopy.Width = 120
    $btnCopy.Add_Click(({
          try { [System.Windows.Forms.Clipboard]::SetText($txt.Text) } catch {}
        }).GetNewClosure())
    [void]$btnBar.Controls.Add($btnCopy)

    $dlg.AcceptButton = $btnClose
    try {
      $themeForDialog = $script:CurrentMainTheme
      if (-not $themeForDialog -and $global:CurrentMainTheme) { $themeForDialog = $global:CurrentMainTheme }
      $scaleForDialog = 1.0
      try {
        if ($script:CurrentMainFontScale) { $scaleForDialog = [double]$script:CurrentMainFontScale }
        elseif ($global:CurrentMainFontScale) { $scaleForDialog = [double]$global:CurrentMainFontScale }
      }
      catch { $scaleForDialog = 1.0 }
      if ($themeForDialog) {
        Set-UiControlRole -Control $lbl -Role 'MutedLabel'
        Set-UiControlRole -Control $btnClose -Role 'PrimaryButton'
        Set-UiControlRole -Control $btnCopy -Role 'SecondaryButton'
        Apply-SchumanTheme -RootControl $dlg -Theme $themeForDialog -FontScale $scaleForDialog
      }
    }
    catch {}
    [void]$dlg.ShowDialog()
  }
  catch {
    try { [System.Windows.Forms.MessageBox]::Show("$safeMessage`r`n`r`n$detailText", $safeTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
  }
  finally {
    try { if ($dlg -and -not $dlg.IsDisposed) { $dlg.Dispose() } } catch {}
  }
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
    [string]$Context = '',
    [string]$ActionName = 'UI Action',
    [scriptblock]$Action,
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $ctx = ("" + $Context).Trim()
  if (-not $ctx) { $ctx = ("" + $ActionName).Trim() }
  if (-not $ctx) { $ctx = 'UI Action' }

  try {
    if (-not $Action) { return $null }
    return (& $Action)
  }
  catch {
    $errText = ("" + $_.Exception.Message).Trim()
    if (-not $errText) { $errText = 'Unhandled UI exception.' }
    Write-UiTrace -Level 'ERROR' -Message ("{0}: {1}" -f $ctx, $errText)
    Show-UiError -Title 'Schuman' -Message ("{0} failed." -f $ctx) -Exception $_.Exception -Context $ctx
    return $null
  }
}

function global:Invoke-UiSafe {
  param(
    [scriptblock]$Action,
    [string]$Context = 'UI Action'
  )
  return (Invoke-SafeUiAction -Context $Context -Action $Action)
}

function global:Test-ControlAlive {
  param([System.Windows.Forms.Control]$Control)
  if (-not $Control) { return $false }
  try {
    if ($Control.IsDisposed) { return $false }
    if (-not $Control.IsHandleCreated) { return $false }
    return $true
  }
  catch {
    return $false
  }
}

function global:Invoke-OnUiThread {
  param(
    [System.Windows.Forms.Control]$Control,
    [scriptblock]$Action,
    [switch]$Synchronous
  )
  if (-not (Test-ControlAlive $Control)) { return }
  if (-not $Action) { return }
  try {
    if ($Control.InvokeRequired) {
      $safeAction = $Action.GetNewClosure()
      $invoker = ([System.Windows.Forms.MethodInvoker]{
            try { & $safeAction } catch { Write-UiTrace -Level 'WARN' -Message ("Invoke-OnUiThread async callback failed: " + $_.Exception.Message) }
          }).GetNewClosure()
      if ($Synchronous) {
        [void]$Control.Invoke($invoker)
      }
      else {
        [void]$Control.BeginInvoke($invoker)
      }
    }
    else {
      & $Action
    }
  }
  catch {
    Write-UiTrace -Level 'WARN' -Message ("Invoke-OnUiThread failed: " + $_.Exception.Message)
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

  $script:ThreadExceptionHandler = [System.Threading.ThreadExceptionEventHandler] {
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

  $script:DomainExceptionHandler = [System.UnhandledExceptionEventHandler] {
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
  Invoke-SafeUiAction -Context $Context -Action $Action | Out-Null
}

Register-WinFormsGlobalExceptionHandling
# Re-apply shared helpers to guarantee canonical wrappers and non-blocking error policy.
. $uiHelpersPath
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

function global:Close-SchumanUiWindows {
  param(
    [System.Windows.Forms.Form]$MainForm = $null,
    [System.Windows.Forms.Form]$GeneratorForm = $null
  )

  $closed = 0
  try {
    if ($GeneratorForm -and -not $GeneratorForm.IsDisposed) {
      try { $GeneratorForm.Close(); $closed++ } catch {}
    }
  }
  catch {}
  try {
    if ($script:DashboardForm -and -not $script:DashboardForm.IsDisposed) {
      try { $script:DashboardForm.Close(); $closed++ } catch {}
      $script:DashboardForm = $null
    }
  }
  catch {}

  $openForms = @()
  try { $openForms = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ }) } catch {}
  foreach ($openForm in $openForms) {
    if (-not $openForm -or $openForm.IsDisposed) { continue }
    if ($MainForm -and ($openForm -eq $MainForm)) { continue }
    if ($GeneratorForm -and ($openForm -eq $GeneratorForm)) { continue }
    $title = ("" + $openForm.Text).Trim()
    if ($title -eq 'Check-in / Check-out Dashboard' -or $title -eq 'Schuman Word Generator') {
      try { $openForm.Close(); $closed++ } catch {}
    }
  }
  return [int]$closed
}

if (-not (Get-Variable -Name SchumanAppShuttingDown -Scope Script -ErrorAction SilentlyContinue)) {
  [bool]$script:SchumanAppShuttingDown = $false
}
function global:Shutdown-SchumanApp {
  param(
    [System.Windows.Forms.Form]$CurrentForm = $null
  )
  if ($script:SchumanAppShuttingDown) { return }
  $script:SchumanAppShuttingDown = $true
  try {
    try { Close-SchumanAllResources -Mode 'All' | Out-Null } catch {}

    $forms = @()
    try { $forms = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ }) } catch { $forms = @() }
    foreach ($openForm in $forms) {
      if (-not $openForm -or $openForm.IsDisposed) { continue }
      if ($CurrentForm -and ($openForm -eq $CurrentForm)) { continue }
      try { $openForm.Close() } catch {}
    }
    if ($CurrentForm -and -not $CurrentForm.IsDisposed) {
      try { $CurrentForm.Close() } catch {}
    }
    try { [System.Windows.Forms.Application]::Exit() } catch {}
  }
  finally {
    $script:SchumanAppShuttingDown = $false
  }
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
    }
    else {
      [void]$parts.Add('"' + ($text -replace '"', '\"') + '"')
    }
  }
  return ($parts -join ' ')
}

function Get-RunningSchumanOperationProcesses {
  param(
    [string[]]$Operations = @('Export', 'DocsGenerate')
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
              ProcessId   = [int]$p.ProcessId
              Operation   = $op
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
      avg_seconds  = [double]$DurationSeconds
      samples      = 1
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
      avg_seconds  = [double]$newAvg
      samples      = [int]$newSamples
      last_seconds = [int]$DurationSeconds
    }
  }

  Save-UiMetrics -Config $Config -Metrics $m
}

function global:Show-ForceUpdateOptionsDialog {
  param(
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $theme = $script:CurrentMainTheme
  if (-not $theme -and $global:CurrentMainTheme) { $theme = $global:CurrentMainTheme }
  $scale = 1.0
  try { if ($script:CurrentMainFontScale) { $scale = [double]$script:CurrentMainFontScale } } catch {}
  if ($scale -le 0) { $scale = 1.0 }

  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Force Update Options'
  $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
  $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
  $dlg.MaximizeBox = $false
  $dlg.MinimizeBox = $false
  $dlg.Size = New-Object System.Drawing.Size(700, 520)
  $dlg.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 26)
  $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
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
    $p.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 34)

    $t = New-Object System.Windows.Forms.Label
    $t.Text = $Title
    $t.AutoSize = $true
    $t.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $t.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 175)
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
  $scopeGroup.Panel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
  [void]$body.Controls.Add($scopeGroup.Panel, 0, 0)

  $piGroup = & $newGroupPanel 'PI Search Mode'
  $piGroup.Panel.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
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
    $chk.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $chk.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    return $chk
  }

  $scopeChoices = @(
    (& $newChoice 'RITM only (fastest)' 'RitmOnly' $true),
    (& $newChoice 'INC only' 'IncOnly' $false),
    (& $newChoice 'All tickets' 'All' $false)
  )
  foreach ($c in $scopeChoices) { [void]$scopeGroup.Flow.Controls.Add($c) }

  $piChoices = @(
    (& $newChoice 'Configuration Item only (fastest)' 'ConfigurationItemOnly' $false),
    (& $newChoice 'Comments only' 'CommentsOnly' $false),
    (& $newChoice 'Comments + Configuration Item' 'CommentsAndCI' $true)
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
  $perfPanel.BackColor = [System.Drawing.Color]::FromArgb(32, 32, 34)
  $perfPanel.Padding = New-Object System.Windows.Forms.Padding(12)
  $perfPanel.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
  [void]$layout.Controls.Add($perfPanel, 0, 2)

  $lblPerf = New-Object System.Windows.Forms.Label
  $lblPerf.Text = 'Performance'
  $lblPerf.AutoSize = $true
  $lblPerf.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
  $lblPerf.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 175)
  $perfPanel.Controls.Add($lblPerf)

  $chkFastMode = New-Object System.Windows.Forms.CheckBox
  $chkFastMode.Text = 'Fast mode (skip deep legal name fallback)'
  $chkFastMode.AutoSize = $true
  $chkFastMode.Checked = $false
  $chkFastMode.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
  $chkFastMode.Location = New-Object System.Drawing.Point(0, 30)
  $perfPanel.Controls.Add($chkFastMode)

  $lblMax = New-Object System.Windows.Forms.Label
  $lblMax.Text = 'Max tickets (0 = all)'
  $lblMax.AutoSize = $true
  $lblMax.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 184)
  $lblMax.Location = New-Object System.Drawing.Point(0, 58)
  $perfPanel.Controls.Add($lblMax)

  $numMaxTickets = New-Object System.Windows.Forms.NumericUpDown
  $numMaxTickets.Minimum = 0
  $numMaxTickets.Maximum = 10000
  $numMaxTickets.Value = 0
  $numMaxTickets.Width = 100
  $numMaxTickets.Location = New-Object System.Drawing.Point(170, 56)
  $numMaxTickets.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 26)
  $numMaxTickets.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
  $perfPanel.Controls.Add($numMaxTickets)

  $hint = New-Object System.Windows.Forms.Label
  $hint.Text = 'Tip: RITM only + Configuration Item only is usually the fastest combination.'
  $hint.AutoSize = $true
  $hint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 126)
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
  $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
  $btnOk.ForeColor = [System.Drawing.Color]::White
  $btnOk.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
  $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
  [void]$buttons.Controls.Add($btnOk)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = 'Cancel'
  $btnCancel.Width = 100
  $btnCancel.Height = 32
  $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(36, 36, 38)
  $btnCancel.ForeColor = [System.Drawing.Color]::FromArgb(230, 230, 230)
  $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(58, 58, 62)
  $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  [void]$buttons.Controls.Add($btnCancel)

  $dlg.AcceptButton = $btnOk
  $dlg.CancelButton = $btnCancel
  try {
    if ($theme) {
      Set-UiControlRole -Control $btnOk -Role 'PrimaryButton'
      Set-UiControlRole -Control $btnCancel -Role 'SecondaryButton'
      Set-UiControlRole -Control $lblPerf -Role 'MutedLabel'
      Set-UiControlRole -Control $hint -Role 'MutedLabel'
      Apply-SchumanTheme -RootControl $dlg -Theme $theme -FontScale $scale
    }
    $res = if ($Owner) { $dlg.ShowDialog($Owner) } else { $dlg.ShowDialog() }
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) {
      return [pscustomobject]@{ ok = $false }
    }

    $scope = @($scopeChoices | Where-Object { $_.Checked } | Select-Object -First 1)
    $scopeValue = if ($scope.Count -gt 0) { "" + $scope[0].Tag } else { 'RitmOnly' }

    $pi = @($piChoices | Where-Object { $_.Checked } | Select-Object -First 1)
    $piValue = if ($pi.Count -gt 0) { "" + $pi[0].Tag } else { 'CommentsAndCI' }

    return [pscustomobject]@{
      ok              = $true
      processingScope = $scopeValue
      piSearchMode    = $piValue
      fastMode        = [bool]$chkFastMode.Checked
      maxTickets      = [int]$numMaxTickets.Value
    }
  }
  finally {
    try { if ($dlg -and -not $dlg.IsDisposed) { $dlg.Dispose() } } catch {}
  }
}

function global:New-FallbackForceUpdateOptionsDialog {
  param(
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $theme = $script:CurrentMainTheme
  if (-not $theme -and $global:CurrentMainTheme) { $theme = $global:CurrentMainTheme }
  $scale = 1.0
  try { if ($script:CurrentMainFontScale) { $scale = [double]$script:CurrentMainFontScale } } catch {}
  if ($scale -le 0) { $scale = 1.0 }

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
  [void]$cmbScope.Items.AddRange(@('RitmOnly', 'IncOnly', 'All'))
  $cmbScope.SelectedItem = 'RitmOnly'
  $cmbScope.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Controls.Add($cmbScope, 1, 0)

  $lblPi = New-Object System.Windows.Forms.Label
  $lblPi.Text = 'PI mode'
  $lblPi.AutoSize = $true
  $root.Controls.Add($lblPi, 0, 1)

  $cmbPi = New-Object System.Windows.Forms.ComboBox
  $cmbPi.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  [void]$cmbPi.Items.AddRange(@('ConfigurationItemOnly', 'CommentsOnly', 'CommentsAndCI'))
  $cmbPi.SelectedItem = 'CommentsAndCI'
  $cmbPi.Dock = [System.Windows.Forms.DockStyle]::Fill
  $root.Controls.Add($cmbPi, 1, 1)

  $chkFast = New-Object System.Windows.Forms.CheckBox
  $chkFast.Text = 'Fast mode'
  $chkFast.Checked = $false
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
  try {
    if ($theme) {
      Set-UiControlRole -Control $btnRun -Role 'PrimaryButton'
      Set-UiControlRole -Control $btnCancel -Role 'SecondaryButton'
      Set-UiControlRole -Control $lblScope -Role 'MutedLabel'
      Set-UiControlRole -Control $lblPi -Role 'MutedLabel'
      Apply-SchumanTheme -RootControl $dlg -Theme $theme -FontScale $scale
    }
    $res = if ($Owner) { $dlg.ShowDialog($Owner) } else { $dlg.ShowDialog() }
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    return @{
      ok              = $true
      processingScope = ("" + $cmbScope.SelectedItem)
      piSearchMode    = ("" + $cmbPi.SelectedItem)
      fastMode        = [bool]$chkFast.Checked
      maxTickets      = 0
    }
  }
  finally {
    try { if ($dlg -and -not $dlg.IsDisposed) { $dlg.Dispose() } } catch {}
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
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)

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
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 12)
  $layout.Controls.Add($status, 0, 1)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 2)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(90, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
  $barTrack.Controls.Add($barFill)

  $animTimer = New-Object System.Windows.Forms.Timer
  $animTimer.Interval = 180

  $args = @(
    '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', $invokeScript,
    '-Operation', 'Export',
    '-ExcelPath', $Excel,
    '-SheetName', $Sheet,
    '-ProcessingScope', 'RitmOnly',
    '-MaxTickets', '40',
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
    Tick        = 0
    Ok          = $false
    Proc        = $proc
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
    [bool]$ExportPdf = $true,
    [int[]]$RowNumbers = @()
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
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)

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
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 12)
  $layout.Controls.Add($status, 0, 1)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 2)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(90, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
  $barTrack.Controls.Add($barFill)

  $args = @(
    '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', $invokeScript,
    '-Operation', 'DocsGenerate',
    '-ExcelPath', $Excel,
    '-SheetName', $Sheet,
    '-TemplatePath', $Template,
    '-OutputDirectory', $Output,
    '-NoPopups'
  )
  if ($ExportPdf) { $args += '-ExportPdf' }
  if ($RowNumbers -and @($RowNumbers).Count -gt 0) {
    $args += @('-RowNumbersCsv', (@($RowNumbers | ForEach-Object { [int]$_ }) -join ','))
  }

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
    Tick        = 0
    Ok          = $false
    Proc        = $proc
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
      }
      catch {
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
  <#
  .SYNOPSIS
  Runs the Force Update export pipeline with a responsive progress dialog.
  .DESCRIPTION
  Launches the Export operation in a child PowerShell process, monitors run logs for
  per-ticket progress, shows live activity history, and allows cancellation.
  .PARAMETER Excel
  Excel workbook path.
  .PARAMETER Sheet
  Worksheet name.
  .PARAMETER ProcessingScope
  Ticket scope filter.
  .PARAMETER PiSearchMode
  PI search strategy.
  .PARAMETER MaxTickets
  Max ticket cap; 0 means all.
  .PARAMETER FastMode
  Enables optimized extraction path.
  .OUTPUTS
  PSCustomObject with ok/durationSec/ticketCount/estimatedSec/errorMessage/cancelled.
  .NOTES
  UI timer callbacks run on UI thread. Child process cancellation is cooperative via kill.
  #>
  param(
    [string]$Excel,
    [string]$Sheet,
    [ValidateSet('Auto', 'RitmOnly', 'IncOnly', 'IncAndRitm', 'All')][string]$ProcessingScope = 'All',
    [ValidateSet('Auto', 'ConfigurationItemOnly', 'CommentsOnly', 'CommentsAndCI')][string]$PiSearchMode = 'Auto',
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
      ok           = $false
      durationSec  = 0
      ticketCount  = 0
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
  $loading.Size = New-Object System.Drawing.Size(760, 460)
  $loading.TopMost = $true
  $loading.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
  $layout.ColumnCount = 1
  $layout.RowCount = 8
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
  $status.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 115)
  $status.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 6)
  $layout.Controls.Add($status, 0, 1)

  $currentTicket = New-Object System.Windows.Forms.Label
  $currentTicket.Text = 'Current ticket: --'
  $currentTicket.AutoSize = $true
  $currentTicket.ForeColor = [System.Drawing.Color]::FromArgb(75, 95, 130)
  $currentTicket.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
  $layout.Controls.Add($currentTicket, 0, 2)

  $timing = New-Object System.Windows.Forms.Label
  $timing.Text = ("Tickets: {0} | ETA: ~{1:mm\:ss} | Elapsed: 00:00" -f $ticketCount, [TimeSpan]::FromSeconds($estimatedSeconds))
  $timing.AutoSize = $true
  $timing.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 96)
  $timing.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
  $layout.Controls.Add($timing, 0, 3)

  $barHost = New-Object System.Windows.Forms.Panel
  $barHost.Dock = [System.Windows.Forms.DockStyle]::Top
  $barHost.Height = 14
  $barHost.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 225)
  $barHost.Padding = New-Object System.Windows.Forms.Padding(1)
  $layout.Controls.Add($barHost, 0, 4)

  $barTrack = New-Object System.Windows.Forms.Panel
  $barTrack.Dock = [System.Windows.Forms.DockStyle]::Fill
  $barTrack.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)
  $barHost.Controls.Add($barTrack)

  $barFill = New-Object System.Windows.Forms.Panel
  $barFill.Size = New-Object System.Drawing.Size(100, 12)
  $barFill.Location = New-Object System.Drawing.Point(0, 0)
  $barFill.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
  $barTrack.Controls.Add($barFill)

  $note = New-Object System.Windows.Forms.Label
  $note.Text = 'This operation rewrites Name/PI/Status/SCTasks in Excel (full refresh).'
  $note.AutoSize = $true
  $note.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 126)
  $note.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
  $layout.Controls.Add($note, 0, 5)

  $historyTitle = New-Object System.Windows.Forms.Label
  $historyTitle.Text = 'Activity / History'
  $historyTitle.AutoSize = $true
  $historyTitle.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 96)
  $historyTitle.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 4)
  $layout.Controls.Add($historyTitle, 0, 6)

  $historyBox = New-Object System.Windows.Forms.ListBox
  $historyBox.Dock = [System.Windows.Forms.DockStyle]::Fill
  $historyBox.Height = 130
  $historyBox.HorizontalScrollbar = $true
  $layout.Controls.Add($historyBox, 0, 7)

  $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
  $buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
  $buttonPanel.Height = 40
  $loading.Controls.Add($buttonPanel)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = 'Cancel'
  $btnCancel.Width = 110
  $btnCancel.Height = 30
  $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $buttonPanel.Controls.Add($btnCancel) | Out-Null

  $args = @(
    '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', $invokeScript,
    '-Operation', 'Export',
    '-ExcelPath', $Excel,
    '-SheetName', $Sheet,
    '-ProcessingScope', $ProcessingScope,
    '-PiSearchMode', $PiSearchMode,
    '-MaxTickets', [string]$MaxTickets,
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
    Tick             = 0
    Ok               = $false
    Proc             = $proc
    DurationSec      = 0
    ProgressDone     = 0
    ProgressTotal    = 0
    ProgressPct      = 0
    RunLogPath       = ''
    LastStatusLine   = ''
    LastLogRefreshMs = -10000
    ExitCode         = -1
    SeenLogLines     = New-Object 'System.Collections.Generic.HashSet[string]'
    Cancelled        = $false
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
            $currentTicket.Text = ("Current ticket: {0}" -f $matches[3])
          }
        }
        $trimLine = ("" + $line).Trim()
        if ($trimLine -and $st.SeenLogLines.Add($trimLine)) {
          $stamp = Get-Date -Format 'HH:mm:ss'
          [void]$historyBox.Items.Add(("[{0}] {1}" -f $stamp, $trimLine))
          while ($historyBox.Items.Count -gt 500) { $historyBox.Items.RemoveAt(0) }
          try { $historyBox.TopIndex = [Math]::Max(0, $historyBox.Items.Count - 1) } catch {}
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
      }
      else {
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
      $throughput = 0.0
      if ($elapsed -gt 0 -and $st.ProgressDone -gt 0) {
        $throughput = [Math]::Round(($st.ProgressDone * 60.0) / [double]$elapsed, 1)
      }
      $timing.Text = ("Tickets: {0} | Progress: {1} | Throughput: {2}/min | ETA: ~{3:mm\:ss} | Elapsed: {4:mm\:ss}" -f $ticketCount, $pctText, $throughput, [TimeSpan]::FromSeconds($etaSec), [TimeSpan]::FromSeconds($elapsed))

      if ($st.ProgressTotal -gt 0) {
        $trackW = [Math]::Max(1, $barTrack.ClientSize.Width)
        $targetW = [int][Math]::Max(14, [Math]::Round(($trackW * $st.ProgressPct) / 100.0))
        $barFill.Width = [Math]::Min($trackW, $targetW)
        $barFill.Left = 0
      }
      else {
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
        }
        else {
          & $resolveRunLog $st
          & $updateProgressFromLog $st
        }
        $st.DurationSec = [int][Math]::Max(1, [Math]::Round($sw.Elapsed.TotalSeconds))
        $loading.Close()
      }
    })

  $btnCancel.Add_Click({
      try {
        $st = $loading.Tag
        if (-not $st) { return }
        $st.Cancelled = $true
        if ($st.Proc -and -not $st.Proc.HasExited) {
          try { $st.Proc.Kill() } catch {}
        }
        $status.Text = 'Cancelling operation...'
      }
      catch {}
    })

  $loading.add_Shown({
      try {
        [void]$loading.Tag.Proc.Start()
        Register-SchumanOwnedProcess -Process $loading.Tag.Proc -Tag 'forceupdate'
        $animTimer.Start()
        $watchTimer.Start()
      }
      catch {
        $loading.Tag.Ok = $false
        $loading.Close()
      }
    })

  $ok = $false
  $durationSec = 0
  $errorMessage = ''
  $cancelled = $false
  try {
    [void]$loading.ShowDialog()
    if ($loading.Tag) {
      $ok = [bool]$loading.Tag.Ok
      $durationSec = [int]$loading.Tag.DurationSec
      $cancelled = [bool]$loading.Tag.Cancelled
      if (-not $ok) {
        if ($cancelled) {
          $errorMessage = 'Export cancelled by user.'
        }
        else {
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
    ok           = $ok
    durationSec  = $durationSec
    ticketCount  = $ticketCount
    estimatedSec = $estimatedSeconds
    errorMessage = $errorMessage
    cancelled    = $cancelled
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

# ---------------------------------------------------------------------------
# Clear-SchumanTempFiles
# Safe cleanup of app-owned temp directories only.
# Returns: [pscustomobject]@{ DeletedFiles; DeletedFolders; Skipped; FreedBytes; Errors }
# ---------------------------------------------------------------------------
function Clear-SchumanTempFiles {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [string]$ProjectRoot = $projectRoot
  )

  $result = [pscustomobject]@{
    DeletedFiles   = 0
    DeletedFolders = 0
    Skipped        = 0
    FreedBytes     = [long]0
    Errors         = [System.Collections.Generic.List[string]]::new()
  }

  # Allowlist of safe temp folders — ONLY within the app boundary or TEMP\Schuman
  $tempTargets = @(
    (Join-Path $env:TEMP 'Schuman'),
    (Join-Path $ProjectRoot 'Temp'),
    (Join-Path $ProjectRoot 'Cache'),
    (Join-Path $ProjectRoot 'system\runs')  # run artefacts, keep last 3 per prefix
  )

  foreach ($targetDir in $tempTargets) {
    if (-not $targetDir -or -not (Test-Path -LiteralPath $targetDir -PathType Container)) { continue }

    # For system\runs - keep only newest 3 per prefix to preserve recent logs
    $isRunsDir = ($targetDir -like '*system*runs*')

    try {
      $items = @(Get-ChildItem -LiteralPath $targetDir -Force -ErrorAction SilentlyContinue)
    }
    catch {
      $result.Errors.Add("Cannot list '$targetDir': $($_.Exception.Message)")
      continue
    }

    if ($isRunsDir) {
      # Group by prefix (text before the timestamp suffix _YYYYMMDD_HHMMSS)
      $groups = $items | Group-Object -Property { ($_.Name -replace '_\d{8}_\d{6}$', '').ToLower() }
      $toDelete = foreach ($g in $groups) {
        # Sort descending by name (timestamp suffix = chronological), keep latest 3
        $sorted = @($g.Group | Sort-Object Name -Descending)
        if ($sorted.Count -gt 3) { $sorted[3..($sorted.Count - 1)] }
      }
      $items = @($toDelete)
    }

    foreach ($item in $items) {
      try {
        if ($item -is [System.IO.DirectoryInfo]) {
          $size = [long]0
          try {
            $size = (Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
              Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
          }
          catch {}
          if ($PSCmdlet.ShouldProcess($item.FullName, 'Delete directory')) {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
          }
          $result.DeletedFolders++
          $result.FreedBytes += [long](if ($size) { $size } else { 0 })
        }
        else {
          $size = [long]0
          try { $size = $item.Length } catch {}
          if ($PSCmdlet.ShouldProcess($item.FullName, 'Delete file')) {
            Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop
          }
          $result.DeletedFiles++
          $result.FreedBytes += [long](if ($size) { $size } else { 0 })
        }
      }
      catch {
        $result.Skipped++
        $result.Errors.Add("Skipped '$($item.FullName)': $($_.Exception.Message)")
      }
    }
  }

  return $result
}

function Get-MainExcelPreferencePath {
  try {
    if ($globalConfig -and $globalConfig.Output -and $globalConfig.Output.SystemRoot -and $globalConfig.Output.DbSubdir) {
      return (Join-Path (Join-Path $globalConfig.Output.SystemRoot $globalConfig.Output.DbSubdir) 'ui-preferences.json')
    }
  }
  catch {}
  return ''
}

function Load-MainExcelPathPreference {
  $path = Get-MainExcelPreferencePath
  if (-not $path -or -not (Test-Path -LiteralPath $path)) { return '' }
  try {
    $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
    if ($json -and $json.PSObject.Properties['excelPath']) {
      return ("" + $json.excelPath).Trim()
    }
  }
  catch {}
  return ''
}

function Save-MainExcelPathPreference {
  param([string]$ExcelPath)
  $safeExcelPath = ("" + $ExcelPath).Trim()
  if (-not $safeExcelPath) { return }
  $path = Get-MainExcelPreferencePath
  if (-not $path) { return }
  try {
    Ensure-Directory -Path (Split-Path -Parent $path) | Out-Null
    $existing = @{}
    if (Test-Path -LiteralPath $path) {
      try {
        $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($json) {
          foreach ($prop in $json.PSObject.Properties) { $existing[$prop.Name] = $prop.Value }
        }
      }
      catch {}
    }
    $existing['excelPath'] = $safeExcelPath
    ($existing | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8
  }
  catch {}
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Schuman — Main'
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.MinimumSize = New-Object System.Drawing.Size(900, 540)
$form.Size = New-Object System.Drawing.Size(960, 580)
$form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0F172A')
$form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10.5)

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = [System.Windows.Forms.DockStyle]::Fill
$root.ColumnCount = 3
$root.RowCount = 3
$root.Padding = New-Object System.Windows.Forms.Padding(24)
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 5)))
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($root)

$hdr = New-Object System.Windows.Forms.Label
$hdr.Text = 'Schuman'
$hdr.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$hdr.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$hdr.AutoSize = $true
$hdr.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$root.Controls.Add($hdr, 1, 0)


$card = New-CardContainer -Title 'Modules'
$card.Border.Dock = [System.Windows.Forms.DockStyle]::Fill
$root.Controls.Add($card.Border, 1, 1)

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = [System.Windows.Forms.DockStyle]::Top
$layout.AutoSize = $true
$layout.ColumnCount = 1
$layout.RowCount = 10
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 0 - Excel label
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 1 - Excel textbox
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 2 - prefs panel
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 3 - main btns (Dashboard/Generate)
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 8)))  # 4 - spacer
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 5 - Force Update
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 6)))  # 6 - spacer
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 7 - Clean temp files
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 8)))  # 8 - spacer
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))  # 9 - Emergency btns
$card.Content.Controls.Add($layout)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = 'Excel file'
$lbl.AutoSize = $true
$lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
$layout.Controls.Add($lbl, 0, 0)

$excelPathRow = New-Object System.Windows.Forms.TableLayoutPanel
$excelPathRow.Dock = [System.Windows.Forms.DockStyle]::Top
$excelPathRow.AutoSize = $true
$excelPathRow.ColumnCount = 2
$excelPathRow.RowCount = 1
[void]$excelPathRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$excelPathRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$excelPathRow.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$layout.Controls.Add($excelPathRow, 0, 1)

$txtExcel = New-Object System.Windows.Forms.TextBox
$txtExcel.Dock = [System.Windows.Forms.DockStyle]::Top
$prefExcelPath = Load-MainExcelPathPreference
$txtExcel.Text = if ($prefExcelPath) { $prefExcelPath } else { $ExcelPath }
$txtExcel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$excelPathRow.Controls.Add($txtExcel, 0, 0)

$btnLoadExcel = New-Object System.Windows.Forms.Button
$btnLoadExcel.Text = 'Load Excel...'
$btnLoadExcel.AutoSize = $true
$btnLoadExcel.Height = 30
$btnLoadExcel.MinimumSize = New-Object System.Drawing.Size(110, 30)
$btnLoadExcel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadExcel.FlatAppearance.BorderSize = 1
$btnLoadExcel.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#334155')
$btnLoadExcel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1E293B')
$btnLoadExcel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$btnLoadExcel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$excelPathRow.Controls.Add($btnLoadExcel, 1, 0)


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
$btnDashboard.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
$btnDashboard.ForeColor = [System.Drawing.Color]::White
$btnDashboard.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnDashboard.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
$btns.Controls.Add($btnDashboard, 0, 0)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = 'Generate'
$btnGenerate.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnGenerate.Height = 46
$btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenerate.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#334155')
$btnGenerate.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1E293B')
$btnGenerate.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$btnGenerate.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnGenerate.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
$btns.Controls.Add($btnGenerate, 1, 0)

$btnForce = New-Object System.Windows.Forms.Button
$btnForce.Text = 'Force Update Excel'
$btnForce.Dock = [System.Windows.Forms.DockStyle]::Top
$btnForce.Height = 40
$btnForce.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnForce.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#334155')
$btnForce.FlatAppearance.BorderSize = 1
$btnForce.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#2563EB')
$btnForce.ForeColor = [System.Drawing.Color]::White
$btnForce.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnForce.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$layout.Controls.Add($btnForce, 0, 5)

# --- Clean Temp Files button (row 7) ---
$btnCleanTemp = New-Object System.Windows.Forms.Button
$btnCleanTemp.Text = 'Clean Temp Files'
$btnCleanTemp.Dock = [System.Windows.Forms.DockStyle]::Top
$btnCleanTemp.Height = 36
$btnCleanTemp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCleanTemp.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#334155')
$btnCleanTemp.FlatAppearance.BorderSize = 1
$btnCleanTemp.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1E293B')
$btnCleanTemp.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#E5E7EB')
$btnCleanTemp.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnCleanTemp.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$layout.Controls.Add($btnCleanTemp, 0, 7)

# --- Emergency buttons panel (row 9) ---
$btnEmergency = New-Object System.Windows.Forms.TableLayoutPanel
$btnEmergency.Dock = [System.Windows.Forms.DockStyle]::Top
$btnEmergency.AutoSize = $true
$btnEmergency.ColumnCount = 1
$btnEmergency.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$layout.Controls.Add($btnEmergency, 0, 9)

$btnCloseCode = New-Object System.Windows.Forms.Button
$btnCloseCode.Text = 'Close All'
$btnCloseCode.Dock = [System.Windows.Forms.DockStyle]::Fill
$btnCloseCode.Height = 36
$btnCloseCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCloseCode.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#DC2626')
$btnCloseCode.FlatAppearance.BorderSize = 1
$btnCloseCode.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#7F1D1D')
$btnCloseCode.ForeColor = [System.Drawing.Color]::White
$btnCloseCode.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 10, [System.Drawing.FontStyle]::Bold)
$btnCloseCode.Margin = New-Object System.Windows.Forms.Padding(0)
$btnEmergency.Controls.Add($btnCloseCode, 0, 0)

$status = New-Object System.Windows.Forms.Label
$status.Text = 'Status: Ready (SSO connected)'
$status.AutoSize = $true
$status.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#94A3B8')
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

$startupOverlay = New-Object System.Windows.Forms.Panel
$startupOverlay.Dock = [System.Windows.Forms.DockStyle]::Fill
$startupOverlay.BackColor = [System.Drawing.Color]::FromArgb(230, 15, 15, 15)
$startupOverlay.Visible = $false
$form.Controls.Add($startupOverlay)
$startupOverlay.BringToFront()

$startupOverlayCenter = New-Object System.Windows.Forms.TableLayoutPanel
$startupOverlayCenter.BackColor = [System.Drawing.Color]::FromArgb(240, 30, 30, 30)
$startupOverlayCenter.ColumnCount = 1
$startupOverlayCenter.RowCount = 4
$startupOverlayCenter.AutoSize = $true
$startupOverlayCenter.Padding = New-Object System.Windows.Forms.Padding(22, 16, 22, 16)
$startupOverlay.Controls.Add($startupOverlayCenter)

$lblStartupTitle = New-Object System.Windows.Forms.Label
$lblStartupTitle.Text = 'Schuman is starting...'
$lblStartupTitle.AutoSize = $true
$lblStartupTitle.Font = New-Object System.Drawing.Font((Get-UiFontNameSafe), 12, [System.Drawing.FontStyle]::Bold)
$lblStartupTitle.ForeColor = [System.Drawing.Color]::FromArgb(235, 235, 235)
$startupOverlayCenter.Controls.Add($lblStartupTitle, 0, 0)

$lblStartupStep = New-Object System.Windows.Forms.Label
$lblStartupStep.Text = 'Checking updates...'
$lblStartupStep.AutoSize = $true
$lblStartupStep.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 4)
$lblStartupStep.ForeColor = [System.Drawing.Color]::FromArgb(195, 195, 195)
$startupOverlayCenter.Controls.Add($lblStartupStep, 0, 1)

$pbStartup = New-Object System.Windows.Forms.ProgressBar
$pbStartup.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$pbStartup.MarqueeAnimationSpeed = 28
$pbStartup.Dock = [System.Windows.Forms.DockStyle]::Top
$pbStartup.Height = 12
$pbStartup.Width = 360
$pbStartup.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 8)
$startupOverlayCenter.Controls.Add($pbStartup, 0, 2)

$lblStartupAnim = New-Object System.Windows.Forms.Label
$lblStartupAnim.Text = 'Loading.'
$lblStartupAnim.AutoSize = $true
$lblStartupAnim.ForeColor = [System.Drawing.Color]::FromArgb(165, 165, 165)
$startupOverlayCenter.Controls.Add($lblStartupAnim, 0, 3)

$script:StartupOverlayAnimTick = 0
$script:StartupOverlayAnimTimer = New-Object System.Windows.Forms.Timer
$script:StartupOverlayAnimTimer.Interval = 230
$script:StartupOverlayAnimTimer.Add_Tick(({
      try {
        if (-not $form -or $form.IsDisposed -or -not $form.IsHandleCreated) { return }
        if (-not $startupOverlay.Visible) { return }
        $script:StartupOverlayAnimTick = [int]$script:StartupOverlayAnimTick + 1
        $lblStartupAnim.Text = ('Loading' + ('.' * (($script:StartupOverlayAnimTick % 3) + 1)))
      }
      catch {}
    }).GetNewClosure())

$startupOverlay.Add_Resize(({
      try {
        if (-not $startupOverlayCenter -or $startupOverlayCenter.IsDisposed) { return }
        $x = [Math]::Max(0, [int](($startupOverlay.ClientSize.Width - $startupOverlayCenter.PreferredSize.Width) / 2))
        $y = [Math]::Max(0, [int](($startupOverlay.ClientSize.Height - $startupOverlayCenter.PreferredSize.Height) / 2))
        $startupOverlayCenter.Location = New-Object System.Drawing.Point($x, $y)
      }
      catch {}
    }).GetNewClosure())

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
    }
    catch {}
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
    }
    catch {}
  }
  else {
    try { $script:MainBusyTimer.Stop() } catch {}
    if ($loadingBar -and -not $loadingBar.IsDisposed) { $loadingBar.Visible = $false }
    $status.Text = 'Status: Ready'
    try {
      Update-MainExcelDependentControls -ExcelReady $script:MainExcelReady
    }
    catch {}
  }
}

$script:UiRoleByControlId = @{}
$script:UiHoverBoundByControlId = @{}
$script:CurrentMainTheme = @{}
$script:CurrentMainFontScale = 1.0
$global:CurrentMainTheme = @{}
$global:CurrentMainFontScale = 1.0
$script:ThemeAppliedStamp = @{}

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
  }
  catch {
    return $Fallback
  }
}

function Get-ShiftedUiColor {
  param([System.Drawing.Color]$Color, [int]$Delta = 0)
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
  $themes['Dark'] = Get-SchumanThemeMidnightEU
  return $themes
}

function Get-SchumanThemeMidnightEU {
  return @{
    FormBG          = Convert-HexToUiColor '#0F172A'
    HeaderBG        = Convert-HexToUiColor '#1E293B'
    CardBG          = Convert-HexToUiColor '#1E293B'
    Border          = Convert-HexToUiColor '#334155'
    Text            = Convert-HexToUiColor '#E5E7EB'
    MutedText       = Convert-HexToUiColor '#94A3B8'
    Primary         = Convert-HexToUiColor '#2563EB'
    PrimaryHover    = Convert-HexToUiColor '#3B82F6'
    PrimaryPressed  = Convert-HexToUiColor '#1D4ED8'
    Secondary       = Convert-HexToUiColor '#1E293B'
    SecondaryHover  = Convert-HexToUiColor '#2B3647'
    Accent          = Convert-HexToUiColor '#2563EB'
    AccentHover     = Convert-HexToUiColor '#3B82F6'
    Danger          = Convert-HexToUiColor '#7F1D1D'
    DangerHover     = Convert-HexToUiColor '#DC2626'
    InputBG         = Convert-HexToUiColor '#0B1220'
    InputBorder     = Convert-HexToUiColor '#334155'
    FocusBorder     = Convert-HexToUiColor '#2563EB'
    GridAltRow      = Convert-HexToUiColor '#172033'
    SelectionBG     = Convert-HexToUiColor '#1E40AF'
    SelectionText   = Convert-HexToUiColor '#FFFFFF'
    Success         = Convert-HexToUiColor '#22C55E'
    Warning         = Convert-HexToUiColor '#F59E0B'
    Error           = Convert-HexToUiColor '#DC2626'
    DisabledBG      = Convert-HexToUiColor '#2B3647'
    DisabledText    = Convert-HexToUiColor '#64748B'
  }
}

function Set-UiControlRole {
  param([System.Windows.Forms.Control]$Control, [string]$Role)
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
  param([System.Windows.Forms.Button]$Button, [hashtable]$Theme, [bool]$Hover = $false)
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
  $pressed = if ($Theme.ContainsKey('PrimaryPressed')) { $Theme.PrimaryPressed } else { Get-ShiftedUiColor -Color $bgHover -Delta -14 }
  $Button.UseVisualStyleBackColor = $false
  $Button.BackColor = $fill
  $Button.ForeColor = Get-ReadableTextColor -Background $fill
  $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $Button.FlatAppearance.BorderSize = 1
  $Button.FlatAppearance.BorderColor = if ($Hover) { $Theme.FocusBorder } else { $Theme.Border }
  $Button.FlatAppearance.MouseOverBackColor = $bgHover
  $Button.FlatAppearance.MouseDownBackColor = $pressed
  if (-not $Button.Enabled -and $Theme.ContainsKey('DisabledBG') -and $Theme.ContainsKey('DisabledText')) {
    $Button.BackColor = $Theme.DisabledBG
    $Button.ForeColor = $Theme.DisabledText
    $Button.FlatAppearance.BorderColor = $Theme.Border
  }
}

function Ensure-UiButtonHoverBinding {
  param([System.Windows.Forms.Button]$Button)
  if (-not $Button) { return }
  $id = [string]$Button.GetHashCode()
  if ($script:UiHoverBoundByControlId.ContainsKey($id)) { return }
  $updateHandler = ${function:Update-UiButtonVisual}
  $Button.Add_MouseEnter(({ param($sender, $eventArgs) try { if ($updateHandler) { & $updateHandler -Button $sender -Theme $script:CurrentMainTheme -Hover:$true } } catch {} }).GetNewClosure())
  $Button.Add_MouseLeave(({ param($sender, $eventArgs) try { if ($updateHandler) { & $updateHandler -Button $sender -Theme $script:CurrentMainTheme -Hover:$false } } catch {} }).GetNewClosure())
  $script:UiHoverBoundByControlId[$id] = $true
}

function Set-ControlThemeSafe {
  param(
    [System.Windows.Forms.Control]$Control,
    [hashtable]$Theme,
    [System.Drawing.Font]$Regular,
    [System.Drawing.Font]$Bold,
    [System.Drawing.Font]$Title
  )
  if (-not $Control -or $Control.IsDisposed -or -not $Theme) { return }
  $role = Get-UiControlRole -Control $Control
  if ($Control -is [System.Windows.Forms.Form]) {
    $Control.BackColor = $Theme.FormBG
    $Control.ForeColor = $Theme.Text
    if ($Regular) { $Control.Font = $Regular }
  }
  elseif ($Control -is [System.Windows.Forms.Panel]) {
    switch ($role) {
      'CardBorder' { $Control.BackColor = $Theme.Border }
      'CardSurface' { $Control.BackColor = $Theme.CardBG }
      'StartupOverlay' { $Control.BackColor = [System.Drawing.Color]::FromArgb(220, 11, 18, 32) }
      'StartupOverlayCard' { $Control.BackColor = $Theme.CardBG }
      default { $Control.BackColor = if ($Control.Parent) { $Control.Parent.BackColor } else { $Theme.FormBG } }
    }
  }
  elseif ($Control -is [System.Windows.Forms.TableLayoutPanel] -or $Control -is [System.Windows.Forms.FlowLayoutPanel]) {
    $Control.BackColor = if ($Control.Parent) { $Control.Parent.BackColor } else { $Theme.FormBG }
  }
  elseif ($Control -is [System.Windows.Forms.GroupBox]) {
    $Control.BackColor = if ($Control.Parent) { $Control.Parent.BackColor } else { $Theme.FormBG }
    $Control.ForeColor = $Theme.Text
    if ($Regular) { $Control.Font = $Regular }
  }
  elseif ($Control -is [System.Windows.Forms.Label]) {
    $Control.ForeColor = if ($role -eq 'MutedLabel' -or $role -eq 'StatusLabel') { $Theme.MutedText } else { $Theme.Text }
    if ($role -eq 'HeaderTitle') {
      if ($Title) { $Control.Font = $Title }
    }
    elseif ($Regular) {
      $Control.Font = $Regular
    }
  }
  elseif ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.RichTextBox] -or $Control -is [System.Windows.Forms.ComboBox]) {
    $Control.BackColor = $Theme.InputBG
    $Control.ForeColor = $Theme.Text
    if ($Regular) { $Control.Font = $Regular }
    if ($Control -is [System.Windows.Forms.TextBox]) { $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle }
    if ($Control -is [System.Windows.Forms.ComboBox]) { $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
  }
  elseif ($Control -is [System.Windows.Forms.CheckBox] -or $Control -is [System.Windows.Forms.RadioButton]) {
    $Control.BackColor = if ($Control.Parent) { $Control.Parent.BackColor } else { $Theme.FormBG }
    $Control.ForeColor = $Theme.Text
    if ($Regular) { $Control.Font = $Regular }
  }
  elseif ($Control -is [System.Windows.Forms.ProgressBar]) {
    $Control.BackColor = $Theme.CardBG
    $Control.ForeColor = $Theme.Primary
  }
  elseif ($Control -is [System.Windows.Forms.DataGridView]) {
    $Control.BackgroundColor = $Theme.CardBG
    $Control.GridColor = $Theme.Border
    $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $Control.EnableHeadersVisualStyles = $false
    $Control.RowHeadersVisible = $false
    $Control.DefaultCellStyle.BackColor = $Theme.InputBG
    $Control.DefaultCellStyle.ForeColor = $Theme.Text
    $Control.DefaultCellStyle.SelectionBackColor = $Theme.SelectionBG
    $Control.DefaultCellStyle.SelectionForeColor = $Theme.SelectionText
    $Control.AlternatingRowsDefaultCellStyle.BackColor = $Theme.GridAltRow
    $Control.AlternatingRowsDefaultCellStyle.ForeColor = $Theme.Text
    $Control.ColumnHeadersDefaultCellStyle.BackColor = $Theme.HeaderBG
    $Control.ColumnHeadersDefaultCellStyle.ForeColor = $Theme.Text
    if ($Bold) { $Control.ColumnHeadersDefaultCellStyle.Font = $Bold }
  }
  elseif ($Control -is [System.Windows.Forms.Button]) {
    if ($Bold) { $Control.Font = $Bold }
    Ensure-UiButtonHoverBinding -Button $Control
    Update-UiButtonVisual -Button $Control -Theme $Theme -Hover:$false
    if (Get-Command -Name Set-UiRoundedButton -ErrorAction SilentlyContinue) {
      Set-UiRoundedButton -Button $Control -Radius 10
    }
  }
}

function Apply-SchumanTheme {
  param([System.Windows.Forms.Control]$RootControl, [hashtable]$Theme, [double]$FontScale = 1.0)
  if (-not $RootControl -or -not $Theme) { return }
  $safeScale = if ($FontScale -le 0) { 1.0 } else { $FontScale }
  $fontName = 'Segoe UI'
  $regular = New-Object System.Drawing.Font($fontName, [single](10.5 * $safeScale), [System.Drawing.FontStyle]::Regular)
  $bold = New-Object System.Drawing.Font($fontName, [single](10.5 * $safeScale), [System.Drawing.FontStyle]::Bold)
  $title = New-Object System.Drawing.Font($fontName, [single](18 * $safeScale), [System.Drawing.FontStyle]::Bold)
  $errors = New-Object System.Collections.Generic.List[string]

  $walk = $null
  $walk = {
    param([System.Windows.Forms.Control]$Ctrl)
    if (-not $Ctrl -or $Ctrl.IsDisposed) { return }
    try {
      Set-ControlThemeSafe -Control $Ctrl -Theme $Theme -Regular $regular -Bold $bold -Title $title
      if ($Ctrl -and $Ctrl.HasChildren) {
        foreach ($child in $Ctrl.Controls) { & $walk $child }
      }
    }
    catch {
      $id = ''
      try { $id = ("" + $Ctrl.Name).Trim() } catch {}
      $errors.Add(("{0}: {1}" -f $(if ($id) { $id } else { $Ctrl.GetType().Name }), $_.Exception.Message)) | Out-Null
    }
  }
  & $walk $RootControl
  if (Get-Command -Name Apply-UiRoundedButtonsRecursive -ErrorAction SilentlyContinue) {
    Apply-UiRoundedButtonsRecursive -Root $RootControl -Radius 10
  }
  if ($errors.Count -gt 0) {
    Write-UiTrace -Level 'WARN' -Message ("Theme application reported {0} control error(s). First: {1}" -f $errors.Count, $errors[0])
  }
}

function Apply-ThemeToControlTree {
  param([System.Windows.Forms.Control]$RootControl, [hashtable]$Theme, [double]$FontScale = 1.0)
  Apply-SchumanTheme -RootControl $RootControl -Theme $Theme -FontScale $FontScale
}

function Sync-ThemeAcrossOpenForms {
  param([switch]$Force)

  if (-not $script:CurrentMainTheme -or $script:CurrentMainTheme.Count -eq 0) { return }
  $scale = 1.0
  try { $scale = [double]$script:CurrentMainFontScale } catch { $scale = 1.0 }
  if ($scale -le 0) { $scale = 1.0 }

  $forms = @()
  try { $forms = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ }) } catch { $forms = @() }
  foreach ($openForm in $forms) {
    if (-not $openForm -or $openForm.IsDisposed) { continue }
    $id = [string]$openForm.GetHashCode()
    $stamp = ("{0}|{1}|{2}" -f $id, $script:CurrentMainTheme.GetHashCode(), [Math]::Round($scale, 3))
    if (-not $Force -and $script:ThemeAppliedStamp.ContainsKey($id) -and $script:ThemeAppliedStamp[$id] -eq $stamp) { continue }
    try {
      Apply-ThemeToControlTree -RootControl $openForm -Theme $script:CurrentMainTheme -FontScale $scale
      if (Get-Command -Name Apply-UiRoundedButtonsRecursive -ErrorAction SilentlyContinue) {
        Apply-UiRoundedButtonsRecursive -Root $openForm -Radius 10
      }
      try { $openForm.Invalidate() } catch {}
      try { $openForm.Refresh() } catch {}
      $script:ThemeAppliedStamp[$id] = $stamp
    }
    catch {
      try { Write-Log -Level WARN -Message ("Theme sync failed on window '{0}': {1}" -f ("" + $openForm.Text), $_.Exception.Message) } catch {}
    }
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
  & $script:SetUiControlRoleHandler -Control $btnLoadExcel -Role 'SecondaryButton'
  & $script:SetUiControlRoleHandler -Control $btnCloseCode -Role 'DangerButton'
  & $script:SetUiControlRoleHandler -Control $card.Border -Role 'CardBorder'
  & $script:SetUiControlRoleHandler -Control $card.Content.Parent -Role 'CardSurface'
  & $script:SetUiControlRoleHandler -Control $hdr -Role 'HeaderTitle'
  & $script:SetUiControlRoleHandler -Control $status -Role 'StatusLabel'
  & $script:SetUiControlRoleHandler -Control $lbl -Role 'MutedLabel'
  & $script:SetUiControlRoleHandler -Control $startupOverlay -Role 'StartupOverlay'
  & $script:SetUiControlRoleHandler -Control $startupOverlayCenter -Role 'StartupOverlayCard'
  & $script:SetUiControlRoleHandler -Control $lblStartupStep -Role 'MutedLabel'
  & $script:SetUiControlRoleHandler -Control $lblStartupAnim -Role 'MutedLabel'
}

function Apply-MainUiTheme {
  param(
    [string]$ThemeName,
    [string]$AccentName,
    [int]$FontScale = 100,
    [bool]$Compact = $false
  )

  $safeTheme = 'Dark'
  $safeScale = 100
  $safeCompact = $false

  $baseTheme = $script:Themes[$safeTheme]
  if (-not $baseTheme -or -not ($baseTheme -is [hashtable])) {
    $safeTheme = 'Dark'
    $baseTheme = $script:Themes['Dark']
  }
  $resolved = @{}
  foreach ($k in $baseTheme.Keys) { $resolved[$k] = $baseTheme[$k] }

  $script:CurrentMainTheme = $resolved
  $script:CurrentMainFontScale = ($safeScale / 100.0)
  $global:CurrentMainTheme = $resolved
  $global:CurrentMainFontScale = ($safeScale / 100.0)
  $script:ThemeAppliedStamp = @{}

  if ($script:ApplyThemeToControlTreeHandler) {
    if ($form -and -not $form.IsDisposed) {
      & $script:ApplyThemeToControlTreeHandler -RootControl $form -Theme $resolved -FontScale ($safeScale / 100.0)
      try { $form.Invalidate() } catch {}
      try { $form.Refresh() } catch {}
    }
    $formsToRefresh = @()
    try { $formsToRefresh = @([System.Windows.Forms.Application]::OpenForms | ForEach-Object { $_ }) } catch {}
    try { Write-Log -Level INFO -Message ("Theme apply '{0}' on {1} open form(s)." -f $safeTheme, $formsToRefresh.Count) } catch {}
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
  try { Sync-ThemeAcrossOpenForms -Force } catch {}

  $moduleHeight = if ($safeCompact) { 38 } else { 46 }
  $dangerHeight = if ($safeCompact) { 32 } else { 36 }
  $forceHeight = if ($safeCompact) { 34 } else { 40 }
  if ($btnDashboard -and -not $btnDashboard.IsDisposed) { $btnDashboard.Height = $moduleHeight }
  if ($btnGenerate -and -not $btnGenerate.IsDisposed) { $btnGenerate.Height = $moduleHeight }
  if ($btnForce -and -not $btnForce.IsDisposed) { $btnForce.Height = $forceHeight }
  if ($btnCloseCode -and -not $btnCloseCode.IsDisposed) { $btnCloseCode.Height = $dangerHeight }
}

try { Apply-MainUiTheme -ThemeName 'Dark' -AccentName '' -FontScale 100 -Compact:$false }
catch { Show-UiError -Title 'Theme' -Message 'Could not apply initial theme.' -Exception $_.Exception }

$script:ApplyMainUiThemeHandler = ${function:Apply-MainUiTheme}
$script:MainExcelReady = $false
$script:ExcelReady = $false
$script:ForceUpdateInProgress = $false
$script:CancelRequested = $false
$script:ExcelRefreshInProgress = $false
$script:ExcelRefreshWorker = $null
$script:ExcelLoadErrorShown = $false
$script:LastExcelRefreshAt = $null
$script:StartupInitDone = $false
$script:ExcelRefreshStartedAt = $null
$script:ExcelRefreshTimeoutSeconds = 0
$script:ExcelRefreshRunId = 0
$script:ActiveExcelRefreshRunId = -1
$script:ExcelRefreshWatchdogTimer = $null

function Update-MainExcelDependentControls {
  param([bool]$ExcelReady, [string]$Reason = '')
  $ready = [bool]$ExcelReady
  $script:MainExcelReady = $ready
  $script:ExcelReady = $ready
  $pathValid = $false
  try {
    $candidate = ("" + $txtExcel.Text).Trim()
    $pathValid = (-not [string]::IsNullOrWhiteSpace($candidate)) -and (Test-Path -LiteralPath $candidate)
  }
  catch { $pathValid = $false }
  if ($btnDashboard -and -not $btnDashboard.IsDisposed) { $btnDashboard.Enabled = (-not $script:MainBusyActive) }
  if ($btnGenerate -and -not $btnGenerate.IsDisposed) { $btnGenerate.Enabled = (-not $script:MainBusyActive) }
  if ($btnForce -and -not $btnForce.IsDisposed) { $btnForce.Enabled = ($pathValid -and -not $script:ExcelRefreshInProgress -and -not $script:ForceUpdateInProgress -and -not $script:MainBusyActive) }
  if ($btnCleanTemp -and -not $btnCleanTemp.IsDisposed) { $btnCleanTemp.Enabled = (-not $script:MainBusyActive) }
  if ($btnLoadExcel -and -not $btnLoadExcel.IsDisposed) { $btnLoadExcel.Enabled = (-not $script:ExcelRefreshInProgress) }
  if (-not $ExcelReady -and $status -and -not $status.IsDisposed) {
    $msg = ("" + $Reason).Trim()
    if (-not $msg) { $msg = 'Excel still loading...' }
    $status.Text = ("Status: {0}" -f $msg)
  }
  elseif ($ExcelReady -and $status -and -not $status.IsDisposed -and -not $script:MainBusyActive) {
    $status.Text = 'Status: Ready'
  }
}

function global:Set-AppBusyState {
  param(
    [bool]$isBusy,
    [string]$reason = 'Working'
  )

  Set-MainBusyState -IsBusy $isBusy -Text $reason
  if (-not $isBusy) {
    try { Update-MainExcelDependentControls -ExcelReady $script:MainExcelReady } catch {}
  }
}

function global:Restore-UiAfterForceUpdate {
  $script:ForceUpdateInProgress = $false
  Set-AppBusyState -isBusy $false -reason 'Ready'
  try {
    $reason = if ($script:MainExcelReady) { 'Ready' } else { 'Excel is empty' }
    Update-MainExcelDependentControls -ExcelReady $script:MainExcelReady -Reason $reason
  }
  catch {}
}

function Update-StartupOverlay {
  param(
    [bool]$Visible,
    [string]$Title = 'Schuman is starting...',
    [string]$Step = 'Loading modules...'
  )
  if (-not $startupOverlay -or $startupOverlay.IsDisposed) { return }
  if ($Visible) {
    Invoke-OnUiThread -Control $form -Action {
      if ($lblStartupTitle -and -not $lblStartupTitle.IsDisposed) { $lblStartupTitle.Text = if ($Title) { $Title } else { 'Schuman is starting...' } }
      if ($lblStartupStep -and -not $lblStartupStep.IsDisposed) { $lblStartupStep.Text = if ($Step) { $Step } else { 'Loading modules...' } }
      $startupOverlay.Visible = $true
      $startupOverlay.BringToFront()
      try {
        $x = [Math]::Max(0, [int](($startupOverlay.ClientSize.Width - $startupOverlayCenter.PreferredSize.Width) / 2))
        $y = [Math]::Max(0, [int](($startupOverlay.ClientSize.Height - $startupOverlayCenter.PreferredSize.Height) / 2))
        $startupOverlayCenter.Location = New-Object System.Drawing.Point($x, $y)
      }
      catch {}
      $script:StartupOverlayAnimTick = 0
      try { $script:StartupOverlayAnimTimer.Start() } catch {}
      try { $startupOverlay.PerformLayout() } catch {}
    }
  }
  else {
    Invoke-OnUiThread -Control $form -Action {
      try { $script:StartupOverlayAnimTimer.Stop() } catch {}
      $startupOverlay.Visible = $false
    }
  }
}

function Show-ExcelLoadErrorOnce {
  param([string]$Message)
  if ($script:ExcelLoadErrorShown) { return }
  $script:ExcelLoadErrorShown = $true
  $safe = ("" + $Message).Trim()
  if (-not $safe) { $safe = 'Excel file is not available. Please select a valid file.' }
  try { [System.Windows.Forms.MessageBox]::Show($safe, 'Excel load', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null } catch {}
}

function Invoke-SchumanExcelRefresh {
  param(
    [string]$ExcelPath,
    [string]$Reason = 'Manual refresh',
    [ValidateSet('Auto', 'RitmOnly', 'IncOnly', 'IncAndRitm', 'All')][string]$ProcessingScope = 'All',
    [ValidateSet('Auto', 'ConfigurationItemOnly', 'CommentsOnly', 'CommentsAndCI')][string]$PiSearchMode = 'Auto',
    [int]$MaxTickets = 0,
    [bool]$FastMode = $true,
    [int]$TimeoutSeconds = 180
  )

  $result = @{
    Success   = $false
    Message   = ''
    UpdatedAt = $null
    ExcelPath = ("" + $ExcelPath).Trim()
    Errors    = @()
  }

  try {
    $safeExcel = ("" + $ExcelPath).Trim()
    if ([string]::IsNullOrWhiteSpace($safeExcel)) {
      $result.Message = 'Excel path is empty.'
      $result.Errors += $result.Message
      return $result
    }
    if (-not (Test-Path -LiteralPath $safeExcel)) {
      $result.Message = ("Excel file not found: {0}" -f $safeExcel)
      $result.Errors += $result.Message
      return $result
    }
    Write-UiTrace -Level 'INFO' -Message ("Excel refresh started. Reason='{0}', Scope='{1}', PiMode='{2}', TimeoutSec={3}" -f $Reason, $ProcessingScope, $PiSearchMode, [int]$TimeoutSeconds)

    $args = @(
      '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
      '-File', $invokeScript,
      '-Operation', 'Export',
      '-ExcelPath', $safeExcel,
      '-SheetName', $SheetName,
      '-ProcessingScope', $ProcessingScope,
      '-PiSearchMode', $PiSearchMode,
      '-MaxTickets', [string]$MaxTickets,
      '-NoPopups'
    )
    if ($FastMode) { $args += '-SkipLegalNameFallback' }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = (Join-Path $PSHOME 'powershell.exe')
    $psi.Arguments = Convert-ToArgumentString -Tokens $args
    $psi.WorkingDirectory = $projectRoot
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError = $false

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    [void]$proc.Start()

    $waitMs = [Math]::Max(1000, ([int]$TimeoutSeconds * 1000))
    $finished = $proc.WaitForExit($waitMs)
    if (-not $finished) {
      try { $proc.Kill() } catch {}
      try { $proc.WaitForExit(1500) | Out-Null } catch {}
      try { $proc.Dispose() } catch {}
      $result.Message = ("Excel refresh timed out after {0} seconds." -f [int]$TimeoutSeconds)
      $result.Errors += $result.Message
      Write-UiTrace -Level 'ERROR' -Message $result.Message
      return $result
    }

    $exitCode = $proc.ExitCode
    try { $proc.Dispose() } catch {}

    if ($exitCode -ne 0) {
      $msg = ("Refresh process exited with code {0}." -f $exitCode)
      $result.Message = $msg
      $result.Errors += $msg
      Write-UiTrace -Level 'ERROR' -Message $msg
      return $result
    }

    $fs = $null
    try {
      $fs = [System.IO.File]::Open($safeExcel, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    }
    catch {
      $result.Message = 'Excel file is locked or unavailable after refresh.'
      $result.Errors += $_.Exception.Message
      return $result
    }
    finally {
      try { if ($fs) { $fs.Close(); $fs.Dispose() } } catch {}
    }

    $result.Success = $true
    $result.UpdatedAt = [DateTime]::Now
    $result.Message = ("Excel refreshed successfully ({0})." -f $Reason)
    Write-UiTrace -Level 'INFO' -Message $result.Message
    return $result
  }
  catch {
    $result.Message = ("" + $_.Exception.Message).Trim()
    if (-not $result.Message) { $result.Message = 'Refresh failed without exception details.' }
    $result.Errors += $result.Message
    Write-UiTrace -Level 'ERROR' -Message ("Excel refresh exception: " + $result.Message)
    return $result
  }
}

function Start-SchumanExcelRefreshAsync {
  param(
    [string]$ExcelPath,
    [string]$Reason = 'Manual refresh',
    [bool]$ShowOverlay = $true,
    [bool]$Startup = $false,
    [ValidateSet('Auto', 'RitmOnly', 'IncOnly', 'IncAndRitm', 'All')][string]$ProcessingScope = 'All',
    [ValidateSet('Auto', 'ConfigurationItemOnly', 'CommentsOnly', 'CommentsAndCI')][string]$PiSearchMode = 'Auto',
    [int]$MaxTickets = 0,
    [bool]$FastMode = $true,
    [scriptblock]$OnCompleted = $null
  )

  $safeExcel = ("" + $ExcelPath).Trim()
  if ($script:ExcelRefreshInProgress) {
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Refresh already in progress...' }
    Write-UiTrace -Level 'WARN' -Message 'Excel refresh request ignored: already in progress.'
    return
  }
  if ([string]::IsNullOrWhiteSpace($safeExcel) -or -not (Test-Path -LiteralPath $safeExcel)) {
    Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel refresh failed. Please load Excel.'
    Show-ExcelLoadErrorOnce -Message 'Excel file is missing. Please click "Load Excel..." and select a valid .xlsx file.'
    return
  }

  $script:ExcelRefreshInProgress = $true
  Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel not ready yet - loading...'
  if ($ShowOverlay) {
    $titleText = if ($Startup) { 'Schuman is starting...' } else { 'Refreshing Excel...' }
    Update-StartupOverlay -Visible $true -Title $titleText -Step 'Checking updates...'
  }
  if ($status -and -not $status.IsDisposed) { $status.Text = ("Status: {0}" -f $Reason) }
  Write-UiTrace -Level 'INFO' -Message ("Excel refresh queued. Startup={0}; Reason='{1}'" -f $Startup, $Reason)

  $worker = New-Object System.ComponentModel.BackgroundWorker
  $worker.WorkerSupportsCancellation = $false
  $script:ExcelRefreshWorker = $worker

  $invokeRefreshHandler = ${function:Invoke-SchumanExcelRefresh}
  $resolvedTimeoutSeconds = if ($Startup) { 120 } else { 300 }
  $script:ExcelRefreshRunId = [int]$script:ExcelRefreshRunId + 1
  $runId = [int]$script:ExcelRefreshRunId
  $script:ActiveExcelRefreshRunId = $runId
  $script:ExcelRefreshStartedAt = [DateTime]::Now
  $script:ExcelRefreshTimeoutSeconds = $resolvedTimeoutSeconds

  if (-not $script:ExcelRefreshWatchdogTimer) {
    $script:ExcelRefreshWatchdogTimer = New-Object System.Timers.Timer
    $script:ExcelRefreshWatchdogTimer.Interval = 1000
    $script:ExcelRefreshWatchdogTimer.AutoReset = $true
    $script:ExcelRefreshWatchdogTimer.add_Elapsed(({
          param($sender, $eventArgs)
          try {
            if (-not $script:ExcelRefreshInProgress) {
              try { if ($script:ExcelRefreshWatchdogTimer) { $script:ExcelRefreshWatchdogTimer.Stop() } } catch {}
              return
            }
            if (-not $script:ExcelRefreshStartedAt -or $script:ExcelRefreshTimeoutSeconds -le 0) { return }
            $elapsed = [int]([DateTime]::Now - $script:ExcelRefreshStartedAt).TotalSeconds
            if ($elapsed -lt $script:ExcelRefreshTimeoutSeconds) { return }

            Write-UiTrace -Level 'ERROR' -Message ("Excel refresh watchdog timeout after {0}s." -f $elapsed)
            $script:ExcelRefreshInProgress = $false
            $script:ExcelRefreshWorker = $null
            $script:ActiveExcelRefreshRunId = -1
            $script:ExcelRefreshStartedAt = $null
            $script:ExcelRefreshTimeoutSeconds = 0
            try { if ($script:ExcelRefreshWatchdogTimer) { $script:ExcelRefreshWatchdogTimer.Stop() } } catch {}

            try {
              Invoke-OnUiThread -Control $form -Action {
                try { Update-StartupOverlay -Visible $false } catch {}
                try { Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel refresh timeout. Please load Excel or retry Force Update.' } catch {}
                if ($status -and -not $status.IsDisposed) {
                  $status.Text = 'Status: Excel refresh timeout. Ready with limited actions.'
                }
              }
            }
            catch {}
          }
          catch {}
        }).GetNewClosure())
  }
  try { $script:ExcelRefreshWatchdogTimer.Start() } catch {}
  $worker.Add_DoWork(({
        param($sender, $eventArgs)
        Write-UiTrace -Level 'INFO' -Message ("Excel refresh DoWork started. RunId={0}" -f $runId)
        Invoke-OnUiThread -Control $form -Action {
          try { if ($lblStartupStep -and -not $lblStartupStep.IsDisposed) { $lblStartupStep.Text = 'Refreshing Excel...' } } catch {}
        }
        $input = $eventArgs.Argument
        $eventArgs.Result = & $invokeRefreshHandler -ExcelPath $input.ExcelPath -Reason $input.Reason -ProcessingScope $input.ProcessingScope -PiSearchMode $input.PiSearchMode -MaxTickets $input.MaxTickets -FastMode $input.FastMode -TimeoutSeconds $input.TimeoutSeconds
      }).GetNewClosure())

  $worker.Add_RunWorkerCompleted(({
        param($sender, $eventArgs)
        try {
          if ($script:ActiveExcelRefreshRunId -ne $runId) {
            Write-UiTrace -Level 'WARN' -Message ("Excel refresh completion ignored (stale run id {0})." -f $runId)
            return
          }

          $script:ExcelRefreshInProgress = $false
          $script:ExcelRefreshWorker = $null
          $script:ActiveExcelRefreshRunId = -1
          $script:ExcelRefreshStartedAt = $null
          $script:ExcelRefreshTimeoutSeconds = 0
          try { if ($script:ExcelRefreshWatchdogTimer) { $script:ExcelRefreshWatchdogTimer.Stop() } } catch {}

          $res = $null
          if ($eventArgs -and $eventArgs.Result) { $res = $eventArgs.Result }
          if ($eventArgs -and $eventArgs.Error) {
            $res = @{
              Success   = $false
              Message   = ("" + $eventArgs.Error.Message).Trim()
              UpdatedAt = $null
              ExcelPath = $safeExcel
              Errors    = @((("" + $eventArgs.Error.Message).Trim()))
            }
          }
          if (-not $res) {
            $res = @{ Success = $false; Message = 'Refresh failed without details.'; UpdatedAt = $null; ExcelPath = $safeExcel; Errors = @('Refresh failed without details.') }
          }

          if ($res.Success) {
            $script:ExcelLoadErrorShown = $false
            $script:LastExcelRefreshAt = $res.UpdatedAt
            Save-MainExcelPathPreference -ExcelPath $res.ExcelPath
            Update-MainExcelDependentControls -ExcelReady $true
            try { if ($lblStartupStep -and -not $lblStartupStep.IsDisposed) { $lblStartupStep.Text = 'Loading modules...' } } catch {}
            if ($status -and -not $status.IsDisposed) {
              $status.Text = ("Status: Ready - Excel refreshed at {0:HH:mm:ss}" -f $res.UpdatedAt)
            }
            Write-UiTrace -Level 'INFO' -Message ("Excel refresh completed successfully. Startup={0}" -f $Startup)
          }
          else {
            Update-MainExcelDependentControls -ExcelReady $false -Reason ("Excel refresh failed: {0}. Please load Excel." -f $res.Message)
            Show-ExcelLoadErrorOnce -Message ("Excel refresh failed.`r`n`r`n{0}`r`n`r`nPlease click 'Load Excel...' and select a valid file." -f $res.Message)
            Write-UiTrace -Level 'ERROR' -Message ("Excel refresh failed. Startup={0}. Reason={1}" -f $Startup, $res.Message)
          }

          if ($ShowOverlay) { Update-StartupOverlay -Visible $false }
          if ($OnCompleted) { & $OnCompleted $res }
        }
        catch {
          try { Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel refresh failed. Please load Excel.' } catch {}
          try { if ($ShowOverlay) { Update-StartupOverlay -Visible $false } } catch {}
        }
      }).GetNewClosure())

  $workerInput = @{
    ExcelPath       = $safeExcel
    Reason          = $Reason
    ProcessingScope = $ProcessingScope
    PiSearchMode    = $PiSearchMode
    MaxTickets      = $MaxTickets
    FastMode        = $FastMode
    TimeoutSeconds  = $resolvedTimeoutSeconds
  }
  Write-UiTrace -Level 'INFO' -Message ("Excel refresh RunWorkerAsync starting. RunId={0}; TimeoutSec={1}" -f $runId, $resolvedTimeoutSeconds)
  $worker.RunWorkerAsync($workerInput)
  Write-UiTrace -Level 'INFO' -Message ("Excel refresh RunWorkerAsync dispatched. RunId={0}" -f $runId)
}

function Start-SchumanStartupInit {
  if ($script:StartupInitDone) { return }
  $script:StartupInitDone = $true

  Update-StartupOverlay -Visible $true -Title 'Schuman is starting...' -Step 'Checking updates...'
  if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Startup initialization in progress...' }
  Write-UiTrace -Level 'INFO' -Message 'Startup initialization started.'
  try {
    $excel = ("" + $txtExcel.Text).Trim()
    if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
      Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel file path is invalid. Please load Excel.'
      Write-UiTrace -Level 'ERROR' -Message 'Startup initialization failed: Excel file path is invalid.'
      return
    }

    $fs = $null
    try {
      $fs = [System.IO.File]::Open($excel, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    }
    finally {
      try { if ($fs) { $fs.Close(); $fs.Dispose() } } catch {}
    }

    Save-MainExcelPathPreference -ExcelPath $excel
    $script:ExcelLoadErrorShown = $false
    $script:LastExcelRefreshAt = [DateTime]::Now
    Update-MainExcelDependentControls -ExcelReady $true
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Ready' }
    Write-UiTrace -Level 'INFO' -Message 'Startup initialization completed successfully (quick Excel validation mode).'
  }
  catch {
    Update-MainExcelDependentControls -ExcelReady $false -Reason 'Excel not ready. Please load Excel.'
    Write-UiTrace -Level 'ERROR' -Message ("Startup initialization exception: " + $_.Exception.Message)
  }
  finally {
    try { Update-StartupOverlay -Visible $false } catch {}
  }
}

$script:DashboardForm = $null
$script:GeneratorForm = $null

function global:Open-SchumanDashboard {
  param()

  Invoke-UiSafe -Context 'Open Dashboard Window' -Action {
    if ($script:DashboardForm -and -not $script:DashboardForm.IsDisposed) {
      try {
        $script:DashboardForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $script:DashboardForm.Activate()
        $script:DashboardForm.BringToFront()
        return
      }
      catch {}
    }

    $excelForDashboard = ''
    try { $excelForDashboard = ("" + $txtExcel.Text).Trim() } catch { $excelForDashboard = '' }
    if (-not $excelForDashboard) { $excelForDashboard = '' }

    try {
      $dash = Resolve-UiForm -UiResult (New-DashboardUI -ExcelPath $excelForDashboard -SheetName $SheetName -Config $globalConfig -RunContext $uiRunContext -InitialSession $startupSession) -UiName 'New-DashboardUI'
      $script:DashboardForm = $dash
      [void]$dash.add_FormClosed(({
            $script:DashboardForm = $null
          }).GetNewClosure())
      [void]$dash.Show()
    }
    catch {
      $detail = ("" + $_.Exception.Message).Trim()
      $stack = ("" + $_.ScriptStackTrace).Trim()
      Write-Log -Level ERROR -Message ("Dashboard open failed: {0} | {1}" -f $detail, $stack)
      [System.Windows.Forms.MessageBox]::Show('Unable to open Dashboard. See log.', 'Schuman') | Out-Null
    }
  }
}

$openModule = {
  param([string]$module)

  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel)) {
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Excel still loading...' }
    $excel = ''
  }

  Set-AppBusyState -isBusy $true -reason ("Opening {0}" -f $module)
  try {
    if ($module -eq 'Dashboard') {
      Open-SchumanDashboard
    }
    else {
      if ($script:GeneratorForm -and -not $script:GeneratorForm.IsDisposed) {
        try {
          if ($script:GeneratorForm.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            $script:GeneratorForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
          }
          $script:GeneratorForm.Show()
          $script:GeneratorForm.BringToFront()
          $script:GeneratorForm.Activate()
          return
        }
        catch {}
      }

      $defaultTemplate = Join-Path $projectRoot $globalConfig.Documents.TemplateFile
      $defaultOutput = Join-Path $projectRoot $globalConfig.Documents.OutputFolder

      $frm = Resolve-UiForm -UiResult (New-GeneratePdfUI -ExcelPath $excel -SheetName $SheetName -TemplatePath $defaultTemplate -OutputPath $defaultOutput -OnOpenDashboard {
          Open-SchumanDashboard
        } -OnGenerate {
          param($argsObj)
          $selectedRows = @()
          try { $selectedRows = @($argsObj.RowNumbers | ForEach-Object { [int]$_ }) } catch { $selectedRows = @() }
          $okRun = Invoke-DocsGenerate -Excel $argsObj.ExcelPath -Sheet $SheetName -Template $argsObj.TemplatePath -Output $argsObj.OutputPath -ExportPdf:[bool]$argsObj.ExportPdf -RowNumbers $selectedRows
          if ($okRun) {
            return [pscustomobject]@{
              ok         = $true
              message    = 'Documents generated successfully.'
              outputPath = $argsObj.OutputPath
              generatedCount = @($selectedRows).Count
            }
          }
          return [pscustomobject]@{
            ok      = $false
            message = 'Document generation failed. Check logs under system/runs.'
            generatedCount = 0
          }
        } -OnCloseAll {
          param($uiObj)
          try { Shutdown-SchumanApp -CurrentForm $uiObj.Form } catch {}
        }) -UiName 'New-GeneratePdfUI'
      $script:GeneratorForm = $frm
      [void]$frm.add_FormClosed(({
            $script:GeneratorForm = $null
          }).GetNewClosure())
      [void]$frm.Show()
    }
  }
  finally {
    Set-AppBusyState -isBusy $false -reason 'Ready'
  }
}

$openDashboardAction = { & $openModule 'Dashboard' }.GetNewClosure()
$openGenerateAction = { & $openModule 'Generate' }.GetNewClosure()
$showForceUpdateOptionsDialogHandler = ${function:Show-ForceUpdateOptionsDialog}
$newFallbackForceUpdateOptionsDialogHandler = ${function:New-FallbackForceUpdateOptionsDialog}
$invokeForceExcelUpdateHandler = ${function:Invoke-ForceExcelUpdate}
$startSchumanStartupInitHandler = ${function:Start-SchumanStartupInit}
$updateMainExcelDependentControlsHandler = ${function:Update-MainExcelDependentControls}
$startSchumanExcelRefreshAsyncHandler = ${function:Start-SchumanExcelRefreshAsync}
$saveMainExcelPathPreferenceHandler = ${function:Save-MainExcelPathPreference}
$showExcelLoadErrorOnceHandler = ${function:Show-ExcelLoadErrorOnce}
$clearSchumanTempFilesHandler = ${function:Clear-SchumanTempFiles}
if (-not $startSchumanStartupInitHandler) { throw 'Startup initialization handler is not available.' }
if (-not $updateMainExcelDependentControlsHandler) { throw 'Excel readiness handler is not available.' }
if (-not $startSchumanExcelRefreshAsyncHandler) { throw 'Excel refresh handler is not available.' }
if (-not $invokeForceExcelUpdateHandler) { throw 'Force update handler is not available.' }
if (-not $saveMainExcelPathPreferenceHandler) { throw 'Excel preference save handler is not available.' }
if (-not $showExcelLoadErrorOnceHandler) { throw 'Excel load error handler is not available.' }
$closeAllAction = {
  try {
    Write-Log -Level INFO -Message 'CLICK: Close All from Main'
    Shutdown-SchumanApp -CurrentForm $form
  }
  catch {
    Show-UiError -Title 'Schuman' -Message 'Could not complete "Close All".' -Exception $_.Exception
  }
}.GetNewClosure()
$forceUpdateAction = {
  $excel = ("" + $txtExcel.Text).Trim()
  if ([string]::IsNullOrWhiteSpace($excel) -or -not (Test-Path -LiteralPath $excel)) {
    & $updateMainExcelDependentControlsHandler -ExcelReady $false -Reason 'Excel file path is invalid. Please load Excel.'
    & $showExcelLoadErrorOnceHandler -Message 'Excel file is missing or invalid. Please click "Load Excel..." and select a valid .xlsx file.'
    return
  }

  if ($script:ExcelRefreshInProgress) {
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Refresh already in progress...' }
    return
  }
  if ($script:ForceUpdateInProgress) {
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Force update already in progress...' }
    return
  }

  Set-AppBusyState -isBusy $true -reason 'Preparing force update options'
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
  }
  finally {
    Set-AppBusyState -isBusy $false -reason 'Ready'
  }

  if (-not $opts -or -not $opts.ok) {
    if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Force update canceled' }
    return
  }

  $script:ForceUpdateInProgress = $true
  $script:CancelRequested = $false
  Set-AppBusyState -isBusy $true -reason ("Refreshing Excel ({0}/{1})..." -f $opts.processingScope, $opts.piSearchMode)
  try {
    $forceResult = & $invokeForceExcelUpdateHandler -Excel $excel -Sheet $SheetName -ProcessingScope $opts.processingScope -PiSearchMode $opts.piSearchMode -MaxTickets ([int]$opts.maxTickets) -FastMode ([bool]$opts.fastMode)
    if ($forceResult -and [bool]$forceResult.cancelled) {
      $script:CancelRequested = $true
      if ($status -and -not $status.IsDisposed) {
        $status.Text = 'Status: Force Update cancelled. Using last loaded Excel data.'
      }
      Write-UiTrace -Level 'INFO' -Message 'Force update cancelled by user. Keeping last loaded Excel state.'
      return
    }
    if ($forceResult -and $forceResult.ok) {
      $script:ExcelLoadErrorShown = $false
      $script:LastExcelRefreshAt = [DateTime]::Now
      $rowsCount = 0
      try {
        $rowsCount = @(Search-DashboardRows -ExcelPath $excel -SheetName $SheetName -SearchText '').Count
      }
      catch { $rowsCount = 0 }
      if ($rowsCount -gt 0) {
        & $updateMainExcelDependentControlsHandler -ExcelReady $true -Reason 'Ready'
        if ($status -and -not $status.IsDisposed) { $status.Text = ("Status: Ready - Force update completed in {0}s" -f [int]$forceResult.durationSec) }
      }
      else {
        & $updateMainExcelDependentControlsHandler -ExcelReady $false -Reason 'Excel is empty'
        if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Excel is empty' }
      }
    }
    else {
      & $updateMainExcelDependentControlsHandler -ExcelReady $script:MainExcelReady -Reason 'Force update failed. Using last loaded Excel data.'
      $msg = if ($forceResult -and $forceResult.errorMessage) { ("" + $forceResult.errorMessage).Trim() } else { 'Force update failed without details.' }
      if ($status -and -not $status.IsDisposed) { $status.Text = ("Status: Force update failed. {0}" -f $msg) }
      Write-UiTrace -Level 'ERROR' -Message ("Force update failed: {0}" -f $msg)
    }
  }
  finally {
    Restore-UiAfterForceUpdate
    $script:CancelRequested = $false
  }
}.GetNewClosure()

& $updateMainExcelDependentControlsHandler -ExcelReady $false -Reason 'Excel not ready yet - loading...'

$txtExcel.Add_TextChanged(({
      if ($script:ExcelRefreshInProgress) { return }
      $candidate = ("" + $txtExcel.Text).Trim()
      if ($candidate -and (Test-Path -LiteralPath $candidate)) {
        & $updateMainExcelDependentControlsHandler -ExcelReady $false -Reason 'Excel path changed. Click "Load Excel..." to refresh.'
      }
      else {
        & $updateMainExcelDependentControlsHandler -ExcelReady $false -Reason 'Excel file path is invalid. Please load Excel.'
      }
    }).GetNewClosure())

$btnLoadExcel.Add_Click(({
      $pickExcelAction = {
        if ($script:ExcelRefreshInProgress) {
          if ($status -and -not $status.IsDisposed) { $status.Text = 'Status: Refresh already in progress...' }
          return
        }
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title = 'Select Excel file'
        $dlg.Filter = 'Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
        $dlg.CheckFileExists = $true
        $dlg.Multiselect = $false
        try {
          $current = ("" + $txtExcel.Text).Trim()
          if ($current -and (Test-Path -LiteralPath $current)) {
            $dlg.InitialDirectory = Split-Path -Parent $current
            $dlg.FileName = Split-Path -Leaf $current
          }
        }
        catch {}
        $res = $dlg.ShowDialog($form)
        if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $selected = ("" + $dlg.FileName).Trim()
        if (-not $selected) { return }
        $txtExcel.Text = $selected
        & $saveMainExcelPathPreferenceHandler -ExcelPath $selected
        & $startSchumanExcelRefreshAsyncHandler -ExcelPath $selected -Reason 'Refreshing Excel from selected file...' -ShowOverlay $true -Startup $false
      }
      Invoke-UiHandler -Context 'Load Excel' -Action $pickExcelAction
    }).GetNewClosure())

$btnDashboard.Add_Click(({
      Write-Log -Level INFO -Message 'CLICK: Open Dashboard from Main'
      Invoke-UiHandler -Context 'Open Dashboard' -Action $openDashboardAction
    }).GetNewClosure())
$btnGenerate.Add_Click(({
      Invoke-UiHandler -Context 'Open Generate' -Action $openGenerateAction
    }).GetNewClosure())
$btnCloseCode.Add_Click(({
      Invoke-UiHandler -Context 'Close All' -Action $closeAllAction
    }).GetNewClosure())
$btnForce.Add_Click(({
      Invoke-UiHandler -Context 'Force Update' -Action $forceUpdateAction
    }).GetNewClosure())
$btnCleanTemp.Add_Click(({
      $cleanTempAction = {
        Set-MainBusyState -IsBusy $true -Text 'Cleaning temp files...'
        try {
          $r = $null
          if ($clearSchumanTempFilesHandler) {
            $r = & $clearSchumanTempFilesHandler -ProjectRoot $projectRoot
          }
          else {
            $cleanupRoot = Join-Path $env:TEMP 'Schuman'
            $deletedFiles = 0
            $deletedFolders = 0
            if (Test-Path -LiteralPath $cleanupRoot) {
              foreach ($child in @(Get-ChildItem -LiteralPath $cleanupRoot -Force -ErrorAction SilentlyContinue)) {
                try {
                  if ($child.PSIsContainer) { Remove-Item -LiteralPath $child.FullName -Recurse -Force -ErrorAction Stop; $deletedFolders++ }
                  else { Remove-Item -LiteralPath $child.FullName -Force -ErrorAction Stop; $deletedFiles++ }
                } catch {}
              }
            }
            $r = [pscustomobject]@{
              DeletedFiles = $deletedFiles
              DeletedFolders = $deletedFolders
              FreedBytes = 0
              Skipped = 0
              Errors = @()
            }
          }
          $kbFreed = [math]::Round($r.FreedBytes / 1KB, 1)
          $errorCount = @($r.Errors).Count
          $summary = "Deleted files: $($r.DeletedFiles). Deleted folders: $($r.DeletedFolders). Skipped: $($r.Skipped). Errors: $errorCount. Freed: ${kbFreed} KB."
          if ($r.Errors.Count -gt 0) {
            $errList = ($r.Errors | Select-Object -First 5) -join "`r`n"
            $summary += "`r`n`r`nFirst errors:`r`n$errList"
          }
          Show-UiInfo -Title 'Clean Temp Files' -Message $summary
          if ($status -and -not $status.IsDisposed) {
            $status.Text = "Status: Temp cleaned - files=$($r.DeletedFiles), folders=$($r.DeletedFolders), skipped=$($r.Skipped), errors=$errorCount"
          }
        }
        finally {
          Set-MainBusyState -IsBusy $false
        }
      }
      Invoke-UiHandler -Context 'Clean Temp Files' -Action $cleanTempAction
    }).GetNewClosure())

[void]$form.add_Shown(({
      Invoke-UiHandler -Context 'StartupInit' -Action {
        & $startSchumanStartupInitHandler
      }
    }).GetNewClosure())


[void]$form.add_FormClosed(({
      try {
        if ($startupSession) { Close-ServiceNowSession -Session $startupSession }
      }
      catch {}
      try { Close-SchumanAllResources -Mode 'All' | Out-Null } catch {}
      try {
        if ($script:MainBusyTimer) {
          $script:MainBusyTimer.Stop()
          $script:MainBusyTimer.Dispose()
        }
      }
      catch {}
      try {
        if ($script:StartupOverlayAnimTimer) {
          $script:StartupOverlayAnimTimer.Stop()
          $script:StartupOverlayAnimTimer.Dispose()
        }
      }
      catch {}
      try {
        if ($script:ExcelRefreshWatchdogTimer) {
          $script:ExcelRefreshWatchdogTimer.Stop()
          $script:ExcelRefreshWatchdogTimer.Dispose()
        }
      }
      catch {}
    }).GetNewClosure())

[void]$form.add_FormClosing(({
      param($sender, $eventArgs)
      try {
        try { Close-SchumanAllResources -Mode 'All' | Out-Null } catch {}
        try { Close-SchumanOpenForms -Except $sender } catch {}
      }
      catch {}
    }).GetNewClosure())

[void]$form.ShowDialog()


