Set-StrictMode -Version Latest

if (-not (Get-Variable -Name UiLogPath -Scope Script -ErrorAction SilentlyContinue)) {
  $script:UiLogPath = Join-Path (Join-Path $env:TEMP 'Schuman') 'schuman-ui.log'
}
if (-not (Get-Variable -Name LogPath -Scope Script -ErrorAction SilentlyContinue)) {
  $script:LogPath = $script:UiLogPath
}
try {
  $uiLogDir = Split-Path -Parent $script:UiLogPath
  if ($uiLogDir -and -not (Test-Path -LiteralPath $uiLogDir)) {
    New-Item -ItemType Directory -Path $uiLogDir -Force | Out-Null
  }
}
catch {}

function global:Write-UiTrace {
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

function global:Show-UiError {
  param(
    [string]$Title = 'Schuman',
    [string]$Message = '',
    [System.Exception]$Exception = $null,
    [string]$Context = '',
    $ErrorRecord = $null,
    [switch]$Blocking
  )

  $safeTitle = if ([string]::IsNullOrWhiteSpace($Title)) { 'Schuman' } else { $Title }
  $safeMessage = ("" + $Message).Trim()
  if (-not $safeMessage -and $Context) { $safeMessage = ("{0} failed." -f $Context) }
  if (-not $safeMessage) { $safeMessage = 'An unexpected UI error occurred.' }

  $ex = $null
  if ($Exception) { $ex = $Exception }
  elseif ($ErrorRecord -and $ErrorRecord.Exception) { $ex = $ErrorRecord.Exception }

  $detail = New-Object System.Collections.Generic.List[string]
  if ($Context) { $detail.Add("Context: $Context") | Out-Null }
  $detail.Add("Message: $safeMessage") | Out-Null
  if ($ex) {
    try { $detail.Add("ExceptionType: $($ex.GetType().FullName)") | Out-Null } catch {}
    try { $detail.Add("ExceptionMessage: $($ex.Message)") | Out-Null } catch {}
    try {
      if ($ex.InnerException) {
        $detail.Add("InnerExceptionType: $($ex.InnerException.GetType().FullName)") | Out-Null
        $detail.Add("InnerExceptionMessage: $($ex.InnerException.Message)") | Out-Null
      }
    }
    catch {}
    try {
      $stack = ("" + $ex.StackTrace).Trim()
      if ($stack) { $detail.Add("StackTrace: $stack") | Out-Null }
    }
    catch {}
  }
  Write-UiTrace -Level 'ERROR' -Message ("{0} | {1}" -f $safeTitle, ($detail -join ' | '))

  if ($Blocking) {
    $uiText = $safeMessage
    if ($Context) { $uiText = ("{0}`r`n`r`nContext: {1}" -f $safeMessage, $Context) }
    try {
      [System.Windows.Forms.MessageBox]::Show($uiText, $safeTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
    catch {}
  }
}

function global:Invoke-SafeUiAction {
  param(
    [scriptblock]$Action,
    [string]$Context = 'UI Action',
    [string]$ActionName = 'UI Action',
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $ctx = ("" + $Context).Trim()
  if (-not $ctx) { $ctx = ("" + $ActionName).Trim() }
  if (-not $ctx) { $ctx = 'UI Action' }
  if (-not $Action) { return $null }

  try {
    return (& $Action)
  }
  catch {
    Show-UiError -Title 'Schuman' -Message ("{0} failed." -f $ctx) -Exception $_.Exception -Context $ctx
    return $null
  }
}

function global:Invoke-UiSafe {
  param(
    [scriptblock]$Action,
    [string]$Context = 'UI Action'
  )
  return (Invoke-SafeUiAction -Action $Action -Context $Context)
}

function global:Invoke-GenerateUiSafe {
  param(
    [hashtable]$UI = $null,
    [scriptblock]$Action,
    [string]$Context = 'Generate UI Action'
  )
  return (Invoke-SafeUiAction -Action $Action -Context $Context)
}

function global:Invoke-Resave {
  param(
    [scriptblock]$Action,
    [string]$Context = 'Resave Action'
  )
  return (Invoke-SafeUiAction -Action $Action -Context $Context)
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
          try { & $safeAction } catch { Show-UiError -Title 'Schuman' -Message 'UI callback failed.' -Exception $_.Exception -Context 'Invoke-OnUiThread callback' }
        }).GetNewClosure()
      if ($Synchronous) {
        [void]$Control.Invoke($invoker)
      }
      else {
        [void]$Control.BeginInvoke($invoker)
      }
      return
    }
    & $Action
  }
  catch {
    Show-UiError -Title 'Schuman' -Message 'UI thread dispatch failed.' -Exception $_.Exception -Context 'Invoke-OnUiThread'
  }
}

function global:Invoke-UiHandler {
  param(
    [string]$Context,
    [scriptblock]$Action
  )
  Invoke-SafeUiAction -Context $Context -Action $Action | Out-Null
}

if (-not (Get-Variable -Name WinFormsExceptionHandlingRegistered -Scope Script -ErrorAction SilentlyContinue)) {
  [bool]$script:WinFormsExceptionHandlingRegistered = $false
}
if (-not (Get-Variable -Name ThreadExceptionHandler -Scope Script -ErrorAction SilentlyContinue)) {
  $script:ThreadExceptionHandler = $null
}
if (-not (Get-Variable -Name DomainExceptionHandler -Scope Script -ErrorAction SilentlyContinue)) {
  $script:DomainExceptionHandler = $null
}

function global:Register-WinFormsGlobalExceptionHandling {
  if ($script:WinFormsExceptionHandlingRegistered) { return }
  try {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
  }
  catch {}

  $script:ThreadExceptionHandler = [System.Threading.ThreadExceptionEventHandler] {
    param($sender, $eventArgs)
    try {
      $ex = $null
      if ($eventArgs) { $ex = $eventArgs.Exception }
      Show-UiError -Title 'Schuman' -Message 'Unexpected UI error.' -Exception $ex -Context 'ThreadException'
    }
    catch {}
  }
  [System.Windows.Forms.Application]::add_ThreadException($script:ThreadExceptionHandler)

  $script:DomainExceptionHandler = [System.UnhandledExceptionEventHandler] {
    param($sender, $eventArgs)
    try {
      $ex = $eventArgs.ExceptionObject -as [System.Exception]
      Show-UiError -Title 'Schuman' -Message 'Unhandled application error.' -Exception $ex -Context 'AppDomain.UnhandledException'
    }
    catch {}
  }
  [AppDomain]::CurrentDomain.add_UnhandledException($script:DomainExceptionHandler)
  $script:WinFormsExceptionHandlingRegistered = $true
}
