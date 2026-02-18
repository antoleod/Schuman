Set-StrictMode -Version Latest

function New-RunContext {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [string]$RunName = 'run'
  )

  $systemRoot = $Config.Output.SystemRoot
  $runsDir = Join-Path $systemRoot $Config.Output.RunsSubdir
  $logsDir = Join-Path $systemRoot $Config.Output.LogsSubdir
  $dbDir = Join-Path $systemRoot $Config.Output.DbSubdir

  New-Item -ItemType Directory -Force -Path $systemRoot, $runsDir, $logsDir, $dbDir | Out-Null

  $runId = Get-Date -Format 'yyyyMMdd_HHmmss'
  $runDir = Join-Path $runsDir ("{0}_{1}" -f $RunName, $runId)
  New-Item -ItemType Directory -Force -Path $runDir | Out-Null

  $ctx = @{
    RunId = $runId
    RunDir = $runDir
    LogPath = Join-Path $runDir 'run.log.txt'
    HistoryLogPath = Join-Path $logsDir 'history.log'
  }

  return $ctx
}

function Write-RunLog {
  param(
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level,
    [Parameter(Mandatory = $true)][string]$Message
  )

  $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
  Write-Host $line
  try { Add-Content -LiteralPath $RunContext.LogPath -Value $line -Encoding UTF8 } catch {}
  try { Add-Content -LiteralPath $RunContext.HistoryLogPath -Value $line -Encoding UTF8 } catch {}
}

function Start-PerfStopwatch {
  return [System.Diagnostics.Stopwatch]::StartNew()
}

function Stop-PerfStopwatch {
  param(
    [Parameter(Mandatory = $true)][System.Diagnostics.Stopwatch]$Stopwatch,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$Label
  )

  $Stopwatch.Stop()
  $elapsedMs = [int64]$Stopwatch.Elapsed.TotalMilliseconds
  Write-RunLog -RunContext $RunContext -Level INFO -Message ("PERF {0} took {1}ms" -f $Label, $elapsedMs)
  return $elapsedMs
}
