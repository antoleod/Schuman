#Requires -Version 5.1
param(
  [string]$ExcelPath = (Join-Path $PSScriptRoot 'Schuman List.xlsx'),
  [string]$SheetName = 'BRU'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
  # WinForms must run in STA.
  if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    $argList = @(
      '-NoProfile'
      '-ExecutionPolicy'
      'Bypass'
      '-STA'
      '-File'
      ('"{0}"' -f $PSCommandPath)
      '-ExcelPath'
      ('"{0}"' -f $ExcelPath)
      '-SheetName'
      ('"{0}"' -f $SheetName)
    )
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argList | Out-Null
    return
  }

  $main = Join-Path $PSScriptRoot 'src\Schuman.Automation\Main.ps1'
  if (-not (Test-Path -LiteralPath $main)) {
    throw "Main UI not found: $main"
  }

  & $main -ExcelPath $ExcelPath -SheetName $SheetName
}
catch {
  $msg = if ($_.Exception) { $_.Exception.Message } else { '' }
  $stack = if ($_.ScriptStackTrace) { $_.ScriptStackTrace } else { '' }
  $line = "[{0}] Schuman-Main failed: {1}`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $msg, $stack
  $logPath = Join-Path (Join-Path $env:TEMP 'Schuman') 'schuman-main-error.log'
  try {
    $logDir = Split-Path -Parent $logPath
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
      New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
  }
  catch {}

  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show(
      ("Schuman no pudo iniciar.`r`n`r`n{0}`r`n`r`nLog: {1}" -f $msg, $logPath),
      'Schuman-Main',
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
  }
  catch {}
  throw
}
