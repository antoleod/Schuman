Set-StrictMode -Version Latest

function Resolve-ExcelPath {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [string]$ExcelPath,
    [string]$StartDirectory
  )

  if ($ExcelPath -and (Test-Path -LiteralPath $ExcelPath)) {
    return (Resolve-Path -LiteralPath $ExcelPath).Path
  }

  $candidateDir = if ($StartDirectory) { $StartDirectory } else { (Get-Location).Path }
  $candidate = Join-Path $candidateDir $Config.Excel.DefaultWorkbook
  if (Test-Path -LiteralPath $candidate) {
    return (Resolve-Path -LiteralPath $candidate).Path
  }

  Add-Type -AssemblyName System.Windows.Forms
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx'
  $dlg.FileName = $Config.Excel.DefaultWorkbook
  $dlg.InitialDirectory = $candidateDir
  $dlg.Title = 'Select Excel planning file'

  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    throw 'No Excel file selected.'
  }

  return $dlg.FileName
}

function Ensure-Directory {
  param([Parameter(Mandatory = $true)][string]$Path)
  New-Item -ItemType Directory -Force -Path $Path | Out-Null
  return $Path
}
