Set-StrictMode -Version Latest

function Get-UiFontName {
  $candidates = @('Segoe UI Variable Text', 'Segoe UI')
  $installed = New-Object System.Drawing.Text.InstalledFontCollection
  foreach ($name in $candidates) {
    if ($installed.Families.Name -contains $name) { return $name }
  }
  return 'Segoe UI'
}

function Stop-UiProcessesByName {
  param(
    [Parameter(Mandatory = $true)][string[]]$ExecutableNames
  )

  $targets = @($ExecutableNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToLowerInvariant() })
  if ($targets.Count -eq 0) {
    return [pscustomobject]@{
      Killed = @()
      Failed = @()
    }
  }

  $killed = New-Object System.Collections.Generic.List[object]
  $failed = New-Object System.Collections.Generic.List[object]
  $procs = @()
  try {
    $procs = @(Get-CimInstance Win32_Process)
  }
  catch {
    return [pscustomobject]@{
      Killed = @()
      Failed = @([pscustomobject]@{
          Id = 0
          Name = 'query'
          Error = $_.Exception.Message
        })
    }
  }

  foreach ($p in $procs) {
    $name = ("" + $p.Name).Trim().ToLowerInvariant()
    if (-not $name) { continue }
    if (-not ($targets -contains $name)) { continue }

    try {
      Stop-Process -Id ([int]$p.ProcessId) -Force -ErrorAction Stop
      $killed.Add([pscustomobject]@{
          Id = [int]$p.ProcessId
          Name = $p.Name
        }) | Out-Null
    }
    catch {
      $failed.Add([pscustomobject]@{
          Id = [int]$p.ProcessId
          Name = $p.Name
          Error = $_.Exception.Message
        }) | Out-Null
    }
  }

  return [pscustomobject]@{
    Killed = @($killed.ToArray())
    Failed = @($failed.ToArray())
  }
}

function Invoke-UiEmergencyClose {
  param(
    [Parameter(Mandatory = $true)][string]$ActionLabel,
    [Parameter(Mandatory = $true)][string[]]$ExecutableNames,
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $confirmText = "This will force-close: {0}`r`nUnsaved changes may be lost.`r`n`r`nContinue?" -f ($ExecutableNames -join ', ')
  $confirmResult = if ($Owner) {
    [System.Windows.Forms.MessageBox]::Show($Owner, $confirmText, $ActionLabel, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
  }
  else {
    [System.Windows.Forms.MessageBox]::Show($confirmText, $ActionLabel, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
  }
  if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
    return [pscustomobject]@{
      Cancelled = $true
      KilledCount = 0
      FailedCount = 0
      Message = 'Operation cancelled.'
    }
  }

  $result = Stop-UiProcessesByName -ExecutableNames $ExecutableNames
  $killedCount = @($result.Killed).Count
  $failedCount = @($result.Failed).Count

  if ($killedCount -eq 0 -and $failedCount -eq 0) {
    $message = 'No matching running process was found.'
  }
  elseif ($failedCount -eq 0) {
    $message = "Closed process(es): $killedCount"
  }
  else {
    $message = "Closed: $killedCount | Failed: $failedCount"
  }

  if ($Owner) {
    [System.Windows.Forms.MessageBox]::Show($Owner, $message, $ActionLabel, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  }
  else {
    [System.Windows.Forms.MessageBox]::Show($message, $ActionLabel, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  }

  return [pscustomobject]@{
    Cancelled = $false
    KilledCount = $killedCount
    FailedCount = $failedCount
    Message = $message
  }
}

function New-CardContainer {
  param(
    [Parameter(Mandatory = $true)][string]$Title,
    [int]$Padding = 16
  )

  $border = New-Object System.Windows.Forms.Panel
  $border.Dock = [System.Windows.Forms.DockStyle]::Top
  $border.BackColor = [System.Drawing.Color]::FromArgb(230, 230, 235)
  $border.Padding = New-Object System.Windows.Forms.Padding(1)
  $border.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
  $border.AutoSize = $true
  $border.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

  $inner = New-Object System.Windows.Forms.Panel
  $inner.Dock = [System.Windows.Forms.DockStyle]::Fill
  $inner.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
  $inner.Padding = New-Object System.Windows.Forms.Padding($Padding)
  $inner.AutoSize = $true
  $inner.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $border.Controls.Add($inner)

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = [System.Windows.Forms.DockStyle]::Top
  $layout.AutoSize = $true
  $layout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $layout.ColumnCount = 1
  $layout.RowCount = 2
  $layout.Margin = New-Object System.Windows.Forms.Padding(0)
  $layout.Padding = New-Object System.Windows.Forms.Padding(0)
  $inner.Controls.Add($layout)

  $titleLabel = New-Object System.Windows.Forms.Label
  $titleLabel.Text = $Title
  $titleLabel.AutoSize = $true
  $titleLabel.Font = New-Object System.Drawing.Font((Get-UiFontName), 11, [System.Drawing.FontStyle]::Bold)
  $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(28, 28, 30)
  $titleLabel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
  $layout.Controls.Add($titleLabel, 0, 0)

  $content = New-Object System.Windows.Forms.Panel
  $content.Dock = [System.Windows.Forms.DockStyle]::Top
  $content.AutoSize = $true
  $content.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
  $content.Margin = New-Object System.Windows.Forms.Padding(0)
  $layout.Controls.Add($content, 0, 1)

  return [pscustomobject]@{
    Border = $border
    Content = $content
  }
}
