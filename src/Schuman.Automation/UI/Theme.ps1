Set-StrictMode -Version Latest

function global:Get-UiFontName {
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
    if ([int]$p.ProcessId -eq $PID) { continue }
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

function Set-UiRoundedButton {
  param(
    [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
    [int]$Radius = 10
  )

  if (-not $Button -or $Button.IsDisposed) { return }
  $safeRadius = [Math]::Max(2, [Math]::Min(30, $Radius))
  $attach = {
    param($btn, [int]$r)
    if (-not $btn -or $btn.IsDisposed) { return }
    $w = [Math]::Max(1, $btn.Width)
    $h = [Math]::Max(1, $btn.Height)
    $arc = [Math]::Min([Math]::Min($w, $h) - 1, $r * 2)
    if ($arc -lt 2) { return }
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $arc, $arc, 180, 90)
    $path.AddArc($w - $arc, 0, $arc, $arc, 270, 90)
    $path.AddArc($w - $arc, $h - $arc, $arc, $arc, 0, 90)
    $path.AddArc(0, $h - $arc, $arc, $arc, 90, 90)
    $path.CloseFigure()
    $oldRegion = $btn.Region
    $btn.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
    try { if ($oldRegion) { $oldRegion.Dispose() } } catch {}
  }

  & $attach $Button $safeRadius
  $Button.Tag = if ($Button.Tag -is [hashtable]) { $Button.Tag } else { @{} }
  if (-not $Button.Tag.ContainsKey('__roundedBound')) {
    $Button.Add_SizeChanged(({
      param($sender, $e)
      try {
        $meta = if ($sender.Tag -is [hashtable]) { $sender.Tag } else { @{} }
        $r = 10
        if ($meta.ContainsKey('__roundRadius')) { $r = [int]$meta['__roundRadius'] }
        & $attach $sender $r
      } catch {}
    }).GetNewClosure())
    $Button.Tag['__roundedBound'] = $true
  }
  $Button.Tag['__roundRadius'] = $safeRadius
}

function Apply-UiRoundedButtonsRecursive {
  param(
    [Parameter(Mandatory = $true)][System.Windows.Forms.Control]$Root,
    [int]$Radius = 10
  )

  if (-not $Root -or $Root.IsDisposed) { return }
  foreach ($ctrl in $Root.Controls) {
    if ($ctrl -is [System.Windows.Forms.Button]) {
      Set-UiRoundedButton -Button $ctrl -Radius $Radius
    }
    if ($ctrl -and $ctrl.HasChildren) {
      Apply-UiRoundedButtonsRecursive -Root $ctrl -Radius $Radius
    }
  }
}

function Invoke-UiEmergencyClose {
  param(
    [Parameter(Mandatory = $true)][string]$ActionLabel,
    [Parameter(Mandatory = $true)][string[]]$ExecutableNames,
    [System.Windows.Forms.IWin32Window]$Owner = $null
  )

  $mode = 'All'
  $labelText = ("" + $ActionLabel).Trim().ToLowerInvariant()
  if ($labelText -match 'codigo') { $mode = 'Code' }
  elseif ($labelText -match 'document') { $mode = 'Documents' }

  $comClosed = 0
  $procClosed = 0
  $procFailed = 0
  try {
    $cleanupCmd = Get-Command -Name Stop-SchumanOwnedResources -ErrorAction SilentlyContinue
    if ($cleanupCmd) {
      $res = Stop-SchumanOwnedResources -Mode $mode
      try { $comClosed = [int]$res.ComClosedCount } catch {}
      try { $procClosed = [int]$res.ProcessClosedCount } catch {}
      try { $procFailed = [int]$res.ProcessFailedCount } catch {}
    }
  } catch {}

  return [pscustomobject]@{
    Cancelled = $false
    KilledCount = [int]$procClosed
    FailedCount = [int]$procFailed
    Message = ("Cleanup done (COM={0}, ProcClosed={1}, ProcFailed={2})" -f $comClosed, $procClosed, $procFailed)
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
