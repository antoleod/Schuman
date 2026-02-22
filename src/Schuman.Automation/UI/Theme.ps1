Set-StrictMode -Version Latest

function global:Get-UiFontName {
  $candidates = @('Segoe UI Variable Text', 'Segoe UI')
  try {
    $installed = New-Object System.Drawing.Text.InstalledFontCollection
    foreach ($name in $candidates) {
      if ($installed.Families.Name -contains $name) { return $name }
    }
  } catch {}
  return 'Segoe UI'
}

function global:Stop-UiProcessesByName {
  param(
    [Parameter(Mandatory = $true)][string[]]$ExecutableNames
  )

  $targets = @($ExecutableNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim().ToLowerInvariant() })
  if ($targets.Count -eq 0) {
    return [pscustomobject]@{ Killed = @(); Failed = @() }
  }

  $killed = New-Object System.Collections.Generic.List[object]
  $failed = New-Object System.Collections.Generic.List[object]

  try {
    $procs = @(Get-CimInstance Win32_Process)
    foreach ($p in $procs) {
      if ([int]$p.ProcessId -eq $PID) { continue }
      $name = ("" + $p.Name).Trim().ToLowerInvariant()
      if (-not $name) { continue }
      if (-not ($targets -contains $name)) { continue }
      try {
        Stop-Process -Id ([int]$p.ProcessId) -Force -ErrorAction Stop
        $killed.Add([pscustomobject]@{ Id = [int]$p.ProcessId; Name = $p.Name }) | Out-Null
      } catch {
        $failed.Add([pscustomobject]@{ Id = [int]$p.ProcessId; Name = $p.Name; Error = $_.Exception.Message }) | Out-Null
      }
    }
  } catch {
    $failed.Add([pscustomobject]@{ Id = 0; Name = 'query'; Error = $_.Exception.Message }) | Out-Null
  }

  return [pscustomobject]@{ Killed = @($killed.ToArray()); Failed = @($failed.ToArray()) }
}

function global:Set-UiRoundedButton {
  param(
    [Parameter(Mandatory = $true)][System.Windows.Forms.Button]$Button,
    [int]$Radius = 10
  )

  if (-not $Button -or $Button.IsDisposed) { return }
  $safeRadius = [Math]::Max(2, [Math]::Min(30, $Radius))

  $applyRound = {
    param([System.Windows.Forms.Button]$btn, [int]$r)
    if (-not $btn -or $btn.IsDisposed) { return }
    $w = [Math]::Max(1, $btn.Width)
    $h = [Math]::Max(1, $btn.Height)
    $arc = [Math]::Min([Math]::Min($w, $h) - 1, $r * 2)
    if ($arc -lt 2) { return }

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    try {
      $path.AddArc(0, 0, $arc, $arc, 180, 90)
      $path.AddArc($w - $arc, 0, $arc, $arc, 270, 90)
      $path.AddArc($w - $arc, $h - $arc, $arc, $arc, 0, 90)
      $path.AddArc(0, $h - $arc, $arc, $arc, 90, 90)
      $path.CloseFigure()
      $oldRegion = $btn.Region
      $btn.Region = New-Object System.Drawing.Region($path)
      try { if ($oldRegion) { $oldRegion.Dispose() } } catch {}
    } finally {
      $path.Dispose()
    }
  }

  $Button.UseVisualStyleBackColor = $false
  & $applyRound $Button $safeRadius

  $meta = if ($Button.Tag -is [hashtable]) { $Button.Tag } else { @{} }
  $Button.Tag = $meta
  $meta['__roundRadius'] = $safeRadius

  if (-not $meta.ContainsKey('__roundedBound')) {
    $Button.Add_SizeChanged(({
      param($sender, $e)
      try {
        $m = if ($sender.Tag -is [hashtable]) { $sender.Tag } else { @{} }
        $r = 10
        if ($m.ContainsKey('__roundRadius')) { $r = [int]$m['__roundRadius'] }
        $fn = ${function:Set-UiRoundedButton}
        if ($fn) { & $fn -Button $sender -Radius $r }
      } catch {}
    }).GetNewClosure())
    $meta['__roundedBound'] = $true
  }
}

function global:Apply-UiRoundedButtonsRecursive {
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

function global:Invoke-UiEmergencyClose {
  param(
    [Parameter(Mandatory = $true)][string]$ActionLabel,
    [Parameter(Mandatory = $true)][string[]]$ExecutableNames,
    [System.Windows.Forms.IWin32Window]$Owner = $null,
    [ValidateSet('Code','Documents','All')][string]$Mode = 'All',
    [System.Windows.Forms.Form]$MainForm = $null,
    [string]$BaseDir = ''
  )

  $modeResolved = $Mode
  $labelText = ("" + $ActionLabel).Trim().ToLowerInvariant()
  if ($modeResolved -eq 'All') {
    if ($labelText -match 'code') { $modeResolved = 'Code' }
    elseif ($labelText -match 'document') { $modeResolved = 'Documents' }
  }

  $comClosed = 0
  $procClosed = 0
  $procFailed = 0

  try {
    $cleanupCmd = Get-Command -Name Stop-SchumanOwnedResources -ErrorAction SilentlyContinue
    if ($cleanupCmd) {
      $res = Stop-SchumanOwnedResources -Mode $modeResolved
      try { $comClosed = [int]$res.ComClosedCount } catch {}
      try { $procClosed = [int]$res.ProcessClosedCount } catch {}
      try { $procFailed = [int]$res.ProcessFailedCount } catch {}
    }
    elseif ($ExecutableNames -and $ExecutableNames.Count -gt 0) {
      $raw = Stop-UiProcessesByName -ExecutableNames $ExecutableNames
      $procClosed = @($raw.Killed).Count
      $procFailed = @($raw.Failed).Count
    }
  } catch {
    $procFailed++
  }

  if ($modeResolved -eq 'Code' -or $modeResolved -eq 'All') {
    try {
      if (Get-Command -Name Close-SchumanOpenForms -ErrorAction SilentlyContinue) {
        Close-SchumanOpenForms
      }
    } catch {}
    try { if ($MainForm -and -not $MainForm.IsDisposed) { $MainForm.Close() } } catch {}
  }

  return [pscustomobject]@{
    Cancelled = $false
    KilledCount = [int]$procClosed
    FailedCount = [int]$procFailed
    Message = ("Cleanup done (COM={0}, ProcClosed={1}, ProcFailed={2})" -f $comClosed, $procClosed, $procFailed)
  }
}

function global:New-CardContainer {
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

  return [pscustomobject]@{ Border = $border; Content = $content }
}

