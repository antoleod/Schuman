Set-StrictMode -Version Latest

function Get-UiFontName {
  $candidates = @('Segoe UI Variable Text', 'Segoe UI')
  $installed = New-Object System.Drawing.Text.InstalledFontCollection
  foreach ($name in $candidates) {
    if ($installed.Families.Name -contains $name) { return $name }
  }
  return 'Segoe UI'
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
