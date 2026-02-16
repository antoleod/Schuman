#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ----------------------------
# Config
# ----------------------------
$BaseDir      = (Get-Location).Path
$ExcelPath    = Join-Path $BaseDir "Schuman List.xlsx"
$TemplatePath = Join-Path $BaseDir "Reception_ITequipment.docx"
$OutDir       = Join-Path $BaseDir "WORD files"
$LogPath      = Join-Path $BaseDir "Generate-Schuman-Words.log"
$PreferredSheet = "BRU"

# ----------------------------
# Logging (never crash)
# ----------------------------
function Write-Log([string]$Message) {
  try {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $LogPath -Value "[$ts] $Message" -Encoding UTF8
  } catch {
    # swallow - log must never kill the app
  }
}

# ----------------------------
# Preflight (UI will also show message)
# ----------------------------
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
if (-not (Test-Path -LiteralPath $ExcelPath)) { Write-Log "Missing Excel: $ExcelPath" }
if (-not (Test-Path -LiteralPath $TemplatePath)) { Write-Log "Missing Template: $TemplatePath" }

# ----------------------------
# Theme
# ----------------------------
$Theme = @{
  Light = @{
    Bg = [Drawing.Color]::FromArgb(245,246,248)
    Card = [Drawing.Color]::White
    Text = [Drawing.Color]::FromArgb(20,20,20)
    Sub  = [Drawing.Color]::FromArgb(100,100,100)
    Border = [Drawing.Color]::FromArgb(228,228,232)
    Accent = [Drawing.Color]::FromArgb(0,122,255)
    Success = [Drawing.Color]::FromArgb(220,245,231)
    Error = [Drawing.Color]::FromArgb(255,230,230)
    BadgeText = [Drawing.Color]::FromArgb(30,30,30)
    Shadow = [Drawing.Color]::FromArgb(235,235,240)
  }
  Dark = @{
    Bg = [Drawing.Color]::FromArgb(20,20,22)
    Card = [Drawing.Color]::FromArgb(32,32,36)
    Text = [Drawing.Color]::FromArgb(240,240,240)
    Sub  = [Drawing.Color]::FromArgb(170,170,170)
    Border = [Drawing.Color]::FromArgb(55,55,60)
    Accent = [Drawing.Color]::FromArgb(10,132,255)
    Success = [Drawing.Color]::FromArgb(32,60,45)
    Error = [Drawing.Color]::FromArgb(70,36,36)
    BadgeText = [Drawing.Color]::FromArgb(230,230,230)
    Shadow = [Drawing.Color]::FromArgb(28,28,32)
  }
}

# ----------------------------
# UI
# ----------------------------
$form = New-Object Windows.Forms.Form
$form.Text = "Schuman Word Generator"
$form.Size = New-Object Drawing.Size(980, 640)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true
$form.MinimumSize = New-Object Drawing.Size(860, 520)
try {
  $prop = $form.GetType().GetProperty("DoubleBuffered", "NonPublic,Instance")
  if ($prop) { $prop.SetValue($form, $true, $null) }
} catch {
  # ignore if DoubleBuffered isn't available in this context
}
$form.Font = New-Object Drawing.Font("Segoe UI", 9)

$script:RowWidgets = @{}  # rowIndex -> object with shimmer panels


$chkDark = New-Object Windows.Forms.CheckBox
$chkDark.Text = "Dark mode"
$chkDark.AutoSize = $true
$chkDark.Checked = $true

$header = New-Object Windows.Forms.Panel
$header.Dock = "Top"
$header.Height = 130
$header.Padding = New-Object Windows.Forms.Padding(16,14,16,14)
$header.BorderStyle = "None"
$form.Controls.Add($header)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.AutoSize = $true
$lblTitle.Font = New-Object Drawing.Font("Segoe UI", 16, [Drawing.FontStyle]::Bold)
$lblTitle.Text = "Reception IT Equipment -> Word (.docx)"
$lblTitle.Location = New-Object Drawing.Point(16, 10)
$header.Controls.Add($lblTitle)

$lblStatus = New-Object Windows.Forms.Label
$lblStatus.AutoSize = $false
$lblStatus.Size = New-Object Drawing.Size(820, 22)
$lblStatus.Location = New-Object Drawing.Point(16, 52)
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 9)
$lblStatus.Text = "Ready."
$lblStatus.Anchor = "Top,Left,Right"
$header.Controls.Add($lblStatus)

$lblCounters = New-Object Windows.Forms.Label
$lblCounters.AutoSize = $false
$lblCounters.Size = New-Object Drawing.Size(820, 16)
$lblCounters.Location = New-Object Drawing.Point(16, 78)
$lblCounters.Text = "Total: 0 | Saved: 0 | Skipped: 0 | Errors: 0"
$lblCounters.Anchor = "Top,Left,Right"
$header.Controls.Add($lblCounters)

$lblElapsed = New-Object Windows.Forms.Label
$lblElapsed.AutoSize = $false
$lblElapsed.Size = New-Object Drawing.Size(260, 16)
$lblElapsed.Location = New-Object Drawing.Point(740, 78)
$lblElapsed.TextAlign = "MiddleRight"
$lblElapsed.Text = "Elapsed: 00:00"
$lblElapsed.Anchor = "Top,Right"
$header.Controls.Add($lblElapsed)

$lblAvg = New-Object Windows.Forms.Label
$lblAvg.AutoSize = $false
$lblAvg.Size = New-Object Drawing.Size(260, 16)
$lblAvg.Location = New-Object Drawing.Point(740, 58)
$lblAvg.TextAlign = "MiddleRight"
$lblAvg.Text = "Avg: --:--"
$lblAvg.Anchor = "Top,Right"
$header.Controls.Add($lblAvg)

$lblEta = New-Object Windows.Forms.Label
$lblEta.AutoSize = $false
$lblEta.Size = New-Object Drawing.Size(260, 16)
$lblEta.Location = New-Object Drawing.Point(740, 40)
$lblEta.TextAlign = "MiddleRight"
$lblEta.Text = "ETA: --:--"
$lblEta.Anchor = "Top,Right"
$header.Controls.Add($lblEta)

$progOverall = New-Object Windows.Forms.ProgressBar
$progOverall.Location = New-Object Drawing.Point(16, 100)
$progOverall.Size = New-Object Drawing.Size(940, 8)
$progOverall.Minimum = 0
$progOverall.Maximum = 100
$progOverall.Value = 0
$progOverall.Anchor = "Top,Left,Right"
$header.Controls.Add($progOverall)

$chkDark.Location = New-Object Drawing.Point(870, 60)
$header.Controls.Add($chkDark)

$listCard = New-Object Windows.Forms.Panel
$listCard.Dock = "Fill"
$listCard.Padding = New-Object Windows.Forms.Padding(14,14,14,14)
$listCard.BorderStyle = "None"
$form.Controls.Add($listCard)

$panel = New-Object Windows.Forms.FlowLayoutPanel
$panel.Dock = "Fill"
$panel.AutoScroll = $true
$panel.WrapContents = $false
$panel.FlowDirection = "TopDown"
$listCard.Controls.Add($panel)

$panel.Add_Resize({
  foreach($k in $script:RowWidgets.Keys){
    $w = $script:RowWidgets[$k]
    $w.Card.Width = [Math]::Max(600, $panel.ClientSize.Width - 24)
  }
})

$footer = New-Object Windows.Forms.Panel
$footer.Dock = "Bottom"
$footer.Height = 56
$footer.Padding = New-Object Windows.Forms.Padding(16,10,16,10)
$footer.BorderStyle = "None"
$form.Controls.Add($footer)

$btnStart = New-Object Windows.Forms.Button
$btnStart.Text = "Start"
$btnStart.Width = 110
$btnStart.Height = 26
$btnStart.FlatStyle = "Flat"
$footer.Controls.Add($btnStart)

$btnStop = New-Object Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Width = 110
$btnStop.Height = 26
$btnStop.Left = 120
$btnStop.FlatStyle = "Flat"
$btnStop.Enabled = $false
$footer.Controls.Add($btnStop)

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = "Open Output Folder"
$btnOpen.Width = 170
$btnOpen.Height = 26
$btnOpen.Left = 240
$btnOpen.FlatStyle = "Flat"
$footer.Controls.Add($btnOpen)

$chkShowWord = New-Object Windows.Forms.CheckBox
$chkShowWord.Text = "Show Word"
$chkShowWord.AutoSize = $true
$chkShowWord.Top = 5
$chkShowWord.Checked = $false
$footer.Controls.Add($chkShowWord)

$chkSaveDocx = New-Object Windows.Forms.CheckBox
$chkSaveDocx.Text = "Save DOCX"
$chkSaveDocx.AutoSize = $true
$chkSaveDocx.Top = 5
$chkSaveDocx.Checked = $true
$footer.Controls.Add($chkSaveDocx)

$chkSavePdf = New-Object Windows.Forms.CheckBox
$chkSavePdf.Text = "Save PDF"
$chkSavePdf.AutoSize = $true
$chkSavePdf.Top = 5
$chkSavePdf.Checked = $true
$footer.Controls.Add($chkSavePdf)

$script:UseUltra = $true
$script:UseFast = $true

$btnOpen.Add_Click({ if (Test-Path -LiteralPath $OutDir) { Start-Process explorer.exe $OutDir } })

function Layout-Footer {
  $x = 16
  $btnStart.Left = $x
  $btnStart.Top = 10
  $x += $btnStart.Width + 10

  $btnStop.Left = $x
  $btnStop.Top = 10
  $x += $btnStop.Width + 10

  $btnOpen.Left = $x
  $btnOpen.Top = 10

  $right = $footer.ClientSize.Width - 16
  foreach($c in @($chkSavePdf,$chkSaveDocx,$chkShowWord)){
    $c.Left = $right - $c.Width
    $c.Top = 12
    $right = $c.Left - 18
  }
}
$footer.Add_Resize({ Layout-Footer })

function Resize-For-Total([int]$Total){
  if($form.WindowState -ne "Normal"){ return }
  $rowHeight = 64
  $rowMargin = 12
  $visible = [Math]::Min([Math]::Max($Total,1), 8)
  $desired = $header.Height + $footer.Height + 40 + ($visible * ($rowHeight + $rowMargin))
  $screen = [Windows.Forms.Screen]::FromControl($form).WorkingArea
  $h = [Math]::Min($desired, $screen.Height - 40)
  $w = [Math]::Min($form.Width, $screen.Width - 40)
  $form.Size = New-Object Drawing.Size($w, $h)
}

function Apply-Theme {
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  $form.BackColor = $t.Bg
  foreach ($c in @($header,$listCard,$footer)) { $c.BackColor = $t.Card; $c.ForeColor = $t.Text }
  $panel.BackColor = $t.Card
  $lblTitle.ForeColor = $t.Text
  $lblStatus.ForeColor = $t.Sub
  $lblCounters.ForeColor = $t.Sub
  $lblElapsed.ForeColor = $t.Sub
  $lblAvg.ForeColor = $t.Sub
  $lblEta.ForeColor = $t.Sub
  foreach ($b in @($btnStart,$btnStop,$btnOpen)) {
    $b.BackColor = $t.Card
    $b.ForeColor = $t.Text
    $b.FlatAppearance.BorderColor = $t.Border
    $b.FlatAppearance.BorderSize = 1
  }
  foreach ($c in @($chkDark,$chkShowWord,$chkSaveDocx,$chkSavePdf)) { $c.ForeColor = $t.Sub }
  foreach ($k in $script:RowWidgets.Keys) {
    $w = $script:RowWidgets[$k]
    $w.Card.BackColor = $t.Bg
    $w.Main.ForeColor = $t.Text
    $w.Sub.ForeColor = $t.Sub
    $w.Host.BackColor = $t.Border
    $w.Fill.BackColor = $t.Accent
    if($w.Badge){ $w.Badge.ForeColor = $t.BadgeText }
  }
}
$chkDark.Add_CheckedChanged({ Apply-Theme })
Apply-Theme
Layout-Footer

# ----------------------------
# Per-row card + shimmer (UI thread only)
# ----------------------------
$script:RowWidgets = @{}  # rowIndex -> object with shimmer panels
function New-RowCard([int]$Row, [string]$FileName) {
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }

  $card = New-Object Windows.Forms.Panel
  $card.Width = [Math]::Max(600, $panel.ClientSize.Width - 24)
  $card.Height = 64
  $card.Margin = New-Object Windows.Forms.Padding(6,6,6,6)
  $card.Padding = New-Object Windows.Forms.Padding(12,10,12,10)
  $card.BackColor = $t.Bg
  $card.BorderStyle = "None"

  $lblMain = New-Object Windows.Forms.Label
  $lblMain.AutoSize = $false
  $lblMain.Width = 600
  $lblMain.Height = 18
  $lblMain.Text = "Row $Row - $FileName"
  $lblMain.ForeColor = $t.Text
  $lblMain.Location = New-Object Drawing.Point(10,10)
  $card.Controls.Add($lblMain)

  $lblBadge = New-Object Windows.Forms.Label
  $lblBadge.AutoSize = $false
  $lblBadge.Width = 80
  $lblBadge.Height = 18
  $lblBadge.TextAlign = "MiddleCenter"
  $lblBadge.Text = "Pending"
  $lblBadge.BackColor = $t.Bg
  $lblBadge.ForeColor = $t.BadgeText
  $lblBadge.Location = New-Object Drawing.Point(840,10)
  $card.Controls.Add($lblBadge)

  $lblSub = New-Object Windows.Forms.Label
  $lblSub.AutoSize = $false
  $lblSub.Width = 600
  $lblSub.Height = 18
  $lblSub.Text = "Pending"
  $lblSub.ForeColor = $t.Sub
  $lblSub.Location = New-Object Drawing.Point(10,32)
  $lblSub.Visible = $true
  $card.Controls.Add($lblSub)

  $barHost = New-Object Windows.Forms.Panel
  $barHost.Width = 300
  $barHost.Height = 8
  $barHost.Location = New-Object Drawing.Point(520, 28)
  $barHost.BackColor = $t.Border
  $card.Controls.Add($barHost)

  $fill = New-Object Windows.Forms.Panel
  $fill.Width = 70
  $fill.Height = 8
  $fill.Left = -70
  $fill.Top = 0
  $fill.BackColor = $t.Accent
  $barHost.Controls.Add($fill)

  $panel.Controls.Add($card)

  return [pscustomobject]@{
    Card=$card; Main=$lblMain; Sub=$lblSub; Host=$barHost; Fill=$fill; Badge=$lblBadge; Running=$false; File=$FileName
  }
}

function Apply-RowDensity {
  foreach($k in $script:RowWidgets.Keys){
    $w = $script:RowWidgets[$k]
    $w.Card.Height = 64
    $w.Sub.Visible = $true
    $w.Sub.Location = New-Object Drawing.Point(10,32)
    $w.Host.Location = New-Object Drawing.Point(520,28)
    if($w.Badge){ $w.Badge.Location = New-Object Drawing.Point(840,10) }
  }
}

$script:AnimPhase = 0.0
$animTimer = New-Object Windows.Forms.Timer
$animTimer.Interval = 20
$animTimer.Add_Tick({
  $script:AnimPhase = ($script:AnimPhase + 0.04)
  if($script:AnimPhase -gt 1.0){ $script:AnimPhase = 0.0 }
  foreach ($k in $script:RowWidgets.Keys) {
    $w = $script:RowWidgets[$k]
    if (-not $w.Running) { continue }
    $span = $w.Host.Width + $w.Fill.Width
    $w.Fill.Left = [int](($script:AnimPhase * $span) - $w.Fill.Width)
  }
})

# Global status dots
$script:Dots = 0
$statusTimer = New-Object Windows.Forms.Timer
$statusTimer.Interval = 250
$statusTimer.Add_Tick({
  if (-not $script:SyncHash.Running) { return }
  $script:Dots = ($script:Dots + 1) % 4
  $lblStatus.Text = "$($script:SyncHash.Status)" + ("." * $script:Dots)
})

# ----------------------------
# Worker communication
# ----------------------------
$script:SyncHash = [hashtable]::Synchronized(@{
  Running = $false
  Cancel  = $false
  Status  = "Ready"
  UiEvents = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
  Result = $null
  Error  = $null
})

$script:PSInstance = $null
$script:RunStarted = $null
$script:LastCounters = @{
  Total = 0
  Saved = 0
  Skipped = 0
  Errors = 0
}

# ----------------------------
# Worker logic (STA runspace) - SAFE PowerShell execution
# ----------------------------
$script:WorkerLogic = {
  param($SyncHash, $Config)

  function WriteLog($Path, $Msg) {
    try {
      $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
      Add-Content -LiteralPath $Path -Value "[$ts] $Msg" -Encoding UTF8
    } catch {}
  }
  function Release-Com($o){ if($o){ try{[Runtime.InteropServices.Marshal]::ReleaseComObject($o)|Out-Null}catch{} } }
  function Normalize([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $n = $s.Trim().ToLowerInvariant()
    $n = $n -replace '^field',''
    $n = $n -replace '[\s_\-]',''
    return $n
  }
  function Sanitize([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $s = $s -replace '[\\/:*?"<>|'']',''
    $s = $s -replace '\s+',' '
    return $s.Trim()
  }
  function Get-UniquePath([string]$Dir, [string]$BaseName){
    $base = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)
    $ext  = [System.IO.Path]::GetExtension($BaseName)
    if([string]::IsNullOrWhiteSpace($ext)){ $ext = ".docx" }
    $candidate = Join-Path $Dir ($base + $ext)
    $i = 2
    while(Test-Path -LiteralPath $candidate){
      $candidate = Join-Path $Dir ("$base ($i)$ext")
      $i++
    }
    return $candidate
  }
  function Inspect-Template($WordApp, [string]$TemplatePath, [string]$LogPath){
    $doc = $null
    try {
      WriteLog $LogPath "Template path: $TemplatePath"
      $doc = $WordApp.Documents.Open($TemplatePath, $false, $true, $false)
      WriteLog $LogPath "Template inspection: opened"

      try {
        $ffNames = @()
        foreach($ff in $doc.FormFields){ $ffNames += $ff.Name }
        WriteLog $LogPath ("Template FormFields (" + $ffNames.Count + "): " + ($ffNames -join ", "))
      } catch { WriteLog $LogPath "Template FormFields: <error reading>" }

      try {
        $ccNames = @()
        foreach($cc in $doc.ContentControls){
          $ccNames += ("Title='" + $cc.Title + "' Tag='" + $cc.Tag + "'")
        }
        WriteLog $LogPath ("Template ContentControls (" + $ccNames.Count + "): " + ($ccNames -join ", "))
      } catch { WriteLog $LogPath "Template ContentControls: <error reading>" }

      try {
        $bmNames = @()
        foreach($bm in $doc.Bookmarks){ $bmNames += $bm.Name }
        WriteLog $LogPath ("Template Bookmarks (" + $bmNames.Count + "): " + ($bmNames -join ", "))
      } catch { WriteLog $LogPath "Template Bookmarks: <error reading>" }

      try {
        $text = [string]$doc.Content.Text
        $matches = [regex]::Matches($text, '\bField[A-Za-z0-9_]+\b')
        $uniq = @($matches | ForEach-Object { $_.Value } | Sort-Object -Unique)
        if($uniq.Count -gt 0){
          WriteLog $LogPath ("Template Text Placeholders (" + $uniq.Count + "): " + ($uniq -join ", "))
        } else {
          WriteLog $LogPath "Template Text Placeholders: <none>"
        }
        $matches2 = [regex]::Matches($text, '(\{\{|\[|<<)\s*Field[A-Za-z0-9_]+\s*(\}\}|\]|>>)')
        $uniq2 = @($matches2 | ForEach-Object { $_.Value } | Sort-Object -Unique)
        if($uniq2.Count -gt 0){
          WriteLog $LogPath ("Template Token Placeholders (" + $uniq2.Count + "): " + ($uniq2 -join ", "))
        }
      } catch { WriteLog $LogPath "Template Text Placeholders: <error scanning>" }
    }
    catch {
      WriteLog $LogPath ("Template inspection failed: " + $_.Exception.Message)
    }
    finally {
      try { if($doc){ $doc.Close($false) | Out-Null } } catch {}
      Release-Com $doc
    }
  }
  function Log-DocPlaceholders($Doc, [string]$LogPath, [string]$Prefix){
    try {
      $text = [string]$Doc.Content.Text
      $matches = [regex]::Matches($text, '\bField[A-Za-z0-9_]+\b')
      $uniq = @($matches | ForEach-Object { $_.Value } | Sort-Object -Unique)
      if($uniq.Count -gt 0){
        WriteLog $LogPath ("${Prefix}: Doc Text Placeholders (" + $uniq.Count + "): " + ($uniq -join ", "))
      } else {
        WriteLog $LogPath "${Prefix}: Doc Text Placeholders: <none>"
      }
    } catch {
      WriteLog $LogPath "${Prefix}: Doc Text Placeholders: <error scanning>"
    }
  }
  function Set-WordPlaceholderValue($Doc, [string]$Key, [string]$Value, [string]$LogPath, [string]$LogPrefix, [bool]$FastMode){
    $changed = $false
    $method = "NotFound"
    $replaceCount = 0

    function Count-Occurrences([string]$Text, [string]$Pattern, [bool]$WholeWord){
      try {
        if([string]::IsNullOrEmpty($Text)){ return 0 }
        $escaped = [regex]::Escape($Pattern)
        $regex = if($WholeWord){ "(?<!\\w)$escaped(?!\\w)" } else { $escaped }
        return ([regex]::Matches($Text, $regex)).Count
      } catch { return 0 }
    }
    function Replace-InRangeFast($Range, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        $text = [string]$Range.Text
        $c = Count-Occurrences -Text $text -Pattern $Pattern -WholeWord $WholeWord
        if($c -gt 0){
          $Count.Value += $c
          $find = $Range.Find
          $find.ClearFormatting()
          $find.Replacement.ClearFormatting()
          $find.Text = $Pattern
          $find.Replacement.Text = $ReplaceValue
          $find.MatchCase = $false
          $find.MatchWholeWord = $WholeWord
          $find.MatchWildcards = $false
          $find.Wrap = 0
          $find.Forward = $true
          [void]$find.Execute($Pattern,$false,$false,$false,$false,$false,$true,1,$false,$ReplaceValue,2)
        }
      } catch {}
    }
    function Replace-InHeadersFooters($DocRef, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        foreach($sec in $DocRef.Sections){
          foreach($hf in @($sec.Headers, $sec.Footers)){
            foreach($item in @($hf.Item(1), $hf.Item(2), $hf.Item(3))){
              try {
                Replace-InRangeFast -Range $item.Range -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
              } catch {}
            }
          }
        }
      } catch {}
    }

    # ContentControls by exact Title/Tag
    try {
      $hits = 0
      foreach($cc in $Doc.ContentControls){
        if($cc.Title -eq $Key -or $cc.Tag -eq $Key){
          $cc.Range.Text = $Value
          $hits++
        }
      }
      if($hits -gt 0){
        $changed = $true
        $method = "ContentControl"
        $replaceCount = $hits
      }
    } catch {}

    # Bookmarks by exact name
    if(-not $changed){
      try {
        if($Doc.Bookmarks.Exists($Key)){
          $bm = $Doc.Bookmarks.Item($Key)
          $bm.Range.Text = $Value
          $Doc.Bookmarks.Add($Key, $bm.Range) | Out-Null
          $changed = $true
          $method = "Bookmark"
          $replaceCount = 1
        }
      } catch {}
    }

    # Legacy FormFields by exact name
    if(-not $changed){
      try {
        $ff = $Doc.FormFields.Item($Key)
        if($ff){
          $ff.Result = $Value
          $changed = $true
          $method = "FormField"
          $replaceCount = 1
        }
      } catch {}
    }

    # Literal Find/Replace for plain text placeholder (fast path)
    if(-not $changed){
      try {
        $tokens = @(
          $Key,
          "{{${Key}}}",
          "{${Key}}",
          "<<${Key}>>",
          "[${Key}]",
          "[[${Key}]]"
        )
        foreach($t in $tokens){
          Replace-InRangeFast -Range $Doc.Content -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
          Replace-InHeadersFooters -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
        }
        if($replaceCount -eq 0){
          foreach($t in $tokens){
            Replace-InRangeFast -Range $Doc.Content -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
            Replace-InHeadersFooters -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
          }
        }
        if($replaceCount -gt 0){
          $changed = $true
          $method = "FindReplace"
        }
      } catch {}
    }

    if(-not $FastMode){
      WriteLog $LogPath ("${LogPrefix}: Set $Key -> $changed via $method (replacements=$replaceCount)")
    }
    return [pscustomobject]@{
      Success = [bool]$changed
      MethodUsed = $method
      ReplaceCount = [int]$replaceCount
    }
  }
  function GetCol($Sheet, [string]$Header){
    $used = $Sheet.UsedRange
    $cols = $used.Columns.Count
    for($c=1;$c -le $cols;$c++){
      $h = ([string]$Sheet.Cells.Item(1,$c).Text).Trim()
      if($h -eq $Header){ return $c }
    }
    return $null
  }
  function Normalize-Header([string]$s){
    if([string]::IsNullOrWhiteSpace($s)){ return "" }
    $n = $s.ToLowerInvariant().Trim()
    $n = [regex]::Replace($n, '[^a-z0-9\s]', '')
    $n = [regex]::Replace($n, '\s+', ' ')
    return $n
  }
  function Find-ColumnByKeywords($Sheet, [string[]]$Keywords){
    $used = $Sheet.UsedRange
    $cols = $used.Columns.Count
    $bestScore = 0
    $bestCol = $null
    for($c=1;$c -le $cols;$c++){
      $h = Normalize-Header ([string]$Sheet.Cells.Item(1,$c).Text)
      if([string]::IsNullOrWhiteSpace($h)){ continue }
      $score = 0
      foreach($k in $Keywords){
        if($h -like "*$k*"){ $score++ }
      }
      if($score -gt $bestScore){
        $bestScore = $score
        $bestCol = $c
      }
    }
    return [pscustomobject]@{ Col=$bestCol; Score=$bestScore }
  }
  function Get-OrCreateColumn($Sheet, [string]$StandardHeader, [string[]]$Keywords){
    $res = Find-ColumnByKeywords $Sheet $Keywords
    if($res.Col -and $res.Score -ge 2){
      return [pscustomobject]@{ Col=$res.Col; Created=$false; Header=$StandardHeader; Score=$res.Score }
    }
    $used = $Sheet.UsedRange
    $newCol = $used.Columns.Count + 1
    $Sheet.Cells.Item(1,$newCol).Value2 = $StandardHeader
    return [pscustomobject]@{ Col=$newCol; Created=$true; Header=$StandardHeader; Score=0 }
  }
  function Get-ProtectionTypeName([int]$pt){
    switch($pt){
      0 { "wdNoProtection" }
      1 { "wdAllowOnlyRevisions" }
      2 { "wdAllowOnlyComments" }
      3 { "wdAllowOnlyFormFields" }
      4 { "wdAllowOnlyReading" }
      5 { "wdAllowOnlyFormFields" }
      default { "Unknown" }
    }
  }

  $excel=$null; $wb=$null; $sheet=$null; $word=$null; $doc=$null; $templateDoc=$null
  $excelCalc = $null
  $saved=0; $skipped=0; $errors=0; $total=0
  $templateInspected = $false

  try {
    if (-not (Test-Path -LiteralPath $Config.ExcelPath)) { throw "Missing Excel: $($Config.ExcelPath)" }
    if (-not (Test-Path -LiteralPath $Config.TemplatePath)) { throw "Missing template: $($Config.TemplatePath)" }
    New-Item -ItemType Directory -Force -Path $Config.OutDir | Out-Null

    # Unblock files if they came from the internet zone
    try { Unblock-File -LiteralPath $Config.ExcelPath -ErrorAction SilentlyContinue } catch {}
    try { Unblock-File -LiteralPath $Config.TemplatePath -ErrorAction SilentlyContinue } catch {}

    WriteLog $Config.LogPath "=== RUN START ==="

    $fast = [bool]($Config.FastMode -or $Config.TurboMode)

    $SyncHash.Status = "Opening Excel"
    WriteLog $Config.LogPath "Opening Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      $excel.ScreenUpdating = $false
      $excel.EnableEvents = $false
      $excelCalc = $excel.Calculation
      $excel.Calculation = -4135 # xlCalculationManual
    } catch {}
    # Open writable so we can update status/PDF path
    $wb = $excel.Workbooks.Open($Config.ExcelPath, $null, $false)
    WriteLog $Config.LogPath "Excel opened"

    try { $sheet = $wb.Worksheets.Item($Config.PreferredSheet) } catch { $sheet = $wb.Worksheets.Item(1) }
    WriteLog $Config.LogPath ("Using sheet: " + $sheet.Name)

    $nameCol   = GetCol $sheet "Name"
    $ticketCol = GetCol $sheet "Ticket"
    $piCol     = GetCol $sheet "PI"
    $equipCol  = GetCol $sheet "Receive ID Equipment" # optional

    $statusInfo = Get-OrCreateColumn $sheet "Export Status" @("status","export","result","done","ok")
    $docxInfo   = Get-OrCreateColumn $sheet "DOCX File"     @("docx","word","generated","output","file")
    $pdfInfo    = Get-OrCreateColumn $sheet "PDF File"      @("pdf","export","generated","output","file")

    $statusCol = $statusInfo.Col
    $docxCol   = $docxInfo.Col
    $pdfCol    = $pdfInfo.Col

    WriteLog $Config.LogPath ("Excel column status: Col=" + $statusCol + " Created=" + $statusInfo.Created + " Score=" + $statusInfo.Score)
    WriteLog $Config.LogPath ("Excel column docx:   Col=" + $docxCol + " Created=" + $docxInfo.Created + " Score=" + $docxInfo.Score)
    WriteLog $Config.LogPath ("Excel column pdf:    Col=" + $pdfCol + " Created=" + $pdfInfo.Created + " Score=" + $pdfInfo.Score)

    if(-not $nameCol -or -not $ticketCol -or -not $piCol){
      throw "Missing required headers. Required: Name, Ticket, PI."
    }

    $used = $sheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $total = [Math]::Max(0, $lastRow - 1)
    WriteLog $Config.LogPath ("Rows detected: " + $total)

    $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Init"; Total=$total })
    $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=0; Skipped=0; Errors=0 })

    $SyncHash.Status = "Opening Word"
    if(-not $fast){ WriteLog $Config.LogPath "Opening Word" }
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Config.ShowWord
    $word.DisplayAlerts = 0
    try {
      $word.ScreenUpdating = $false
      $word.Options.ConfirmConversions = $false
      $word.Options.SaveNormalPrompt = $false
      $word.Options.BackgroundSave = $false
      $word.Options.AllowFastSave = $false
      $word.Options.UpdateLinksAtOpen = $false
      $word.Options.CheckSpellingAsYouType = $false
      $word.Options.CheckGrammarAsYouType = $false
    } catch {}
    try {
      # 3 = msoAutomationSecurityForceDisable (avoid macro prompts)
      $word.AutomationSecurity = 3
    } catch {}
    if(-not $fast){ WriteLog $Config.LogPath "Word opened" }

    $wdFormatDOCX = 16
    $wdFormatPDF = 17

    if($Config.TurboMode){
      try {
        $templateDoc = $word.Documents.Open($Config.TemplatePath, $false, $true, $false)
        if(-not $fast){ WriteLog $Config.LogPath "Template opened in memory (Turbo)" }
      } catch {
        WriteLog $Config.LogPath ("Turbo template open failed: " + $_.Exception.Message)
        $templateDoc = $null
      }
    }

    if(-not $templateInspected -and -not $fast){
      Inspect-Template -WordApp $word -TemplatePath $Config.TemplatePath -LogPath $Config.LogPath
      $templateInspected = $true
    }

    for($r=2;$r -le $lastRow;$r++){
      if($SyncHash.Cancel){ break }

      $name   = [string]$sheet.Cells.Item($r,$nameCol).Text
      $ticket = [string]$sheet.Cells.Item($r,$ticketCol).Text
      $pi     = [string]$sheet.Cells.Item($r,$piCol).Text

      if([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($ticket) -and [string]::IsNullOrWhiteSpace($pi)){
        continue
      }
      if([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($ticket) -or [string]::IsNullOrWhiteSpace($pi)){
        $skipped++
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowSkip"; Row=$r; File="Skipped (missing data)" })
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors })
        continue
      }

      $equipment = "Laptop"
      if($equipCol){
        $tmp = [string]$sheet.Cells.Item($r,$equipCol).Text
        if(-not [string]::IsNullOrWhiteSpace($tmp)){ $equipment = $tmp.Trim() }
      }

      $safeTicket = Sanitize $ticket
      $safeName   = Sanitize $name
      if([string]::IsNullOrWhiteSpace($safeTicket)){ $safeTicket="UNKNOWN_TICKET" }
      if([string]::IsNullOrWhiteSpace($safeName)){ $safeName="UNKNOWN_NAME" }

      $fileName = "${safeTicket}_${safeName}.docx"
      $filePath = Get-UniquePath -Dir $Config.OutDir -BaseName $fileName

      # UI row start
      $rowStart = Get-Date
      $SyncHash.Status = "Saving $fileName"
      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStart"; Row=$r; File=$fileName })
      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Creating file..." })
      WriteLog $Config.LogPath "Row ${r} start: $fileName"
      if(-not $fast){
        WriteLog $Config.LogPath "Row ${r} values: Name='$name' Ticket='$ticket' PI='$pi' Equipment='$equipment'"
      }

      try {
        if($Config.TurboMode){
          $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Creating from template..." })
          $doc = $word.Documents.Add()
          if($templateDoc){
            try {
              $doc.Range.FormattedText = $templateDoc.Range.FormattedText
            } catch {
              $doc = $word.Documents.Add($Config.TemplatePath)
            }
          } else {
            $doc = $word.Documents.Add($Config.TemplatePath)
          }
          if(-not $fast){ WriteLog $Config.LogPath "Row ${r}: Doc created in memory" }
        } else {
          # overwrite target by copying template first (avoids SaveAs/SaveAs2 COM issues)
          if(Test-Path -LiteralPath $filePath){ Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue }
          Copy-Item -LiteralPath $Config.TemplatePath -Destination $filePath -Force

          WriteLog $Config.LogPath "Row ${r}: Opening copied doc"
          $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Opening file..." })
          # Open existing file (not template) so we can Save() directly
          $doc = $word.Documents.Open($filePath, $false, $false, $false)
          WriteLog $Config.LogPath "Row ${r}: Doc opened"
        }
        if($SyncHash.Cancel){
          try { $doc.Close($false) | Out-Null } catch {}
          Release-Com $doc; $doc=$null
          break
        }
        if(-not $fast){
          Log-DocPlaceholders -Doc $doc -LogPath $Config.LogPath -Prefix "Row ${r}"
        }

        if(-not $Config.UltraMode){
          try {
            $prot = $doc.ProtectionType
            $protName = Get-ProtectionTypeName $prot
            WriteLog $Config.LogPath "Row ${r}: ProtectionType=$prot ($protName)"
            if($prot -ne 0){
              WriteLog $Config.LogPath "Row ${r}: Document protected, attempting unprotect"
              try {
                $doc.Unprotect() | Out-Null
              } catch {
                WriteLog $Config.LogPath ("Row ${r}: Unprotect failed: " + $_.Exception.Message)
              }
              try {
                $prot2 = $doc.ProtectionType
                $prot2Name = Get-ProtectionTypeName $prot2
                WriteLog $Config.LogPath "Row ${r}: Unprotect attempted, ProtectionType now=$prot2 ($prot2Name)"
              } catch {
                WriteLog $Config.LogPath "Row ${r}: Unprotect attempted, re-check failed"
              }
            }
          } catch { WriteLog $Config.LogPath "Row ${r}: Protection check failed" }
        }

        # Forced mapping (extend this hashtable to support additional placeholders)
        $forced = @{
          "FieldDisplayName"  = $name
          "FieldTicketNumber" = $ticket
          "FieldPINumber"     = $pi
          "FieldITEquipment"  = $equipment
        }

        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Filling fields..." })
        foreach($key in $forced.Keys){
          if($SyncHash.Cancel){ break }
          [void](Set-WordPlaceholderValue -Doc $doc -Key $key -Value $forced[$key] -LogPath $Config.LogPath -LogPrefix "Row ${r}" -FastMode $fast)
        }
        if($SyncHash.Cancel){
          try { $doc.Close($false) | Out-Null } catch {}
          Release-Com $doc; $doc=$null
          break
        }

        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Saving..." })
        if($Config.ExportDocx){
          if($Config.TurboMode){
            $doc.SaveAs2([string]$filePath, $wdFormatDOCX)
          } else {
            WriteLog $Config.LogPath "Row ${r}: Saving"
            $doc.Save()
            WriteLog $Config.LogPath "Row ${r}: Saved"
          }
          $sheet.Cells.Item($r,$docxCol).Value2 = [string]$filePath
        } else {
          $sheet.Cells.Item($r,$docxCol).Value2 = ""
        }

        if($Config.ExportPdf){
          # Export PDF
          $pdfPath = [System.IO.Path]::ChangeExtension([string]$filePath, ".pdf")
          try {
            $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStage"; Row=$r; Stage="Exporting PDF..." })
            $doc.ExportAsFixedFormat($pdfPath, $wdFormatPDF)
            WriteLog $Config.LogPath "Row ${r}: PDF saved -> $pdfPath"
            $sheet.Cells.Item($r,$statusCol).Value2 = "OK"
            $sheet.Cells.Item($r,$pdfCol).Value2 = $pdfPath
          } catch {
            WriteLog $Config.LogPath "Row ${r}: PDF export failed -> $($_.Exception.Message)"
            $sheet.Cells.Item($r,$statusCol).Value2 = "FAILED: PDF export"
            $sheet.Cells.Item($r,$pdfCol).Value2 = ""
          }
        } else {
          if($sheet.Cells.Item($r,$statusCol).Value2 -ne "FAILED: PDF export"){
            $sheet.Cells.Item($r,$statusCol).Value2 = "OK"
          }
          $sheet.Cells.Item($r,$pdfCol).Value2 = ""
        }

        $doc.Close($false)
        Release-Com $doc; $doc=$null

        $saved++
        try {
          $ms = [int]((Get-Date) - $rowStart).TotalMilliseconds
        if(-not $fast){
          WriteLog $Config.LogPath ("Row ${r}: Done in ${ms} ms")
        }
        } catch {}
        WriteLog $Config.LogPath "Saved: $filePath"
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowDone"; Row=$r; File=$fileName; Ok=$true })
      }
      catch {
        $errors++
        WriteLog $Config.LogPath "Row $r ERROR: $($_.Exception.Message)"
        try {
          if($statusCol){ $sheet.Cells.Item($r,$statusCol).Value2 = ("FAILED: " + $_.Exception.Message) }
          if($docxCol){ $sheet.Cells.Item($r,$docxCol).Value2 = [string]$filePath }
          if($pdfCol){ $sheet.Cells.Item($r,$pdfCol).Value2 = "" }
        } catch {}
        try { if($doc){ $doc.Close($false) | Out-Null } } catch {}
        Release-Com $doc; $doc=$null
        $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowDone"; Row=$r; File=$fileName; Ok=$false; Error=$_.Exception.Message })
      }

      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="Counters"; Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors })
    }

    try {
      $wb.Save()
      WriteLog $Config.LogPath "Excel saved"
    } catch {
      WriteLog $Config.LogPath ("Excel save failed: " + $_.Exception.Message)
    }

    $SyncHash.Result = [pscustomobject]@{ Total=$total; Saved=$saved; Skipped=$skipped; Errors=$errors }
    WriteLog $Config.LogPath "=== RUN END === Total=$total Saved=$saved Skipped=$skipped Errors=$errors"
  }
  catch {
    $SyncHash.Error = $_.Exception
    WriteLog $Config.LogPath ("FATAL: " + $_.Exception.ToString())
  }
  finally {
    try { if($wb){ $wb.Close($false) | Out-Null } } catch {}
    try { if($excel -and $excelCalc -ne $null){ $excel.Calculation = $excelCalc } } catch {}
    try { if($excel){ $excel.Quit() | Out-Null } } catch {}
    try { if($word){ $word.Quit() | Out-Null } } catch {}

    Release-Com $templateDoc
    Release-Com $sheet
    Release-Com $wb
    Release-Com $excel
    Release-Com $word

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $SyncHash.Running = $false
  }
}

# ----------------------------
# UI Timer consumes events
# ----------------------------
$uiTimer = New-Object Windows.Forms.Timer
$uiTimer.Interval = 100
$uiTimer.Add_Tick({
  while($script:SyncHash.UiEvents.Count -gt 0){
    $p = $script:SyncHash.UiEvents.Dequeue()
    switch($p.Type){
      "Init" {
        $panel.Controls.Clear()
        $script:RowWidgets.Clear()
        $lblCounters.Text = "Total: $($p.Total) | Saved: 0 | Skipped: 0 | Errors: 0"
        if($p.Total -gt 0){
          $progOverall.Maximum = $p.Total
          $progOverall.Value = 0
        }
        Resize-For-Total -Total $p.Total
      }
      "Counters" {
        $lblCounters.Text = "Total: $($p.Total) | Saved: $($p.Saved) | Skipped: $($p.Skipped) | Errors: $($p.Errors)"
        $done = [int]$p.Saved + [int]$p.Skipped + [int]$p.Errors
        if($progOverall.Maximum -ne $p.Total){ $progOverall.Maximum = [Math]::Max(1, [int]$p.Total) }
        $progOverall.Value = [Math]::Min($progOverall.Maximum, $done)
        $script:LastCounters.Total = [int]$p.Total
        $script:LastCounters.Saved = [int]$p.Saved
        $script:LastCounters.Skipped = [int]$p.Skipped
        $script:LastCounters.Errors = [int]$p.Errors
      }
      "RowStart" {
        if(-not $script:RowWidgets.ContainsKey($p.Row)){
          $script:RowWidgets[$p.Row] = New-RowCard -Row $p.Row -FileName $p.File
        }
        $w = $script:RowWidgets[$p.Row]
        $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
        $w.Sub.Text = "Creating..."
        $w.Running = $true
        if($w.Badge){
          $w.Badge.Text = "Working"
          $w.Badge.BackColor = $t.Border
        }
      }
      "RowStage" {
        if($script:RowWidgets.ContainsKey($p.Row)){
          $w = $script:RowWidgets[$p.Row]
          $w.Sub.Text = $p.Stage
        }
      }
      "RowDone" {
        $w = $script:RowWidgets[$p.Row]
        $w.Running = $false
        $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
        if($p.Ok){
          $w.Sub.Text = "Done"
          $w.Fill.Left = 0
          $w.Fill.Width = $w.Host.Width
          $w.Card.BackColor = $t.Success
          if($w.Badge){
            $w.Badge.Text = "Done"
            $w.Badge.BackColor = $t.Success
          }
        } else {
          $w.Sub.Text = "ERROR: $($p.Error)"
          $w.Fill.Left = 0
          $w.Fill.Width = 60
          $w.Card.BackColor = $t.Error
          if($w.Badge){
            $w.Badge.Text = "Error"
            $w.Badge.BackColor = $t.Error
          }
        }
      }
      "RowSkip" {
        if(-not $script:RowWidgets.ContainsKey($p.Row)){
          $script:RowWidgets[$p.Row] = New-RowCard -Row $p.Row -FileName $p.File
        }
        $w = $script:RowWidgets[$p.Row]
        $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
        $w.Running = $false
        $w.Sub.Text = "Skipped"
        $w.Fill.Left = 0
        $w.Fill.Width = 60
        $w.Card.BackColor = $t.Border
        if($w.Badge){
          $w.Badge.Text = "Skipped"
          $w.Badge.BackColor = $t.Border
        }
      }
    }
  }

  if(-not $script:SyncHash.Running -and $script:PSInstance){
    $uiTimer.Stop()
    $animTimer.Stop()
    $statusTimer.Stop()
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
    $btnOpen.Enabled = $true

    $script:PSInstance.Dispose()
    $script:PSInstance = $null

    $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  foreach ($k in $script:RowWidgets.Keys) {
    $w = $script:RowWidgets[$k]
    if($w.Running){
      $w.Running = $false
      $w.Sub.Text = "Stopped"
      $w.Fill.Left = 0
      $w.Fill.Width = 60
      $w.Card.BackColor = $t.Border
      if($w.Badge){
        $w.Badge.Text = "Stopped"
        $w.Badge.BackColor = $t.Border
      }
    } elseif($w.Badge -and ($w.Badge.Text -eq "Pending" -or $w.Badge.Text -eq "Working")){
      $exists = $false
      try {
        $exists = Test-Path -LiteralPath (Join-Path $OutDir $w.File)
      } catch {}
      if($exists){
        $w.Badge.Text = "OK"
        $w.Badge.BackColor = $t.Success
        $w.Sub.Text = "Saved"
      } else {
        $w.Badge.Text = "Stopped"
        $w.Badge.BackColor = $t.Border
        $w.Sub.Text = "Stopped"
      }
    }
  }

    if($script:SyncHash.Error){
      $lblStatus.Text = "FAILED: " + $script:SyncHash.Error.Message
      [Windows.Forms.MessageBox]::Show($script:SyncHash.Error.Message, "Error") | Out-Null
    } else {
      $res = $script:SyncHash.Result
      $lblStatus.Text = "Completed. Saved=$($res.Saved) Skipped=$($res.Skipped) Errors=$($res.Errors)"
      [Windows.Forms.MessageBox]::Show($lblStatus.Text + "`r`n`r`nOutput: $OutDir", "Done") | Out-Null
    }
  }
  if($script:RunStarted){
    $elapsed = (Get-Date) - $script:RunStarted
    $lblElapsed.Text = ("Elapsed: " + $elapsed.ToString("mm\:ss"))
    $done = [int]$script:LastCounters.Saved + [int]$script:LastCounters.Skipped + [int]$script:LastCounters.Errors
    if($done -gt 0){
      $avgSec = [Math]::Max(1, [int]($elapsed.TotalSeconds / $done))
      $avgTs = [TimeSpan]::FromSeconds($avgSec)
      $lblAvg.Text = ("Avg: " + $avgTs.ToString("mm\:ss"))
      $remaining = [Math]::Max(0, ([int]$script:LastCounters.Total - $done))
      $etaTs = [TimeSpan]::FromSeconds($avgSec * $remaining)
      $lblEta.Text = ("ETA: " + $etaTs.ToString("mm\:ss"))
    } else {
      $lblAvg.Text = "Avg: --:--"
      if([int]$script:LastCounters.Total -gt 0){
        $lblEta.Text = "ETA: calculating..."
      } else {
        $lblEta.Text = "ETA: --:--"
      }
    }
  }
})

# ----------------------------
# Buttons
# ----------------------------
$btnStart.Add_Click({
  if(-not (Test-Path -LiteralPath $ExcelPath)){
    [Windows.Forms.MessageBox]::Show("Excel not found:`r`n$ExcelPath","Error") | Out-Null
    return
  }
  if(-not (Test-Path -LiteralPath $TemplatePath)){
    [Windows.Forms.MessageBox]::Show("Template not found:`r`n$TemplatePath","Error") | Out-Null
    return
  }
  if(-not $chkSaveDocx.Checked -and -not $chkSavePdf.Checked){
    [Windows.Forms.MessageBox]::Show("Select at least one output: DOCX or PDF.","Error") | Out-Null
    return
  }

  $btnStart.Enabled = $false
  $btnStop.Enabled = $true
  $btnOpen.Enabled = $false

  $panel.Controls.Clear()
  $script:RowWidgets.Clear()

  $script:SyncHash.Cancel = $false
  $script:SyncHash.Error = $null
  $script:SyncHash.Result = $null
  $script:SyncHash.Running = $true
  $script:SyncHash.Status = "Starting"
  $script:RunStarted = Get-Date

  Write-Log "User clicked Start."

  # STA runspace for Office COM
  $rs = [RunspaceFactory]::CreateRunspace()
  $rs.ApartmentState = "STA"
  $rs.ThreadOptions = "ReuseThread"
  $rs.Open()

  $script:PSInstance = [PowerShell]::Create()
  $script:PSInstance.Runspace = $rs
  $script:PSInstance.AddScript($script:WorkerLogic) | Out-Null
  $script:PSInstance.AddArgument($script:SyncHash) | Out-Null
  $script:PSInstance.AddArgument(@{
    ExcelPath = $ExcelPath
    TemplatePath = $TemplatePath
    OutDir = $OutDir
    LogPath = $LogPath
    PreferredSheet = $PreferredSheet
    ShowWord = $chkShowWord.Checked
    ExportPdf = $chkSavePdf.Checked
    ExportDocx = $chkSaveDocx.Checked
    FastMode = $script:UseFast
    TurboMode = $script:UseUltra
    UltraMode = $script:UseUltra
  }) | Out-Null

  $script:PSInstance.BeginInvoke() | Out-Null

  $uiTimer.Start()
  $animTimer.Start()
  $statusTimer.Start()
})

$btnStop.Add_Click({
  $script:SyncHash.Cancel = $true
  $script:SyncHash.Status = "Stopping"
  if($script:PSInstance){
    try { $script:PSInstance.Stop() | Out-Null } catch {}
  }
  Write-Log "User clicked Stop."
})

# Start timers idle-safe
$animTimer.Start()
$animTimer.Stop()
$statusTimer.Start()
$statusTimer.Stop()

Apply-Theme

# Start app
[Windows.Forms.Application]::EnableVisualStyles()
[Windows.Forms.Application]::Run($form)
