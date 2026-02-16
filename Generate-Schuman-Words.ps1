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
  }
  Dark = @{
    Bg = [Drawing.Color]::FromArgb(20,20,22)
    Card = [Drawing.Color]::FromArgb(32,32,36)
    Text = [Drawing.Color]::FromArgb(240,240,240)
    Sub  = [Drawing.Color]::FromArgb(170,170,170)
    Border = [Drawing.Color]::FromArgb(55,55,60)
    Accent = [Drawing.Color]::FromArgb(10,132,255)
  }
}

# ----------------------------
# UI
# ----------------------------
$form = New-Object Windows.Forms.Form
$form.Text = "Schuman Word Generator"
$form.Size = New-Object Drawing.Size(1040, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
try {
  $prop = $form.GetType().GetProperty("DoubleBuffered", "NonPublic,Instance")
  if ($prop) { $prop.SetValue($form, $true, $null) }
} catch {
  # ignore if DoubleBuffered isn't available in this context
}
$form.Font = New-Object Drawing.Font("Segoe UI", 9)

$chkDark = New-Object Windows.Forms.CheckBox
$chkDark.Text = "Dark mode"
$chkDark.AutoSize = $true
$chkDark.Checked = $true

$header = New-Object Windows.Forms.Panel
$header.Location = New-Object Drawing.Point(14, 14)
$header.Size = New-Object Drawing.Size(996, 106)
$header.Padding = New-Object Windows.Forms.Padding(16,14,16,14)
$header.BorderStyle = "FixedSingle"
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
$lblStatus.Font = New-Object Drawing.Font("Segoe UI", 10)
$lblStatus.Text = "Ready."
$header.Controls.Add($lblStatus)

$lblCounters = New-Object Windows.Forms.Label
$lblCounters.AutoSize = $false
$lblCounters.Size = New-Object Drawing.Size(820, 18)
$lblCounters.Location = New-Object Drawing.Point(16, 76)
$lblCounters.Text = "Total: 0 | Saved: 0 | Skipped: 0 | Errors: 0"
$header.Controls.Add($lblCounters)

$chkDark.Location = New-Object Drawing.Point(870, 60)
$header.Controls.Add($chkDark)

$listCard = New-Object Windows.Forms.Panel
$listCard.Location = New-Object Drawing.Point(14, 132)
$listCard.Size = New-Object Drawing.Size(996, 496)
$listCard.Padding = New-Object Windows.Forms.Padding(14,14,14,14)
$listCard.BorderStyle = "FixedSingle"
$form.Controls.Add($listCard)

$panel = New-Object Windows.Forms.FlowLayoutPanel
$panel.Dock = "Fill"
$panel.AutoScroll = $true
$panel.WrapContents = $false
$panel.FlowDirection = "TopDown"
$listCard.Controls.Add($panel)

$footer = New-Object Windows.Forms.Panel
$footer.Location = New-Object Drawing.Point(14, 642)
$footer.Size = New-Object Drawing.Size(996, 46)
$footer.Padding = New-Object Windows.Forms.Padding(16,10,16,10)
$footer.BorderStyle = "FixedSingle"
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
$chkShowWord.Left = 430
$chkShowWord.Top = 5
$chkShowWord.Checked = $false
$footer.Controls.Add($chkShowWord)

$btnOpen.Add_Click({ if (Test-Path -LiteralPath $OutDir) { Start-Process explorer.exe $OutDir } })

function Apply-Theme {
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }
  $form.BackColor = $t.Bg
  foreach ($c in @($header,$listCard,$footer)) { $c.BackColor = $t.Card; $c.ForeColor = $t.Text }
  $panel.BackColor = $t.Card
  $lblTitle.ForeColor = $t.Text
  $lblStatus.ForeColor = $t.Sub
  $lblCounters.ForeColor = $t.Sub
  foreach ($b in @($btnStart,$btnStop,$btnOpen)) {
    $b.BackColor = $t.Card
    $b.ForeColor = $t.Text
    $b.FlatAppearance.BorderColor = $t.Border
    $b.FlatAppearance.BorderSize = 1
  }
  foreach ($c in @($chkDark,$chkShowWord)) { $c.ForeColor = $t.Sub }
}
$chkDark.Add_CheckedChanged({ Apply-Theme })
Apply-Theme

# ----------------------------
# Per-row card + shimmer (UI thread only)
# ----------------------------
$script:RowWidgets = @{}  # rowIndex -> object with shimmer panels
function New-RowCard([int]$Row, [string]$FileName) {
  $t = if ($chkDark.Checked) { $Theme.Dark } else { $Theme.Light }

  $card = New-Object Windows.Forms.Panel
  $card.Width = 940
  $card.Height = 64
  $card.Margin = New-Object Windows.Forms.Padding(6,6,6,6)
  $card.Padding = New-Object Windows.Forms.Padding(12,10,12,10)
  $card.BackColor = $t.Bg
  $card.BorderStyle = "FixedSingle"

  $lblMain = New-Object Windows.Forms.Label
  $lblMain.AutoSize = $false
  $lblMain.Width = 600
  $lblMain.Height = 18
  $lblMain.Text = "Row $Row - $FileName"
  $lblMain.ForeColor = $t.Text
  $lblMain.Location = New-Object Drawing.Point(10,10)
  $card.Controls.Add($lblMain)

  $lblSub = New-Object Windows.Forms.Label
  $lblSub.AutoSize = $false
  $lblSub.Width = 600
  $lblSub.Height = 18
  $lblSub.Text = "Pending"
  $lblSub.ForeColor = $t.Sub
  $lblSub.Location = New-Object Drawing.Point(10,32)
  $card.Controls.Add($lblSub)

  $barHost = New-Object Windows.Forms.Panel
  $barHost.Width = 300
  $barHost.Height = 8
  $barHost.Location = New-Object Drawing.Point(630, 26)
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
    Card=$card; Main=$lblMain; Sub=$lblSub; Host=$barHost; Fill=$fill; Running=$false
  }
}

$animTimer = New-Object Windows.Forms.Timer
$animTimer.Interval = 30
$animTimer.Add_Tick({
  foreach ($k in $script:RowWidgets.Keys) {
    $w = $script:RowWidgets[$k]
    if (-not $w.Running) { continue }
    $w.Fill.Left += 12
    if ($w.Fill.Left -gt $w.Host.Width) { $w.Fill.Left = -$w.Fill.Width }
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
  function Set-WordPlaceholderValue($Doc, [string]$Key, [string]$Value, [string]$LogPath, [string]$LogPrefix){
    $changed = $false
    $method = "NotFound"
    $replaceCount = 0

    function Replace-InRange($Range, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        $text = [string]$Range.Text
        if([string]::IsNullOrEmpty($text)){ return }
        $escaped = [regex]::Escape($Pattern)
        $regex = if($WholeWord){ "(?<!\\w)$escaped(?!\\w)" } else { $escaped }
        $m = [regex]::Matches($text, $regex)
        if($m.Count -gt 0){
          $Count.Value += $m.Count
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

    function Replace-InStoryRanges($DocRef, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      $storyTypes = @(1,6,7,9,10,11,12)
      foreach($stype in $storyTypes){
        try {
          $range = $DocRef.StoryRanges.Item($stype)
          while($range -ne $null){
            Replace-InRange -Range $range -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
            $range = $range.NextStoryRange
          }
        } catch {}
      }
    }

    function Replace-InTables($DocRef, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        foreach($tbl in $DocRef.Tables){
          foreach($cell in $tbl.Range.Cells){
            $r = $cell.Range
            if($r.End -gt $r.Start){ $r.End = $r.End - 1 }
            Replace-InRange -Range $r -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
          }
        }
      } catch {}
    }

    function Replace-InShapes($DocRef, [string]$Pattern, [string]$ReplaceValue, [bool]$WholeWord, [ref]$Count){
      try {
        foreach($sh in $DocRef.Shapes){
          try {
            if($sh.TextFrame -and $sh.TextFrame.HasText -eq -1){
              $r = $sh.TextFrame.TextRange
              Replace-InRange -Range $r -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
            }
          } catch {}
        }
      } catch {}
      try {
        foreach($sec in $DocRef.Sections){
          foreach($hf in @($sec.Headers, $sec.Footers)){
            foreach($item in @($hf.Item(1), $hf.Item(2), $hf.Item(3))){
              try {
                foreach($sh in $item.Shapes){
                  try {
                    if($sh.TextFrame -and $sh.TextFrame.HasText -eq -1){
                      $r = $sh.TextFrame.TextRange
                      Replace-InRange -Range $r -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord -Count $Count
                    }
                  } catch {}
                }
              } catch {}
            }
          }
        }
      } catch {}
      try {
        foreach($ish in $DocRef.InlineShapes){
          try {
            if($ish.TextEffect){ continue }
          } catch {}
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

    # Literal Find/Replace for plain text placeholder (exact key and token variants)
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
          Replace-InStoryRanges -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
          Replace-InTables -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
          Replace-InShapes -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $true -Count ([ref]$replaceCount)
        }
        if($replaceCount -eq 0){
          foreach($t in $tokens){
            Replace-InStoryRanges -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
            Replace-InTables -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
            Replace-InShapes -DocRef $Doc -Pattern $t -ReplaceValue $Value -WholeWord $false -Count ([ref]$replaceCount)
          }
        }
        if($replaceCount -gt 0){
          $changed = $true
          $method = "FindReplace"
        }
      } catch {}
    }

    WriteLog $LogPath ("${LogPrefix}: Set $Key -> $changed via $method (replacements=$replaceCount)")
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

  $excel=$null; $wb=$null; $sheet=$null; $word=$null; $doc=$null
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

    $SyncHash.Status = "Opening Excel"
    WriteLog $Config.LogPath "Opening Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
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
    WriteLog $Config.LogPath "Opening Word"
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Config.ShowWord
    $word.DisplayAlerts = 0
    try {
      $word.Options.ConfirmConversions = $false
      $word.Options.SaveNormalPrompt = $false
      $word.Options.BackgroundSave = $false
      $word.Options.AllowFastSave = $false
      $word.Options.UpdateLinksAtOpen = $false
    } catch {}
    try {
      # 3 = msoAutomationSecurityForceDisable (avoid macro prompts)
      $word.AutomationSecurity = 3
    } catch {}
    WriteLog $Config.LogPath "Word opened"

    $wdFormatDOCX = 16
    $wdFormatPDF = 17

    if(-not $templateInspected){
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
      $SyncHash.Status = "Saving $fileName"
      $SyncHash.UiEvents.Enqueue([pscustomobject]@{ Type="RowStart"; Row=$r; File=$fileName })
      WriteLog $Config.LogPath "Row ${r} start: $fileName"
      WriteLog $Config.LogPath "Row ${r} values: Name='$name' Ticket='$ticket' PI='$pi' Equipment='$equipment'"

      try {
        # overwrite target by copying template first (avoids SaveAs/SaveAs2 COM issues)
        if(Test-Path -LiteralPath $filePath){ Remove-Item -LiteralPath $filePath -Force -ErrorAction SilentlyContinue }
        Copy-Item -LiteralPath $Config.TemplatePath -Destination $filePath -Force

        WriteLog $Config.LogPath "Row ${r}: Opening copied doc"
        # Open existing file (not template) so we can Save() directly
        $doc = $word.Documents.Open($filePath, $false, $false, $false)
        WriteLog $Config.LogPath "Row ${r}: Doc opened"
        Log-DocPlaceholders -Doc $doc -LogPath $Config.LogPath -Prefix "Row ${r}"

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

        # Forced mapping (extend this hashtable to support additional placeholders)
        $forced = @{
          "FieldDisplayName"  = $name
          "FieldTicketNumber" = $ticket
          "FieldPINumber"     = $pi
          "FieldITEquipment"  = $equipment
        }

        foreach($key in $forced.Keys){
          [void](Set-WordPlaceholderValue -Doc $doc -Key $key -Value $forced[$key] -LogPath $Config.LogPath -LogPrefix "Row ${r}")
        }

        WriteLog $Config.LogPath "Row ${r}: Saving"
        $doc.Save()
        WriteLog $Config.LogPath "Row ${r}: Saved"

        # Export PDF
        $pdfPath = [System.IO.Path]::ChangeExtension([string]$filePath, ".pdf")
        try {
          $doc.ExportAsFixedFormat($pdfPath, $wdFormatPDF)
          WriteLog $Config.LogPath "Row ${r}: PDF saved -> $pdfPath"
          $sheet.Cells.Item($r,$statusCol).Value2 = "OK"
          $sheet.Cells.Item($r,$pdfCol).Value2 = $pdfPath
        } catch {
          WriteLog $Config.LogPath "Row ${r}: PDF export failed -> $($_.Exception.Message)"
          $sheet.Cells.Item($r,$statusCol).Value2 = "FAILED: PDF export"
          $sheet.Cells.Item($r,$pdfCol).Value2 = ""
        }

        $sheet.Cells.Item($r,$docxCol).Value2 = [string]$filePath

        $doc.Close($false)
        Release-Com $doc; $doc=$null

        $saved++
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
    try { if($excel){ $excel.Quit() | Out-Null } } catch {}
    try { if($word){ $word.Quit() | Out-Null } } catch {}

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
      }
      "Counters" {
        $lblCounters.Text = "Total: $($p.Total) | Saved: $($p.Saved) | Skipped: $($p.Skipped) | Errors: $($p.Errors)"
      }
      "RowStart" {
        if(-not $script:RowWidgets.ContainsKey($p.Row)){
          $script:RowWidgets[$p.Row] = New-RowCard -Row $p.Row -FileName $p.File
        }
        $w = $script:RowWidgets[$p.Row]
        $w.Sub.Text = "Saving..."
        $w.Running = $true
      }
      "RowDone" {
        $w = $script:RowWidgets[$p.Row]
        $w.Running = $false
        if($p.Ok){
          $w.Sub.Text = "Saved"
          $w.Fill.Left = 0
          $w.Fill.Width = $w.Host.Width
        } else {
          $w.Sub.Text = "ERROR: $($p.Error)"
          $w.Fill.Left = 0
          $w.Fill.Width = 60
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

    if($script:SyncHash.Error){
      $lblStatus.Text = "FAILED: " + $script:SyncHash.Error.Message
      [Windows.Forms.MessageBox]::Show($script:SyncHash.Error.Message, "Error") | Out-Null
    } else {
      $res = $script:SyncHash.Result
      $lblStatus.Text = "Completed. Saved=$($res.Saved) Skipped=$($res.Skipped) Errors=$($res.Errors)"
      [Windows.Forms.MessageBox]::Show($lblStatus.Text + "`r`n`r`nOutput: $OutDir", "Done") | Out-Null
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
  }) | Out-Null

  $script:PSInstance.BeginInvoke() | Out-Null

  $uiTimer.Start()
  $animTimer.Start()
  $statusTimer.Start()
})

$btnStop.Add_Click({
  $script:SyncHash.Cancel = $true
  $script:SyncHash.Status = "Stopping"
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
