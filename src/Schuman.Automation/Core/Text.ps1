Set-StrictMode -Version Latest

function Get-TicketType {
  param([Parameter(Mandatory = $true)][string]$Ticket)
  if ($Ticket -match '^RITM\d{6,8}$') { return 'RITM' }
  if ($Ticket -match '^INC\d{6,8}$') { return 'INC' }
  if ($Ticket -match '^SCTASK\d{6,8}$') { return 'SCTASK' }
  return 'UNKNOWN'
}

function global:Test-ClosedState {
  param(
    [string]$StateLabel,
    [string]$StateValue
  )

  $label = ("" + $StateLabel).Trim().ToLowerInvariant()
  $value = ("" + $StateValue).Trim().ToLowerInvariant()

  if ($value -in @('3','4','7')) { return $true }
  if ($label -in @('3','4','7')) { return $true }
  if ($value -match 'closed|complete|resolved|cancel') { return $true }
  if ($label -match 'closed|complete|resolved|cancel') { return $true }
  if ($label -match 'cerrad|complet|resuelt|cancelad') { return $true }
  if ($label -match 'ferm|chius|annul') { return $true }
  return $false
}

function Get-CompletionStatus {
  param(
    [Parameter(Mandatory = $true)][string]$Ticket,
    [string]$StateLabel,
    [string]$StateValue,
    [int]$OpenTasks
  )

  if ($OpenTasks -gt 0) { return 'Pending' }
  if ($Ticket -like 'INC*' -and (("" + $StateLabel) -match '(?i)hold')) { return 'Hold' }

  if ($Ticket -like 'RITM*') {
    if ($StateValue -in @('3','4','7')) { return 'Complete' }
    return 'Pending'
  }

  if (Test-ClosedState -StateLabel $StateLabel -StateValue $StateValue) {
    return 'Complete'
  }

  return 'Pending'
}

function Get-DetectedPiFromActivityText {
  param([string]$ActivityText)
  if ([string]::IsNullOrWhiteSpace($ActivityText)) { return "" }

  $patterns = @(
    '\b(?:02PI20[A-Z0-9_-]*|ITEC(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*|MUST(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*|EDPSBRUN[A-Z0-9_-]*\d[A-Z0-9_-]*|PRESBRUN[A-Z0-9_-]*\d[A-Z0-9_-]*|[A-Z]{3,}BRUN[A-Z0-9_-]*\d[A-Z0-9_-]*)\b',
    '\b(?:MUST|ITEC|EDPS|PRES)\s*[-_ ]?\s*BRUN\s*[-_ ]?\s*\d{6,}\b',
    '\b02\s*PI\s*20\s*\d{6,}\b'
  )

  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $vals = New-Object System.Collections.Generic.List[string]
  foreach ($p in $patterns) {
    $matches = [regex]::Matches($ActivityText, $p, 'IgnoreCase')
    foreach ($m in $matches) {
      $v = ("" + $m.Value).Trim().ToUpperInvariant()
      if ($v -match '\s|-' ) { $v = ($v -replace '[\s-]+', '') }
      if ($v -notmatch '\d') { continue }
      if ($v -and $seen.Add($v)) {
        [void]$vals.Add($v)
      }
    }
  }

  if ($vals.Count -gt 1) {
    $valsArr = @($vals)
    $pi02 = @($valsArr | Where-Object { $_ -match '^02PI20' })
    if ($pi02.Count -gt 0) {
      $keep = New-Object System.Collections.Generic.List[string]
      foreach ($v in $valsArr) {
        $isDomainBrun = $v -match '^(MUST|ITEC|EDPS|PRES)[A-Z_]*BRUN'
        if (-not $isDomainBrun) { [void]$keep.Add($v); continue }
        $dDigits = ($v -replace '\D', '')
        $dropDomain = $false
        if ($dDigits) {
          foreach ($p2 in $pi02) {
            $pDigits = ($p2 -replace '\D', '')
            if ($pDigits -and $pDigits.Contains($dDigits)) { $dropDomain = $true; break }
          }
        }
        if (-not $dropDomain) { [void]$keep.Add($v) }
      }
      $vals = $keep
    }
  }

  if ($vals.Count -eq 0) { return "" }
  return ($vals -join ", ")
}

function Get-DetectedMachineHintFromText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

  $patterns = @(
    '(?im)\b(?:machine|device|computer|hostname|serial|asset|tag|pi)\b[^A-Za-z0-9]{0,30}([A-Z0-9][A-Z0-9_-]{5,})',
    '(?im)\b([A-Z]{3,}BRUN[0-9A-Z_-]{4,})\b',
    '(?im)\b([A-Z]{2,}[0-9]{6,})\b'
  )

  foreach ($p in $patterns) {
    $m = [regex]::Match($Text, $p)
    if (-not $m.Success -or $m.Groups.Count -lt 2) { continue }
    $v = ("" + $m.Groups[1].Value).Trim().ToUpperInvariant()
    if ($v -match '^(STATE|NUMBER|REQUEST|RITM|SCTASK|INC|TASK|USER|ARRIVAL|CLOSED|COMPLETE|FACILITIES|SERVICE|LOGISTICS|SUPPORT|DESK|LOCAL)$') { continue }
    if ($v -match '^(SCTASK|RITM|INC)\d{5,}$') { continue }
    if ($v.Length -lt 6) { continue }
    if ($v -notmatch '\d') { continue }
    return $v
  }
  return ""
}

function Get-FirstPiToken([string]$PiText) {
  if ([string]::IsNullOrWhiteSpace($PiText)) { return "" }
  $parts = @($PiText -split ',')
  foreach ($p in $parts) {
    $v = ("" + $p).Trim()
    if ($v) { return $v }
  }
  return ""
}

function Resolve-ConfidentPiFromSource {
  param(
    [string]$PiListText,
    [string]$SourceText
  )

  if ([string]::IsNullOrWhiteSpace($PiListText)) {
    return [pscustomobject]@{
      selected  = ""
      ambiguous = $false
      reason    = "empty"
    }
  }

  $candidates = @($PiListText -split ',' | ForEach-Object { ("" + $_).Trim() } | Where-Object { $_ })
  if ($candidates.Count -le 1) {
    return [pscustomobject]@{
      selected  = if ($candidates.Count -eq 1) { $candidates[0] } else { "" }
      ambiguous = $false
      reason    = "single"
    }
  }

  $src = "" + $SourceText
  if ([string]::IsNullOrWhiteSpace($src)) {
    return [pscustomobject]@{
      selected  = $candidates[0]
      ambiguous = $true
      reason    = "no_source_text"
    }
  }

  $scores = @{}
  foreach ($c in $candidates) {
    $scores[$c] = 0
    $rx = [regex]::Escape($c)
    $ms = [regex]::Matches($src, $rx, 'IgnoreCase')
    foreach ($m in $ms) {
      $start = [Math]::Max(0, $m.Index - 120)
      $len = [Math]::Min(240, $src.Length - $start)
      if ($len -le 0) { continue }
      $ctx = $src.Substring($start, $len)

      if ($ctx -match '(?i)prepare[\s\W_]*device|new[\s\W_]*user') { $scores[$c] += 4 }
      if ($ctx -match '(?i)\b(machine|device|hostname|serial|asset|tag|pi)\b') { $scores[$c] += 3 }
      if ($ctx -match '(?i)\b(assigned|delivered|configured|ready)\b') { $scores[$c] += 2 }

      if ($ctx -match '(?i)\b(old|previous|former|replaced|replace|returned|obsolete|wrong)\b') { $scores[$c] -= 4 }
      if ($ctx -match '(?i)\b(history|audit|closed complete)\b') { $scores[$c] -= 1 }
    }
  }

  $ordered = @($scores.GetEnumerator() | Sort-Object -Property Value -Descending)
  if ($ordered.Count -eq 0) {
    return [pscustomobject]@{
      selected  = $candidates[0]
      ambiguous = $true
      reason    = "no_scores"
    }
  }

  $best = "" + $ordered[0].Key
  $bestScore = [int]$ordered[0].Value
  $secondScore = if ($ordered.Count -gt 1) { [int]$ordered[1].Value } else { -999 }

  if (($bestScore -ge 3) -and (($bestScore - $secondScore) -ge 2)) {
    return [pscustomobject]@{
      selected  = $best
      ambiguous = $false
      reason    = "scored"
    }
  }

  return [pscustomobject]@{
    selected  = ($candidates -join ", ")
    ambiguous = $true
    reason    = "ambiguous"
  }
}

function Get-DetectedPiFromText {
  param([string]$Text)
  return Get-DetectedPiFromActivityText -ActivityText $Text
}
