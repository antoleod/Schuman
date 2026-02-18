Set-StrictMode -Version Latest

function Get-TicketType {
  param([Parameter(Mandatory = $true)][string]$Ticket)
  if ($Ticket -match '^RITM\d{6,8}$') { return 'RITM' }
  if ($Ticket -match '^INC\d{6,8}$') { return 'INC' }
  if ($Ticket -match '^SCTASK\d{6,8}$') { return 'SCTASK' }
  return 'UNKNOWN'
}

function Test-ClosedState {
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

function Get-DetectedPiFromText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $patterns = @(
    '\b(?:02PI20[A-Z0-9_-]*|ITEC(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*|MUST(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*)\b',
    '\b(?:MUST|ITEC|EDPS|PRES)\s*[-_ ]?\s*BRUN\s*[-_ ]?\s*\d{6,}\b',
    '\b02\s*PI\s*20\s*\d{6,}\b'
  )

  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $vals = New-Object System.Collections.Generic.List[string]

  foreach ($pattern in $patterns) {
    $matches = [regex]::Matches($Text, $pattern, 'IgnoreCase')
    foreach ($m in $matches) {
      $v = ("" + $m.Value).Trim().ToUpperInvariant()
      $v = $v -replace '[\s-]+', ''
      if ($v -and $seen.Add($v)) { [void]$vals.Add($v) }
    }
  }

  return ($vals -join ', ')
}
