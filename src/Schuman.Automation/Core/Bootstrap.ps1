Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Initialize-SchumanEnvironment {
  param(
    [string]$ProjectRoot,
    [string]$ConfigPath = (Join-Path $ProjectRoot 'configs\appsettings.json')
  )

  if (-not (Test-Path -LiteralPath $ProjectRoot)) {
    throw "Project root not found: $ProjectRoot"
  }

  $cfg = @{
    ServiceNow = @{
      BaseUrl = 'https://europarl.service-now.com'
      LoginUrl = 'https://europarl.service-now.com/nav_to.do'
      WebViewProfileRoot = (Join-Path $env:LOCALAPPDATA 'Schuman\WebView2')
      JsonTimeoutMs = 12000
      QueryRetryCount = 3
      QueryRetryDelayMs = 800
    }
    Excel = @{
      DefaultWorkbook = 'Schuman List.xlsx'
      DefaultSheet = 'BRU'
      StopScanAfterEmptyRows = 50
      MaxRowsAfterFirstTicket = 300
    }
    Output = @{
      SystemRoot = (Join-Path $ProjectRoot 'system')
      RunsSubdir = 'runs'
      LogsSubdir = 'logs'
      DbSubdir = 'db'
    }
    Documents = @{
      TemplateFile = 'Reception_ITequipment.docx'
      OutputFolder = 'WORD files'
    }
  }

  if (Test-Path -LiteralPath $ConfigPath) {
    try {
      $json = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
      $cfg = Merge-Hashtable -Base $cfg -Overlay (ConvertTo-Hashtable $json)
    } catch {
      throw "Failed to parse config file '$ConfigPath': $($_.Exception.Message)"
    }
  }

  # Normalize configured paths for consistency.
  $cfg.ServiceNow.WebViewProfileRoot = Expand-ConfigPath -ProjectRoot $ProjectRoot -PathValue $cfg.ServiceNow.WebViewProfileRoot
  $cfg.Output.SystemRoot = Expand-ConfigPath -ProjectRoot $ProjectRoot -PathValue $cfg.Output.SystemRoot

  return $cfg
}

function Expand-ConfigPath {
  param(
    [Parameter(Mandatory = $true)][string]$ProjectRoot,
    [Parameter(Mandatory = $true)][string]$PathValue
  )

  $expanded = [Environment]::ExpandEnvironmentVariables($PathValue)
  if ([System.IO.Path]::IsPathRooted($expanded)) {
    return $expanded
  }

  return (Join-Path $ProjectRoot $expanded)
}

function ConvertTo-Hashtable {
  param([Parameter(Mandatory = $true)]$InputObject)

  if ($null -eq $InputObject) { return $null }

  if ($InputObject -is [System.Collections.IDictionary]) {
    $map = @{}
    foreach ($key in $InputObject.Keys) {
      $map[$key] = ConvertTo-Hashtable $InputObject[$key]
    }
    return $map
  }

  if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($item in $InputObject) {
      [void]$list.Add((ConvertTo-Hashtable $item))
    }
    return @($list)
  }

  if ($InputObject -is [pscustomobject]) {
    $map = @{}
    foreach ($p in $InputObject.PSObject.Properties) {
      $map[$p.Name] = ConvertTo-Hashtable $p.Value
    }
    return $map
  }

  return $InputObject
}

function Merge-Hashtable {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Base,
    [Parameter(Mandatory = $true)][hashtable]$Overlay
  )

  $result = @{}
  foreach ($k in $Base.Keys) { $result[$k] = $Base[$k] }

  foreach ($k in $Overlay.Keys) {
    if ($result.ContainsKey($k) -and ($result[$k] -is [hashtable]) -and ($Overlay[$k] -is [hashtable])) {
      $result[$k] = Merge-Hashtable -Base $result[$k] -Overlay $Overlay[$k]
    } else {
      $result[$k] = $Overlay[$k]
    }
  }

  return $result
}
