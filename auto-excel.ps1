# =============================================================================
# SNOW_ExportTicketsToJSON_UI.ps1  (PowerShell 5.1)
# =============================================================================
# PURPOSE
# -------
# Export ServiceNow ticket data to JSON (and optionally write back into Excel)
# WITHOUT using the official ServiceNow API.
#
# HOW IT WORKS (high-level)
# -------------------------
# 1) User selects an Excel file.
# 2) Script opens a WebView2 window for interactive SSO login (no password storage).
# 3) Script reads ticket numbers from Excel (INC/RITM/SCTASK).
# 4) For each ticket, it runs JavaScript inside the authenticated WebView2 session:
#      - Calls <table>.do?JSONv2 endpoints (internal SNOW endpoints, fast)
#      - Resolves user + CI display values (sys_user / cmdb_ci)
#      - Resolves state label via sys_choice
# 5) Writes one JSON per ticket + one combined JSON file.
# 6) Optionally writes results back to Excel columns (Name / New Phone / Action finished?).
#
# IMPORTANT NOTES
# ---------------
# - “NO API” here means: no token-based REST Table API usage, no ServiceNow API keys.
# - This script depends on your current SNOW permissions. If JSONv2 endpoints/fields are ACL-blocked,
#   extraction may fail or return empty fields.
# - WebView2 is loaded from Teams Meeting Add-in (corporate friendly, no install step).
# =============================================================================

param(
  [string]$ExcelPath,
  [string]$SheetName = "BRU",
  [string]$TicketHeader = "Number",
  [int]$TicketColumn = 4,
  [string]$NameHeader = "Name",
  [string]$PhoneHeader = "PI",
  [string]$ActionHeader = "Estado de RITM",
  [string]$DetectedPIHeader = "Detected PI / Machine",
  # >>> CHANGE DEFAULT EXCEL NAME HERE if the planning file is renamed again <<<
  [string]$DefaultExcelName = "Schuman List.xlsx",
  [string]$DefaultStartDir = $PSScriptRoot
)

# ----------------------------
# Global error behavior
# ----------------------------
# Stop on first unhandled error so we can reliably show a single error popup and log it.
$ErrorActionPreference = "Stop"

# ----------------------------
# WinForms assemblies
# ----------------------------
# Used for:
# - OpenFileDialog (pick Excel file)
# - Login window (WebView2 embedded)
# - MessageBox at end / errors
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# =============================================================================
# CONFIG
# =============================================================================
# ----------------------------
# ServiceNow base URL
# ----------------------------
# Change this if your SNOW instance differs.
$InstanceBaseUrl = "https://europarl.service-now.com"
$ServiceNowBase = $InstanceBaseUrl

# Consistent URL templates
$RitmRecordUrlTemplate         = "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_req_item.do%3Fsys_id%3D{0}%26sysparm_view%3D"
$SctaskRecordUrlTemplate       = "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D{0}"
$SctaskListByNumberUrlTemplate = "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_task_list.do%3Fsysparm_query%3Dnumber%3D{0}"

# Closed-state handling for SCTASK open/closed detection
$ClosedTaskStates = @(
  "Closed Complete",
  "Closed Incomplete",
  "Closed Skipped",
  "Complete",
  "Closed"
)

# Activity Stream PI/Machine extraction behavior
$EnableActivityStreamSearch = $true
$WriteNotFoundText = $false

# Entry point to begin SNOW navigation.
# nav_to.do is a safe starting URL that typically triggers the app shell after SSO.
$LoginUrl = "$ServiceNowBase/nav_to.do"

# ----------------------------
# Excel sheet + ticket column settings
# ----------------------------
# SheetName/TicketHeader/TicketColumn are provided via param().

# ----------------------------
# Excel write-back settings (autofill)
# ----------------------------
# If true, script will open Excel in write mode and fill empty cells with extracted values.
$WriteBackExcel = $true
$KillExcelBeforeOpen = $true
$StopScanAfterEmptyRows = 50
$MaxRowsAfterFirstTicket = 300
$ReadProgressEveryRows = 2000
$DebugActivityTicket = "RITM0427680"
$DebugActivityMaxChars = 4000

# NameHeader/PhoneHeader/ActionHeader are provided via param().

# ----------------------------
# Output folder
# ----------------------------
# Create a unique per-run folder under %TEMP% so multiple runs don't overwrite each other.
$RunId  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutDir = Join-Path $env:TEMP "SNOW_Tickets_Export_$RunId"

# Ensure output folder exists.
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# Log file path.
$LogPath = Join-Path $OutDir "run.log.txt"

# Combined JSON file path.
$AllJson = Join-Path $OutDir "tickets_export.json"

# =============================================================================
# LOGGING
# =============================================================================
function Log([string]$level, [string]$msg) {
  # Build a timestamped log line.
  $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$level] $msg"

  # Console output for visibility.
  Write-Host $line

  # Also write to run log file (best effort; do not crash if file write fails).
  try { Add-Content -Path $LogPath -Value $line } catch {}
}

function Close-ExcelProcessesIfRequested {
  param([string]$Reason = "")
  if (-not $KillExcelBeforeOpen) { return }
  try {
    $procs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    if ($procs.Count -eq 0) {
      Log "INFO" "No EXCEL process to kill. $Reason"
      return
    }
    Log "INFO" "Killing EXCEL processes: $($procs.Count). $Reason"
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 400
  }
  catch {
    Log "ERROR" "Failed to kill EXCEL processes. $Reason error='$($_.Exception.Message)'"
  }
}

Log "INFO" "Output folder: $OutDir"

# =============================================================================
# EXCEL FILE PICKER (UI)
# =============================================================================
function Pick-ExcelFile {
  param(
    [string]$ExcelPath,
    [string]$DefaultStartDir,
    [string]$DefaultExcelName
  )

  if ($ExcelPath -and (Test-Path -LiteralPath $ExcelPath)) {
    return $ExcelPath
  }

  $defaultCandidate = Join-Path $DefaultStartDir $DefaultExcelName
  if (Test-Path -LiteralPath $defaultCandidate) {
    return $defaultCandidate
  }

  # Windows file picker dialog
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
  $dlg.Title  = "Select Excel planning file"
  $dlg.Multiselect = $false
  $dlg.InitialDirectory = $DefaultStartDir
  $dlg.FileName = $DefaultExcelName

  # If user cancels, stop execution (this is a required input).
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    throw "No Excel file selected."
  }

  return $dlg.FileName
}

# =============================================================================
# WEBVIEW2 LOADER (from Teams Meeting Add-in)
# =============================================================================
function Load-WebView2FromTeams {
  Log "INFO" "Searching WebView2 DLLs in Teams Meeting Add-in..."

  # Determine architecture folder (x64 vs x86).
  $arch = if ($env:PROCESSOR_ARCHITECTURE -match "64") { "x64" } else { "x86" }

  # Teams meeting add-in location often contains WebView2 assemblies.
  $base = "$env:LOCALAPPDATA\Microsoft\TeamsMeetingAdd-in"

  # Select the newest version folder (Teams updates frequently).
  $dir = Get-ChildItem $base -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $dir) { throw "Teams Meeting Add-in not found." }

  # Load WebView2 assemblies from that folder.
  Add-Type -Path (Join-Path $dir.FullName "$arch\Microsoft.Web.WebView2.WinForms.dll")
  Add-Type -Path (Join-Path $dir.FullName "$arch\Microsoft.Web.WebView2.Core.dll")

  Log "INFO" "Loaded WebView2 from: $($dir.FullName)\$arch"
}

# =============================================================================
# WEBVIEW2 PROFILE (persistent cookies / SSO session)
# =============================================================================
# Use a stable UserDataFolder so SSO cookies persist across runs.
$script:WV2Profile = Join-Path $env:LOCALAPPDATA "SNOW_Export_UI\WebView2\$env:USERNAME"
New-Item -ItemType Directory -Force -Path $script:WV2Profile | Out-Null
Log "INFO" "Using WebView2 profile: $script:WV2Profile"

# =============================================================================
# WEBVIEW2 HELPERS
# =============================================================================
function New-ReadyWebView2($Parent) {
  # Create WebView2 control.
  $wv = New-Object Microsoft.Web.WebView2.WinForms.WebView2

  # Provide custom CreationProperties to set the persistent profile path.
  $props = New-Object Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
  $props.UserDataFolder = $script:WV2Profile
  $wv.CreationProperties = $props

  # Fill parent container.
  $wv.Dock = "Fill"
  $Parent.Controls.Add($wv)

  return $wv
}

function Ensure-WV2($wv) {
  # Ensure CoreWebView2 is created (async).
  # We pump WinForms events in a loop to avoid freezing the UI.
  $t = $wv.EnsureCoreWebView2Async()
  while (-not $t.IsCompleted) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 50
  }
  if ($t.IsFaulted) { throw $t.Exception.InnerException }
}

function ExecJS($wv, [string]$Js, [int]$TimeoutMs = 12000) {
  # Execute JS in current WebView2 document context.
  # Return the raw result string from WebView2 (often JSON-escaped).
  $task = $wv.ExecuteScriptAsync($Js)

  # Wait with timeout while processing WinForms events.
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while (-not $task.IsCompleted -and $sw.ElapsedMilliseconds -lt $TimeoutMs) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 50
  }

  # If not completed, faulted, or timed out -> return null (caller handles).
  if (-not $task.IsCompleted -or $task.IsFaulted) { return $null }

  return $task.GetAwaiter().GetResult()
}

function Parse-WV2Json([string]$Raw) {
  # WebView2 often returns JSON as a quoted string containing JSON (double-encoded).
  # This function tries to decode it safely.
  if ([string]::IsNullOrWhiteSpace($Raw) -or $Raw -eq 'null') { return $null }
  try {
    $o = $Raw | ConvertFrom-Json
    if ($o -is [string]) { $o = $o | ConvertFrom-Json }
    return $o
  }
  catch {
    return $null
  }
}

# =============================================================================
# LOGIN SSO (interactive WebView2 window)
# =============================================================================
function Connect-ServiceNowSSO {
  param([string]$StartUrl)

  Log "INFO" "Starting SSO login..."

  # Build a WinForms window to host WebView2 for interactive login.
  $form = New-Object Windows.Forms.Form
  $form.Text = "ServiceNow Login (SSO) - complete login then wait for GREEN status"
  $form.Size = New-Object Drawing.Size(1000, 720)
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  # Status label (top). Shows current URL/title and color indicates login state.
  $lbl = New-Object Windows.Forms.Label
  $lbl.Dock = "Top"
  $lbl.Height = 60
  $lbl.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
  $lbl.Text = "Loading..."
  $form.Controls.Add($lbl)

  # Panel to host WebView2.
  $panel = New-Object Windows.Forms.Panel
  $panel.Dock = "Fill"
  $form.Controls.Add($panel)

  # Create WebView2 (with persistent profile).
  $wv = New-ReadyWebView2 $panel

  # Show form and initialize WV2 core.
  $form.Show() | Out-Null
  Ensure-WV2 $wv

  # Navigate to start URL (SNOW nav_to.do).
  $wv.Source = $StartUrl

  # Poll login state for up to 240 seconds.
  $ok = $false
  $sw = [Diagnostics.Stopwatch]::StartNew()

  while ($form.Visible -and -not $ok -and $sw.Elapsed.TotalSeconds -lt 240) {

    # JS logic: detect whether we're on IdP, on SNOW, and “logged in” indicators.
    $js = @"
(function(){
  try{
    var href = location.href || "";
    var title = document.title || "";
    var host = "";
    try{ host = (new URL(href)).host; }catch(e){}
    var isIdp = /idp-lux\.extranet\.ep\.europa\.eu/i.test(host) || /F5Networks/i.test(href);
    var isSnow = /service-now\.com/i.test(host);
    var hasLogin = !!document.querySelector('form#login,input#user_name,input#username,input[type=password]');
    var hasNOW = (typeof window.NOW !== 'undefined') || (typeof window.g_user !== 'undefined');
    var domLogged = !!document.querySelector('sn-polaris-layout, now-global-nav, sn-appshell-root, now-avatar, [aria-label*="profile" i], [aria-label*="user" i]');
    var logged = isSnow && !hasLogin && (hasNOW || domLogged);
    return JSON.stringify({href:href,title:title,host:host,isIdp:isIdp,isSnow:isSnow,logged:logged,hasLogin:hasLogin});
  }catch(e){ return JSON.stringify({error:""+e}); }
})();
"@

    # Execute JS and parse its JSON output.
    $o = Parse-WV2Json (ExecJS $wv $js 4000)
    if ($o) {
      $lbl.Text = "URL: $($o.href)`r`nTITLE: $($o.title)"

      # Color logic:
      # - Red: at Identity Provider page
      # - Green: logged in to SNOW
      # - Orange: in-between / still loading / not logged
      if ($o.isIdp -eq $true) {
        $lbl.ForeColor = [Drawing.Color]::Red
      }
      elseif ($o.logged -eq $true) {
        $lbl.ForeColor = [Drawing.Color]::Green
        $ok = $true
      }
      else {
        $lbl.ForeColor = [Drawing.Color]::DarkOrange
      }
    }

    Start-Sleep -Milliseconds 250
  }

  # If login not confirmed, close and abort.
  if (-not $ok) {
    try { $form.Close() } catch {}
    throw "SSO not confirmed (timeout/closed)."
  }

  # Login confirmed: keep WV2 alive for later calls.
  Log "INFO" "SSO confirmed. Reusing same WebView2 instance."

  # Hide the form (we still keep the WV2 session).
  $null = $form.Hide()

  return @{ Form = $form; Wv = $wv }
}

# =============================================================================
# EXCEL: READ TICKETS
# =============================================================================
function Read-TicketsFromExcel {
  param(
    [string]$ExcelPath,
    [string]$TicketHeader,
    [string]$SheetName,
    [int]$TicketColumn
  )

  Log "INFO" "Opening Excel: $ExcelPath"
  Close-ExcelProcessesIfRequested -Reason "before read open"
  Log "INFO" "Creating Excel COM (read)..."

  # Excel COM object: keep invisible.
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  try { $excel.AskToUpdateLinks = $false } catch {}
  try { $excel.EnableEvents = $false } catch {}

  # Open workbook read-only to safely read tickets without locking for edits.
  Log "INFO" "Excel COM created (read), opening workbook..."
  $wb = $excel.Workbooks.Open($ExcelPath, $null, $true) # read-only
  Log "INFO" "Workbook opened (read)."

  # Get worksheet by name.
  Log "INFO" "Opening worksheet '$SheetName'..."
  $ws = $wb.Worksheets.Item($SheetName)
  Log "INFO" "Worksheet opened."

  # --- Build header map from first row ---
  # map[headerText] = columnNumber
  $map = @{}
  Log "INFO" "Building header map..."
  $cols = $ws.UsedRange.Columns.Count
  for ($c = 1; $c -le $cols; $c++) {
    $h = ("" + $ws.Cells.Item(1, $c).Text).Trim()
    if ($h) { $map[$h] = $c }
  }
  Log "INFO" "Header map built. Columns=$cols"

  # --- Determine ticket column ---
  $ticketCol = $null
  if ($TicketHeader -and $map.ContainsKey($TicketHeader)) {
    $ticketCol = $map[$TicketHeader]
  }
  elseif ($TicketColumn -gt 0) {
    $ticketCol = $TicketColumn
  }
  else {
    throw "Missing header '$TicketHeader' and no TicketColumn provided. Found: $($map.Keys -join ', ')"
  }
  Log "INFO" "Ticket column resolved: $ticketCol"

  # --- Collect ticket numbers ---
  # Use HashSet to avoid duplicates.
  $tickets = New-Object System.Collections.Generic.HashSet[string]
  $xlUp = -4162
  $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ticketCol).End($xlUp).Row
  if ($rows -lt 2) { $rows = 2 }
  Log "INFO" "Scanning ticket rows 2..$rows"

  $emptyStreak = 0
  $firstFoundRow = $null
  $ticketRange = $ws.Range($ws.Cells.Item(2, $ticketCol), $ws.Cells.Item($rows, $ticketCol))
  $ticketVals = $ticketRange.Value2
  $countRows = if ($ticketVals -is [System.Array]) { $ticketVals.GetLength(0) } else { 1 }
  for ($i = 1; $i -le $countRows; $i++) {
    $r = $i + 1
    $raw = if ($ticketVals -is [System.Array]) { $ticketVals[$i, 1] } else { $ticketVals }
    $t = ("" + $raw).Trim()

    # Accept INC/RITM/SCTASK + 6-8 digits.
    if ($t -match '^(INC|RITM|SCTASK)\d{6,8}$') {
      [void]$tickets.Add($t)
      if (-not $firstFoundRow) { $firstFoundRow = $r }
      $emptyStreak = 0
    }
    elseif ([string]::IsNullOrWhiteSpace($t)) {
      $emptyStreak++
      if ($tickets.Count -gt 0 -and $emptyStreak -ge $StopScanAfterEmptyRows) {
        Log "INFO" "Stopping read scan at row=$r after $emptyStreak consecutive empty rows."
        break
      }
    }
    else {
      $emptyStreak = 0
    }

    if ($firstFoundRow -and (($r - $firstFoundRow) -ge $MaxRowsAfterFirstTicket) -and ($emptyStreak -ge 10)) {
      Log "INFO" "Stopping read scan at row=$r after first ticket window ($MaxRowsAfterFirstTicket rows)."
      break
    }

    if (($r % $ReadProgressEveryRows) -eq 0) {
      Log "INFO" "Read progress row=$r/$rows found=$($tickets.Count)"
    }
  }
  Log "INFO" "Ticket scan completed. Found=$($tickets.Count)"

  # --- Cleanup COM objects to prevent Excel.exe zombie processes ---
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ticketRange) | Out-Null
  $wb.Close($false) | Out-Null
  $excel.Quit() | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()

  # Return as array (comma ensures array even if 1 item)
  return , @($tickets)
}

# =============================================================================
# Helper: decide if an Excel cell can be overwritten
# =============================================================================
function Is-EmptyOrPlaceholder([string]$CellText, [string]$Ticket) {
  # Consider "empty" or "ticket itself" as placeholder values
  # so we can autofill without overwriting real data.
  $t = ($CellText + "").Trim()
  if (-not $t) { return $true }
  if ($t -eq $Ticket) { return $true }
  return $false
}

function Get-OrCreateHeaderColumn {
  param($ws, [hashtable]$map, [string]$Header)
  if ($map.ContainsKey($Header)) {
    return [int]$map[$Header]
  }
  $newCol = [int]$ws.UsedRange.Columns.Count + 1
  $ws.Cells.Item(1, $newCol) = $Header
  $map[$Header] = $newCol
  return $newCol
}

function Build-RitmRecordUrl([string]$SysId) {
  if ([string]::IsNullOrWhiteSpace($SysId)) { return "" }
  return [string]::Format($RitmRecordUrlTemplate, $SysId.Trim())
}

function Build-SCTaskRecordUrl([string]$SysId) {
  if ([string]::IsNullOrWhiteSpace($SysId)) { return "" }
  return [string]::Format($SctaskRecordUrlTemplate, $SysId.Trim())
}

function Build-SCTaskFallbackUrl([string]$TaskNumber) {
  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return "" }
  $safeNumber = [System.Uri]::EscapeDataString($TaskNumber.Trim())
  return [string]::Format($SctaskListByNumberUrlTemplate, $safeNumber)
}

function Build-SCTaskBestUrl([string]$SysId, [string]$TaskNumber) {
  $u = Build-SCTaskRecordUrl $SysId
  if (-not [string]::IsNullOrWhiteSpace($u)) { return $u }
  return Build-SCTaskFallbackUrl $TaskNumber
}

function Get-DetectedPiFromActivityText {
  param([string]$ActivityText)
  if (-not $EnableActivityStreamSearch) { return "" }
  if ([string]::IsNullOrWhiteSpace($ActivityText)) {
    return $(if ($WriteNotFoundText) { "Not found" } else { "" })
  }

  $matches = [regex]::Matches(
    $ActivityText,
    '\b(?:02PI20[A-Z0-9_-]*|ITECBRUN[A-Z0-9_-]*|MUSTBRUN[A-Z0-9_-]*)\b',
    'IgnoreCase'
  )

  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $vals = New-Object System.Collections.Generic.List[string]
  foreach ($m in $matches) {
    $v = ("" + $m.Value).Trim().ToUpperInvariant()
    if ($v -and $seen.Add($v)) {
      [void]$vals.Add($v)
    }
  }

  if ($vals.Count -eq 0) {
    return $(if ($WriteNotFoundText) { "Not found" } else { "" })
  }
  return ($vals -join ", ")
}

function Set-ExcelHyperlinkSafe {
  param(
    $ws,
    [int]$Row,
    [int]$Col,
    [string]$DisplayText,
    [string]$Url,
    [string]$TicketForLog
  )
  $cell = $ws.Cells.Item($Row, $Col)
  try {
    if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() }
  } catch {}

  $missing = [System.Type]::Missing
  try {
    $cell.Value2 = $DisplayText
    $null = $ws.Hyperlinks.Add($cell, $Url, $missing, $missing, $DisplayText)
  }
  catch {
    try {
      $cell.Value2 = $DisplayText
      $null = $cell.Hyperlinks.Add($cell, $Url, $missing, $missing, $DisplayText)
    }
    catch {
      Log "ERROR" "Hyperlink failed for $TicketForLog at row=$Row col=$Col url='$Url' error='$($_.Exception.Message)'"
      $cell.Value2 = $Url
    }
  }
}

function Get-RitmActivityTextFromUiPage {
  param($wv, [string]$RitmSysId)

  if ([string]::IsNullOrWhiteSpace($RitmSysId)) { return "" }
  $ritmUrl = Build-RitmRecordUrl -SysId $RitmSysId
  if ([string]::IsNullOrWhiteSpace($ritmUrl)) { return "" }

  try {
    $wv.CoreWebView2.Navigate($ritmUrl)
  }
  catch {
    Log "ERROR" "RITM UI navigate failed sys_id='$RitmSysId': $($_.Exception.Message)"
    return ""
  }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 15000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function collectFromDoc(doc, out, seen){
      if (!doc) return;
      var selectors = [
        'h-card-wrapper activities-form',
        'h-card-wrapper .activities-form',
        '.activities-form',
        'activities-form',
        '.sn-widget-textblock-body',
        '.sn-widget-textblock-body_formatted',
        '.sn-card-component_accent-bar',
        '.sn-card-component_accent-bar_dark'
      ];
      for (var si = 0; si < selectors.length; si++) {
        var nodes = doc.querySelectorAll(selectors[si]);
        for (var ni = 0; ni < nodes.length; ni++) {
          var t = s(nodes[ni].innerText || nodes[ni].textContent || '');
          if (t && !seen[t]) { seen[t] = true; out.push(t); }
        }
      }
      var bodyTxt = s(doc.body && (doc.body.innerText || doc.body.textContent));
      if (bodyTxt && !seen[bodyTxt]) { seen[bodyTxt] = true; out.push(bodyTxt); }
    }
    var shellOut = [];
    var shellSeen = {};
    collectFromDoc(document, shellOut, shellSeen);
    var shellText = shellOut.join(' ');

    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var frameText = '';
    var frameReady = false;
    if (frame && frame.contentDocument) {
      var fdoc = frame.contentDocument;
      var readyState = s(fdoc.readyState || '');
      var frameOut = [];
      var frameSeen = {};
      collectFromDoc(fdoc, frameOut, frameSeen);
      frameText = frameOut.join(' ');
      if (readyState === 'complete' && frameText.length > 50) frameReady = true;
    }

    return JSON.stringify({
      ok:true,
      text: frameText ? frameText : shellText,
      frame_text: frameText,
      shell_text: shellText,
      frame_ready: frameReady
    });
  } catch(e) {
    return JSON.stringify({ ok:false, error:''+e, text:'' });
  }
})();
"@
  $o = $null
  $frameReady = $false
  for ($attempt = 1; $attempt -le 12; $attempt++) {
    $o = Parse-WV2Json (ExecJS $wv $js 7000)
    if ($o -and $o.PSObject.Properties["frame_ready"] -and $o.frame_ready -eq $true) {
      $frameReady = $true
      break
    }
    Start-Sleep -Milliseconds 500
  }
  if (-not $o) { return "" }
  if ($o.ok -ne $true -and $o.PSObject.Properties["error"]) {
    Log "ERROR" "RITM UI activity read failed sys_id='$RitmSysId': $($o.error)"
  }
  if ($o.PSObject.Properties["frame_text"]) {
    $ft = "" + $o.frame_text
    Log "INFO" "RITM UI frame text length sys_id='$RitmSysId': $($ft.Length) ready=$frameReady"
  }
  if ($o.PSObject.Properties["shell_text"]) {
    $st = "" + $o.shell_text
    Log "INFO" "RITM UI shell text length sys_id='$RitmSysId': $($st.Length)"
  }
  if ($o.PSObject.Properties["frame_text"]) {
    $ft = ("" + $o.frame_text)
    if (-not [string]::IsNullOrWhiteSpace($ft)) { return $ft }
  }
  if ($o.PSObject.Properties["text"]) { return ("" + $o.text) }
  return ""
}

# =============================================================================
# EXCEL: WRITE BACK RESULTS (optional)
# =============================================================================
function Write-BackToExcel {
  param(
    [string]$ExcelPath,
    [string]$SheetName,
    [string]$TicketHeader,
    [int]$TicketColumn,
    [string]$NameHeader,
    [string]$PhoneHeader,
    [string]$ActionHeader,
    [string]$DetectedPIHeader,
    [hashtable]$ResultMap
  )

  Log "INFO" "Writing back to Excel: $ExcelPath"
  Close-ExcelProcessesIfRequested -Reason "before write open"
  Log "INFO" "Creating Excel COM (write)..."

  # Open Excel in write mode.
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  try { $excel.AskToUpdateLinks = $false } catch {}
  try { $excel.EnableEvents = $false } catch {}
  Log "INFO" "Excel COM created (write), opening workbook..."
  $wb = $excel.Workbooks.Open($ExcelPath, $null, $false)
  Log "INFO" "Workbook opened (write)."
  $ws = $wb.Worksheets.Item($SheetName)

  # --- Build header map ---
  $map = @{}
  $cols = $ws.UsedRange.Columns.Count
  for ($c = 1; $c -le $cols; $c++) {
    $h = ("" + $ws.Cells.Item(1, $c).Text).Trim()
    if ($h) { $map[$h] = $c }
  }

  # --- Determine ticket column ---
  $ticketCol = $null
  if ($TicketHeader -and $map.ContainsKey($TicketHeader)) {
    $ticketCol = $map[$TicketHeader]
  }
  elseif ($TicketColumn -gt 0) {
    $ticketCol = $TicketColumn
  }
  else {
    throw "Missing header '$TicketHeader' and no TicketColumn provided. Found: $($map.Keys -join ', ')"
  }

  # --- Validate required headers for write-back ---
  if (-not $map.ContainsKey($NameHeader))   { throw "Missing header '$NameHeader'." }
  if (-not $map.ContainsKey($PhoneHeader))  { throw "Missing header '$PhoneHeader'." }
  if (-not $map.ContainsKey($ActionHeader)) { throw "Missing header '$ActionHeader'." }
  $detectedPiCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header $DetectedPIHeader
  $sctaskFirstCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "SCTask 1"
  $sctaskColMap = @{ 1 = $sctaskFirstCol }
  $sctaskColNums = New-Object System.Collections.Generic.List[int]
  foreach ($k in @($map.Keys)) {
    if ($k -match '^SCTask\s+(\d+)$') {
      [void]$sctaskColNums.Add([int]$matches[1])
      if (-not $sctaskColMap.ContainsKey([int]$matches[1])) {
        $sctaskColMap[[int]$matches[1]] = [int]$map[$k]
      }
    }
  }
  if (-not $sctaskColNums.Contains(1)) { [void]$sctaskColNums.Add(1) }

  # --- Iterate rows and fill values (only if empty/placeholder) ---
  $xlUp = -4162
  $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ticketCol).End($xlUp).Row
  if ($rows -lt 2) { $rows = 2 }
  Log "INFO" "Write-back scan rows 2..$rows"
  $emptyStreak = 0
  $firstFoundRow = $null
  $ticketRange = $ws.Range($ws.Cells.Item(2, $ticketCol), $ws.Cells.Item($rows, $ticketCol))
  $ticketVals = $ticketRange.Value2
  $countRows = if ($ticketVals -is [System.Array]) { $ticketVals.GetLength(0) } else { 1 }
  for ($i = 1; $i -le $countRows; $i++) {
    $r = $i + 1
    $rawTicket = if ($ticketVals -is [System.Array]) { $ticketVals[$i, 1] } else { $ticketVals }
    $ticket = ("" + $rawTicket).Trim()
    if (-not $ticket) {
      $emptyStreak++
      if ($emptyStreak -ge $StopScanAfterEmptyRows) {
        Log "INFO" "Stopping write-back scan at row=$r after $emptyStreak consecutive empty rows."
        break
      }
      continue
    }
    if ($ticket -match '^(INC|RITM|SCTASK)\d{6,8}$' -and -not $firstFoundRow) { $firstFoundRow = $r }
    $emptyStreak = 0
    if ($firstFoundRow -and (($r - $firstFoundRow) -ge $MaxRowsAfterFirstTicket) -and ($emptyStreak -ge 10)) {
      Log "INFO" "Stopping write-back scan at row=$r after first ticket window ($MaxRowsAfterFirstTicket rows)."
      break
    }
    if (-not $ResultMap.ContainsKey($ticket)) { continue }

    $res = $ResultMap[$ticket]
    if ($res.ok -ne $true) { continue }

    # Fill "Name" (affected_user)
    $nameCell = "" + $ws.Cells.Item($r, $map[$NameHeader]).Text
    if (Is-EmptyOrPlaceholder $nameCell $ticket) {
      $ws.Cells.Item($r, $map[$NameHeader]) = $res.affected_user
    }

    # Fill "New Phone" (configuration_item)
    $phoneCell = "" + $ws.Cells.Item($r, $map[$PhoneHeader]).Text
    if (Is-EmptyOrPlaceholder $phoneCell $ticket) {
      $ws.Cells.Item($r, $map[$PhoneHeader]) = $res.configuration_item
    }

    # Fill "Action finished?" (status)
    $actionCell = "" + $ws.Cells.Item($r, $map[$ActionHeader]).Text
    if (Is-EmptyOrPlaceholder $actionCell $ticket) {
      $statusOut = $res.status

      # Localized label normalization (example from German UI)
      if ($statusOut -eq "Abgebrochen") { $statusOut = "Cancelled" }

      $ws.Cells.Item($r, $map[$ActionHeader]) = $statusOut
    }

    # Fill "Detected PI / Machine"
    $detectedPiOut = ""
    if ($res.PSObject.Properties["detected_pi_machine"]) {
      $detectedPiOut = ("" + $res.detected_pi_machine).Trim()
    }
    $detectedPiCell = "" + $ws.Cells.Item($r, $detectedPiCol).Text
    if (Is-EmptyOrPlaceholder $detectedPiCell $ticket -and $detectedPiOut) {
      $ws.Cells.Item($r, $detectedPiCol) = $detectedPiOut
    }

    # Fill open SCTASK hyperlinks as "SCTask 1", "SCTask 2", ...
    $openTasks = @()
    if ($res.PSObject.Properties["open_task_items"] -and $res.open_task_items) {
      $openTasks = @($res.open_task_items)
    }

    if ($ticket -like "RITM*") {
      if ($openTasks.Count -gt 0) {
        for ($idx = 1; $idx -le $openTasks.Count; $idx++) {
          if (-not $sctaskColMap.ContainsKey($idx)) {
            $sctaskColMap[$idx] = Get-OrCreateHeaderColumn -ws $ws -map $map -Header ("SCTask " + $idx)
            if (-not $sctaskColNums.Contains($idx)) { [void]$sctaskColNums.Add($idx) }
          }
          $taskObj = $openTasks[$idx - 1]
          $taskNumber = if ($taskObj.PSObject.Properties["number"]) { "" + $taskObj.number } else { "" }
          $taskSysId = if ($taskObj.PSObject.Properties["sys_id"]) { "" + $taskObj.sys_id } else { "" }
          $taskUrl = Build-SCTaskBestUrl -SysId $taskSysId -TaskNumber $taskNumber
          if ([string]::IsNullOrWhiteSpace($taskUrl)) {
            $taskUrl = Build-SCTaskFallbackUrl -TaskNumber $taskNumber
          }
          Set-ExcelHyperlinkSafe -ws $ws -Row $r -Col $sctaskColMap[$idx] -DisplayText ("SCTask " + $idx) -Url $taskUrl -TicketForLog $ticket
        }
        foreach ($idx in @($sctaskColNums)) {
          if ($idx -le $openTasks.Count) { continue }
          if (-not $sctaskColMap.ContainsKey($idx)) { continue }
          $cell = $ws.Cells.Item($r, $sctaskColMap[$idx])
          try { if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() } } catch {}
          $cell.Value2 = ""
        }
      }
      else {
        foreach ($idx in @($sctaskColNums)) {
          if (-not $sctaskColMap.ContainsKey($idx)) { continue }
          $cell = $ws.Cells.Item($r, $sctaskColMap[$idx])
          try { if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() } } catch {}
          $cell.Value2 = if ($idx -eq 1) { "No open tasks." } else { "" }
        }
      }
    }
    else {
      foreach ($idx in @($sctaskColNums)) {
        if (-not $sctaskColMap.ContainsKey($idx)) { continue }
        $cell = $ws.Cells.Item($r, $sctaskColMap[$idx])
        try { if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() } } catch {}
        $cell.Value2 = ""
      }
    }
  }

  # Save changes
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ticketRange) | Out-Null
  $wb.Save()
  $wb.Close($false) | Out-Null
  $excel.Quit() | Out-Null

  # Cleanup COM objects
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()

  Log "INFO" "Excel updated."
}

# =============================================================================
# Ticket number -> ServiceNow table mapping
# =============================================================================
function Ticket-ToTable([string]$ticket) {
  # Determine the SNOW table for JSONv2 query based on ticket prefix:
  # - INC...   => incident
  # - RITM...  => sc_req_item
  # - SCTASK.. => sc_task
  if ($ticket -like "INC*")  { return "incident" }
  if ($ticket -like "RITM*") { return "sc_req_item" }
  return "sc_task"
}

# =============================================================================
# EXTRACTION: JSONv2 via JavaScript inside authenticated WebView2
# =============================================================================
function Extract-Ticket_JSONv2 {
  param($wv, [string]$Ticket)

  # Determine which table we query for this ticket.
  $table = Ticket-ToTable $Ticket
  $closedStatesJson = ($ClosedTaskStates | ConvertTo-Json -Compress)
  $enableActivitySearchJs = if ($EnableActivityStreamSearch) { "true" } else { "false" }

  # ---------------------------------------------------------------------------
  # JS block executed inside WebView2
  #
  # Why JS?
  # - Runs in the authenticated browser session (SSO cookies)
  # - Can call internal SNOW endpoints with credentials
  #
  # What it does:
  # 1) Fetch record by number from the correct table using JSONv2
  # 2) Determine user field depending on table:
  #    - RITM: requested_for
  #    - INC : caller_id
  #    - If not found and request exists: fetch sc_request and read requested_for
  # 3) Resolve sys_id -> display name for user and CI
  # 4) Resolve state label via sys_choice (best display label)
  # 5) Return normalized object for PowerShell to parse
  # ---------------------------------------------------------------------------
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? "" : (""+x).trim(); }

    function pickRec(obj){
      return (obj && obj.records && obj.records[0]) ? obj.records[0] :
             (obj && obj.result && obj.result[0]) ? obj.result[0] :
             (Array.isArray(obj) && obj[0]) ? obj[0] : null;
    }

    function looksSysId(v){ return /^[0-9a-f]{32}$/i.test(s(v)); }

    function httpGetText(url){
      try {
        var x = new XMLHttpRequest();
        x.open('GET', url, false);     // synchronous request (simple + reliable here)
        x.withCredentials = true;      // ensure cookies/SSO
        x.send(null);
        return (x.status>=200 && x.status<300) ? (x.responseText||"") : "";
      } catch(e){ return ""; }
    }

    function httpGetJsonV2(path){
      var txt = httpGetText(path);
      if (!txt) return null;
      try { return JSON.parse(txt); } catch(e){ return null; }
    }

    function resolveUserFromId(id){
      if (!looksSysId(id)) return "";
      var u = '/sys_user.do?JSONv2&sysparm_limit=1&sysparm_query=' + encodeURIComponent('sys_id=' + id);
      var obj = httpGetJsonV2(u);
      var rec = pickRec(obj);
      if (!rec) return "";
      // Prefer display name, fallback to user_name / employee_number
      return s(rec.name || rec.user_name || rec.employee_number || "");
    }

    function resolveCiFromId(id){
      if (!looksSysId(id)) return "";
      var u = '/cmdb_ci.do?JSONv2&sysparm_limit=1&sysparm_query=' + encodeURIComponent('sys_id=' + id);
      var obj = httpGetJsonV2(u);
      var rec = pickRec(obj);
      if (!rec) return "";
      return s(rec.name || rec.display_name || rec.u_name || "");
    }

    function resolveStateLabel(table, value){
      if (!s(value)) return "";
      var u = '/sys_choice.do?JSONv2&sysparm_limit=1&sysparm_query=' +
              encodeURIComponent('name=' + table + '^element=state^value=' + value);
      var obj = httpGetJsonV2(u);
      var rec = pickRec(obj);
      if (!rec) return "";
      return s(rec.label || rec.value || "");
    }

    var CLOSED_TASK_STATES = $closedStatesJson;

    function normalizeStateToken(v){
      return s(v).toLowerCase().replace(/[\s_-]+/g, ' ').trim();
    }

    function buildClosedStateSets(){
      var labels = {};
      var values = {};
      if (!Array.isArray(CLOSED_TASK_STATES)) return { labels:labels, values:values };
      for (var i = 0; i < CLOSED_TASK_STATES.length; i++) {
        var n = normalizeStateToken(CLOSED_TASK_STATES[i]);
        if (!n) continue;
        if (/^\d+$/.test(n)) values[n] = true; else labels[n] = true;
      }
      return { labels:labels, values:values };
    }

    var CLOSED_STATE_SETS = buildClosedStateSets();

    function isTaskOpen(stateValue, stateLabel){
      var sv = normalizeStateToken(stateValue);
      var sl = normalizeStateToken(stateLabel);
      if (sv === '3' || sv === '4' || sv === '7') return false;
      if (sv && CLOSED_STATE_SETS.values[sv]) return false;
      if (sv && CLOSED_STATE_SETS.labels[sv]) return false;
      if (sl && CLOSED_STATE_SETS.labels[sl]) return false;
      return true;
    }

    function getOpenCatalogTasks(reqItemSysId, ritmNumber){
      var rowsAll = [];

      if (looksSysId(reqItemSysId)) {
        var q1 = 'request_item=' + reqItemSysId;
        var p1 = '/sc_task.do?JSONv2&sysparm_limit=200&sysparm_query=' + encodeURIComponent(q1);
        var o1 = httpGetJsonV2(p1);
        var r1 = (o1 && o1.records) ? o1.records : ((o1 && o1.result) ? o1.result : []);
        if (Array.isArray(r1)) rowsAll = rowsAll.concat(r1);
      }

      // Fallback: dot-walk by RITM number in case request_item sys_id query is ACL-limited.
      if (s(ritmNumber)) {
        var q2 = 'request_item.number=' + ritmNumber;
        var p2 = '/sc_task.do?JSONv2&sysparm_limit=200&sysparm_query=' + encodeURIComponent(q2);
        var o2 = httpGetJsonV2(p2);
        var r2 = (o2 && o2.records) ? o2.records : ((o2 && o2.result) ? o2.result : []);
        if (Array.isArray(r2)) rowsAll = rowsAll.concat(r2);
      }

      // De-duplicate by task number/sys_id.
      var seen = {};
      var rows = [];
      for (var i0 = 0; i0 < rowsAll.length; i0++) {
        var k = s(rowsAll[i0].sys_id || rowsAll[i0].number || ("idx_" + i0));
        if (!seen[k]) {
          seen[k] = true;
          rows.push(rowsAll[i0]);
        }
      }

      if (!Array.isArray(rows)) return [];
      var openTasks = [];
      for (var i = 0; i < rows.length; i++) {
        var stVal = s(rows[i].state || "");
        var stLabel = resolveStateLabel('sc_task', stVal);
        if (isTaskOpen(stVal, stLabel)) {
          openTasks.push({
            number: s(rows[i].number || ""),
            sys_id: s(rows[i].sys_id || ""),
            state_value: stVal,
            state_label: stLabel
          });
        }
      }
      return openTasks;
    }

    var activityRetrievalError = "";

    function getRows(path, query){
      var p = path + '?JSONv2&sysparm_limit=200&sysparm_query=' + encodeURIComponent(query);
      var o = httpGetJsonV2(p);
      var rows = (o && o.records) ? o.records : ((o && o.result) ? o.result : []);
      return Array.isArray(rows) ? rows : [];
    }

    function getRitmActivitiesText(reqItemSysId, ritmNumber){
      var out = [];

      // Preferred: backend journal entries for this RITM
      try {
        if (looksSysId(reqItemSysId)) {
          var jAll = [];
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'name=sc_req_item^element_id=' + reqItemSysId));
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'element_id=' + reqItemSysId));
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'element_id=' + reqItemSysId + '^elementINcomments,work_notes'));
          for (var i1 = 0; i1 < jAll.length; i1++) {
            var v1 = s(jAll[i1].value || jAll[i1].message || jAll[i1].comments || jAll[i1].work_notes || "");
            if (v1) out.push(v1);
          }
        }
      } catch(ej) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";journal") : "journal"; }

      // UI container fallback: h-card-wrapper activities-form
      try {
        var selectors = [
          'h-card-wrapper activities-form',
          'h-card-wrapper .activities-form',
          '.activities-form',
          'activities-form',
          '.sn-widget-textblock-body',
          '.sn-widget-textblock-body_formatted',
          '.sn-card-component_accent-bar',
          '.sn-card-component_accent-bar_dark'
        ];
        var seenText = {};
        for (var si = 0; si < selectors.length; si++) {
          var nodes = document.querySelectorAll(selectors[si]);
          for (var ni = 0; ni < nodes.length; ni++) {
            var tx = s(nodes[ni].innerText || nodes[ni].textContent || "");
            if (tx && !seenText[tx]) {
              seenText[tx] = true;
              out.push(tx);
            }
          }
        }
      } catch(ed) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";dom") : "dom"; }

      // Also attempt inside the SNOW main iframe when present.
      try {
        var f = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
        var fdoc = (f && f.contentDocument) ? f.contentDocument : null;
        if (fdoc) {
          var selectorsF = [
            'h-card-wrapper activities-form',
            'h-card-wrapper .activities-form',
            '.activities-form',
            'activities-form',
            '.sn-widget-textblock-body',
            '.sn-widget-textblock-body_formatted',
            '.sn-card-component_accent-bar',
            '.sn-card-component_accent-bar_dark'
          ];
          var seenTextF = {};
          for (var sf = 0; sf < selectorsF.length; sf++) {
            var nodesF = fdoc.querySelectorAll(selectorsF[sf]);
            for (var nf = 0; nf < nodesF.length; nf++) {
              var txf = s(nodesF[nf].innerText || nodesF[nf].textContent || "");
              if (txf && !seenTextF[txf]) {
                seenTextF[txf] = true;
                out.push(txf);
              }
            }
          }
        }
      } catch(ef) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";dom_iframe") : "dom_iframe"; }

      // Custom activity table provided by you
      var activityTable = '/activity_ee4a85aa3bcf3e14ca382f37f4e45a20.do';
      var aAll = [];
      try {
        if (looksSysId(reqItemSysId)) {
          aAll = aAll.concat(getRows(activityTable, 'request_item=' + reqItemSysId));
          aAll = aAll.concat(getRows(activityTable, 'element_id=' + reqItemSysId));
          aAll = aAll.concat(getRows(activityTable, 'parent=' + reqItemSysId));
        }
        if (s(ritmNumber)) {
          aAll = aAll.concat(getRows(activityTable, 'number=' + ritmNumber));
          aAll = aAll.concat(getRows(activityTable, 'documentkey=' + ritmNumber));
        }
        for (var i2 = 0; i2 < aAll.length; i2++) {
          var v2 = s(
            aAll[i2].value ||
            aAll[i2].message ||
            aAll[i2].comments ||
            aAll[i2].work_notes ||
            aAll[i2].text ||
            aAll[i2].description ||
            aAll[i2].u_message ||
            ""
          );
          if (v2) out.push(v2);
        }
      } catch(ea) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";custom_activity") : "custom_activity"; }
      return out.join(' ');
    }

    function extractMachineFromActivityText(activityText){
      var txt = s(activityText);
      if (!txt) return "";

      var m = txt.match(/\b(?:02PI20[A-Z0-9_-]*|ITECBRUN[A-Z0-9_-]*|MUSTBRUN[A-Z0-9_-]*)\b/i);
      return m ? s(m[0]).toUpperCase() : "";
    }

    // --- Main record fetch ---
    var q1 = 'number=' + '$Ticket';
    var p1 = '/$table.do?JSONv2&sysparm_limit=1&sysparm_query=' + encodeURIComponent(q1);
    var o1 = httpGetJsonV2(p1);
    var r1 = pickRec(o1);
    if (!r1) return JSON.stringify({ ok:false, reason:'not_found', ticket:'$Ticket', table:'$table', query:p1 });

    // --- Determine affected user display value ---
    var userDisplay = s(r1.requested_for || r1.caller_id || "");

    // If still empty and record links to sc_request via sys_id, try sc_request.requested_for
    if (!userDisplay && r1.request && looksSysId(r1.request)) {
      var reqObj = httpGetJsonV2('/sc_request.do?JSONv2&sysparm_limit=1&sysparm_query=' +
                                encodeURIComponent('sys_id=' + r1.request));
      var reqRec = pickRec(reqObj);
      if (reqRec) userDisplay = s(reqRec.requested_for || "");
    }

    // Resolve sys_id to a user display string if needed
    var userName = looksSysId(userDisplay) ? resolveUserFromId(userDisplay) : userDisplay;
    if (!userName) userName = userDisplay;

    // --- Determine CI display value ---
    var ciVal = s(r1.configuration_item || r1.cmdb_ci || "");
    if (looksSysId(ciVal)) {
      var ciName = resolveCiFromId(ciVal);
      if (ciName) ciVal = ciName;
    }

    // --- RITM-specific enrichments ---
    var openTaskCount = 0;
    var openTasks = [];
    var acts = "";
    if ('$table' === 'sc_req_item') {
      openTasks = getOpenCatalogTasks(s(r1.sys_id || ""), '$Ticket');
      openTaskCount = openTasks.length;

      // If activities contain a PI machine id, prefer that value for CI output.
      if ($enableActivitySearchJs) {
        acts = getRitmActivitiesText(s(r1.sys_id || ""), '$Ticket');
        var piFromActivity = extractMachineFromActivityText(acts);
        if (piFromActivity) {
          ciVal = piFromActivity;
        }
      }
    }

    // --- Determine status (state) ---
    var stVal = s(r1.state || "");
    var stLabel = resolveStateLabel('$table', stVal);

    // Fallback mapping if label lookup fails
    var stMap = {
      'sc_req_item': {'1':'Open','2':'In Progress','3':'Closed Complete','4':'Closed Incomplete','7':'Cancelled'},
      'incident':    {'1':'New','2':'In Progress','3':'On Hold','6':'Resolved','7':'Closed','8':'Canceled'},
      'sc_task':     {'1':'Open','2':'Work in Progress','3':'Closed Complete','4':'Closed Incomplete','7':'Cancelled'}
    };
    var stFallback = (stMap['$table'] && stMap['$table'][stVal]) ? stMap['$table'][stVal] : "";
    var stOut = stLabel ? stLabel : (stFallback ? stFallback : stVal);
    if ('$table' === 'sc_req_item' && openTaskCount > 0) {
      stOut = 'Open:' + openTaskCount;
    }

    // Return normalized object
    return JSON.stringify({
      ok:true,
      ticket:'$Ticket',
      table:'$table',
      sys_id:s(r1.sys_id || ""),
      affected_user:userName,
      configuration_item:ciVal,
      status:stOut,
      status_value:stVal,
      status_label:stLabel,
      open_tasks:openTaskCount,
      open_task_items:openTasks,
      activity_text:acts,
      activity_error:activityRetrievalError,
      query:p1
    });
  } catch(e) {
    // If anything breaks inside JS, return structured error.
    return JSON.stringify({ ok:false, reason:'exception', error:""+e, ticket:'$Ticket', table:'$table' });
  }
})();
"@

  # Execute JS and parse
  $o = Parse-WV2Json (ExecJS $wv $js 12000)
  if ($o) { return $o }

  # If no response (timeout/failure), return a PowerShell object with minimal info.
  return [pscustomobject]@{
    ok                 = $false
    reason             = "no_js_response"
    ticket             = $Ticket
    table              = $table
    sys_id             = ""
    affected_user      = ""
    configuration_item = ""
    open_tasks         = 0
    open_task_items    = @()
    activity_text      = ""
    activity_error     = ""
    href               = "" + $wv.Source
  }
}

# =============================================================================
# MAIN EXECUTION
# =============================================================================
$session = $null

try {
  # 1) Load WebView2 DLLs from Teams add-in
  Load-WebView2FromTeams

  # 2) Select Excel file
  $ExcelPath = Pick-ExcelFile -ExcelPath $ExcelPath -DefaultStartDir $DefaultStartDir -DefaultExcelName $DefaultExcelName
  Log "INFO" "Excel selected: $ExcelPath"

  # 3) Interactive SSO login, then keep WebView2 session
  $session = Connect-ServiceNowSSO -StartUrl $LoginUrl
  $wv = $session.Wv

  # 4) Read tickets list from Excel
  $tickets = Read-TicketsFromExcel -ExcelPath $ExcelPath -TicketHeader $TicketHeader -SheetName $SheetName -TicketColumn $TicketColumn
  Log "INFO" "Tickets found: $($tickets.Count)"

  if ($tickets.Count -eq 0) {
    throw "No valid tickets found in Excel (INC/RITM/SCTASK + 6-8 digits)."
  }

  # 5) For each ticket: extract + export JSON
  $results = New-Object System.Collections.Generic.List[object]
  $i = 0

  foreach ($t in $tickets) {
    $i++
    Log "INFO" "[$i/$($tickets.Count)] Open + extract: $t"

    # Extract fields via JSONv2 in authenticated session
    $r = Extract-Ticket_JSONv2 -wv $wv -Ticket $t

    if (($r.ok -eq $true) -and ($t -like "RITM*")) {
      try {
        $activityText = if ($r.PSObject.Properties["activity_text"]) { "" + $r.activity_text } else { "" }
        $uiActivityText = ""
        $activityError = if ($r.PSObject.Properties["activity_error"]) { "" + $r.activity_error } else { "" }
        Log "INFO" "$t activity backend text length: $($activityText.Length)"
        if ($activityError) {
          Log "ERROR" "$t activity retrieval issue: $activityError"
        }
        $detectedPi = Get-DetectedPiFromActivityText -ActivityText $activityText
        if ([string]::IsNullOrWhiteSpace($detectedPi) -and $EnableActivityStreamSearch) {
          $ritmSysId = if ($r.PSObject.Properties["sys_id"]) { "" + $r.sys_id } else { "" }
          $uiActivityText = Get-RitmActivityTextFromUiPage -wv $wv -RitmSysId $ritmSysId
          Log "INFO" "$t activity UI text length: $($uiActivityText.Length)"
          if (-not [string]::IsNullOrWhiteSpace($uiActivityText)) {
            Log "INFO" "$t activity UI fallback text collected (len=$($uiActivityText.Length))"
            $detectedPi = Get-DetectedPiFromActivityText -ActivityText $uiActivityText
          }
        }
        if ($t -eq $DebugActivityTicket) {
          $backendDump = $activityText
          if ($backendDump.Length -gt $DebugActivityMaxChars) { $backendDump = $backendDump.Substring(0, $DebugActivityMaxChars) }
          $uiDump = $uiActivityText
          if ($uiDump.Length -gt $DebugActivityMaxChars) { $uiDump = $uiDump.Substring(0, $DebugActivityMaxChars) }
          $r | Add-Member -NotePropertyName activity_text_backend_debug -NotePropertyValue $backendDump -Force
          $r | Add-Member -NotePropertyName activity_text_ui_debug -NotePropertyValue $uiDump -Force
          Log "INFO" "$t backend debug text: [$backendDump]"
          Log "INFO" "$t UI debug text: [$uiDump]"
        }
        if ([string]::IsNullOrWhiteSpace($detectedPi)) {
          Log "INFO" "$t PI not found in activity stream"
        }
        $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue $detectedPi -Force
      }
      catch {
        Log "ERROR" "$t activity parsing failed: $($_.Exception.Message)"
        $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue $(if ($WriteNotFoundText) { "Not found" } else { "" }) -Force
      }

      $openTaskItems = @()
      if ($r.PSObject.Properties["open_task_items"] -and $r.open_task_items) {
        $openTaskItems = @($r.open_task_items)
      }
      Log "INFO" "$t open SCTASK count: $($openTaskItems.Count)"
      foreach ($ot in $openTaskItems) {
        $taskNo = if ($ot.PSObject.Properties["number"]) { "" + $ot.number } else { "" }
        $taskSys = if ($ot.PSObject.Properties["sys_id"]) { "" + $ot.sys_id } else { "" }
        $taskUrl = Build-SCTaskBestUrl -SysId $taskSys -TaskNumber $taskNo
        Log "INFO" "$t open SCTASK number='$taskNo' sys_id='$taskSys' url='$taskUrl'"
      }
    }
    elseif ($r.ok -eq $true) {
      $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue "" -Force
    }

    # Log quick summary line
    $status = if ($r.ok -eq $true) { "OK" } else { "FAIL" }
    $reason = if ($r -and $r.PSObject.Properties["reason"]) { "" + $r.reason } else { "" }
    $userOut = if ($r -and $r.PSObject.Properties["affected_user"]) { "" + $r.affected_user } else { "" }
    $ciOut = if ($r -and $r.PSObject.Properties["configuration_item"]) { "" + $r.configuration_item } else { "" }
    $urlOut = if ($r -and $r.PSObject.Properties["query"]) { "" + $r.query } elseif ($r -and $r.PSObject.Properties["href"]) { "" + $r.href } else { "" }
    Log "INFO" "$t => $status reason=$reason user='$userOut' ci='$ciOut' url='$urlOut'"

    # Save per-ticket JSON file
    $perPath = Join-Path $OutDir ("ticket_" + $t + ".json")
    $jsonPer = ($r | ConvertTo-Json -Depth 6) -replace '\\u0027', "'"
    Set-Content -Path $perPath -Value $jsonPer -Encoding UTF8

    # Add to in-memory list for combined export + write-back map
    $results.Add($r) | Out-Null
  }

  # 6) Save combined JSON
  $jsonAll = ($results | ConvertTo-Json -Depth 6) -replace '\\u0027', "'"
  Set-Content -Path $AllJson -Value $jsonAll -Encoding UTF8
  Log "INFO" "ALL JSON: $AllJson"
  Log "INFO" "DONE. Logs: $LogPath"

  # 7) Optional write-back to Excel
  if ($WriteBackExcel) {
    # Build a map ticket -> result object
    $map = @{}
    foreach ($r in $results) { $map[$r.ticket] = $r }

    Write-BackToExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
      -NameHeader $NameHeader -PhoneHeader $PhoneHeader -ActionHeader $ActionHeader -DetectedPIHeader $DetectedPIHeader -ResultMap $map
  }

  # 8) Final success popup
  [System.Windows.Forms.MessageBox]::Show(
    "Export complete.`r`nFolder: $OutDir`r`nAll JSON: $AllJson",
    "SNOW Export",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
  ) | Out-Null
}
catch {
  # Any exception: log + show popup
  Log "ERROR" $_.Exception.Message

  [System.Windows.Forms.MessageBox]::Show(
    $_.Exception.Message,
    "SNOW Export ERROR",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
  ) | Out-Null
}
finally {
  # Cleanup: close the hidden login form gracefully.
  # We do NOT brutally kill WebView2 processes; dispose the form if possible.
  try {
    if ($session -and $session.Form) {
      $session.Form.Close()
      $session.Form.Dispose()
    }
  }
  catch {}

  # Force GC to reduce lingering COM/WebView2 references.
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
