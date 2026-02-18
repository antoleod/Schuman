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
# - â€œNO APIâ€ here means: no token-based REST Table API usage, no ServiceNow API keys.
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
  [string]$SCTasksHeader = "SCTasks",
  [switch]$DashboardMode,
  [switch]$NoPopups,
  [ValidateSet("Auto","RitmOnly","IncAndRitm","All")]
  [string]$ProcessingScope = "Auto",
  [int]$MaxTickets = 0,
  [switch]$QuickMode,
  [switch]$SmartMode,
  [switch]$TurboMode,
  [switch]$SkipActivityScan,
  [switch]$NoUiFallback,
  [switch]$NoWriteBack,
  # >>> CHANGE DEFAULT EXCEL NAME HERE if the planning file is renamed again <<<
  [string]$DefaultExcelName = "Schuman List.xlsx",
  [string]$DefaultStartDir = $PSScriptRoot,
  [string]$RitmScanDbPath = (Join-Path (Join-Path $PSScriptRoot "system\db") "ritm-scan-db.json")
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
$IncidentRecordUrlTemplate     = "$InstanceBaseUrl/nav_to.do?uri=%2Fincident.do%3Fsys_id%3D{0}%26sysparm_view%3D"
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
$FastMode = $true
$EnableUiFallbackActivitySearch = $true
$UiFallbackMinBackendChars = 200
$DebugActivityTicket = ""
$DebugActivityMaxChars = 4000
$ForceUpdateDetectedPI = $true
$ForceUpdateNameFromLegal = $true
$WritePerTicketJson = $true
$VerboseTicketLogging = $true
$ExtractJsTimeoutMs = 12000
$ExtractRetryCount = 3
$ExtractRetryDelayMs = 1200
$EnableSctaskRowExpansion = $false
$DisableSysChoiceLookup = $false

if ($FastMode) {
  $EnableUiFallbackActivitySearch = $true
  $WritePerTicketJson = $false
  $VerboseTicketLogging = $false
  $ReadProgressEveryRows = 5000
  $ExtractJsTimeoutMs = 12000
  $ExtractRetryCount = 4
  $ExtractRetryDelayMs = 1500
  $UiFallbackMinBackendChars = 120
}

if ($QuickMode) {
  if (-not $PSBoundParameters.ContainsKey("ProcessingScope")) { $ProcessingScope = "RitmOnly" }
  if ($MaxTickets -le 0) { $MaxTickets = 30 }
  $SkipActivityScan = $true
  $NoWriteBack = $true
  $ExtractRetryCount = [Math]::Min($ExtractRetryCount, 2)
  $ExtractRetryDelayMs = [Math]::Min($ExtractRetryDelayMs, 500)
  Log "INFO" ("Quick mode enabled: scope={0}, maxTickets={1}, skipActivity={2}, writeBackDisabled={3}" -f $ProcessingScope, $MaxTickets, $true, $true)
}

if ($SmartMode) {
  Log "INFO" "Smart mode enabled: caching user lookups and skipping heavy activity scan for INC tickets."
}

if ($TurboMode) {
  if (-not $PSBoundParameters.ContainsKey("ProcessingScope")) { $ProcessingScope = "RitmOnly" }
  $NoUiFallback = $true
  $ExtractRetryCount = [Math]::Min($ExtractRetryCount, 2)
  $ExtractRetryDelayMs = [Math]::Min($ExtractRetryDelayMs, 500)
  $ExtractJsTimeoutMs = [Math]::Min($ExtractJsTimeoutMs, 9000)
  $DisableSysChoiceLookup = $true
  Log "INFO" "Turbo mode enabled: RITM-first filtering, state cache, no UI fallback, reduced retries/timeouts."
}

if ($SkipActivityScan) {
  $EnableActivityStreamSearch = $false
  $EnableUiFallbackActivitySearch = $false
  Log "INFO" "Speed mode: activity-stream PI scan disabled (-SkipActivityScan)."
}
if ($NoWriteBack) {
  $WriteBackExcel = $false
  Log "INFO" "Speed mode: Excel write-back disabled (-NoWriteBack)."
}
if ($NoUiFallback) {
  $EnableUiFallbackActivitySearch = $false
  Log "INFO" "Speed mode: UI fallback disabled (-NoUiFallback)."
}

# NameHeader/PhoneHeader/ActionHeader are provided via param().

# ----------------------------
# Output folder
# ----------------------------
# Create a unique per-run folder under %TEMP% so multiple runs don't overwrite each other.
$RunId  = Get-Date -Format "yyyyMMdd_HHmmss"
$SystemRoot = Join-Path $PSScriptRoot "system"
$SystemLogsDir = Join-Path $SystemRoot "logs"
$SystemDbDir = Join-Path $SystemRoot "db"
$SystemRunsDir = Join-Path $SystemRoot "runs"
$OutDir = Join-Path $SystemRunsDir "SNOW_Tickets_Export_$RunId"

# Ensure system folders exist.
New-Item -ItemType Directory -Force -Path $SystemRoot | Out-Null
New-Item -ItemType Directory -Force -Path $SystemLogsDir | Out-Null
New-Item -ItemType Directory -Force -Path $SystemDbDir | Out-Null
New-Item -ItemType Directory -Force -Path $SystemRunsDir | Out-Null
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

# Log file path.
$LogPath = Join-Path $OutDir "run.log.txt"
$HistoryLogPath = Join-Path $SystemLogsDir "auto-excel.history.log"
$ScriptBuildTag = "auto-excel build 2026-02-17 18:30 inc-hold-status-ux"
$DashboardDefaultCheckInNote = "Deliver all credentials to the new user"
$DashboardDefaultCheckOutNote = "Laptop has been delivered.`r`nFirst login made with the user.`r`nOutlook, Teams and Jabber successfully tested."

# Combined JSON file path.
$AllJson = Join-Path $OutDir "tickets_export.json"
$script:RitmScanDbLastSuccessfulUtc = ""
$script:RitmScanRunSuccessful = $false

# =============================================================================
# LOGGING
# =============================================================================
function Log([string]$level, [string]$msg) {
  # Build a timestamped log line.
  $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$level] $msg"

  # Also write to run log file (best effort; do not crash if file write fails).
  try { Add-Content -Path $LogPath -Value $line } catch {}
  # Persistent history across runs.
  try { Add-Content -Path $HistoryLogPath -Value $line } catch {}
}

function Start-PerfTimer {
  return [System.Diagnostics.Stopwatch]::StartNew()
}

function Stop-PerfTimer {
  param(
    [System.Diagnostics.Stopwatch]$Timer,
    [string]$Label
  )
  if (-not $Timer) { return 0 }
  $Timer.Stop()
  $ms = [int64]$Timer.Elapsed.TotalMilliseconds
  Log "INFO" "PERF $Label took ${ms}ms"
  return $ms
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
Log "INFO" $ScriptBuildTag
Log "INFO" "Incremental scan DB path: $RitmScanDbPath"

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

function Ensure-SnowReady {
  param($wv, [int]$MaxWaitMs = 12000)
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt $MaxWaitMs) {
    $diag = Parse-WV2Json (ExecJS $wv "JSON.stringify({href:location.href,ready:document.readyState})" 2500)
    if ($diag -and $diag.PSObject.Properties["href"]) {
      $href = ("" + $diag.href).ToLowerInvariant()
      $ready = if ($diag.PSObject.Properties["ready"]) { "" + $diag.ready } else { "" }
      if ($href -like "*service-now.com*" -and ($ready -eq "complete" -or $ready -eq "interactive")) {
        return $true
      }
    }
    Start-Sleep -Milliseconds 300
  }
  return $false
}

function Resolve-UserDisplayNameFromSysId {
  param($wv, [string]$UserSysId)
  if ([string]::IsNullOrWhiteSpace($UserSysId)) { return "" }
  if ($UserSysId -notmatch '^[0-9a-fA-F]{32}$') { return "" }

  $safeId = $UserSysId.Trim()
  $js = @"
(function(){
  try {
    var q = encodeURIComponent('sys_id=$safeId');
    var p = '/sys_user.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_query=' + q;
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) return JSON.stringify({ok:false});
    var o = JSON.parse(x.responseText || "{}");
    var r = (o && o.records && o.records[0]) ? o.records[0] : ((o && o.result && o.result[0]) ? o.result[0] : null);
    if (!r) return JSON.stringify({ok:false});
    var n = (r.name || r.user_name || r.display_name || "").toString().trim();
    return JSON.stringify({ok:true,name:n});
  } catch(e){
    return JSON.stringify({ok:false});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 6000)
  if ($o -and $o.ok -eq $true -and $o.PSObject.Properties["name"]) {
    return ("" + $o.name).Trim()
  }
  return ""
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

    # JS logic: detect whether we're on IdP, on SNOW, and â€œlogged inâ€ indicators.
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

function Is-InvalidUserDisplay([string]$UserText) {
  $u = ("" + $UserText).Trim()
  if (-not $u) { return $true }
  if ($u -match '^[0-9a-fA-F]{32}$') { return $true }
  if ($u -match '(?i)^new\b.*\buser$') { return $true }
  if ($u -match '(?i)^explanation$') { return $true }
  if ($u -match '(?i)^n/?a$') { return $true }
  return $false
}

function Get-CompletionStatusForExcel {
  param([string]$Ticket, $Res)
  if (-not $Res) { return "Pending" }

  if ($Ticket -like "INC*") {
    $sInc = ""
    if ($Res.PSObject.Properties["status"]) { $sInc = ("" + $Res.status) }
    if (-not $sInc -and $Res.PSObject.Properties["status_label"]) { $sInc = ("" + $Res.status_label) }
    if (-not $sInc -and $Res.PSObject.Properties["status_value"]) { $sInc = ("" + $Res.status_value) }
    $stInc = $sInc.Trim().ToLowerInvariant()
    if ($stInc -match 'hold') { return "Hold" }
  }

  if (($Ticket -like "RITM*") -or ($Ticket -like "INC*")) {
    $openTasks = 0
    if ($Res.PSObject.Properties["open_tasks"]) {
      try { $openTasks = [int]$Res.open_tasks } catch { $openTasks = 0 }
    }
    if ($Res.PSObject.Properties["open_task_items"] -and $Res.open_task_items) {
      try {
        $itemsCount = @($Res.open_task_items).Count
        if ($itemsCount -gt $openTasks) { $openTasks = $itemsCount }
      } catch {}
    }
    if ($openTasks -gt 0) { return "Pending" }

    $statusText = ""
    if ($Res.PSObject.Properties["status"]) { $statusText = ("" + $Res.status) }
    if (-not $statusText -and $Res.PSObject.Properties["status_label"]) { $statusText = ("" + $Res.status_label) }
    if (-not $statusText -and $Res.PSObject.Properties["status_value"]) { $statusText = ("" + $Res.status_value) }
    $st = $statusText.Trim().ToLowerInvariant()

    $statusValue = ""
    if ($Res.PSObject.Properties["status_value"]) { $statusValue = ("" + $Res.status_value).Trim().ToLowerInvariant() }

    if ($Ticket -like "RITM*") {
      # RITM is strict: only explicit closed numeric states are considered Complete.
      # 1=Open, 2=In Progress, 3=Closed Complete, 4=Closed Incomplete, 7=Cancelled
      if ($statusValue -in @("3", "4", "7")) { return "Complete" }
      if ($statusValue -in @("1", "2")) { return "Pending" }

      # Fallback only when numeric value is unavailable.
      if ($st -match '^(closed(\s+(complete|incomplete|skipped))?|complete|completed|resolved|cancelled|canceled)$') { return "Complete" }
      return "Pending"
    }

    if ($st -match 'closed|close|complete|completed|resolved|cancel') { return "Complete" }
    return "Pending"
  }

  if ($Res.PSObject.Properties["status"]) { return ("" + $Res.status) }
  return "Pending"
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

function Get-DashboardUserDirectory {
  param(
    [string]$ExcelPath,
    [string]$SheetName
  )

  Close-ExcelProcessesIfRequested -Reason "before dashboard user directory read"
  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -ws $ws
    $ritmCol = Resolve-DashboardRitmColumn -HeaderMap $map -Required
    $nameCols = Resolve-DashboardNameColumns -HeaderMap $map
    $requestedForCols = @($nameCols.RequestedForCols)
    $firstNameCols = @($nameCols.FirstNameCols)
    $lastNameCols = @($nameCols.LastNameCols)
    if (($requestedForCols.Count -eq 0) -and ($firstNameCols.Count -eq 0) -and ($lastNameCols.Count -eq 0)) { return @() }

    $xlUp = -4162
    $rows = [int]$ws.Cells.Item($ws.Rows.Count, $ritmCol).End($xlUp).Row
    if ($rows -lt 2) { return @() }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $out = New-Object System.Collections.Generic.List[string]
    for ($r = 2; $r -le $rows; $r++) {
      $ritm = ("" + $ws.Cells.Item($r, $ritmCol).Text).Trim().ToUpperInvariant()
      if (-not ($ritm -match '^RITM\d{6,8}$')) { continue }

      $requestedFor = ""
      foreach ($c in $requestedForCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $requestedFor = $v; break }
      }
      $firstName = ""
      foreach ($c in $firstNameCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $firstName = $v; break }
      }
      $lastName = ""
      foreach ($c in $lastNameCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $lastName = $v; break }
      }

      $name = $requestedFor
      if ([string]::IsNullOrWhiteSpace($name)) { $name = (($firstName + " " + $lastName).Trim()) }
      if ([string]::IsNullOrWhiteSpace($name)) { continue }

      if ($seen.Add($name)) { [void]$out.Add($name) }
    }

    return @($out | Sort-Object)
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Build-RitmByNumberUrl([string]$RitmNumber) {
  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return "" }
  $query = "number=" + $RitmNumber.Trim().ToUpperInvariant()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_req_item_list.do%3Fsysparm_query%3D$safeQuery"
}

function Build-RitmBestUrl([string]$SysId, [string]$RitmNumber) {
  $u = Build-RitmRecordUrl -SysId $SysId
  if (-not [string]::IsNullOrWhiteSpace($u)) { return $u }
  return Build-RitmByNumberUrl -RitmNumber $RitmNumber
}

function Build-IncidentRecordUrl([string]$SysId) {
  if ([string]::IsNullOrWhiteSpace($SysId)) { return "" }
  return [string]::Format($IncidentRecordUrlTemplate, $SysId.Trim())
}

function Build-IncidentByNumberUrl([string]$IncNumber) {
  if ([string]::IsNullOrWhiteSpace($IncNumber)) { return "" }
  $query = "number=" + $IncNumber.Trim().ToUpperInvariant()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return "$InstanceBaseUrl/nav_to.do?uri=%2Fincident_list.do%3Fsysparm_query%3D$safeQuery"
}

function Build-IncidentBestUrl([string]$SysId, [string]$IncNumber) {
  $u = Build-IncidentRecordUrl -SysId $SysId
  if (-not [string]::IsNullOrWhiteSpace($u)) { return $u }
  return Build-IncidentByNumberUrl -IncNumber $IncNumber
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

function Build-SCTaskListByNumbersUrl {
  param([string[]]$TaskNumbers)
  $nums = @($TaskNumbers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
  if ($nums.Count -eq 0) { return "" }
  if ($nums.Count -eq 1) { return Build-SCTaskFallbackUrl -TaskNumber $nums[0] }
  $query = "numberIN" + ($nums -join ",")
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_task_list.do%3Fsysparm_query%3D$safeQuery"
}

function Build-SCTaskListByRitmUrl {
  param([string]$RitmNumber)
  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return "" }
  $query = "request_item.number=" + $RitmNumber.Trim()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_task_list.do%3Fsysparm_query%3D$safeQuery"
}

function Build-SCTaskRecordByNumberUrl {
  param([string]$TaskNumber)
  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return "" }
  $query = "number=" + $TaskNumber.Trim().ToUpperInvariant()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return "$InstanceBaseUrl/nav_to.do?uri=%2Fsc_task.do%3Fsysparm_query%3D$safeQuery"
}

function Get-DetectedPiFromActivityText {
  param([string]$ActivityText)
  if (-not $EnableActivityStreamSearch) { return "" }
  if ([string]::IsNullOrWhiteSpace($ActivityText)) {
    return $(if ($WriteNotFoundText) { "Not found" } else { "" })
  }

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

  # Preference rule:
  # if short-domain PI (e.g., MUSTBRUN2420165) and long PI (02PI20...) point to same digits,
  # keep 02PI20... and drop short-domain token.
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

  if ($vals.Count -eq 0) {
    return $(if ($WriteNotFoundText) { "Not found" } else { "" })
  }
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

  # Accept only when clearly better; otherwise keep all to avoid wrong single PI.
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

function Get-LegalNameFromText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return "" }

  $patterns = @(
    '(?im)\bLegal\s*name\s*[:\-\s]*([A-Za-z''\- ]{3,})',
    '(?is)\bLegal\s*name\b[\s\S]{0,200}?\bvalue\s*=\s*["'']([^"'']{3,})["'']',
    '(?is)<input[^>]*\bvalue\s*=\s*["'']([^"'']{3,})["''][^>]*\bplaceholder\s*=\s*["'']Joe Willy Smith["'']',
    '(?im)\bLegal\s*name\b[\r\n\t ]+([A-Z][A-Za-z''\- ]{3,})'
  )

  foreach ($p in $patterns) {
    $m = [regex]::Match($Text, $p)
    if ($m.Success -and $m.Groups.Count -gt 1) {
      $v = ("" + $m.Groups[1].Value).Trim()
      if ($v -and -not (Is-InvalidUserDisplay $v)) { return $v }
    }
  }
  return ""
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

function Get-RecordActivityTextFromUiPage {
  param(
    $wv,
    [string]$RecordSysId,
    [string]$Table
  )

  if ([string]::IsNullOrWhiteSpace($RecordSysId)) { return "" }
  $recordUrl = ""
  if ($Table -eq "incident") {
    $recordUrl = Build-IncidentRecordUrl -SysId $RecordSysId
  }
  else {
    $recordUrl = Build-RitmRecordUrl -SysId $RecordSysId
  }
  if ([string]::IsNullOrWhiteSpace($recordUrl)) { return "" }

  try {
    $wv.CoreWebView2.Navigate($recordUrl)
  }
  catch {
    Log "ERROR" "Record UI navigate failed table='$Table' sys_id='$RecordSysId': $($_.Exception.Message)"
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
      try {
        var li = doc.querySelector('input[placeholder="Joe Willy Smith"]');
        var lv = s(li && (li.value || li.getAttribute('value') || ""));
        if (lv && !seen["LEGAL_NAME:"+lv]) {
          seen["LEGAL_NAME:"+lv] = true;
          out.push("LEGAL_NAME:" + lv);
        }
      } catch(el) {}
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
    Log "ERROR" "Record UI activity read failed table='$Table' sys_id='$RecordSysId': $($o.error)"
  }
  if ($o.PSObject.Properties["frame_text"]) {
    $ft = "" + $o.frame_text
    if ($VerboseTicketLogging) {
      Log "INFO" "Record UI frame text length table='$Table' sys_id='$RecordSysId': $($ft.Length) ready=$frameReady"
    }
  }
  if ($o.PSObject.Properties["shell_text"]) {
    $st = "" + $o.shell_text
    if ($VerboseTicketLogging) {
      Log "INFO" "Record UI shell text length table='$Table' sys_id='$RecordSysId': $($st.Length)"
    }
  }
  if ($o.PSObject.Properties["frame_text"]) {
    $ft = ("" + $o.frame_text)
    if (-not [string]::IsNullOrWhiteSpace($ft)) { return $ft }
  }
  if ($o.PSObject.Properties["text"]) { return ("" + $o.text) }
  return ""
}

function Get-LegalNameFromUiForm {
  param(
    $wv,
    [string]$RecordSysId,
    [string]$Table = "sc_req_item"
  )

  if ([string]::IsNullOrWhiteSpace($RecordSysId)) { return "" }
  $recordUrl = if ($Table -eq "incident") { Build-IncidentRecordUrl -SysId $RecordSysId } else { Build-RitmRecordUrl -SysId $RecordSysId }
  if ([string]::IsNullOrWhiteSpace($recordUrl)) { return "" }

  try { $wv.CoreWebView2.Navigate($recordUrl) } catch { return "" }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 12000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function extractFromDoc(doc){
      if (!doc) return '';

      function valid(v){
        var t = s(v);
        if (!t) return '';
        if (/^explanation$/i.test(t)) return '';
        return t;
      }

      // Exact known pattern from catalog variable input
      var exact = doc.querySelector('input[placeholder=\"Joe Willy Smith\"],textarea[placeholder=\"Joe Willy Smith\"]');
      var ev = valid(exact && (exact.value || (exact.getAttribute && exact.getAttribute('value')) || ''));
      if (ev) return ev;

      // Label-driven lookup: "Legal name" -> associated control via aria-labelledby / for
      var labels = doc.querySelectorAll('label, span, div');
      for (var i=0; i<labels.length; i++) {
        var lt = s(labels[i].innerText || labels[i].textContent || '');
        if (!/legal\s*name/i.test(lt)) continue;
        var lid = s(labels[i].id || '');
        var fr = s(labels[i].getAttribute && labels[i].getAttribute('for'));
        var cands = [];
        if (lid) cands = cands.concat([].slice.call(doc.querySelectorAll('[aria-labelledby=\"'+lid+'\"], [aria-describedby=\"'+lid+'\"]')));
        if (fr) cands = cands.concat([].slice.call(doc.querySelectorAll('#'+fr)));
        for (var j=0; j<cands.length; j++) {
          var v = valid(cands[j].value || cands[j].getAttribute && cands[j].getAttribute('value') || cands[j].innerText || '');
          if (v) return v;
        }
      }

      // Fallback: any ni.* input carrying Joe Willy Smith placeholder
      var vars = doc.querySelectorAll('input[id^=\"ni.\"],textarea[id^=\"ni.\"]');
      for (var vi=0; vi<vars.length; vi++) {
        var ph = s(vars[vi].getAttribute && vars[vi].getAttribute('placeholder'));
        if (!/Joe Willy Smith/i.test(ph)) continue;
        var vv = valid(vars[vi].value || (vars[vi].getAttribute && vars[vi].getAttribute('value')) || '');
        if (vv) return vv;
      }

      return '';
    }

    var out = extractFromDoc(document);
    if (out) return JSON.stringify({ok:true, legal_name:out});
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    out = extractFromDoc(fdoc);
    return JSON.stringify({ok:true, legal_name:out});
  } catch(e){
    return JSON.stringify({ok:false, legal_name:''});
  }
})();
"@
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    $o = Parse-WV2Json (ExecJS $wv $js 8000)
    if ($o -and $o.ok -eq $true -and $o.PSObject.Properties["legal_name"]) {
      $ln = ("" + $o.legal_name).Trim()
      if ($ln -and -not (Is-InvalidUserDisplay $ln)) { return $ln }
    }
    Start-Sleep -Milliseconds 400
  }
  return ""
}

function Get-RitmTaskListTextFromUiPage {
  param(
    $wv,
    [string]$RitmNumber
  )

  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return "" }
  $taskListUrl = Build-SCTaskListByRitmUrl -RitmNumber $RitmNumber
  if ([string]::IsNullOrWhiteSpace($taskListUrl)) { return "" }

  try { $wv.CoreWebView2.Navigate($taskListUrl) } catch { return "" }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 10000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function collectFromDoc(doc){
      if (!doc) return '';
      var out = [];
      var seen = {};
      var sels = [
        'table',
        '.list2_body',
        '.list2_body tr',
        '.list2 td',
        '.vt',
        '.list_decoration',
        '.linked',
        'a'
      ];
      for (var si = 0; si < sels.length; si++) {
        var nodes = doc.querySelectorAll(sels[si]);
        for (var ni = 0; ni < nodes.length; ni++) {
          var t = s(nodes[ni].innerText || nodes[ni].textContent || '');
          if (t && !seen[t]) { seen[t] = true; out.push(t); }
        }
      }
      var b = s(doc.body && (doc.body.innerText || doc.body.textContent));
      if (b && !seen[b]) { seen[b] = true; out.push(b); }
      return out.join(' ');
    }
    var shell = collectFromDoc(document);
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    var ft = collectFromDoc(fdoc);
    return JSON.stringify({ ok:true, text: ft ? ft : shell, frame_text:ft, shell_text:shell });
  } catch(e){
    return JSON.stringify({ ok:false, text:'' });
  }
})();
"@
  $o = $null
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    $o = Parse-WV2Json (ExecJS $wv $js 6000)
    $txt = if ($o -and $o.PSObject.Properties["text"]) { ("" + $o.text) } else { "" }
    if (-not [string]::IsNullOrWhiteSpace($txt)) { break }
    Start-Sleep -Milliseconds 350
  }
  if (-not $o) { return "" }
  if ($o.PSObject.Properties["text"]) { return ("" + $o.text) }
  return ""
}

function Get-TaskNumbersFromText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
  $matches = [regex]::Matches($Text, '\bSCTASK\d{6,}\b', 'IgnoreCase')
  $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($m in $matches) {
    $v = ("" + $m.Value).Trim().ToUpperInvariant()
    if ($v -and $seen.Add($v)) { [void]$out.Add($v) }
  }
  return ,@($out)
}

function Get-SCTaskRecordTextFromUiPage {
  param(
    $wv,
    [string]$TaskNumber
  )

  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return "" }
  $taskUrl = Build-SCTaskFallbackUrl -TaskNumber $TaskNumber
  if ([string]::IsNullOrWhiteSpace($taskUrl)) { return "" }

  try { $wv.CoreWebView2.Navigate($taskUrl) } catch { return "" }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 10000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function collect(doc){
      if (!doc) return '';
      var out = [];
      var seen = {};
      var sels = [
        '.sn-widget-textblock-body',
        '.sn-widget-textblock-body_formatted',
        '.sn-card-component_accent-bar',
        '.sn-card-component_accent-bar_dark',
        '.activities-form',
        'activities-form',
        'table',
        'input',
        'textarea',
        'label'
      ];
      for (var si = 0; si < sels.length; si++) {
        var nodes = doc.querySelectorAll(sels[si]);
        for (var ni = 0; ni < nodes.length; ni++) {
          var t = s(nodes[ni].innerText || nodes[ni].textContent || nodes[ni].value || (nodes[ni].getAttribute && nodes[ni].getAttribute('value')) || '');
          if (t && !seen[t]) { seen[t] = true; out.push(t); }
        }
      }
      var b = s(doc.body && (doc.body.innerText || doc.body.textContent));
      if (b && !seen[b]) out.push(b);
      return out.join(' ');
    }
    var shell = collect(document);
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    var ft = collect(fdoc);
    return JSON.stringify({ok:true, text: ft ? ft : shell});
  } catch(e){
    return JSON.stringify({ok:false, text:''});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 7000)
  if (-not $o) { return "" }
  if ($o.PSObject.Properties["text"]) { return ("" + $o.text) }
  return ""
}

function Get-SCTaskActivityTextFromUiPage {
  param(
    $wv,
    [string]$TaskNumber
  )

  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return "" }
  $taskUrl = Build-SCTaskRecordByNumberUrl -TaskNumber $TaskNumber
  if ([string]::IsNullOrWhiteSpace($taskUrl)) { return "" }

  try { $wv.CoreWebView2.Navigate($taskUrl) } catch { return "" }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 12000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function collectActivities(doc){
      if (!doc) return '';
      var out = [];
      var seen = {};
      var sels = [
        'h-card-wrapper activities-form',
        'h-card-wrapper .activities-form',
        '.activities-form',
        'activities-form',
        '.sn-widget-textblock-body',
        '.sn-widget-textblock-body_formatted',
        '.sn-card-component_accent-bar',
        '.sn-card-component_accent-bar_dark',
        '.activity-stream-text',
        '.activity-stream',
        '.journal',
        '.journal_field',
        '[data-stream-entry]'
      ];
      for (var si = 0; si < sels.length; si++) {
        var nodes = doc.querySelectorAll(sels[si]);
        for (var ni = 0; ni < nodes.length; ni++) {
          var t = s(nodes[ni].innerText || nodes[ni].textContent || '');
          if (t && !seen[t]) { seen[t] = true; out.push(t); }
        }
      }
      return out.join(' ');
    }
    var shell = collectActivities(document);
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    var ft = collectActivities(fdoc);
    return JSON.stringify({ok:true, text: ft ? ft : shell});
  } catch(e){
    return JSON.stringify({ok:false, text:''});
  }
})();
"@
  $o = $null
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    $o = Parse-WV2Json (ExecJS $wv $js 7000)
    $txt = if ($o -and $o.PSObject.Properties["text"]) { ("" + $o.text) } else { "" }
    if (-not [string]::IsNullOrWhiteSpace($txt)) { break }
    Start-Sleep -Milliseconds 350
  }
  if (-not $o) { return "" }
  if ($o.PSObject.Properties["text"]) { return ("" + $o.text) }
  return ""
}

function Get-SCTaskNumbersFromBackendByRitm {
  param(
    $wv,
    [string]$RitmNumber
  )
  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return @() }
  $safeRitm = $RitmNumber.Trim().ToUpperInvariant()
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRows(o){
      if (o && Array.isArray(o.records)) return o.records;
      if (o && Array.isArray(o.result)) return o.result;
      return [];
    }
    var q = encodeURIComponent('request_item.number=$safeRitm');
    var p = '/sc_task.do?JSONv2&sysparm_limit=200&sysparm_display_value=true&sysparm_query=' + q;
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) return JSON.stringify({ok:false, tasks:[]});
    var o = JSON.parse(x.responseText || '{}');
    var rows = pickRows(o);
    var seen = {};
    var out = [];
    for (var i=0; i<rows.length; i++) {
      var n = s(rows[i].number || '');
      if (!n) continue;
      var u = n.toUpperCase();
      if (!/^SCTASK\d{6,}$/.test(u)) continue;
      if (seen[u]) continue;
      seen[u] = true;
      out.push(u);
    }
    return JSON.stringify({ok:true, tasks:out});
  } catch(e){
    return JSON.stringify({ok:false, tasks:[]});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 7000)
  if (-not $o) { return @() }
  if (-not $o.PSObject.Properties["tasks"]) { return @() }
  return @($o.tasks)
}

# =============================================================================
# DASHBOARD: CHECK-IN / CHECK-OUT (NEW FEATURE)
# =============================================================================
function Get-ExcelHeaderMap {
  param($ws)
  $map = @{}
  $cols = $ws.UsedRange.Columns.Count
  for ($c = 1; $c -le $cols; $c++) {
    $h = ("" + $ws.Cells.Item(1, $c).Text).Trim()
    if ($h -and -not $map.ContainsKey($h)) { $map[$h] = $c }
  }
  return $map
}

function Resolve-HeaderColumn {
  param(
    [hashtable]$HeaderMap,
    [string[]]$Patterns,
    [string]$LogicalName = "",
    [switch]$Required
  )

  foreach ($k in $HeaderMap.Keys) {
    foreach ($p in $Patterns) {
      if (("" + $k) -match $p) { return [int]$HeaderMap[$k] }
    }
  }

  if ($Required) {
    throw "Dashboard missing required column for '$LogicalName'. Headers found: $($HeaderMap.Keys -join ', ')"
  }
  return $null
}

function Resolve-DashboardRitmColumn {
  param(
    [hashtable]$HeaderMap,
    [switch]$Required
  )

  foreach ($k in $HeaderMap.Keys) {
    if (("" + $k) -match '(?i)^\s*ritm\s*$') { return [int]$HeaderMap[$k] }
  }
  foreach ($k in $HeaderMap.Keys) {
    if (("" + $k) -match '(?i)^\s*request\s*item\s*$') { return [int]$HeaderMap[$k] }
  }
  foreach ($k in $HeaderMap.Keys) {
    if (("" + $k) -match '(?i)^\s*number\s*$') { return [int]$HeaderMap[$k] }
  }
  foreach ($k in $HeaderMap.Keys) {
    $hk = ("" + $k)
    if (($hk -match '(?i)\britm\b') -and ($hk -notmatch '(?i)estado|status|action|finished|state')) {
      return [int]$HeaderMap[$k]
    }
  }

  if ($Required) {
    throw "Dashboard missing required RITM column. Headers found: $($HeaderMap.Keys -join ', ')"
  }
  return $null
}

function Resolve-DashboardNameColumns {
  param([hashtable]$HeaderMap)

  $requestedForCols = @()
  $firstNameCols = @()
  $lastNameCols = @()
  foreach ($k in $HeaderMap.Keys) {
    $kText = "" + $k
    if ($kText -match '(?i)^requested\s*for$|^name$|employee|user') { $requestedForCols += [int]$HeaderMap[$k] }
    if ($kText -match '(?i)^first\s*name$') { $firstNameCols += [int]$HeaderMap[$k] }
    if ($kText -match '(?i)^last\s*name$') { $lastNameCols += [int]$HeaderMap[$k] }
  }
  return [pscustomobject]@{
    RequestedForCols = @($requestedForCols | Select-Object -Unique)
    FirstNameCols    = @($firstNameCols | Select-Object -Unique)
    LastNameCols     = @($lastNameCols | Select-Object -Unique)
  }
}

function Resolve-DashboardSctaskColumn {
  param([hashtable]$HeaderMap)

  foreach ($k in $HeaderMap.Keys) {
    $hk = ("" + $k).Trim()
    if ($hk -match '(?i)split') { continue }
    if ($hk -match '(?i)^\s*sctasks?\s*$') { return [int]$HeaderMap[$k] }
  }
  foreach ($k in $HeaderMap.Keys) {
    $hk = ("" + $k).Trim()
    if ($hk -match '(?i)split') { continue }
    if ($hk -match '(?i)^sc\s*task(s)?\s*$') { return [int]$HeaderMap[$k] }
  }
  foreach ($k in $HeaderMap.Keys) {
    $hk = ("" + $k).Trim()
    if ($hk -match '(?i)split') { continue }
    if ($hk -match '(?i)\bsctasks?\b|\bsc\s*task\b') { return [int]$HeaderMap[$k] }
  }
  return $null
}

function Ensure-DashboardExcelColumns {
  param(
    [string]$ExcelPath,
    [string]$SheetName
  )

  Log "INFO" "Dashboard: ensuring required Excel columns for '$ExcelPath' sheet '$SheetName'"
  Close-ExcelProcessesIfRequested -Reason "before dashboard ensure columns"

  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $false)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -ws $ws
    $ritmCol = Resolve-DashboardRitmColumn -HeaderMap $map -Required
    $nameCols = Resolve-DashboardNameColumns -HeaderMap $map
    $requestedForCols = @($nameCols.RequestedForCols)
    $firstNameCols = @($nameCols.FirstNameCols)
    $lastNameCols = @($nameCols.LastNameCols)
    if (($requestedForCols.Count -eq 0) -and ($firstNameCols.Count -eq 0) -and ($lastNameCols.Count -eq 0)) {
      throw "Dashboard requires one of: Requested for, Name, First Name/Last Name."
    }

    $statusCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Dashboard Status"
    $presentCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Present Time"
    $closedCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Closed Time"
    $sctaskCol = Resolve-DashboardSctaskColumn -HeaderMap $map

    $wb.Save()

    return [pscustomobject]@{
      RITM          = $ritmCol
      RequestedFor  = if ($requestedForCols.Count -gt 0) { [int]$requestedForCols[0] } else { $null }
      FirstName     = if ($firstNameCols.Count -gt 0) { [int]$firstNameCols[0] } else { $null }
      LastName      = if ($lastNameCols.Count -gt 0) { [int]$lastNameCols[0] } else { $null }
      SCTASK        = $sctaskCol
      Status        = $statusCol
      PresentTime   = $presentCol
      ClosedTime    = $closedCol
    }
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Search-DashboardRows {
  param(
    [string]$ExcelPath,
    [string]$SheetName,
    [string]$SearchText
  )

  $query = ("" + $SearchText).Trim()

  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -ws $ws
    $ritmCol = Resolve-DashboardRitmColumn -HeaderMap $map -Required
    $nameCols = Resolve-DashboardNameColumns -HeaderMap $map
    $requestedForCols = @($nameCols.RequestedForCols)
    $firstNameCols = @($nameCols.FirstNameCols)
    $lastNameCols = @($nameCols.LastNameCols)
    if (($requestedForCols.Count -eq 0) -and ($firstNameCols.Count -eq 0) -and ($lastNameCols.Count -eq 0)) {
      throw "Dashboard requires one of: Requested for, Name, First Name/Last Name."
    }

    $sctaskCol = Resolve-DashboardSctaskColumn -HeaderMap $map
    $sctaskSplitCol = Resolve-HeaderColumn -HeaderMap $map -Patterns @('(?i)^sctask\s*split$')
    $statusCol = Resolve-HeaderColumn -HeaderMap $map -Patterns @('(?i)^dashboard\s*status$')
    $presentCol = Resolve-HeaderColumn -HeaderMap $map -Patterns @('(?i)^present\s*time$')
    $closedCol = Resolve-HeaderColumn -HeaderMap $map -Patterns @('(?i)^closed\s*time$')
    $pdfCol = Resolve-HeaderColumn -HeaderMap $map -Patterns @('(?i)^pdf\s*file$')

    $used = $ws.UsedRange
    $rows = 0
    try { $rows = [int]($used.Row + $used.Rows.Count - 1) } catch { $rows = 0 }
    $cols = 0
    try { $cols = [int]$used.Columns.Count } catch { $cols = 0 }
    if ($rows -lt 2) { return @() }
    if ($cols -lt 1) { $cols = [int]$ws.UsedRange.Columns.Count }

    $out = New-Object System.Collections.Generic.List[object]
    $seenRitm = @{}
    for ($r = 2; $r -le $rows; $r++) {
      $ritm = ("" + $ws.Cells.Item($r, $ritmCol).Text).Trim().ToUpperInvariant()
      $splitFlag = if ($sctaskSplitCol) { ("" + $ws.Cells.Item($r, $sctaskSplitCol).Text).Trim().ToUpperInvariant() } else { "" }
      if ($splitFlag -eq "AUTO") { continue }
      if (($ritm -match '^RITM\d{6,8}$') -and $seenRitm.ContainsKey($ritm)) { continue }

      # Collect full row text so search works even when name headers vary in Excel.
      $rowTexts = New-Object System.Collections.Generic.List[string]
      for ($c = 1; $c -le $cols; $c++) {
        $cv = ("" + $ws.Cells.Item($r, $c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($cv)) { [void]$rowTexts.Add($cv) }
      }
      $rowBlob = ($rowTexts -join " ")
      if ([string]::IsNullOrWhiteSpace($rowBlob)) { continue }
      if ((-not [string]::IsNullOrWhiteSpace($query)) -and ($rowBlob.IndexOf($query, [System.StringComparison]::OrdinalIgnoreCase) -lt 0)) { continue }

      $requestedFor = ""
      foreach ($c in $requestedForCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $requestedFor = $v; break }
      }
      $firstName = ""
      foreach ($c in $firstNameCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $firstName = $v; break }
      }
      $lastName = ""
      foreach ($c in $lastNameCols) {
        $v = ("" + $ws.Cells.Item($r, [int]$c).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $lastName = $v; break }
      }
      if ([string]::IsNullOrWhiteSpace($requestedFor)) {
        $requestedFor = (($firstName + " " + $lastName).Trim())
      }
      if ([string]::IsNullOrWhiteSpace($requestedFor)) {
        foreach ($txt in $rowTexts) {
          if ($txt -match '^(?i)(RITM|SCTASK|INC)\d{6,8}$') { continue }
          if ($txt -match '^\d{4}-\d{2}-\d{2}') { continue }
          if ($txt -match '^\d{1,2}:\d{2}(:\d{2})?$') { continue }
          if ($txt.Length -lt 3) { continue }
          $requestedFor = $txt
          break
        }
      }
      $sctask = if ($sctaskCol) { ("" + $ws.Cells.Item($r, $sctaskCol).Text).Trim() } else { "" }

      $status = if ($statusCol) { ("" + $ws.Cells.Item($r, $statusCol).Text).Trim() } else { "" }
      $presentTime = ""
      if ($presentCol) {
        try {
          $pv = $ws.Cells.Item($r, $presentCol).Value2
          if ($pv -is [double] -or $pv -is [int]) {
            $presentTime = ([datetime]::FromOADate([double]$pv)).ToString("yyyy-MM-dd HH:mm")
          }
          else {
            $pt = ("" + $ws.Cells.Item($r, $presentCol).Text).Trim()
            if ($pt -eq "########") { $pt = "" }
            $tmpDt = $null
            if ([datetime]::TryParse($pt, [ref]$tmpDt)) { $presentTime = $tmpDt.ToString("yyyy-MM-dd HH:mm") } else { $presentTime = $pt }
          }
        } catch {}
      }
      $closedTime = ""
      if ($closedCol) {
        try {
          $cv = $ws.Cells.Item($r, $closedCol).Value2
          if ($cv -is [double] -or $cv -is [int]) {
            $closedTime = ([datetime]::FromOADate([double]$cv)).ToString("yyyy-MM-dd HH:mm")
          }
          else {
            $ct = ("" + $ws.Cells.Item($r, $closedCol).Text).Trim()
            if ($ct -eq "########") { $ct = "" }
            $tmpDt2 = $null
            if ([datetime]::TryParse($ct, [ref]$tmpDt2)) { $closedTime = $tmpDt2.ToString("yyyy-MM-dd HH:mm") } else { $closedTime = $ct }
          }
        } catch {}
      }
      $pdfPath = if ($pdfCol) { ("" + $ws.Cells.Item($r, $pdfCol).Text).Trim() } else { "" }
      $lastUpdated = ""
      if (-not [string]::IsNullOrWhiteSpace($closedTime)) {
        $lastUpdated = $closedTime
      }
      elseif (-not [string]::IsNullOrWhiteSpace($presentTime)) {
        $lastUpdated = $presentTime
      }

      $out.Add([pscustomobject]@{
        Row           = [int]$r
        RequestedFor  = $requestedFor
        FirstName     = $firstName
        LastName      = $lastName
        RITM          = $ritm
        SctaskSplit   = $splitFlag
        SCTASK        = $sctask
        DashboardStatus = $status
        PresentTime   = $presentTime
        ClosedTime    = $closedTime
        LastUpdated   = $lastUpdated
        PdfPath       = $pdfPath
      }) | Out-Null
      if ($ritm -match '^RITM\d{6,8}$') { $seenRitm[$ritm] = $true }
    }

    Log "INFO" "Dashboard search query='$query' => matches=$($out.Count)"
    return @($out.ToArray())
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Update-DashboardExcelRow {
  param(
    [string]$ExcelPath,
    [string]$SheetName,
    [int]$RowIndex,
    [string]$DashboardStatus,
    [string]$TimestampHeader,
    [string]$TaskNumberToWrite
  )

  Close-ExcelProcessesIfRequested -Reason "before dashboard write"
  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $false)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -ws $ws
    $statusCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Dashboard Status"
    $presentCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Present Time"
    $closedCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "Closed Time"
    $sctaskCol = Resolve-DashboardSctaskColumn -HeaderMap $map

    $ws.Cells.Item($RowIndex, $statusCol) = $DashboardStatus

    $timestampCol = if ($TimestampHeader -eq "Present Time") { $presentCol } elseif ($TimestampHeader -eq "Closed Time") { $closedCol } else { $null }
    if ($timestampCol) {
      $currentTs = ("" + $ws.Cells.Item($RowIndex, $timestampCol).Text).Trim()
      if ([string]::IsNullOrWhiteSpace($currentTs)) {
        $ws.Cells.Item($RowIndex, $timestampCol) = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
      }
      else {
        Log "INFO" "Dashboard timestamp preserved at row=$RowIndex col='$TimestampHeader' value='$currentTs'"
      }
    }

    if ($sctaskCol -and -not [string]::IsNullOrWhiteSpace($TaskNumberToWrite)) {
      $existingTask = ("" + $ws.Cells.Item($RowIndex, $sctaskCol).Text).Trim()
      if ([string]::IsNullOrWhiteSpace($existingTask)) {
        $ws.Cells.Item($RowIndex, $sctaskCol) = $TaskNumberToWrite
      }
    }

    $wb.Save()
    return $true
  }
  catch {
    Log "ERROR" "Dashboard Excel update failed at row=${RowIndex}: $($_.Exception.Message)"
    return $false
  }
  finally {
    try { if ($wb) { $wb.Close($false) | Out-Null } } catch {}
    try { if ($excel) { $excel.Quit() | Out-Null } } catch {}
    try { if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } } catch {}
    try { if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } } catch {}
    try { if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Get-SCTaskCandidatesForRitm {
  param(
    $wv,
    [string]$RitmNumber
  )

  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return @() }
  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)
  $safeRitm = $RitmNumber.Trim().ToUpperInvariant()
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRows(o){
      if (o && Array.isArray(o.records)) return o.records;
      if (o && Array.isArray(o.result)) return o.result;
      return [];
    }
    var p = '/sc_task.do?JSONv2&sysparm_limit=200&sysparm_display_value=true&sysparm_query=' + encodeURIComponent('request_item.number=$safeRitm');
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) return JSON.stringify({ok:false, tasks:[]});
    var o = JSON.parse(x.responseText || '{}');
    var rows = pickRows(o);
    var out = [];
    for (var i=0; i<rows.length; i++) {
      var r = rows[i] || {};
      var num = s(r.number || '');
      if (!num) continue;
      out.push({
        number: num.toUpperCase(),
        sys_id: s(r.sys_id || ''),
        state: s(r.state || ''),
        state_value: s(r.state_value || ''),
        short_description: s(r.short_description || ''),
        sys_updated_on: s(r.sys_updated_on || '')
      });
    }
    return JSON.stringify({ok:true, tasks:out});
  } catch(e){
    return JSON.stringify({ok:false, tasks:[]});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 9000)
  if (-not $o -or $o.ok -ne $true -or -not $o.PSObject.Properties["tasks"]) { return @() }
  return @($o.tasks)
}

function Test-SCTaskClosedState {
  param(
    [string]$StateValue,
    [string]$StateLabel
  )
  $sv = ("" + $StateValue).Trim().ToLowerInvariant()
  $sl = ("" + $StateLabel).Trim().ToLowerInvariant()

  if ($sv -in @("3", "4", "7")) { return $true }
  if ($sv -match 'closed|complete|cancel') { return $true }
  if ($sl -match 'closed|complete|cancel') { return $true }
  return $false
}

function Get-SCTaskByNumber {
  param(
    $wv,
    [string]$TaskNumber
  )
  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return $null }
  $safeTask = $TaskNumber.Trim().ToUpperInvariant()
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRec(o){
      return (o && o.records && o.records[0]) ? o.records[0] :
             (o && o.result && o.result[0]) ? o.result[0] : null;
    }
    var p = '/sc_task.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_query=' + encodeURIComponent('number=$safeTask');
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) return JSON.stringify({ok:false});
    var o = JSON.parse(x.responseText || '{}');
    var r = pickRec(o);
    if (!r) return JSON.stringify({ok:false});
    return JSON.stringify({
      ok:true,
      number:s(r.number || ''),
      sys_id:s(r.sys_id || ''),
      state:s(r.state || ''),
      state_value:s(r.state_value || ''),
      short_description:s(r.short_description || ''),
      sys_updated_on:s(r.sys_updated_on || '')
    });
  } catch(e){
    return JSON.stringify({ok:false});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 7000)
  if (-not $o -or $o.ok -ne $true) { return $null }
  return $o
}

function Get-OpenSCTasksForRitmFallback {
  param(
    $wv,
    [string]$RitmNumber
  )

  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return @() }

  # 1) Backend relation query (fast path)
  $candidates = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  $open = New-Object System.Collections.Generic.List[object]
  foreach ($t in $candidates) {
    $num = if ($t.PSObject.Properties["number"]) { ("" + $t.number).Trim().ToUpperInvariant() } else { "" }
    $sid = if ($t.PSObject.Properties["sys_id"]) { ("" + $t.sys_id).Trim() } else { "" }
    $st = if ($t.PSObject.Properties["state"]) { ("" + $t.state).Trim() } else { "" }
    $sv = if ($t.PSObject.Properties["state_value"]) { ("" + $t.state_value).Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($num)) { continue }
    if (-not (Test-SCTaskClosedState -StateValue $sv -StateLabel $st)) {
      $open.Add([pscustomobject]@{
        number = $num
        sys_id = $sid
        state_value = $sv
        state_label = $st
      }) | Out-Null
    }
  }
  if ($open.Count -gt 0) { return @($open.ToArray()) }

  # 2) UI list fallback: discover task numbers visible in list
  $taskListText = Get-RitmTaskListTextFromUiPage -wv $wv -RitmNumber $ritm
  $taskNumbers = @(Get-TaskNumbersFromText -Text $taskListText)
  if ($taskNumbers.Count -eq 0) { return @() }

  $openUi = New-Object System.Collections.Generic.List[object]
  foreach ($n in $taskNumbers) {
    $row = Get-SCTaskByNumber -wv $wv -TaskNumber $n
    if ($row) {
      $st = if ($row.PSObject.Properties["state"]) { ("" + $row.state).Trim() } else { "" }
      $sv = if ($row.PSObject.Properties["state_value"]) { ("" + $row.state_value).Trim() } else { "" }
      if (-not (Test-SCTaskClosedState -StateValue $sv -StateLabel $st)) {
        $openUi.Add([pscustomobject]@{
          number = ("" + $row.number).Trim().ToUpperInvariant()
          sys_id = ("" + $row.sys_id).Trim()
          state_value = $sv
          state_label = $st
        }) | Out-Null
      }
    }
  }
  return @($openUi.ToArray())
}

function Test-DashboardTaskStateOpen {
  param(
    [string]$StateText,
    [string]$StateValue
  )
  $s = ("" + $StateText).Trim().ToLowerInvariant()
  $v = ("" + $StateValue).Trim().ToLowerInvariant()
  if ($v -eq "1") { return $true }
  if ($s -eq "1") { return $true }
  if ($s -match '^\s*open\s*$|^\s*new\s*$') { return $true }
  return $false
}

function Test-DashboardTaskStateInProgress {
  param(
    [string]$StateText,
    [string]$StateValue
  )
  $s = ("" + $StateText).Trim().ToLowerInvariant()
  $v = ("" + $StateValue).Trim().ToLowerInvariant()
  if ($v -eq "2") { return $true }
  if ($s -eq "2") { return $true }
  return ($s -match 'work\s*in\s*progress|in\s*progress')
}

function Invoke-ServiceNowDomUpdate {
  param(
    $wv,
    [string]$Table,
    [string]$SysId,
    [string]$TargetStateLabel,
    [string]$WorkNote
  )

  if ([string]::IsNullOrWhiteSpace($SysId)) { return $false }
  $recordUrl = if ($Table -eq "sc_req_item") { Build-RitmRecordUrl -SysId $SysId } else { Build-SCTaskRecordUrl -SysId $SysId }
  if ([string]::IsNullOrWhiteSpace($recordUrl)) { return $false }

  try { $wv.CoreWebView2.Navigate($recordUrl) } catch { return $false }
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 15000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $wv "document.readyState==='complete'" 2000)
    if ($isReady -eq $true) { break }
  }

  # Wait for inner SNOW form controls (including iframe) before trying to update state.
  $probeJs = @"
(function(){
  try{
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function hasState(doc){
      if (!doc) return false;
      var sels = [
        'select#sc_task\\.state',
        'select#sc_req_item\\.state',
        'select[name=\"sc_task.state\"]',
        'select[name=\"sc_req_item.state\"]',
        'select[name=\"state\"]',
        'select[id$=\".state\"]'
      ];
      for (var i=0; i<sels.length; i++) { if (doc.querySelector(sels[i])) return true; }
      var all = doc.querySelectorAll('select');
      for (var j=0; j<all.length; j++) {
        var idn = s(all[j].id) + ' ' + s(all[j].name);
        if (/state/i.test(idn)) return true;
      }
      return false;
    }
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    var ok = hasState(document) || hasState(fdoc);
    return JSON.stringify({ok:ok});
  } catch(e){ return JSON.stringify({ok:false}); }
})();
"@
  $probeSw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($probeSw.ElapsedMilliseconds -lt 12000) {
    $po = Parse-WV2Json (ExecJS $wv $probeJs 2500)
    if ($po -and $po.ok -eq $true) { break }
    Start-Sleep -Milliseconds 300
  }

  $targetJson = ($TargetStateLabel | ConvertTo-Json -Compress)
  $noteJson = ($WorkNote | ConvertTo-Json -Compress)
  $sysIdJson = ($SysId | ConvertTo-Json -Compress)
  $js = @"
(function(){
  try {
    var target = $targetJson;
    var note = $noteJson;
    var table = '$Table';
    var targetSysId = $sysIdJson;
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function norm(x){ return s(x).toLowerCase().replace(/[\s_-]+/g,' ').trim(); }
    function targetAlias(label){
      var t = norm(label);
      if (t.indexOf('appoint') >= 0 || t.indexOf('appoin') >= 0) return 'appointment';
      if (t.indexOf('work in progress') >= 0) return 'wip';
      if (t.indexOf('closed complete') >= 0) return 'closed_complete';
      return t;
    }
    function stateLooksLikeTarget(targetLabel, currentLabel, currentValue){
      var t = targetAlias(targetLabel);
      var cl = norm(currentLabel);
      var cv = norm(currentValue);
      if (!t) return false;
      if (t === 'appointment') {
        return (cl.indexOf('appoint') >= 0 || cl.indexOf('appoin') >= 0);
      }
      if (t === 'wip') {
        return (cl.indexOf('work in progress') >= 0 || cv === '2');
      }
      if (t === 'closed_complete') {
        return (cl.indexOf('closed complete') >= 0 || cv === '3');
      }
      return (cl === t || cl.indexOf(t) >= 0 || cv === t);
    }
    function resolveChoiceTargetValue(tableName, targetLabel){
      try {
        var p = '/sys_choice.do?JSONv2&sysparm_limit=200&sysparm_query=' + encodeURIComponent('name=' + tableName + '^element=state');
        var x = new XMLHttpRequest();
        x.open('GET', p, false);
        x.withCredentials = true;
        x.send(null);
        if (!(x.status>=200 && x.status<300)) return '';
        var o = JSON.parse(x.responseText || '{}');
        var rows = (o && o.records) ? o.records : ((o && o.result) ? o.result : []);
        if (!Array.isArray(rows)) return '';
        var t = targetAlias(targetLabel);
        for (var i=0; i<rows.length; i++) {
          var lbl = norm(rows[i].label || '');
          if (t === 'appointment' && (lbl.indexOf('appoint') >= 0 || lbl.indexOf('appoin') >= 0)) return s(rows[i].value || '');
          if (t === 'wip' && lbl.indexOf('work in progress') >= 0) return s(rows[i].value || '');
          if (t === 'closed_complete' && lbl.indexOf('closed complete') >= 0) return s(rows[i].value || '');
          if (lbl === norm(targetLabel)) return s(rows[i].value || '');
        }
      } catch(e){}
      return '';
    }
    function docs(){
      var out = [];
      var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
      if (frame && frame.contentDocument) out.push(frame.contentDocument);
      out.push(document);
      return out;
    }
    function findState(doc){
      var sels = [
        'select#sc_task\\.state',
        'select#sc_req_item\\.state',
        'select[name=\"sc_task.state\"]',
        'select[name=\"sc_req_item.state\"]',
        'select[name=\"state\"]',
        'select[id$=\".state\"]'
      ];
      for (var i=0; i<sels.length; i++) {
        var el = doc.querySelector(sels[i]);
        if (el) return el;
      }
      var all = doc.querySelectorAll('select');
      for (var j=0; j<all.length; j++) {
        var idn = s(all[j].id) + ' ' + s(all[j].name);
        if (/state/i.test(idn)) return all[j];
      }
      return null;
    }
    function setState(el, stateLabel){
      if (!el || !stateLabel) return true;
      var wanted = norm(stateLabel);
      var wantedAlias = targetAlias(stateLabel);
      var wantedValue = resolveChoiceTargetValue(table, stateLabel);
      var best = null;
      for (var i=0; i<el.options.length; i++) {
        var t = norm(el.options[i].text);
        if (t === wanted) { best = el.options[i].value; break; }
      }
      if (best === null) {
        for (var j=0; j<el.options.length; j++) {
          var t2 = norm(el.options[j].text);
          if (t2.indexOf(wanted) >= 0) { best = el.options[j].value; break; }
        }
      }
      if (best === null && wantedAlias === 'appointment') {
        for (var k=0; k<el.options.length; k++) {
          var t3 = norm(el.options[k].text);
          if (t3.indexOf('appoint') >= 0 || t3.indexOf('appoin') >= 0) { best = el.options[k].value; break; }
        }
      }
      if (best === null && wantedValue) {
        for (var z=0; z<el.options.length; z++) {
          if (s(el.options[z].value) === s(wantedValue)) { best = el.options[z].value; break; }
        }
      }
      if (best === null) return false;
      el.value = best;
      el.dispatchEvent(new Event('change', { bubbles:true }));
      return true;
    }
    function verifyPersistedState(sysId, tableName, targetLabel){
      try {
        if (!sysId || !tableName || !targetLabel) return {ok:false, reason:'missing_inputs'};
        var lastLabel = '';
        var lastValue = '';
        for (var a=0; a<4; a++) {
          var p = '/' + tableName + '.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_query=' + encodeURIComponent('sys_id=' + sysId);
          var x = new XMLHttpRequest();
          x.open('GET', p, false);
          x.withCredentials = true;
          x.send(null);
          if (x.status>=200 && x.status<300) {
            var o = {};
            try { o = JSON.parse(x.responseText || '{}'); } catch(e0) { o = {}; }
            var r = (o && o.records && o.records[0]) ? o.records[0] : ((o && o.result && o.result[0]) ? o.result[0] : null);
            if (r) {
              lastValue = s(r.state_value || r.state || '');
              lastLabel = s(r.state_label || '');
              if (!lastLabel) {
                try { lastLabel = s(resolveStateLabel(tableName, lastValue)); } catch(e1) {}
              }
              if (stateLooksLikeTarget(targetLabel, lastLabel, lastValue)) {
                return {ok:true, state_label:lastLabel, state_value:lastValue};
              }
            }
          }
        }
        return {ok:false, reason:'state_not_persisted', state_label:lastLabel, state_value:lastValue};
      } catch(e){
        return {ok:false, reason:'verify_exception', error:''+e};
      }
    }
    function readDomCurrentState(){
      try {
        function stateFromDoc(doc){
          if (!doc) return {label:'', value:''};
          var sels = [
            'select#sc_task\\.state',
            'select#sc_req_item\\.state',
            'select[name=\"sc_task.state\"]',
            'select[name=\"sc_req_item.state\"]',
            'select[name=\"state\"]',
            'select[id$=\".state\"]'
          ];
          var el = null;
          for (var i=0; i<sels.length; i++) { el = doc.querySelector(sels[i]); if (el) break; }
          if (!el) {
            var all = doc.querySelectorAll('select');
            for (var j=0; j<all.length; j++) {
              var idn = s(all[j].id) + ' ' + s(all[j].name);
              if (/state/i.test(idn)) { el = all[j]; break; }
            }
          }
          if (!el) return {label:'', value:''};
          var lbl = '';
          try {
            var idx = (typeof el.selectedIndex === 'number') ? el.selectedIndex : -1;
            if (idx >= 0 && el.options && el.options[idx]) lbl = s(el.options[idx].text || '');
          } catch(e0) {}
          return {label:lbl, value:s(el.value || '')};
        }
        var d1 = stateFromDoc(document);
        if (d1.label || d1.value) return d1;
        var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
        var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
        return stateFromDoc(fdoc);
      } catch(e){
        return {label:'', value:''};
      }
    }
    function findNotes(doc){
      var sels = [
        'textarea#activity-stream-work_notes-textarea',
        'textarea[name=\"work_notes\"]',
        'textarea[id*=\"work_notes\"]',
        'textarea[id$=\".work_notes\"]',
        'textarea[aria-label*=\"Work notes\"]'
      ];
      for (var i=0; i<sels.length; i++) {
        var el = doc.querySelector(sels[i]);
        if (el) return el;
      }
      return null;
    }
    function setNotes(el, txt){
      if (!txt) return true;
      if (!el) return false;
      if (typeof el.value !== 'undefined') {
        el.value = txt;
      } else if (el.isContentEditable) {
        el.innerText = txt;
      } else {
        return false;
      }
      el.dispatchEvent(new Event('input', { bubbles:true }));
      el.dispatchEvent(new Event('change', { bubbles:true }));
      return true;
    }
    function clickSave(doc){
      var sels = [
        '#sysverb_update',
        '#sysverb_update_and_stay',
        'button#sysverb_update',
        'button[name=\"sysverb_update\"]',
        'input#sysverb_update',
        'input[name=\"sysverb_update\"]',
        'button#sysverb_update_and_stay',
        'button[name=\"sysverb_update_and_stay\"]',
        'button.activity-submit',
        'button[data-action=\"save\"]',
        'button[aria-label*=\"Update\" i]',
        'button[title*=\"Update\" i]',
        'input[value=\"Update\"]'
      ];
      for (var i=0; i<sels.length; i++) {
        var btn = doc.querySelector(sels[i]);
        if (!btn) continue;
        btn.click();
        return true;
      }
      return false;
    }
    function tryGFormUpdate(doc){
      try {
        if (!doc || !doc.defaultView) return {ok:false};
        var w = doc.defaultView;
        var gf = w.g_form;
        if (!gf) return {ok:false};
        var ctl = null;
        try { ctl = gf.getControl('state'); } catch(e0) { ctl = null; }
        if (!ctl || !ctl.options || ctl.options.length===0) return {ok:false};
        var wanted = norm(target);
        var best = null;
        var wantedAlias = targetAlias(target);
        var wantedValue = resolveChoiceTargetValue(table, target);
        for (var i=0; i<ctl.options.length; i++) {
          var t = norm(ctl.options[i].text);
          if (t === wanted || t.indexOf(wanted) >= 0) { best = s(ctl.options[i].value); break; }
        }
        if (!best && wantedAlias === 'appointment') {
          for (var j=0; j<ctl.options.length; j++) {
            var t2 = norm(ctl.options[j].text);
            if (t2.indexOf('appoint') >= 0 || t2.indexOf('appoin') >= 0) { best = s(ctl.options[j].value); break; }
          }
        }
        if (!best && wantedValue) {
          for (var k=0; k<ctl.options.length; k++) {
            if (s(ctl.options[k].value) === s(wantedValue)) { best = s(ctl.options[k].value); break; }
          }
        }
        if (!best) return {ok:false};
        try { gf.setValue('state', best); } catch(e1) { return {ok:false}; }
        if (note) {
          try { gf.setValue('work_notes', note); } catch(e2) {}
        }
        try {
          if (typeof w.gsftSubmit === 'function') {
            w.gsftSubmit(null, gf.getFormElement(), 'sysverb_update');
            return {ok:true};
          }
        } catch(e3) {}
        try {
          if (typeof gf.save === 'function') { gf.save(); return {ok:true}; }
        } catch(e4) {}
        return {ok:false};
      } catch(e){
        return {ok:false};
      }
    }
    var stateApplied = false;
    var notesApplied = (note === '');
    var saved = false;
    var ds = docs();
    var gformSaved = false;
    for (var di=0; di<ds.length; di++) {
      var d = ds[di];
      if (!gformSaved) {
        var g = tryGFormUpdate(d);
        if (g && g.ok) { gformSaved = true; break; }
      }
      if (!stateApplied) {
        var st = findState(d);
        if (st) stateApplied = setState(st, target);
      }
      if (!notesApplied) {
        var wn = findNotes(d);
        if (wn) notesApplied = setNotes(wn, note);
      }
      if ((stateApplied && notesApplied) && !saved) {
        saved = clickSave(d);
      }
      if (stateApplied && notesApplied && saved) break;
    }
    var okFinal = gformSaved || (stateApplied && notesApplied && saved);
    var verify = {ok:false, reason:'not_run'};
    if (okFinal) {
      verify = verifyPersistedState(targetSysId, table, target);
      if (!verify.ok) {
        var domState = readDomCurrentState();
        if (stateLooksLikeTarget(target, domState.label, domState.value)) {
          verify.ok = true;
          verify.reason = 'dom_state_match';
          verify.state_label = s(domState.label || '');
          verify.state_value = s(domState.value || '');
        } else {
          okFinal = false;
        }
      }
    }
    return JSON.stringify({
      ok:okFinal,
      state_applied:stateApplied,
      notes_applied:notesApplied,
      saved:saved,
      gform_saved:gformSaved,
      verify_ok:verify.ok,
      verify_reason:s(verify.reason || ''),
      verify_state_label:s(verify.state_label || ''),
      verify_state_value:s(verify.state_value || ''),
      verify_error:s(verify.error || '')
    });
  } catch(e){
    return JSON.stringify({ok:false, error:''+e});
  }
})();
"@
  $o = $null
  for ($attempt = 1; $attempt -le 3; $attempt++) {
    $o = Parse-WV2Json (ExecJS $wv $js 12000)
    if ($o -and $o.ok -eq $true) { break }
    Start-Sleep -Milliseconds 900
  }
  if (-not $o -or $o.ok -ne $true) {
    $detail = ""
    if ($o -and $o.PSObject.Properties["error"]) { $detail = "" + $o.error }
    if (-not $detail -and $o) {
      $sa = if ($o.PSObject.Properties["state_applied"]) { "" + $o.state_applied } else { "n/a" }
      $na = if ($o.PSObject.Properties["notes_applied"]) { "" + $o.notes_applied } else { "n/a" }
      $sv = if ($o.PSObject.Properties["saved"]) { "" + $o.saved } else { "n/a" }
      $gf = if ($o.PSObject.Properties["gform_saved"]) { "" + $o.gform_saved } else { "n/a" }
      $vr = if ($o.PSObject.Properties["verify_reason"]) { "" + $o.verify_reason } else { "n/a" }
      $vl = if ($o.PSObject.Properties["verify_state_label"]) { "" + $o.verify_state_label } else { "" }
      $vv = if ($o.PSObject.Properties["verify_state_value"]) { "" + $o.verify_state_value } else { "" }
      $detail = "state_applied=$sa notes_applied=$na saved=$sv gform_saved=$gf verify_reason=$vr verify_state='$vl/$vv'"
    }
    if (-not $detail) { $detail = "unknown dom update failure" }
    Log "ERROR" "Dashboard ServiceNow DOM update failed table='$Table' sys_id='$SysId' state='$TargetStateLabel' detail='$detail'"
    return $false
  }
  Start-Sleep -Milliseconds 1200
  return $true
}

function Invoke-DashboardCheckIn {
  param(
    $wv,
    [string]$ExcelPath,
    [string]$SheetName,
    $RowItem,
    [string]$WorkNote
  )

  $ritm = ("" + $RowItem.RITM).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return [pscustomobject]@{ ok = $false; message = "Selected row has no RITM." } }

  Log "INFO" "Dashboard CHECK-IN started for $ritm row=$($RowItem.Row)"
  $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  if ($tasks.Count -eq 0) {
    $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm)
  }
  if ($tasks.Count -eq 0) { return [pscustomobject]@{ ok = $false; message = "No SCTASK found for $ritm." } }

  $openTasks = @($tasks | Where-Object {
    $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
    $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
    Test-DashboardTaskStateOpen -StateText $st -StateValue $sv
  })
  if ($openTasks.Count -eq 0) {
    $diag = @($tasks | ForEach-Object {
      $n = if ($_.PSObject.Properties["number"]) { ("" + $_.number) } else { "" }
      $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
      $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
      "$n[$st/$sv]"
    })
    Log "INFO" "Dashboard CHECK-IN no open task for $ritm candidates=$($diag -join ', ')"
    return [pscustomobject]@{ ok = $false; message = "No task in state 'Open' for $ritm." }
  }
  $task = $openTasks[0]

  if (-not (Invoke-ServiceNowDomUpdate -wv $wv -Table "sc_task" -SysId ("" + $task.sys_id) -TargetStateLabel "Work in Progress" -WorkNote $WorkNote)) {
    return [pscustomobject]@{ ok = $false; message = "ServiceNow update failed for task $($task.number)." }
  }

  $excelOk = Update-DashboardExcelRow -ExcelPath $ExcelPath -SheetName $SheetName -RowIndex ([int]$RowItem.Row) `
    -DashboardStatus "Checked-In" -TimestampHeader "Present Time" -TaskNumberToWrite ("" + $task.number)
  if (-not $excelOk) { return [pscustomobject]@{ ok = $false; message = "ServiceNow updated, but Excel write failed." } }

  Log "INFO" "Dashboard CHECK-IN completed for $ritm task=$($task.number)"
  return [pscustomobject]@{ ok = $true; message = "Checked-In: $ritm ($($task.number))" }
}

function Get-DashboardCheckInCandidate {
  param(
    $wv,
    [string]$RitmNumber
  )
  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return $null }

  $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  if ($tasks.Count -eq 0) { $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm) }
  if ($tasks.Count -eq 0) { return $null }

  foreach ($t in $tasks) {
    $st = if ($t.PSObject.Properties["state"]) { ("" + $t.state) } elseif ($t.PSObject.Properties["state_label"]) { ("" + $t.state_label) } else { "" }
    $sv = if ($t.PSObject.Properties["state_value"]) { ("" + $t.state_value) } else { "" }
    if (Test-DashboardTaskStateOpen -StateText $st -StateValue $sv) {
      return [pscustomobject]@{
        number = if ($t.PSObject.Properties["number"]) { ("" + $t.number).Trim().ToUpperInvariant() } else { "" }
        state_text = ("" + $st).Trim()
        state_value = ("" + $sv).Trim()
      }
    }
  }
  return $null
}

function Get-DashboardCheckOutCandidate {
  param(
    $wv,
    [string]$RitmNumber
  )
  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return $null }

  $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  if ($tasks.Count -eq 0) { $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm) }
  if ($tasks.Count -eq 0) { return $null }

  foreach ($t in $tasks) {
    $st = if ($t.PSObject.Properties["state"]) { ("" + $t.state) } elseif ($t.PSObject.Properties["state_label"]) { ("" + $t.state_label) } else { "" }
    $sv = if ($t.PSObject.Properties["state_value"]) { ("" + $t.state_value) } else { "" }
    if (Test-DashboardTaskStateOpen -StateText $st -StateValue $sv) {
      return [pscustomobject]@{
        number = if ($t.PSObject.Properties["number"]) { ("" + $t.number).Trim().ToUpperInvariant() } else { "" }
        state_text = ("" + $st).Trim()
        state_value = ("" + $sv).Trim()
      }
    }
  }
  foreach ($t in $tasks) {
    $st = if ($t.PSObject.Properties["state"]) { ("" + $t.state) } elseif ($t.PSObject.Properties["state_label"]) { ("" + $t.state_label) } else { "" }
    $sv = if ($t.PSObject.Properties["state_value"]) { ("" + $t.state_value) } else { "" }
    if (Test-DashboardTaskStateInProgress -StateText $st -StateValue $sv) {
      return [pscustomobject]@{
        number = if ($t.PSObject.Properties["number"]) { ("" + $t.number).Trim().ToUpperInvariant() } else { "" }
        state_text = ("" + $st).Trim()
        state_value = ("" + $sv).Trim()
      }
    }
  }
  return $null
}

function Invoke-DashboardCheckOut {
  param(
    $wv,
    [string]$ExcelPath,
    [string]$SheetName,
    $RowItem,
    [string]$WorkNote
  )

  $ritm = ("" + $RowItem.RITM).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return [pscustomobject]@{ ok = $false; message = "Selected row has no RITM." } }

  Log "INFO" "Dashboard CHECK-OUT started for $ritm row=$($RowItem.Row)"
  $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  if ($tasks.Count -eq 0) {
    $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm)
  }
  if ($tasks.Count -eq 0) { return [pscustomobject]@{ ok = $false; message = "No SCTASK found for $ritm." } }

  $openTasks = @($tasks | Where-Object {
    $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
    $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
    Test-DashboardTaskStateOpen -StateText $st -StateValue $sv
  })
  $wipTasks = @($tasks | Where-Object {
    $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
    $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
    Test-DashboardTaskStateInProgress -StateText $st -StateValue $sv
  })
  $task = $null
  if ($openTasks.Count -gt 0) { $task = $openTasks[0] } elseif ($wipTasks.Count -gt 0) { $task = $wipTasks[0] }
  if (-not $task) {
    return [pscustomobject]@{ ok = $false; message = "No task in Open/Work in Progress for $ritm." }
  }

  if (-not (Invoke-ServiceNowDomUpdate -wv $wv -Table "sc_task" -SysId ("" + $task.sys_id) -TargetStateLabel "Appointment" -WorkNote $WorkNote)) {
    return [pscustomobject]@{ ok = $false; message = "ServiceNow update failed for task $($task.number)." }
  }
  # Temporary test mode:
  # do not close parent RITM during CHECK-OUT; keep flow in Appointment for iterative testing.

  $excelOk = Update-DashboardExcelRow -ExcelPath $ExcelPath -SheetName $SheetName -RowIndex ([int]$RowItem.Row) `
    -DashboardStatus "Appointment" -TimestampHeader "Closed Time" -TaskNumberToWrite ("" + $task.number)
  if (-not $excelOk) { return [pscustomobject]@{ ok = $false; message = "ServiceNow updated, but Excel write failed." } }

  Log "INFO" "Dashboard CHECK-OUT completed for $ritm task=$($task.number)"
  return [pscustomobject]@{ ok = $true; message = "Appointment: $ritm ($($task.number))" }
}

function Invoke-DashboardRecalculateStatus {
  param(
    $wv,
    [string]$ExcelPath,
    [string]$SheetName,
    $RowItem
  )

  $ritm = ("" + $RowItem.RITM).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) {
    return [pscustomobject]@{ ok = $false; message = "Selected row has no RITM." }
  }

  Log "INFO" "Dashboard RE-CALCULATE started for $ritm row=$($RowItem.Row)"
  $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
  if ($tasks.Count -eq 0) {
    $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm)
  }
  $openTasks = @($tasks | Where-Object {
    $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
    $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
    Test-DashboardTaskStateOpen -StateText $st -StateValue $sv
  })
  $wipTasks = @($tasks | Where-Object {
    $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
    $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
    Test-DashboardTaskStateInProgress -StateText $st -StateValue $sv
  })

  $ritmRes = Extract-Ticket_JSONv2 -wv $wv -Ticket $ritm
  $ritmStateRaw = ""
  if ($ritmRes -and $ritmRes.PSObject.Properties["status"]) { $ritmStateRaw = "" + $ritmRes.status }
  if ([string]::IsNullOrWhiteSpace($ritmStateRaw) -and $ritmRes -and $ritmRes.PSObject.Properties["status_label"]) { $ritmStateRaw = "" + $ritmRes.status_label }
  if ([string]::IsNullOrWhiteSpace($ritmStateRaw) -and $ritmRes -and $ritmRes.PSObject.Properties["status_value"]) { $ritmStateRaw = "" + $ritmRes.status_value }
  $ritmState = $ritmStateRaw.Trim().ToLowerInvariant()
  $ritmClosed = ($ritmState -match 'closed|close|complete|completed|resolved|cancel')

  $targetStatus = ""
  $tsHeader = ""
  if (($openTasks.Count -gt 0) -or ($wipTasks.Count -gt 0)) {
    $targetStatus = "Checked-In"
    $tsHeader = "Present Time"
  }
  elseif ($ritmClosed) {
    $targetStatus = "Completed"
    $tsHeader = "Closed Time"
  }
  else {
    return [pscustomobject]@{
      ok = $false
      message = "No deterministic state for $ritm. RITM state='$ritmStateRaw', open=$($openTasks.Count), in-progress=$($wipTasks.Count)."
    }
  }

  $taskNum = ""
  if ($openTasks.Count -gt 0) { $taskNum = "" + $openTasks[0].number }
  elseif ($wipTasks.Count -gt 0) { $taskNum = "" + $wipTasks[0].number }
  elseif ($tasks.Count -gt 0) { $taskNum = "" + $tasks[0].number }

  $excelOk = Update-DashboardExcelRow -ExcelPath $ExcelPath -SheetName $SheetName -RowIndex ([int]$RowItem.Row) `
    -DashboardStatus $targetStatus -TimestampHeader $tsHeader -TaskNumberToWrite $taskNum
  if (-not $excelOk) {
    return [pscustomobject]@{ ok = $false; message = "SNOW checked, but Excel update failed for $ritm." }
  }

  Log "INFO" "Dashboard RE-CALCULATE completed for $ritm => status='$targetStatus' open=$($openTasks.Count) wip=$($wipTasks.Count) ritm_state='$ritmStateRaw'"
  return [pscustomobject]@{ ok = $true; message = "Recalculated: $ritm => $targetStatus" }
}

function Test-DashboardRowOpenLocal {
  param($RowItem)
  if (-not $RowItem) { return $false }

  $status = ("" + $RowItem.DashboardStatus).Trim().ToLowerInvariant()
  if ($status -match 'completed|complete|closed|cerrado') { return $false }
  return $true
}

function Open-DashboardRowInServiceNow {
  param(
    $wv,
    $RowItem
  )

  $ritm = ("" + $RowItem.RITM).Trim().ToUpperInvariant()
  if ([string]::IsNullOrWhiteSpace($ritm)) { return }
  $url = Build-RitmByNumberUrl -RitmNumber $ritm
  if ([string]::IsNullOrWhiteSpace($url)) { return }
  try { $wv.CoreWebView2.Navigate($url) } catch {}
  try { Start-Process $url | Out-Null } catch {}
}

function Get-DashboardSettingsPath {
  $base = Join-Path $env:APPDATA "SchumanDashboard"
  try { New-Item -ItemType Directory -Path $base -Force | Out-Null } catch {}
  return (Join-Path $base "settings.json")
}

function Load-DashboardSettings {
  param(
    [string]$DefaultCheckIn,
    [string]$DefaultCheckOut
  )
  $out = [pscustomobject]@{
    CheckInNoteTemplate = $DefaultCheckIn
    CheckOutNoteTemplate = $DefaultCheckOut
  }
  $path = Get-DashboardSettingsPath
  if (-not (Test-Path -LiteralPath $path)) { return $out }
  try {
    $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
    if (-not [string]::IsNullOrWhiteSpace($raw)) {
      $o = $raw | ConvertFrom-Json
      if ($o -and $o.PSObject.Properties["CheckInNoteTemplate"]) {
        $v = ("" + $o.CheckInNoteTemplate).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $out.CheckInNoteTemplate = $v }
      }
      if ($o -and $o.PSObject.Properties["CheckOutNoteTemplate"]) {
        $v2 = ("" + $o.CheckOutNoteTemplate).Trim()
        if (-not [string]::IsNullOrWhiteSpace($v2)) { $out.CheckOutNoteTemplate = $v2 }
      }
    }
  }
  catch {
    Log "ERROR" "Dashboard settings load failed: $($_.Exception.Message)"
  }
  return $out
}

function Save-DashboardSettings {
  param([pscustomobject]$Settings)
  try {
    $path = Get-DashboardSettingsPath
    $payload = [pscustomobject]@{
      CheckInNoteTemplate = ("" + $Settings.CheckInNoteTemplate)
      CheckOutNoteTemplate = ("" + $Settings.CheckOutNoteTemplate)
    }
    $json = $payload | ConvertTo-Json -Depth 4
    Set-Content -LiteralPath $path -Value $json -Encoding UTF8
    return $true
  }
  catch {
    Log "ERROR" "Dashboard settings save failed: $($_.Exception.Message)"
    return $false
  }
}

function Show-CheckInOutDashboard {
  param(
    $wv,
    [string]$ExcelPath,
    [string]$SheetName
  )

  [void](Ensure-DashboardExcelColumns -ExcelPath $ExcelPath -SheetName $SheetName)
  try {
    $settingsLoaded = Load-DashboardSettings -DefaultCheckIn "Deliver all credentials to the new user" -DefaultCheckOut "Equipment returned / handover completed"
    if ($settingsLoaded) {
      $script:DashboardDefaultCheckInNote = "" + $settingsLoaded.CheckInNoteTemplate
      $script:DashboardDefaultCheckOutNote = "" + $settingsLoaded.CheckOutNoteTemplate
    }
  } catch {}

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Check-in / Check-out Dashboard"
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(1366, 768)
  $form.MinimumSize = New-Object System.Drawing.Size(1366, 768)
  $form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
  $form.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $form.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)

  $root = New-Object System.Windows.Forms.TableLayoutPanel
  $root.Dock = "Fill"
  $root.Padding = New-Object System.Windows.Forms.Padding(12)
  $root.RowCount = 1
  $root.ColumnCount = 3
  $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 22)))
  $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 56)))
  $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 22)))
  $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $form.Controls.Add($root)

  $leftPanel = New-Object System.Windows.Forms.Panel
  $leftPanel.Dock = "Fill"
  $leftPanel.Padding = New-Object System.Windows.Forms.Padding(10)
  $leftPanel.Margin = New-Object System.Windows.Forms.Padding(0,0,10,0)
  $leftPanel.AutoScroll = $true
  $leftPanel.BackColor = [System.Drawing.Color]::FromArgb(30,30,32)
  $root.Controls.Add($leftPanel, 0, 0)

  $leftGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $leftGrid.Dock = "Top"
  $leftGrid.AutoSize = $true
  $leftGrid.RowCount = 11
  $leftGrid.ColumnCount = 1
  $leftGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $leftPanel.Controls.Add($leftGrid)

  $lblSearch = New-Object System.Windows.Forms.Label
  $lblSearch.Text = "Search (Display Name / Last Name):"
  $lblSearch.AutoSize = $true
  $lblSearch.ForeColor = [System.Drawing.Color]::FromArgb(178,178,182)
  $leftGrid.Controls.Add($lblSearch, 0, 0)

  $txtSearch = New-Object System.Windows.Forms.TextBox
  $txtSearch.Dock = "Top"
  $txtSearch.Height = 28
  $txtSearch.BackColor = [System.Drawing.Color]::FromArgb(34,34,36)
  $txtSearch.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $txtSearch.BorderStyle = "FixedSingle"
  $txtSearch.Margin = New-Object System.Windows.Forms.Padding(0,4,0,8)
  $leftGrid.Controls.Add($txtSearch, 0, 1)

  $lblStatusFilter = New-Object System.Windows.Forms.Label
  $lblStatusFilter.Text = "Status Filter:"
  $lblStatusFilter.AutoSize = $true
  $lblStatusFilter.ForeColor = [System.Drawing.Color]::FromArgb(178,178,182)
  $leftGrid.Controls.Add($lblStatusFilter, 0, 2)

  $cmbStatus = New-Object System.Windows.Forms.ComboBox
  $cmbStatus.Dock = "Top"
  $cmbStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
  $cmbStatus.BackColor = [System.Drawing.Color]::FromArgb(34,34,36)
  $cmbStatus.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $cmbStatus.FlatStyle = "Flat"
  [void]$cmbStatus.Items.AddRange(@("All","Open Only","Checked-In","Appointment","Completed / Closed"))
  $cmbStatus.SelectedIndex = 0
  $cmbStatus.Margin = New-Object System.Windows.Forms.Padding(0,4,0,8)
  $leftGrid.Controls.Add($cmbStatus, 0, 3)

  $btnSearch = New-Object System.Windows.Forms.Button
  $btnSearch.Text = "Search"
  $btnSearch.Dock = "Top"
  $btnSearch.Height = 34
  $btnSearch.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)
  $leftGrid.Controls.Add($btnSearch, 0, 4)

  $btnRefresh = New-Object System.Windows.Forms.Button
  $btnRefresh.Text = "Refresh"
  $btnRefresh.Dock = "Top"
  $btnRefresh.Height = 30
  $btnRefresh.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)
  $leftGrid.Controls.Add($btnRefresh, 0, 5)

  $btnClear = New-Object System.Windows.Forms.Button
  $btnClear.Text = "Clear"
  $btnClear.Dock = "Top"
  $btnClear.Height = 30
  $leftGrid.Controls.Add($btnClear, 0, 6)

  $pagerFlow = New-Object System.Windows.Forms.FlowLayoutPanel
  $pagerFlow.Dock = "Top"
  $pagerFlow.Height = 34
  $pagerFlow.WrapContents = $false
  $pagerFlow.FlowDirection = "LeftToRight"
  $pagerFlow.Margin = New-Object System.Windows.Forms.Padding(0,8,0,0)
  $leftGrid.Controls.Add($pagerFlow, 0, 7)

  $btnPrevPage = New-Object System.Windows.Forms.Button
  $btnPrevPage.Text = "< Prev"
  $btnPrevPage.Width = 80
  $btnPrevPage.Height = 28
  $pagerFlow.Controls.Add($btnPrevPage)

  $btnNextPage = New-Object System.Windows.Forms.Button
  $btnNextPage.Text = "Next >"
  $btnNextPage.Width = 80
  $btnNextPage.Height = 28
  $pagerFlow.Controls.Add($btnNextPage)

  $lblPage = New-Object System.Windows.Forms.Label
  $lblPage.Text = "Page 1/1"
  $lblPage.AutoSize = $true
  $lblPage.Margin = New-Object System.Windows.Forms.Padding(12,7,0,0)
  $lblPage.ForeColor = [System.Drawing.Color]::FromArgb(170,170,170)
  $pagerFlow.Controls.Add($lblPage)

  $lblHint = New-Object System.Windows.Forms.Label
  $lblHint.Text = "Type to filter automatically."
  $lblHint.AutoSize = $true
  $lblHint.ForeColor = [System.Drawing.Color]::FromArgb(120,180,255)
  $lblHint.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)
  $leftGrid.Controls.Add($lblHint, 0, 8)

  $btnSettings = New-Object System.Windows.Forms.Button
  $btnSettings.Text = "⚙"
  $btnSettings.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 12)
  $btnSettings.Size = New-Object System.Drawing.Size(36, 36)
  $btnSettings.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
  $btnSettings.Location = New-Object System.Drawing.Point(10, 10)
  $leftPanel.Controls.Add($btnSettings)
  $leftPanel.Add_Resize({
    $btnSettings.Location = New-Object System.Drawing.Point(10, [Math]::Max(10, $leftPanel.ClientSize.Height - 46))
  })
  $form.Add_Shown({
    $btnSettings.Location = New-Object System.Drawing.Point(10, [Math]::Max(10, $leftPanel.ClientSize.Height - 46))
  })

  $btnStyle = {
    param($b, [bool]$accent = $false)
    $b.FlatStyle = "Flat"
    $b.FlatAppearance.BorderSize = 1
    if ($accent) {
      $b.BackColor = [System.Drawing.Color]::FromArgb(0,122,255)
      $b.ForeColor = [System.Drawing.Color]::FromArgb(245,245,245)
      $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(0,122,255)
      $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(20,138,255)
      $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(0,106,226)
    }
    else {
      $b.BackColor = [System.Drawing.Color]::FromArgb(36,36,38)
      $b.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
      $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(58,58,62)
      $b.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(44,44,48)
      $b.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(32,32,34)
    }
  }
  & $btnStyle $btnSearch $true
  & $btnStyle $btnRefresh $false
  & $btnStyle $btnClear $false
  & $btnStyle $btnPrevPage $false
  & $btnStyle $btnNextPage $false
  & $btnStyle $btnSettings $false

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = "Fill"
  $grid.ReadOnly = $true
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.AllowUserToAddRows = $false
  $grid.AllowUserToDeleteRows = $false
  $grid.AllowUserToResizeRows = $false
  $grid.ScrollBars = "Vertical"
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EnableHeadersVisualStyles = $false
  $grid.BackgroundColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $grid.GridColor = [System.Drawing.Color]::FromArgb(58,58,62)
  $grid.BorderStyle = "None"
  $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30,30,32)
  $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(178,178,182)
  $grid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(42,54,72)
  $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30,30,32)
  $grid.RowHeadersVisible = $false
  $grid.RowTemplate.Height = 30
  $grid.ColumnHeadersHeight = 34
  $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing

  $centerPanel = New-Object System.Windows.Forms.Panel
  $centerPanel.Dock = "Fill"
  $centerPanel.Margin = New-Object System.Windows.Forms.Padding(0)
  $centerPanel.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
  $root.Controls.Add($centerPanel, 1, 0)

  $centerGrid = New-Object System.Windows.Forms.TableLayoutPanel
  $centerGrid.Dock = "Fill"
  $centerGrid.RowCount = 2
  $centerGrid.ColumnCount = 1
  $centerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  $centerGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  $centerPanel.Controls.Add($centerGrid)

  $lblResults = New-Object System.Windows.Forms.Label
  $lblResults.Text = "Results"
  $lblResults.AutoSize = $true
  $lblResults.ForeColor = [System.Drawing.Color]::FromArgb(178,178,182)
  $lblResults.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)
  $centerGrid.Controls.Add($lblResults, 0, 0)
  $centerGrid.Controls.Add($grid, 0, 1)

  $lblComment = New-Object System.Windows.Forms.Label
  $lblComment.Text = "Work Note:"
  $lblComment.Margin = New-Object System.Windows.Forms.Padding(0,8,0,2)
  $lblComment.AutoSize = $true
  $lblComment.ForeColor = [System.Drawing.Color]::FromArgb(170,170,170)

  $txtComment = New-Object System.Windows.Forms.TextBox
  $txtComment.Height = 120
  $txtComment.Dock = "Top"
  $txtComment.Multiline = $true
  $txtComment.ScrollBars = "Vertical"
  $txtComment.Text = $DashboardDefaultCheckInNote
  $txtComment.BackColor = [System.Drawing.Color]::FromArgb(37,37,38)
  $txtComment.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $txtComment.BorderStyle = "FixedSingle"

  $btnUseCheckInNote = New-Object System.Windows.Forms.Button
  $btnUseCheckInNote.Text = "Use Check-In Note"
  $btnUseCheckInNote.Height = 30
  $btnUseCheckInNote.Dock = "Top"
  $btnUseCheckInNote.Margin = New-Object System.Windows.Forms.Padding(0,6,0,6)

  $btnUseCheckOutNote = New-Object System.Windows.Forms.Button
  $btnUseCheckOutNote.Text = "Use Check-Out Note"
  $btnUseCheckOutNote.Height = 30
  $btnUseCheckOutNote.Dock = "Top"
  $btnUseCheckOutNote.Margin = New-Object System.Windows.Forms.Padding(0,0,0,10)

  $btnCheckIn = New-Object System.Windows.Forms.Button
  $btnCheckIn.Text = "CHECK-IN"
  $btnCheckIn.Height = 36
  $btnCheckIn.Dock = "Top"
  $btnCheckIn.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)

  $btnCheckOut = New-Object System.Windows.Forms.Button
  $btnCheckOut.Text = "CHECK-OUT"
  $btnCheckOut.Height = 36
  $btnCheckOut.Dock = "Top"
  $btnCheckOut.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)

  $btnGeneratePdf = New-Object System.Windows.Forms.Button
  $btnGeneratePdf.Text = "Generate PDF"
  $btnGeneratePdf.Height = 34
  $btnGeneratePdf.Dock = "Top"
  $btnGeneratePdf.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)

  $btnOpen = New-Object System.Windows.Forms.Button
  $btnOpen.Text = "Open in ServiceNow"
  $btnOpen.Height = 34
  $btnOpen.Dock = "Top"
  $btnOpen.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)

  $btnRecalc = New-Object System.Windows.Forms.Button
  $btnRecalc.Text = "Recalculate from SNOW"
  $btnRecalc.Height = 32
  $btnRecalc.Dock = "Top"
  $btnRecalc.Margin = New-Object System.Windows.Forms.Padding(0,0,0,8)

  $lblSelected = New-Object System.Windows.Forms.Label
  $lblSelected.Text = "Selected ticket"
  $lblSelected.AutoSize = $true
  $lblSelected.ForeColor = [System.Drawing.Color]::FromArgb(178,178,182)

  $lblSelectedName = New-Object System.Windows.Forms.Label
  $lblSelectedName.Text = "Display Name: -"
  $lblSelectedName.AutoSize = $true
  $lblSelectedName.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

  $lblSelectedRitm = New-Object System.Windows.Forms.Label
  $lblSelectedRitm.Text = "RITM: -"
  $lblSelectedRitm.AutoSize = $true
  $lblSelectedRitm.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

  $lblSelectedStatus = New-Object System.Windows.Forms.Label
  $lblSelectedStatus.Text = "SCTask: -"
  $lblSelectedStatus.AutoSize = $true
  $lblSelectedStatus.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

  $lblSelectedUpdated = New-Object System.Windows.Forms.Label
  $lblSelectedUpdated.Text = "Last Updated: -"
  $lblSelectedUpdated.AutoSize = $true
  $lblSelectedUpdated.Margin = New-Object System.Windows.Forms.Padding(0,4,0,10)

  $lblPdfPath = New-Object System.Windows.Forms.Label
  $lblPdfPath.Text = "PDF path:"
  $lblPdfPath.AutoSize = $true

  $txtPdfPath = New-Object System.Windows.Forms.TextBox
  $txtPdfPath.ReadOnly = $true
  $txtPdfPath.Dock = "Top"
  $txtPdfPath.BackColor = [System.Drawing.Color]::FromArgb(37,37,38)
  $txtPdfPath.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
  $txtPdfPath.BorderStyle = "FixedSingle"
  $txtPdfPath.Margin = New-Object System.Windows.Forms.Padding(0,4,0,10)

  & $btnStyle $btnUseCheckInNote $false
  & $btnStyle $btnUseCheckOutNote $false
  & $btnStyle $btnCheckIn $true
  & $btnStyle $btnCheckOut $false
  & $btnStyle $btnGeneratePdf $false
  & $btnStyle $btnOpen $false
  & $btnStyle $btnRecalc $false
  $btnCheckIn.Enabled = $false
  $btnCheckOut.Enabled = $false
  $btnGeneratePdf.Enabled = $false
  $btnRecalc.Enabled = $false
  $btnOpen.Enabled = $false

  # Simplified dashboard UX: hide legacy/manual action controls.
  $btnSearch.Visible = $false
  $btnRefresh.Visible = $false
  $btnGeneratePdf.Visible = $false
  $btnRecalc.Visible = $false
  $btnOpen.Visible = $false
  $lblComment.Visible = $false
  $txtComment.Visible = $false
  $btnUseCheckInNote.Visible = $false
  $btnUseCheckOutNote.Visible = $false
  $btnCheckIn.Visible = $false
  $btnCheckOut.Visible = $false
  $lblPdfPath.Visible = $false
  $txtPdfPath.Visible = $false

  $lblStatus = New-Object System.Windows.Forms.Label
  $lblStatus.Text = "Type Display/Last name."
  $lblStatus.AutoSize = $true
  $lblStatus.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)
  $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170,170,170)

  $rightPanel = New-Object System.Windows.Forms.Panel
  $rightPanel.Dock = "Fill"
  $rightPanel.Padding = New-Object System.Windows.Forms.Padding(10)
  $rightPanel.Margin = New-Object System.Windows.Forms.Padding(10,0,0,0)
  $rightPanel.AutoScroll = $true
  $rightPanel.BackColor = [System.Drawing.Color]::FromArgb(30,30,32)
  $root.Controls.Add($rightPanel, 2, 0)

  $rightStack = New-Object System.Windows.Forms.FlowLayoutPanel
  $rightStack.Dock = "Top"
  $rightStack.FlowDirection = "TopDown"
  $rightStack.WrapContents = $false
  $rightStack.AutoSize = $true
  $rightStack.Width = 280
  $rightStack.Padding = New-Object System.Windows.Forms.Padding(0)
  $rightPanel.Controls.Add($rightStack)

  $rightStack.Controls.Add($lblSelected)
  $rightStack.Controls.Add($lblSelectedName)
  $rightStack.Controls.Add($lblSelectedRitm)
  $rightStack.Controls.Add($lblSelectedStatus)
  $rightStack.Controls.Add($lblSelectedUpdated)
  $rightStack.Controls.Add($lblPdfPath)
  $rightStack.Controls.Add($txtPdfPath)
  $rightStack.Controls.Add($lblComment)
  $rightStack.Controls.Add($txtComment)
  $rightStack.Controls.Add($btnUseCheckInNote)
  $rightStack.Controls.Add($btnUseCheckOutNote)
  $rightStack.Controls.Add($btnCheckIn)
  $rightStack.Controls.Add($btnCheckOut)
  $rightStack.Controls.Add($btnGeneratePdf)
  $rightStack.Controls.Add($btnOpen)
  $rightStack.Controls.Add($btnRecalc)
  $rightStack.Controls.Add($lblStatus)

  $state = [pscustomobject]@{
    Rows = @()
    AllRows = @()
    LastSearch = ""
    UserDirectory = @()
    TaskCache = @{}
    LastOpenedRitm = ""
    PendingTaskResolve = @{}
    MaxRenderedRows = 800
    PageSize = 120
    CurrentPage = 1
    TotalPages = 1
    EnableBackgroundTaskResolve = $false
    EnableAutoResolveOnSelection = $false
  }

  $grid.AutoGenerateColumns = $false
  $grid.Columns.Clear()
  $colRow = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRow.Name = "Row"
  $colRow.HeaderText = "Row"
  $colRow.Visible = $false
  [void]$grid.Columns.Add($colRow)

  $colRequestedFor = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRequestedFor.Name = "Display Name"
  $colRequestedFor.HeaderText = "Display Name"
  $colRequestedFor.FillWeight = 38
  [void]$grid.Columns.Add($colRequestedFor)

  $colRitm = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colRitm.Name = "RITM"
  $colRitm.HeaderText = "RITM"
  $colRitm.FillWeight = 22
  [void]$grid.Columns.Add($colRitm)

  $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colStatus.Name = "SCTask"
  $colStatus.HeaderText = "SCTask"
  $colStatus.FillWeight = 20
  [void]$grid.Columns.Add($colStatus)

  $colUpdated = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $colUpdated.Name = "Last Updated"
  $colUpdated.HeaderText = "Last Updated"
  $colUpdated.FillWeight = 20
  [void]$grid.Columns.Add($colUpdated)

  $getTaskSummary = {
    param($rowItem)
    if (-not $rowItem) { return "No Open Task" }
    $ritm = ("" + $rowItem.RITM).Trim().ToUpperInvariant()
    if ($state.TaskCache.ContainsKey($ritm)) {
      $tasks = @($state.TaskCache[$ritm])
      if ($tasks.Count -eq 0) { return "No Open Task" }
      if ($tasks.Count -eq 1) { return ("" + $tasks[0].number) }
      return "Multiple Tasks ($($tasks.Count))"
    }
    $txt = ("" + $rowItem.SCTASK).Trim()
    if ([string]::IsNullOrWhiteSpace($txt)) {
      if ($state.PendingTaskResolve.ContainsKey($ritm)) { return "Resolving..." }
      return "No Open Task"
    }
    if ($txt -match ',' -or $txt -match ';') {
      $parts = @($txt -split '[,;]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
      if ($parts.Count -gt 1) { return "Multiple Tasks ($($parts.Count))" }
    }
    return $txt
  }

  $formatLastUpdated = {
    param([string]$v)
    $txt = ("" + $v).Trim()
    if ([string]::IsNullOrWhiteSpace($txt) -or $txt -eq "########") { return "-" }
    $dt = $null
    if ([datetime]::TryParse($txt, [ref]$dt)) { return $dt.ToString("yyyy-MM-dd HH:mm") }
    return $txt
  }

  $updatePagerUi = {
    param([int]$totalCount)
    if ($state.PageSize -lt 1) { $state.PageSize = 120 }
    $pages = 1
    if ($totalCount -gt 0) {
      $pages = [int][Math]::Ceiling($totalCount / [double]$state.PageSize)
      if ($pages -lt 1) { $pages = 1 }
    }
    $state.TotalPages = $pages
    if ($state.CurrentPage -lt 1) { $state.CurrentPage = 1 }
    if ($state.CurrentPage -gt $state.TotalPages) { $state.CurrentPage = $state.TotalPages }
    $lblPage.Text = "Page $($state.CurrentPage)/$($state.TotalPages)"
    $btnPrevPage.Enabled = ($state.CurrentPage -gt 1) -and (-not $form.UseWaitCursor)
    $btnNextPage.Enabled = ($state.CurrentPage -lt $state.TotalPages) -and (-not $form.UseWaitCursor)
  }

  $bindRowsToGrid = {
    param($rows)
    $state.Rows = @($rows)
    $totalRows = $state.Rows.Count
    & $updatePagerUi $totalRows
    $startIndex = ($state.CurrentPage - 1) * $state.PageSize
    if ($startIndex -lt 0) { $startIndex = 0 }
    if ($startIndex -ge $totalRows) {
      $state.CurrentPage = 1
      $startIndex = 0
      & $updatePagerUi $totalRows
    }

    $grid.SuspendLayout()
    try {
      $grid.Rows.Clear()
      $renderRows = @()
      if ($totalRows -gt 0) {
        $renderRows = @($state.Rows | Select-Object -Skip $startIndex -First $state.PageSize)
      }
      foreach ($x in @($renderRows)) {
        [void]$grid.Rows.Add(
          ("" + $x.Row),
          ("" + $x.RequestedFor),
          ("" + $x.RITM),
          (& $getTaskSummary $x),
          (& $formatLastUpdated ("" + $x.LastUpdated))
        )
        if ($state.EnableBackgroundTaskResolve) {
          & $enqueueTaskResolve $x
        }
      }
      $grid.ClearSelection()
      if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
        $grid.CurrentCell = $grid.Rows[0].Cells["Display Name"]
      }
    }
    finally {
      $grid.ResumeLayout()
    }
    $sel = & $getSelectedRow
    if ($sel) {
      $lblSelectedName.Text = "Display Name: " + ("" + $sel.RequestedFor)
      $lblSelectedRitm.Text = "RITM: " + ("" + $sel.RITM)
      $lblSelectedStatus.Text = "SCTask: " + (& $getTaskSummary $sel)
      $lblSelectedUpdated.Text = "Last Updated: " + (& $formatLastUpdated ("" + $sel.LastUpdated))
      $txtPdfPath.Text = ("" + $sel.PdfPath)
    } else {
      $lblSelectedName.Text = "Display Name: -"
      $lblSelectedRitm.Text = "RITM: -"
      $lblSelectedStatus.Text = "SCTask: -"
      $lblSelectedUpdated.Text = "Last Updated: -"
      $txtPdfPath.Text = ""
    }
    & $updateActionButtons
  }

  $getSelectedRow = {
    if ($grid.SelectedRows.Count -eq 0) { return $null }
    $selected = $grid.SelectedRows[0]
    if (-not $selected) { return $null }
    $rowNum = 0
    $rowTxt = ""
    try { $rowTxt = ("" + $selected.Cells["Row"].Value).Trim() } catch { return $null }
    $okParse = [int]::TryParse($rowTxt, [ref]$rowNum)
    if (-not $okParse) { return $null }
    foreach ($item in $state.Rows) {
      if ([int]$item.Row -eq $rowNum) { return $item }
    }
    return $null
  }

  $updateActionButtons = {
    $sel = & $getSelectedRow
    $hasValidRitm = $false
    if ($sel) {
      $ritmTxt = ("" + $sel.RITM).Trim().ToUpperInvariant()
      $hasValidRitm = ($ritmTxt -match '^RITM\d{6,8}$')
    }
    $canAct = $hasValidRitm -and (-not $actionState.IsRunning)
    $btnCheckIn.Enabled = $canAct
    $btnCheckOut.Enabled = $canAct
    $btnGeneratePdf.Enabled = $canAct
    $btnRecalc.Enabled = $canAct
    $btnOpen.Enabled = $canAct
  }

  $getVisibleRows = {
    $rows = @($state.AllRows)
    $q = ("" + $txtSearch.Text).Trim()
    if (-not [string]::IsNullOrWhiteSpace($q)) {
      $rows = @($rows | Where-Object {
        ((("" + $_.RequestedFor).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) -or
        ((("" + $_.LastName).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) -or
        ((("" + $_.FirstName).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) -or
        ((("" + $_.RITM).IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0))
      })
    }
    $statusSel = ("" + $cmbStatus.SelectedItem).Trim()
    switch ($statusSel) {
      "Open Only" {
        $rows = @($rows | Where-Object { Test-DashboardRowOpenLocal -RowItem $_ })
      }
      "Checked-In" {
        $rows = @($rows | Where-Object { ("" + $_.DashboardStatus).Trim().ToLowerInvariant() -eq "checked-in" })
      }
      "Appointment" {
        $rows = @($rows | Where-Object { ("" + $_.DashboardStatus).Trim().ToLowerInvariant() -eq "appointment" })
      }
      "Completed / Closed" {
        $rows = @($rows | Where-Object {
          $s = ("" + $_.DashboardStatus).Trim().ToLowerInvariant()
          $s -match 'completed|complete|closed|cerrado'
        })
      }
      default { }
    }
    return @($rows)
  }

  $updateSearchUserSuggestions = {
    # Search is now a plain TextBox for smoother typing under large datasets.
    return
  }

  $setBusyUi = {
    param([bool]$On, [string]$Message = "Working...")
    $grid.Enabled = -not $On
    $txtSearch.Enabled = -not $On
    $cmbStatus.Enabled = -not $On
    $btnClear.Enabled = -not $On
    $btnPrevPage.Enabled = -not $On
    $btnNextPage.Enabled = -not $On
    $btnSettings.Enabled = -not $On
    if ($On) { $lblStatus.Text = $Message }
    $form.UseWaitCursor = $On
    & $updatePagerUi $state.Rows.Count
    [System.Windows.Forms.Application]::DoEvents()
  }

  $performSearch = {
    param([switch]$ReloadFromExcel, [switch]$KeepPage)
    try {
      $q = ("" + $txtSearch.Text).Trim()
      if ($ReloadFromExcel -or (-not $state.AllRows) -or ($state.AllRows.Count -eq 0)) {
        & $setBusyUi $true "Loading dashboard data..."
        # Load once from Excel, then filter locally for smooth typing.
        $state.AllRows = @(Search-DashboardRows -ExcelPath $ExcelPath -SheetName $SheetName -SearchText "")
        $state.TaskCache.Clear()
        $state.PendingTaskResolve.Clear()
        try { while ($taskResolveQueue.Count -gt 0) { [void]$taskResolveQueue.Dequeue() } } catch {}
        & $setBusyUi $false
      }
      if (-not $KeepPage) { $state.CurrentPage = 1 }
      $rows = & $getVisibleRows
      $state.LastSearch = $q
      & $bindRowsToGrid $rows
      $shown = $grid.Rows.Count
      $from = 0
      $to = 0
      if ($rows.Count -gt 0) {
        $from = (($state.CurrentPage - 1) * $state.PageSize) + 1
        $to = $from + $shown - 1
      }
      $filterNote = ""
      if ($cmbStatus.SelectedItem -and ("" + $cmbStatus.SelectedItem) -ne "All") {
        $filterNote = " (" + $cmbStatus.SelectedItem + ")"
      }
      if ([string]::IsNullOrWhiteSpace($q)) {
        if ($rows.Count -eq 0) {
          $lblStatus.Text = "Results: 0$filterNote"
        } else {
          $lblStatus.Text = "Results: $from-$to / $($rows.Count)$filterNote"
        }
      } else {
        if ($rows.Count -eq 0) {
          $lblStatus.Text = "Results: 0 for '$q'$filterNote"
        } else {
          $lblStatus.Text = "Results: $from-$to / $($rows.Count) for '$q'$filterNote"
        }
      }
    }
    catch {
      & $setBusyUi $false
      $errMsg = $_.Exception.Message
      $errPos = $_.InvocationInfo.PositionMessage
      Log "ERROR" "Dashboard search failed: $errMsg | $errPos"
      [System.Windows.Forms.MessageBox]::Show(
        "Search failed: $errMsg`r`n$errPos",
        "Dashboard Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
      ) | Out-Null
    }
  }

  $searchDebounce = New-Object System.Windows.Forms.Timer
  $searchDebounce.Interval = 320
  $searchDebounce.Add_Tick({
    $searchDebounce.Stop()
    & $performSearch
  })

  $txtSearch.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
      $e.SuppressKeyPress = $true
      $searchDebounce.Stop()
      & $performSearch
    }
  })
  $txtSearch.Add_TextChanged({
    $searchDebounce.Stop()
    $searchDebounce.Start()
  })
  $cmbStatus.Add_SelectedIndexChanged({
    $searchDebounce.Stop()
    & $performSearch
  })
  $btnPrevPage.Add_Click({
    if ($state.CurrentPage -le 1) { return }
    $state.CurrentPage--
    & $performSearch -KeepPage
  })
  $btnNextPage.Add_Click({
    if ($state.CurrentPage -ge $state.TotalPages) { return }
    $state.CurrentPage++
    & $performSearch -KeepPage
  })
  $grid.Add_SelectionChanged({
    $sel = & $getSelectedRow
    if ($sel) {
      $lblSelectedName.Text = "Display Name: " + ("" + $sel.RequestedFor)
      $lblSelectedRitm.Text = "RITM: " + ("" + $sel.RITM)
      $lblSelectedStatus.Text = "SCTask: " + (& $getTaskSummary $sel)
      $lblSelectedUpdated.Text = "Last Updated: " + (& $formatLastUpdated ("" + $sel.LastUpdated))
      $txtPdfPath.Text = ("" + $sel.PdfPath)
    } else {
      $lblSelectedName.Text = "Display Name: -"
      $lblSelectedRitm.Text = "RITM: -"
      $lblSelectedStatus.Text = "SCTask: -"
      $lblSelectedUpdated.Text = "Last Updated: -"
      $txtPdfPath.Text = ""
    }
    & $updateActionButtons
    if ($state.EnableAutoResolveOnSelection) {
      $selectedResolveTimer.Stop()
      $selectedResolveTimer.Start()
    }
  })
  $grid.Add_CellClick({
    param($sender, $e)
    try {
      if ($e.RowIndex -lt 0) { return }
      $col = "" + $grid.Columns[$e.ColumnIndex].Name
      if ($col -ne "RITM") { return }
      $row = & $getSelectedRow
      if (-not $row) { return }
      $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
      if (-not ($ritm -match '^RITM\d{6,8}$')) { return }
      if ($state.LastOpenedRitm -eq $ritm) { return }
      $state.LastOpenedRitm = $ritm
      Open-DashboardRowInServiceNow -wv $wv -RowItem $row
    } catch {}
  })
  $grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    try {
      $row = & $getSelectedRow
      if (-not $row) { return }
      $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
      if (-not ($ritm -match '^RITM\d{6,8}$')) { return }

      & $setBusyUi $true "Resolving tasks..."
      $tasks = @(& $resolveOpenTasksForRow $row)
      & $setBusyUi $false
      if ($tasks.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
          "No open CSC task found for this RITM.",
          "Dashboard",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
      }

      $selectedTask = $null
      if ($tasks.Count -eq 1) {
        $selectedTask = $tasks[0]
      } else {
        $selectedTask = & $showTaskSelectionDialog $ritm $tasks
      }
      if (-not $selectedTask) { return }
      Log "INFO" "Dashboard task selected ritm='$ritm' task='$($selectedTask.number)'"

      $taskUrl = Build-SCTaskBestUrl -SysId ("" + $selectedTask.sys_id) -TaskNumber ("" + $selectedTask.number)
      if (-not [string]::IsNullOrWhiteSpace($taskUrl)) {
        try { $wv.CoreWebView2.Navigate($taskUrl) } catch {}
        try { Start-Process $taskUrl | Out-Null } catch {}
      }

      $choice = [System.Windows.Forms.MessageBox]::Show(
        "What do you want to do?`r`n`r`nYes = Check-In`r`nNo = Check-Out`r`nCancel = Cancel",
        "Action",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
      )
      if ($choice -eq [System.Windows.Forms.DialogResult]::Cancel) { return }

      $isCheckIn = ($choice -eq [System.Windows.Forms.DialogResult]::Yes)
      $template = if ($isCheckIn) { "" + $DashboardDefaultCheckInNote } else { "" + $DashboardDefaultCheckOutNote }
      $noteTitle = if ($isCheckIn) { "Edit Check-In Note" } else { "Edit Check-Out Note" }
      $noteText = & $showNoteEditorDialog $noteTitle $template
      if ($null -eq $noteText) { return }

      & $setBusyUi $true "Submitting update..."
      try {
        $targetState = if ($isCheckIn) { "Work in Progress" } else { "Appointment" }
        $excelState = if ($isCheckIn) { "Checked-In" } else { "Appointment" }
        $tsHeader = if ($isCheckIn) { "Present Time" } else { "Closed Time" }
        Log "INFO" "Dashboard SN update start ritm='$ritm' task='$($selectedTask.number)' target='$targetState'"
        if (-not (Invoke-ServiceNowDomUpdate -wv $wv -Table "sc_task" -SysId ("" + $selectedTask.sys_id) -TargetStateLabel $targetState -WorkNote ("" + $noteText))) {
          throw "ServiceNow update failed for task $($selectedTask.number)."
        }
        Log "INFO" "Dashboard SN update end ritm='$ritm' task='$($selectedTask.number)'"
        Log "INFO" "Dashboard Excel update start ritm='$ritm' row=$($row.Row)"
        $excelOk = Update-DashboardExcelRow -ExcelPath $ExcelPath -SheetName $SheetName -RowIndex ([int]$row.Row) -DashboardStatus $excelState -TimestampHeader $tsHeader -TaskNumberToWrite ("" + $selectedTask.number)
        if (-not $excelOk) { throw "ServiceNow updated, but Excel write failed." }
        Log "INFO" "Dashboard Excel update end ritm='$ritm' row=$($row.Row)"
        $lblStatus.Text = if ($isCheckIn) { "Checked-In: $ritm ($($selectedTask.number))" } else { "Appointment: $ritm ($($selectedTask.number))" }
        $state.TaskCache.Remove($ritm) | Out-Null
        & $performSearch -ReloadFromExcel
      }
      catch {
        Log "ERROR" "Dashboard double-click action failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
          $_.Exception.Message,
          "Dashboard Error",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
      }
      finally {
        & $setBusyUi $false
      }
    }
    catch {
      Log "ERROR" "Dashboard double-click flow failed: $($_.Exception.Message)"
    }
  })
  $btnRefresh.Add_Click({
    if ([string]::IsNullOrWhiteSpace($state.LastSearch)) {
      $lblStatus.Text = "No previous search."
      return
    }
    $txtSearch.Text = $state.LastSearch
    & $performSearch -ReloadFromExcel
  })
  $btnClear.Add_Click({
    $searchDebounce.Stop()
    $txtSearch.Text = ""
    $state.LastSearch = ""
    $state.CurrentPage = 1
    & $performSearch -KeepPage:$false
  })
  $btnUseCheckInNote.Add_Click({
    $txtComment.Text = $DashboardDefaultCheckInNote
  })
  $btnUseCheckOutNote.Add_Click({
    $txtComment.Text = $DashboardDefaultCheckOutNote
  })

  $btnSettings.Add_Click({
    try {
      $dlg = New-Object System.Windows.Forms.Form
      $dlg.Text = "Settings"
      $dlg.StartPosition = "CenterParent"
      $dlg.Size = New-Object System.Drawing.Size(760, 480)
      $dlg.MinimumSize = New-Object System.Drawing.Size(640, 420)
      $dlg.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
      $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
      $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

      $lay = New-Object System.Windows.Forms.TableLayoutPanel
      $lay.Dock = "Fill"
      $lay.Padding = New-Object System.Windows.Forms.Padding(12)
      $lay.RowCount = 6
      $lay.ColumnCount = 1
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
      $lay.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
      $dlg.Controls.Add($lay)

      $lblIn = New-Object System.Windows.Forms.Label
      $lblIn.Text = "Check-In Note Template"
      $lblIn.AutoSize = $true
      $lay.Controls.Add($lblIn, 0, 0)

      $txtIn = New-Object System.Windows.Forms.TextBox
      $txtIn.Multiline = $true
      $txtIn.ScrollBars = "Vertical"
      $txtIn.Dock = "Fill"
      $txtIn.BackColor = [System.Drawing.Color]::FromArgb(37,37,38)
      $txtIn.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
      $txtIn.Text = ("" + $DashboardDefaultCheckInNote)
      $lay.Controls.Add($txtIn, 0, 1)

      $lblOut = New-Object System.Windows.Forms.Label
      $lblOut.Text = "Check-Out Note Template"
      $lblOut.AutoSize = $true
      $lay.Controls.Add($lblOut, 0, 2)

      $txtOut = New-Object System.Windows.Forms.TextBox
      $txtOut.Multiline = $true
      $txtOut.ScrollBars = "Vertical"
      $txtOut.Dock = "Fill"
      $txtOut.BackColor = [System.Drawing.Color]::FromArgb(37,37,38)
      $txtOut.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
      $txtOut.Text = ("" + $DashboardDefaultCheckOutNote)
      $lay.Controls.Add($txtOut, 0, 3)

      $btnFlow = New-Object System.Windows.Forms.FlowLayoutPanel
      $btnFlow.Dock = "Fill"
      $btnFlow.FlowDirection = "RightToLeft"
      $btnFlow.WrapContents = $false
      $btnFlow.AutoSize = $true

      $btnSaveSettings = New-Object System.Windows.Forms.Button
      $btnSaveSettings.Text = "Save"
      $btnSaveSettings.Width = 100
      & $btnStyle $btnSaveSettings $true
      $btnCancelSettings = New-Object System.Windows.Forms.Button
      $btnCancelSettings.Text = "Cancel"
      $btnCancelSettings.Width = 100
      & $btnStyle $btnCancelSettings $false
      $btnResetSettings = New-Object System.Windows.Forms.Button
      $btnResetSettings.Text = "Reset defaults"
      $btnResetSettings.Width = 130
      & $btnStyle $btnResetSettings $false
      $btnFlow.Controls.AddRange(@($btnSaveSettings, $btnCancelSettings, $btnResetSettings))
      $lay.Controls.Add($btnFlow, 0, 5)

      $btnResetSettings.Add_Click({
        $txtIn.Text = "Deliver all credentials to the new user"
        $txtOut.Text = "Equipment returned / handover completed"
      })
      $btnCancelSettings.Add_Click({ $dlg.Close() })
      $btnSaveSettings.Add_Click({
        $newIn = ("" + $txtIn.Text).Trim()
        $newOut = ("" + $txtOut.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($newIn)) { $newIn = "Deliver all credentials to the new user" }
        if ([string]::IsNullOrWhiteSpace($newOut)) { $newOut = "Equipment returned / handover completed" }
        $ok = Save-DashboardSettings -Settings ([pscustomobject]@{
          CheckInNoteTemplate = $newIn
          CheckOutNoteTemplate = $newOut
        })
        if ($ok) {
          $script:DashboardDefaultCheckInNote = $newIn
          $script:DashboardDefaultCheckOutNote = $newOut
          $txtComment.Text = $newIn
          $dlg.Close()
        }
      })
      [void]$dlg.ShowDialog($form)
    }
    catch {
      Log "ERROR" "Dashboard settings dialog failed: $($_.Exception.Message)"
    }
  })

  $resolveOpenTasksForRow = {
    param($row)
    if (-not $row) { return @() }
    $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
    if (-not ($ritm -match '^RITM\d{6,8}$')) { return @() }
    if ($state.TaskCache.ContainsKey($ritm)) { return @($state.TaskCache[$ritm]) }
    try {
      $state.PendingTaskResolve[$ritm] = $true
      Log "INFO" "Dashboard task resolution start ritm='$ritm'"
      $tasks = @(Get-SCTaskCandidatesForRitm -wv $wv -RitmNumber $ritm)
      if ($tasks.Count -eq 0) { $tasks = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $ritm) }
      $open = @($tasks | Where-Object {
        $st = if ($_.PSObject.Properties["state"]) { ("" + $_.state) } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label) } else { "" }
        $sv = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value) } else { "" }
        (Test-DashboardTaskStateOpen -StateText $st -StateValue $sv) -or (Test-DashboardTaskStateInProgress -StateText $st -StateValue $sv)
      } | ForEach-Object {
        [pscustomobject]@{
          number = if ($_.PSObject.Properties["number"]) { ("" + $_.number).Trim().ToUpperInvariant() } else { "" }
          sys_id = if ($_.PSObject.Properties["sys_id"]) { ("" + $_.sys_id).Trim() } else { "" }
          state_text = if ($_.PSObject.Properties["state"]) { ("" + $_.state).Trim() } elseif ($_.PSObject.Properties["state_label"]) { ("" + $_.state_label).Trim() } else { "" }
          state_value = if ($_.PSObject.Properties["state_value"]) { ("" + $_.state_value).Trim() } else { "" }
          short_description = if ($_.PSObject.Properties["short_description"]) { ("" + $_.short_description).Trim() } else { "" }
          updated = if ($_.PSObject.Properties["sys_updated_on"]) { ("" + $_.sys_updated_on).Trim() } else { "" }
        }
      })
      $state.TaskCache[$ritm] = @($open)
      Log "INFO" "Dashboard task resolution end ritm='$ritm' open_count=$($open.Count)"
      return @($open)
    }
    finally {
      if ($state.PendingTaskResolve.ContainsKey($ritm)) { $state.PendingTaskResolve.Remove($ritm) | Out-Null }
    }
  }

  $refreshSelectedTaskDisplay = {
    $row = & $getSelectedRow
    if (-not $row) { return }
    $summary = & $getTaskSummary $row
    $lblSelectedStatus.Text = "SCTask: $summary"
    if ($grid.SelectedRows.Count -gt 0) {
      try { $grid.SelectedRows[0].Cells["SCTask"].Value = $summary } catch {}
    }
  }

  $resolveSelectedTaskIfNeeded = {
    $row = & $getSelectedRow
    if (-not $row) { return }
    $ritm = ("" + $row.RITM).Trim().ToUpperInvariant()
    if (-not ($ritm -match '^RITM\d{6,8}$')) { return }
    if ($state.TaskCache.ContainsKey($ritm)) { & $refreshSelectedTaskDisplay; return }
    try {
      $lblStatus.Text = "Resolving tasks..."
      [void](& $resolveOpenTasksForRow $row)
      & $refreshSelectedTaskDisplay
      $lblStatus.Text = "Ready."
    }
    catch {
      Log "ERROR" "Dashboard selected-task resolve failed: $($_.Exception.Message)"
    }
  }

  $selectedResolveTimer = New-Object System.Windows.Forms.Timer
  $selectedResolveTimer.Interval = 220
  $selectedResolveTimer.Add_Tick({
    $selectedResolveTimer.Stop()
    if ($form.UseWaitCursor) { return }
    & $resolveSelectedTaskIfNeeded
  })

  $taskResolveQueue = New-Object System.Collections.Queue
  $enqueueTaskResolve = {
    param($rowItem)
    if (-not $rowItem) { return }
    $ritm = ("" + $rowItem.RITM).Trim().ToUpperInvariant()
    if (-not ($ritm -match '^RITM\d{6,8}$')) { return }
    if ($state.TaskCache.ContainsKey($ritm)) { return }
    if ($state.PendingTaskResolve.ContainsKey($ritm)) { return }
    $taskResolveQueue.Enqueue($rowItem)
    $state.PendingTaskResolve[$ritm] = $true
  }

  $taskResolveTimer = New-Object System.Windows.Forms.Timer
  $taskResolveTimer.Interval = 180
  $taskResolveTimer.Add_Tick({
    if ($form.UseWaitCursor) { return }
    if ($taskResolveQueue.Count -eq 0) { return }
    $item = $null
    try { $item = $taskResolveQueue.Dequeue() } catch { return }
    if (-not $item) { return }
    $ritm = ("" + $item.RITM).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($ritm)) { return }
    try {
      [void](& $resolveOpenTasksForRow $item)
      for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
        $rnum = ("" + $grid.Rows[$i].Cells["Row"].Value).Trim()
        if ($rnum -eq ("" + $item.Row)) {
          $grid.Rows[$i].Cells["SCTask"].Value = (& $getTaskSummary $item)
          break
        }
      }
      $sel = & $getSelectedRow
      if ($sel -and ((("" + $sel.RITM).Trim().ToUpperInvariant()) -eq $ritm)) {
        & $refreshSelectedTaskDisplay
      }
    }
    catch {
      Log "ERROR" "Dashboard background task resolve failed ritm='$ritm': $($_.Exception.Message)"
    }
  })
  if ($state.EnableBackgroundTaskResolve) {
    $taskResolveTimer.Start()
  }

  $showTaskSelectionDialog = {
    param([string]$ritm, [object[]]$tasks)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select Task for $ritm"
    $dlg.StartPosition = "CenterParent"
    $dlg.Size = New-Object System.Drawing.Size(920, 420)
    $dlg.MinimumSize = New-Object System.Drawing.Size(820, 320)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
    $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $list = New-Object System.Windows.Forms.ListView
    $list.Dock = "Fill"
    $list.View = "Details"
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.HideSelection = $false
    [void]$list.Columns.Add("Task", 150)
    [void]$list.Columns.Add("State", 120)
    [void]$list.Columns.Add("Short description", 420)
    [void]$list.Columns.Add("Updated", 160)
    foreach ($t in @($tasks)) {
      $item = New-Object System.Windows.Forms.ListViewItem(("" + $t.number))
      [void]$item.SubItems.Add(("" + $t.state_text))
      [void]$item.SubItems.Add(("" + $t.short_description))
      [void]$item.SubItems.Add(("" + $t.updated))
      $item.Tag = $t
      [void]$list.Items.Add($item)
    }

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = "Bottom"
    $flow.Height = 44
    $flow.FlowDirection = "RightToLeft"
    $flow.Padding = New-Object System.Windows.Forms.Padding(8)
    $btnSelectTask = New-Object System.Windows.Forms.Button
    $btnSelectTask.Text = "Select"
    $btnSelectTask.Width = 100
    & $btnStyle $btnSelectTask $true
    $btnCancelTask = New-Object System.Windows.Forms.Button
    $btnCancelTask.Text = "Cancel"
    $btnCancelTask.Width = 100
    & $btnStyle $btnCancelTask $false
    $flow.Controls.AddRange(@($btnSelectTask, $btnCancelTask))
    $dlg.Controls.Add($list)
    $dlg.Controls.Add($flow)

    $selected = $null
    $selectAction = {
      if ($list.SelectedItems.Count -eq 0) { return }
      $selected = $list.SelectedItems[0].Tag
      $dlg.Close()
    }
    $btnSelectTask.Add_Click($selectAction)
    $list.Add_DoubleClick($selectAction)
    $btnCancelTask.Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog($form)
    return $selected
  }

  $showNoteEditorDialog = {
    param([string]$title, [string]$templateText)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $title
    $dlg.StartPosition = "CenterParent"
    $dlg.Size = New-Object System.Drawing.Size(720, 420)
    $dlg.MinimumSize = New-Object System.Drawing.Size(620, 320)
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(24,24,26)
    $dlg.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Dock = "Fill"
    $tb.Multiline = $true
    $tb.ScrollBars = "Vertical"
    $tb.BackColor = [System.Drawing.Color]::FromArgb(37,37,38)
    $tb.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
    $tb.Text = "" + $templateText
    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = "Bottom"
    $flow.Height = 44
    $flow.FlowDirection = "RightToLeft"
    $flow.Padding = New-Object System.Windows.Forms.Padding(8)
    $btnSubmitNote = New-Object System.Windows.Forms.Button
    $btnSubmitNote.Text = "Submit"
    $btnSubmitNote.Width = 100
    & $btnStyle $btnSubmitNote $true
    $btnCancelNote = New-Object System.Windows.Forms.Button
    $btnCancelNote.Text = "Cancel"
    $btnCancelNote.Width = 100
    & $btnStyle $btnCancelNote $false
    $flow.Controls.AddRange(@($btnSubmitNote, $btnCancelNote))
    $dlg.Controls.Add($tb)
    $dlg.Controls.Add($flow)
    $result = $null
    $btnSubmitNote.Add_Click({ $result = "" + $tb.Text; $dlg.Close() })
    $btnCancelNote.Add_Click({ $dlg.Close() })
    [void]$dlg.ShowDialog($form)
    return $result
  }

  $actionButtons = @($btnCheckIn, $btnCheckOut, $btnGeneratePdf, $btnOpen, $btnRecalc)
  $spinnerFrames = @("|","/","-","\\")
  $actionState = [pscustomobject]@{
    IsRunning = $false
    Button = $null
    OriginalText = ""
    RunningText = ""
    Index = 0
  }
  $spinnerTimer = New-Object System.Windows.Forms.Timer
  $spinnerTimer.Interval = 120
  $spinnerTimer.Add_Tick({
    if (-not $actionState.IsRunning -or -not $actionState.Button) { return }
    $actionState.Index = ($actionState.Index + 1) % $spinnerFrames.Count
    $glyph = $spinnerFrames[$actionState.Index]
    $actionState.Button.Text = "$($actionState.RunningText) $glyph"
  })
  $invokeDashboardAction = {
    param(
      [System.Windows.Forms.Button]$Button,
      [string]$RunningText,
      [scriptblock]$Action
    )
    if ($actionState.IsRunning) { return }
    $actionState.IsRunning = $true
    $actionState.Button = $Button
    $actionState.OriginalText = $Button.Text
    $actionState.RunningText = $RunningText
    $actionState.Index = 0
    foreach ($b in $actionButtons) { $b.Enabled = $false }
    $Button.Text = "$RunningText $($spinnerFrames[0])"
    $spinnerTimer.Start()
    try {
      & $Action
    }
    finally {
      $spinnerTimer.Stop()
      if ($actionState.Button) {
        $actionState.Button.Text = $actionState.OriginalText
      }
      $actionState.IsRunning = $false
      $actionState.Button = $null
      & $updateActionButtons
    }
  }

  $btnOpen.Add_Click({
    & $invokeDashboardAction $btnOpen "Opening..." {
      try {
        $row = & $getSelectedRow
        if (-not $row) {
          [System.Windows.Forms.MessageBox]::Show(
            "Select one row first.",
            "Dashboard",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
          ) | Out-Null
          return
        }
        Open-DashboardRowInServiceNow -wv $wv -RowItem $row
      }
      catch {
        Log "ERROR" "Dashboard open-in-SNOW failed: $($_.Exception.Message)"
      }
    }
  })

  $btnGeneratePdf.Add_Click({
    & $invokeDashboardAction $btnGeneratePdf "Generating..." {
      try {
        $row = & $getSelectedRow
        if (-not $row) {
          [System.Windows.Forms.MessageBox]::Show(
            "Select one row first.",
            "Dashboard",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
          ) | Out-Null
          return
        }
        $genScript = Join-Path $PSScriptRoot "Generate-pdf.ps1"
        if (-not (Test-Path -LiteralPath $genScript)) {
          [System.Windows.Forms.MessageBox]::Show(
            "Generate-pdf.ps1 not found next to auto-excel.ps1.",
            "Generate PDF",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
          return
        }
        Start-Process -FilePath "powershell.exe" -ArgumentList @(
          "-NoProfile",
          "-ExecutionPolicy", "Bypass",
          "-File", $genScript,
          "-ExcelPath", $ExcelPath,
          "-PreferredSheet", $SheetName
        ) | Out-Null
        $lblStatus.Text = "PDF generator launched."
      }
      catch {
        Log "ERROR" "Dashboard generate-pdf launch failed: $($_.Exception.Message)"
      }
    }
  })

  $btnRecalc.Add_Click({
    & $invokeDashboardAction $btnRecalc "Recalculating..." {
      try {
        $row = & $getSelectedRow
        if (-not $row) {
          [System.Windows.Forms.MessageBox]::Show(
            "Select one row first.",
            "Dashboard",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
          ) | Out-Null
          return
        }

        $res = Invoke-DashboardRecalculateStatus -wv $wv -ExcelPath $ExcelPath -SheetName $SheetName -RowItem $row
        if ($res.ok -eq $true) {
          $lblStatus.Text = "" + $res.message
          & $performSearch -ReloadFromExcel
        }
        else {
          [System.Windows.Forms.MessageBox]::Show(
            "" + $res.message,
            "Recalculate",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
        }
      }
      catch {
        Log "ERROR" "Dashboard RE-CALCULATE failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
          "Recalculate failed: $($_.Exception.Message)",
          "Recalculate Error",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
      }
    }
  })

  $btnCheckIn.Add_Click({
    & $invokeDashboardAction $btnCheckIn "Checking-In..." {
      try {
        $row = & $getSelectedRow
        if (-not $row) {
          [System.Windows.Forms.MessageBox]::Show(
            "Select one row first.",
            "Dashboard",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
          ) | Out-Null
          return
        }
        $note = ("" + $txtComment.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($note)) {
          $note = $DashboardDefaultCheckInNote
          $txtComment.Text = $note
        }
        $ritmSel = ("" + $row.RITM).Trim().ToUpperInvariant()
        $candIn = Get-DashboardCheckInCandidate -wv $wv -RitmNumber $ritmSel
        if (-not $candIn) {
          [System.Windows.Forms.MessageBox]::Show(
            "No Open task detected for $ritmSel.",
            "Check-In",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
          return
        }
        $confirmIn = [System.Windows.Forms.MessageBox]::Show(
          "RITM: $ritmSel`r`nSCTASK: $($candIn.number)`r`nCurrent: $($candIn.state_text) [$($candIn.state_value)]`r`nTarget: Work in Progress`r`n`r`nContinue?",
          "Confirm Check-In",
          [System.Windows.Forms.MessageBoxButtons]::YesNo,
          [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirmIn -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $res = Invoke-DashboardCheckIn -wv $wv -ExcelPath $ExcelPath -SheetName $SheetName -RowItem $row -WorkNote $note
        if ($res.ok -eq $true) {
          $lblStatus.Text = "" + $res.message
          & $performSearch -ReloadFromExcel
        }
        else {
          [System.Windows.Forms.MessageBox]::Show(
            "" + $res.message,
            "Check-In Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
        }
      }
      catch {
        Log "ERROR" "Dashboard CHECK-IN failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
          "Check-In failed: $($_.Exception.Message)",
          "Check-In Error",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
      }
    }
  })

  $btnCheckOut.Add_Click({
    & $invokeDashboardAction $btnCheckOut "Checking-Out..." {
      try {
        $row = & $getSelectedRow
        if (-not $row) {
          [System.Windows.Forms.MessageBox]::Show(
            "Select one row first.",
            "Dashboard",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
          ) | Out-Null
          return
        }
        $note = ("" + $txtComment.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($note)) {
          $note = $DashboardDefaultCheckOutNote
          $txtComment.Text = $note
        }
        $ritmSel = ("" + $row.RITM).Trim().ToUpperInvariant()
        $candOut = Get-DashboardCheckOutCandidate -wv $wv -RitmNumber $ritmSel
        if (-not $candOut) {
          [System.Windows.Forms.MessageBox]::Show(
            "No Open/Work in Progress task detected for $ritmSel.",
            "Check-Out",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
          return
        }
        $confirmOut = [System.Windows.Forms.MessageBox]::Show(
          "RITM: $ritmSel`r`nSCTASK: $($candOut.number)`r`nCurrent: $($candOut.state_text) [$($candOut.state_value)]`r`nTarget: Appointment (task only, parent RITM unchanged)`r`n`r`nContinue?",
          "Confirm Check-Out",
          [System.Windows.Forms.MessageBoxButtons]::YesNo,
          [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirmOut -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $res = Invoke-DashboardCheckOut -wv $wv -ExcelPath $ExcelPath -SheetName $SheetName -RowItem $row -WorkNote $note
        if ($res.ok -eq $true) {
          $lblStatus.Text = "" + $res.message
          & $performSearch -ReloadFromExcel
        }
        else {
          [System.Windows.Forms.MessageBox]::Show(
            "" + $res.message,
            "Check-Out Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
          ) | Out-Null
        }
      }
      catch {
        Log "ERROR" "Dashboard CHECK-OUT failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
          "Check-Out failed: $($_.Exception.Message)",
          "Check-Out Error",
          [System.Windows.Forms.MessageBoxButtons]::OK,
          [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
      }
    }
  })

  # Initial load: keep grid empty by design.
  try {
    & $performSearch
    $lblStatus.Text = "Ready. Rows loaded: $($state.AllRows.Count)."
  } catch {}

  [void]$form.ShowDialog()
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
    [string]$SCTasksHeader,
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
  $sctasksCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header $SCTasksHeader
  $sctaskSplitCol = Get-OrCreateHeaderColumn -ws $ws -map $map -Header "SCTask Split"
  $sctaskExpansionRequests = New-Object System.Collections.Generic.List[object]

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

    # Make Number clickable for RITM/INC records.
    if (($ticket -like "RITM*") -or ($ticket -like "INC*")) {
      $sysIdOut = if ($res.PSObject.Properties["sys_id"]) { ("" + $res.sys_id).Trim() } else { "" }
      $ticketUrl = ""
      if ($ticket -like "RITM*") {
        $ticketUrl = Build-RitmBestUrl -SysId $sysIdOut -RitmNumber $ticket
      }
      elseif ($ticket -like "INC*") {
        $ticketUrl = Build-IncidentBestUrl -SysId $sysIdOut -IncNumber $ticket
      }
      if (-not [string]::IsNullOrWhiteSpace($ticketUrl)) {
        Set-ExcelHyperlinkSafe -ws $ws -Row $r -Col $ticketCol -DisplayText $ticket -Url $ticketUrl -TicketForLog $ticket
      }
    }

    # Fill "Name" (affected_user)
    $nameCell = "" + $ws.Cells.Item($r, $map[$NameHeader]).Text
    $nameOut = ("" + $res.affected_user).Trim()
    $legalNameOut = if ($res.PSObject.Properties["legal_name"]) { ("" + $res.legal_name).Trim() } else { "" }
    if (($ticket -like "RITM*") -and $legalNameOut -and $ForceUpdateNameFromLegal) {
      $nameOut = $legalNameOut
    }
    elseif (($ticket -like "RITM*") -and $legalNameOut -and (Is-InvalidUserDisplay $nameOut)) {
      $nameOut = $legalNameOut
    }
    if ((Is-EmptyOrPlaceholder $nameCell $ticket) -or (($ticket -like "RITM*") -and $legalNameOut -and $ForceUpdateNameFromLegal)) {
      $ws.Cells.Item($r, $map[$NameHeader]) = $nameOut
    }

    # Determine detected PI/machine (if present)
    $detectedPiOut = ""
    if ($res.PSObject.Properties["detected_pi_machine"]) {
      $detectedPiOut = ("" + $res.detected_pi_machine).Trim()
    }

    # Fill PI column:
    # - Prefer detected PI for RITM
    # - Else use configuration_item (original behavior)
    $phoneCell = "" + $ws.Cells.Item($r, $map[$PhoneHeader]).Text
    $phoneOut = ("" + $res.configuration_item).Trim()
    if ((($ticket -like "RITM*") -or ($ticket -like "INC*")) -and $detectedPiOut) {
      $phoneOut = $detectedPiOut
    }
    if (($ticket -like "RITM*") -and $ForceUpdateDetectedPI) {
      $ws.Cells.Item($r, $map[$PhoneHeader]) = $detectedPiOut
      Log "INFO" "$ticket PI force-updated in '$PhoneHeader' => '$detectedPiOut'"
    }
    elseif ((($ticket -like "INC*")) -and $detectedPiOut -and $ForceUpdateDetectedPI) {
      $ws.Cells.Item($r, $map[$PhoneHeader]) = $phoneOut
      Log "INFO" "$ticket PI force-updated in '$PhoneHeader' => '$phoneOut'"
    }
    elseif (Is-EmptyOrPlaceholder $phoneCell $ticket) {
      $ws.Cells.Item($r, $map[$PhoneHeader]) = $phoneOut
    }

    # Fill "Action finished?" (status)
    $actionCell = "" + $ws.Cells.Item($r, $map[$ActionHeader]).Text
    $statusOut = Get-CompletionStatusForExcel -Ticket $ticket -Res $res
    if (($ticket -like "INC*") -or ($ticket -like "RITM*")) {
      # Always refresh INC/RITM status so stale values are corrected.
      $ws.Cells.Item($r, $map[$ActionHeader]) = $statusOut
    }
    elseif (Is-EmptyOrPlaceholder $actionCell $ticket) {
      $ws.Cells.Item($r, $map[$ActionHeader]) = $statusOut
    }

    # Fill open SCTASK(s) into single "SCTasks" column
    $openTasks = @()
    if ($res.PSObject.Properties["open_task_items"] -and $res.open_task_items) {
      $openTasks = @($res.open_task_items)
    }

    if ($ticket -like "RITM*") {
      if ($openTasks.Count -gt 0) {
        $openTaskNumbers = New-Object System.Collections.Generic.List[string]
        $firstTaskSysId = ""
        foreach ($taskObj in $openTasks) {
          $taskNumber = if ($taskObj.PSObject.Properties["number"]) { ("" + $taskObj.number).Trim() } else { "" }
          if (-not $firstTaskSysId -and $taskObj.PSObject.Properties["sys_id"]) {
            $firstTaskSysId = ("" + $taskObj.sys_id).Trim()
          }
          if ($taskNumber) { [void]$openTaskNumbers.Add($taskNumber) }
        }
        $existingSplit = ("" + $ws.Cells.Item($r, $sctaskSplitCol).Text).Trim()
        if ($openTaskNumbers.Count -gt 0) {
          $firstTask = $openTaskNumbers[0]
          $taskUrl = Build-SCTaskBestUrl -SysId $firstTaskSysId -TaskNumber $firstTask
          if (-not [string]::IsNullOrWhiteSpace($taskUrl)) {
            Set-ExcelHyperlinkSafe -ws $ws -Row $r -Col $sctasksCol -DisplayText $firstTask -Url $taskUrl -TicketForLog $ticket
          }
          else {
            $ws.Cells.Item($r, $sctasksCol) = $firstTask
          }
          $ws.Cells.Item($r, $sctaskSplitCol) = "PARENT"
        }
        else {
          $ws.Cells.Item($r, $sctasksCol) = "Open tasks: $($openTasks.Count)"
          $ws.Cells.Item($r, $sctaskSplitCol) = "PARENT"
        }

        if ($EnableSctaskRowExpansion -and ($openTaskNumbers.Count -gt 1)) {
          $sctaskExpansionRequests.Add([pscustomobject]@{
            Row = [int]$r
            Ticket = $ticket
            TaskItems = @($openTasks)
            NameValue = ("" + $ws.Cells.Item($r, $map[$NameHeader]).Text)
            PhoneValue = ("" + $ws.Cells.Item($r, $map[$PhoneHeader]).Text)
            ActionValue = ("" + $ws.Cells.Item($r, $map[$ActionHeader]).Text)
            TicketUrl = if ($ticket -like "RITM*") { Build-RitmBestUrl -SysId $sysIdOut -RitmNumber $ticket } else { "" }
          }) | Out-Null
        }
        elseif ($openTaskNumbers.Count -eq 1) {
          # Remove stale AUTO split rows left from previous runs when only one task remains.
          $ws.Cells.Item($r, $sctaskSplitCol) = "PARENT"
        }
      }
      else {
        $cell = $ws.Cells.Item($r, $sctasksCol)
        try { if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() } } catch {}
        $ws.Cells.Item($r, $sctasksCol) = "No open tasks."
        $ws.Cells.Item($r, $sctaskSplitCol) = ""
      }
    }
    else {
      $cell = $ws.Cells.Item($r, $sctasksCol)
      try { if ($cell.Hyperlinks.Count -gt 0) { $cell.Hyperlinks.Delete() } } catch {}
      $ws.Cells.Item($r, $sctasksCol) = ""
      $ws.Cells.Item($r, $sctaskSplitCol) = ""
    }
  }

  # Expand RITM rows: one Excel row per open SCTASK with direct hyperlink.
  if ($EnableSctaskRowExpansion -and ($sctaskExpansionRequests.Count -gt 0)) {
    $expansionSorted = @($sctaskExpansionRequests | Sort-Object -Property Row -Descending)
    foreach ($req in $expansionSorted) {
      $baseRow = [int]$req.Row
      $ticket = ("" + $req.Ticket).Trim().ToUpperInvariant()
      $taskObjs = @($req.TaskItems)
      if ($taskObjs.Count -le 1) { continue }

      # Delete previous AUTO rows directly under parent to avoid duplicates between runs.
      $scan = $baseRow + 1
      while ($true) {
        $ticketTxt = ("" + $ws.Cells.Item($scan, $ticketCol).Text).Trim().ToUpperInvariant()
        $splitTxt = ("" + $ws.Cells.Item($scan, $sctaskSplitCol).Text).Trim().ToUpperInvariant()
        if (($ticketTxt -eq $ticket) -and ($splitTxt -eq "AUTO")) {
          $ws.Rows.Item($scan).Delete() | Out-Null
          continue
        }
        break
      }

      # Build task list and keep parent row = first task.
      $taskNums = New-Object System.Collections.Generic.List[string]
      $taskSysMap = @{}
      foreach ($to in $taskObjs) {
        $tn = if ($to.PSObject.Properties["number"]) { ("" + $to.number).Trim().ToUpperInvariant() } else { "" }
        $ts = if ($to.PSObject.Properties["sys_id"]) { ("" + $to.sys_id).Trim() } else { "" }
        if ([string]::IsNullOrWhiteSpace($tn)) { continue }
        if (-not $taskSysMap.ContainsKey($tn)) {
          $taskSysMap[$tn] = $ts
          [void]$taskNums.Add($tn)
        }
      }
      if ($taskNums.Count -le 1) { continue }

      for ($iTask = 1; $iTask -lt $taskNums.Count; $iTask++) {
        $insRow = $baseRow + $iTask
        $ws.Rows.Item($insRow).Insert() | Out-Null

        # Keep key values aligned with parent.
        $ws.Cells.Item($insRow, $ticketCol) = $ticket
        $ws.Cells.Item($insRow, $map[$NameHeader]) = "" + $req.NameValue
        $ws.Cells.Item($insRow, $map[$PhoneHeader]) = "" + $req.PhoneValue
        $ws.Cells.Item($insRow, $map[$ActionHeader]) = "" + $req.ActionValue
        $ws.Cells.Item($insRow, $sctaskSplitCol) = "AUTO"

        if (-not [string]::IsNullOrWhiteSpace($req.TicketUrl)) {
          Set-ExcelHyperlinkSafe -ws $ws -Row $insRow -Col $ticketCol -DisplayText $ticket -Url ("" + $req.TicketUrl) -TicketForLog $ticket
        }

        $tn2 = $taskNums[$iTask]
        $ts2 = if ($taskSysMap.ContainsKey($tn2)) { "" + $taskSysMap[$tn2] } else { "" }
        $tu2 = Build-SCTaskBestUrl -SysId $ts2 -TaskNumber $tn2
        if (-not [string]::IsNullOrWhiteSpace($tu2)) {
          Set-ExcelHyperlinkSafe -ws $ws -Row $insRow -Col $sctasksCol -DisplayText $tn2 -Url $tu2 -TicketForLog $ticket
        }
        else {
          $ws.Cells.Item($insRow, $sctasksCol) = $tn2
        }
      }
    }
  }
  elseif (-not $EnableSctaskRowExpansion) {
    # Cleanup legacy AUTO rows from previous runs to prevent duplicated RITM lines.
    for ($scan = $rows; $scan -ge 2; $scan--) {
      $splitTxt = ("" + $ws.Cells.Item($scan, $sctaskSplitCol).Text).Trim().ToUpperInvariant()
      if ($splitTxt -eq "AUTO") {
        try { $ws.Rows.Item($scan).Delete() | Out-Null } catch {}
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

function Convert-SnowTimestampToUtc {
  param([string]$Value)
  $txt = ("" + $Value).Trim()
  if ([string]::IsNullOrWhiteSpace($txt)) { return $null }

  $dt = [datetime]::MinValue
  $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
  $formats = @(
    "yyyy-MM-dd HH:mm:ss",
    "yyyy-MM-ddTHH:mm:ss",
    "yyyy-MM-ddTHH:mm:ssZ",
    "yyyy-MM-ddTHH:mm:ss.fffZ"
  )
  if ([datetime]::TryParseExact($txt, $formats, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$dt)) {
    return $dt.ToUniversalTime()
  }
  if ([datetime]::TryParse($txt, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$dt)) {
    return $dt.ToUniversalTime()
  }
  return $null
}

function Convert-ToUtcStampText {
  param([datetime]$UtcDate)
  if (-not $UtcDate) { return "" }
  return $UtcDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
}

function Compare-TimestampValues {
  param(
    [string]$Left,
    [string]$Right
  )
  $l = ("" + $Left).Trim()
  $r = ("" + $Right).Trim()
  if (-not $l -and -not $r) { return 0 }
  if (-not $l) { return -1 }
  if (-not $r) { return 1 }

  $ld = Convert-SnowTimestampToUtc -Value $l
  $rd = Convert-SnowTimestampToUtc -Value $r
  if ($ld -and $rd) {
    if ($ld -lt $rd) { return -1 }
    if ($ld -gt $rd) { return 1 }
    return 0
  }
  return [string]::CompareOrdinal($l, $r)
}

function Test-IncrementalTaskClosedState {
  param([string]$State)
  $s = ("" + $State).Trim().ToLowerInvariant()
  if (-not $s) { return $false }
  if ($s -in @("3", "4", "7")) { return $true }
  if ($s -match 'closed|complete|cancel') { return $true }
  return $false
}

function Get-IncrementalTaskOpenCount {
  param([object[]]$Tasks)
  $count = 0
  foreach ($task in @($Tasks)) {
    $state = if ($task.PSObject.Properties["state"]) { ("" + $task.state).Trim() } else { "" }
    $activeRaw = if ($task.PSObject.Properties["active"]) { ("" + $task.active).Trim().ToLowerInvariant() } else { "" }
    $isActive = $activeRaw -in @("true", "1", "yes")
    if ((-not (Test-IncrementalTaskClosedState -State $state)) -or $isActive) {
      $count++
    }
  }
  return [int]$count
}

function Get-IncrementalTasksUpdatedOnMax {
  param([object[]]$Tasks)
  $maxRaw = ""
  foreach ($task in @($Tasks)) {
    $updated = if ($task.PSObject.Properties["sys_updated_on"]) { ("" + $task.sys_updated_on).Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($updated)) { continue }
    if ([string]::IsNullOrWhiteSpace($maxRaw)) { $maxRaw = $updated; continue }
    if ((Compare-TimestampValues -Left $updated -Right $maxRaw) -gt 0) { $maxRaw = $updated }
  }
  return $maxRaw
}

function Load-RitmScanDatabase {
  param([string]$Path)
  $map = @{}
  $script:RitmScanDbLastSuccessfulUtc = ""
  if (-not (Test-Path -LiteralPath $Path)) { return $map }
  try {
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $map }
    $doc = $raw | ConvertFrom-Json -ErrorAction Stop
    if ($doc -and $doc.PSObject.Properties["last_successful_scan_utc"]) {
      $script:RitmScanDbLastSuccessfulUtc = ("" + $doc.last_successful_scan_utc).Trim()
    }

    $rows = @()
    if ($doc -is [System.Array]) { $rows = @($doc) }
    elseif ($doc -and $doc.PSObject.Properties["items"]) { $rows = @($doc.items) }

    foreach ($row in $rows) {
      $ritm = if ($row -and $row.PSObject.Properties["RITM"]) { ("" + $row.RITM).Trim().ToUpperInvariant() } else { "" }
      if (-not ($ritm -match '^RITM\d{6,8}$')) { continue }
      $map[$ritm] = [pscustomobject]@{
        RITM                = $ritm
        LastScanUtc         = if ($row.PSObject.Properties["LastScanUtc"]) { ("" + $row.LastScanUtc).Trim() } else { "" }
        RITM_UpdatedOn      = if ($row.PSObject.Properties["RITM_UpdatedOn"]) { ("" + $row.RITM_UpdatedOn).Trim() } else { "" }
        Tasks_UpdatedOn_Max = if ($row.PSObject.Properties["Tasks_UpdatedOn_Max"]) { ("" + $row.Tasks_UpdatedOn_Max).Trim() } else { "" }
        OpenTaskCount       = if ($row.PSObject.Properties["OpenTaskCount"]) { [int]$row.OpenTaskCount } else { 0 }
        Status              = if ($row.PSObject.Properties["Status"]) { ("" + $row.Status).Trim() } else { "" }
        SkipReason          = if ($row.PSObject.Properties["SkipReason"]) { ("" + $row.SkipReason).Trim() } else { "" }
      }
    }
  }
  catch {
    Log "ERROR" "Failed to load incremental scan DB '$Path': $($_.Exception.Message)"
  }
  return $map
}

function Save-RitmScanDatabase {
  param(
    [string]$Path,
    [hashtable]$Map,
    [string]$LastSuccessfulScanUtc = ""
  )
  try {
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
      New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
    $items = @()
    foreach ($k in @($Map.Keys | Sort-Object)) {
      $items += $Map[$k]
    }
    $payload = [pscustomobject]@{
      version = 1
      generated_utc = (Convert-ToUtcStampText -UtcDate (Get-Date).ToUniversalTime())
      last_successful_scan_utc = if ([string]::IsNullOrWhiteSpace($LastSuccessfulScanUtc)) { $script:RitmScanDbLastSuccessfulUtc } else { ("" + $LastSuccessfulScanUtc).Trim() }
      items = $items
    }
    $json = $payload | ConvertTo-Json -Depth 6
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
    Log "INFO" "Incremental scan DB saved: '$Path' entries=$($items.Count)"
    return $true
  }
  catch {
    Log "ERROR" "Failed to save incremental scan DB '$Path': $($_.Exception.Message)"
    return $false
  }
}

function Get-RitmIncrementalSnapshot {
  param(
    $wv,
    [string]$RitmNumber
  )
  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if (-not ($ritm -match '^RITM\d{6,8}$')) {
    return [pscustomobject]@{ ok = $false; reason = "invalid_ritm"; ritm = $ritm; ritm_updated_on = ""; ritm_state = ""; tasks = @(); tasks_updated_on_max = ""; open_task_count = 0 }
  }

  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRec(obj){
      return (obj && obj.records && obj.records[0]) ? obj.records[0] :
             (obj && obj.result && obj.result[0]) ? obj.result[0] : null;
    }
    function pickRows(obj){
      if (obj && Array.isArray(obj.records)) return obj.records;
      if (obj && Array.isArray(obj.result)) return obj.result;
      return [];
    }
    function get(path){
      var x = new XMLHttpRequest();
      x.open('GET', path, false);
      x.withCredentials = true;
      x.send(null);
      if (!(x.status>=200 && x.status<300)) return null;
      try { return JSON.parse(x.responseText || '{}'); } catch(e){ return null; }
    }

    var qR = encodeURIComponent('number=$ritm');
    var pR = '/sc_req_item.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_fields=number,sys_updated_on,state&sysparm_query=' + qR;
    var oR = get(pR);
    var r = pickRec(oR);
    if (!r) return JSON.stringify({ok:false, reason:'ritm_not_found', ritm:'$ritm', tasks:[]});

    var qT = encodeURIComponent('request_item.number=$ritm');
    var pT = '/sc_task.do?JSONv2&sysparm_limit=200&sysparm_display_value=true&sysparm_fields=sys_updated_on,state,active&sysparm_query=' + qT;
    var oT = get(pT);
    var rows = pickRows(oT);
    var tasks = [];
    for (var i = 0; i < rows.length; i++) {
      var t = rows[i] || {};
      tasks.push({
        sys_updated_on: s(t.sys_updated_on || ''),
        state: s(t.state || ''),
        active: s(t.active || '')
      });
    }
    return JSON.stringify({
      ok:true,
      ritm:'$ritm',
      ritm_updated_on:s(r.sys_updated_on || ''),
      ritm_state:s(r.state || ''),
      tasks:tasks
    });
  } catch(e){
    return JSON.stringify({ok:false, reason:'snapshot_exception', error:''+e, ritm:'$ritm', tasks:[]});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 9000)
  if (-not $o) {
    return [pscustomobject]@{ ok = $false; reason = "snapshot_no_response"; ritm = $ritm; ritm_updated_on = ""; ritm_state = ""; tasks = @(); tasks_updated_on_max = ""; open_task_count = 0 }
  }

  $tasks = if ($o.PSObject.Properties["tasks"] -and $o.tasks) { @($o.tasks) } else { @() }
  $tasksMax = Get-IncrementalTasksUpdatedOnMax -Tasks $tasks
  $openCount = Get-IncrementalTaskOpenCount -Tasks $tasks
  return [pscustomobject]@{
    ok                   = ($o.ok -eq $true)
    reason               = if ($o.PSObject.Properties["reason"]) { ("" + $o.reason).Trim() } else { "" }
    ritm                 = $ritm
    ritm_updated_on      = if ($o.PSObject.Properties["ritm_updated_on"]) { ("" + $o.ritm_updated_on).Trim() } else { "" }
    ritm_state           = if ($o.PSObject.Properties["ritm_state"]) { ("" + $o.ritm_state).Trim() } else { "" }
    tasks                = $tasks
    tasks_updated_on_max = $tasksMax
    open_task_count      = [int]$openCount
  }
}

function Get-RitmStatusFromSnapshot {
  param([object]$Snapshot)
  if (-not $Snapshot) { return "" }
  $openCount = 0
  try { $openCount = [int]$Snapshot.open_task_count } catch { $openCount = 0 }
  if ($openCount -gt 0) { return "Open:$openCount" }
  if ($Snapshot.PSObject.Properties["ritm_state"]) { return ("" + $Snapshot.ritm_state).Trim() }
  return ""
}

function Get-RitmIncrementalDecision {
  param(
    [string]$Ritm,
    [object]$Snapshot,
    [object]$Previous
  )
  if (-not $Snapshot -or $Snapshot.ok -ne $true) {
    $reason = if ($Snapshot -and $Snapshot.PSObject.Properties["reason"]) { ("" + $Snapshot.reason).Trim() } else { "snapshot_failed" }
    return [pscustomobject]@{ process = $true; skip_reason = ""; process_reason = $reason }
  }
  if (-not $Previous) {
    return [pscustomobject]@{ process = $true; skip_reason = ""; process_reason = "first_scan" }
  }

  $prevRitmUpdated = if ($Previous.PSObject.Properties["RITM_UpdatedOn"]) { ("" + $Previous.RITM_UpdatedOn).Trim() } else { "" }
  $prevTasksUpdated = if ($Previous.PSObject.Properties["Tasks_UpdatedOn_Max"]) { ("" + $Previous.Tasks_UpdatedOn_Max).Trim() } else { "" }
  if ([string]::IsNullOrWhiteSpace($prevRitmUpdated)) {
    return [pscustomobject]@{ process = $true; skip_reason = ""; process_reason = "missing_previous_ritm_updated_on" }
  }
  if ([string]::IsNullOrWhiteSpace($prevTasksUpdated)) {
    return [pscustomobject]@{ process = $true; skip_reason = ""; process_reason = "missing_previous_tasks_updated_on_max" }
  }

  $ritmCmp = Compare-TimestampValues -Left ("" + $Snapshot.ritm_updated_on) -Right $prevRitmUpdated
  $taskCmp = Compare-TimestampValues -Left ("" + $Snapshot.tasks_updated_on_max) -Right $prevTasksUpdated
  if (($ritmCmp -le 0) -and ($taskCmp -le 0)) {
    $openCount = 0
    try { $openCount = [int]$Snapshot.open_task_count } catch { $openCount = 0 }
    $skipReason = if ($openCount -eq 0) { "unchanged_all_tasks_closed" } else { "unchanged" }
    return [pscustomobject]@{ process = $false; skip_reason = $skipReason; process_reason = "" }
  }
  return [pscustomobject]@{ process = $true; skip_reason = ""; process_reason = "changed_timestamps" }
}

function Update-RitmScanDbEntry {
  param(
    [hashtable]$Db,
    [string]$Ritm,
    [object]$Snapshot,
    [string]$Status,
    [string]$SkipReason
  )
  $ritmKey = ("" + $Ritm).Trim().ToUpperInvariant()
  if (-not ($ritmKey -match '^RITM\d{6,8}$')) { return }

  $openTaskCount = 0
  try {
    if ($Snapshot -and $Snapshot.PSObject.Properties["open_task_count"]) { $openTaskCount = [int]$Snapshot.open_task_count }
  } catch { $openTaskCount = 0 }

  $Db[$ritmKey] = [pscustomobject]@{
    RITM                = $ritmKey
    LastScanUtc         = (Convert-ToUtcStampText -UtcDate (Get-Date).ToUniversalTime())
    RITM_UpdatedOn      = if ($Snapshot -and $Snapshot.PSObject.Properties["ritm_updated_on"]) { ("" + $Snapshot.ritm_updated_on).Trim() } else { "" }
    Tasks_UpdatedOn_Max = if ($Snapshot -and $Snapshot.PSObject.Properties["tasks_updated_on_max"]) { ("" + $Snapshot.tasks_updated_on_max).Trim() } else { "" }
    OpenTaskCount       = $openTaskCount
    Status              = ("" + $Status).Trim()
    SkipReason          = ("" + $SkipReason).Trim()
  }
}

function Convert-UtcStampToSnowQueryDate {
  param([string]$UtcStamp)
  $dt = Convert-SnowTimestampToUtc -Value $UtcStamp
  if (-not $dt) { return "" }
  return $dt.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
}

function Get-RitmNumbersUpdatedSince {
  param(
    $wv,
    [string]$SinceUtc
  )
  $sinceSnow = Convert-UtcStampToSnowQueryDate -UtcStamp $SinceUtc
  if ([string]::IsNullOrWhiteSpace($sinceSnow)) { return @() }
  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRows(o){
      if (o && Array.isArray(o.records)) return o.records;
      if (o && Array.isArray(o.result)) return o.result;
      return [];
    }
    function get(path){
      var x = new XMLHttpRequest();
      x.open('GET', path, false);
      x.withCredentials = true;
      x.send(null);
      if (!(x.status>=200 && x.status<300)) return null;
      try { return JSON.parse(x.responseText || '{}'); } catch(e){ return null; }
    }
    var out = [];
    var seen = {};
    var limit = 200;
    for (var page = 0; page < 40; page++) {
      var offset = page * limit;
      var q = encodeURIComponent('sys_updated_on>=$sinceSnow');
      var p = '/sc_req_item.do?JSONv2&sysparm_limit=' + limit + '&sysparm_offset=' + offset + '&sysparm_fields=number&sysparm_query=' + q;
      var o = get(p);
      var rows = pickRows(o);
      if (!rows || rows.length === 0) break;
      for (var i = 0; i < rows.length; i++) {
        var n = s(rows[i].number || '').toUpperCase();
        if (!/^RITM\d{6,8}$/.test(n)) continue;
        if (seen[n]) continue;
        seen[n] = true;
        out.push(n);
      }
      if (rows.length < limit) break;
    }
    return JSON.stringify({ok:true, items:out});
  } catch(e){
    return JSON.stringify({ok:false, items:[]});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 12000)
  if (-not $o -or $o.ok -ne $true -or -not $o.PSObject.Properties["items"]) { return @() }
  return @($o.items)
}

function Get-RitmNumbersFromUpdatedTasksSince {
  param(
    $wv,
    [string]$SinceUtc
  )
  $sinceSnow = Convert-UtcStampToSnowQueryDate -UtcStamp $SinceUtc
  if ([string]::IsNullOrWhiteSpace($sinceSnow)) { return @() }
  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRows(o){
      if (o && Array.isArray(o.records)) return o.records;
      if (o && Array.isArray(o.result)) return o.result;
      return [];
    }
    function get(path){
      var x = new XMLHttpRequest();
      x.open('GET', path, false);
      x.withCredentials = true;
      x.send(null);
      if (!(x.status>=200 && x.status<300)) return null;
      try { return JSON.parse(x.responseText || '{}'); } catch(e){ return null; }
    }
    var out = [];
    var seen = {};
    var limit = 200;
    for (var page = 0; page < 40; page++) {
      var offset = page * limit;
      var q = encodeURIComponent('sys_updated_on>=$sinceSnow');
      var p = '/sc_task.do?JSONv2&sysparm_display_value=true&sysparm_limit=' + limit + '&sysparm_offset=' + offset + '&sysparm_fields=request_item,request_item.number&sysparm_query=' + q;
      var o = get(p);
      var rows = pickRows(o);
      if (!rows || rows.length === 0) break;
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i] || {};
        var cands = [
          s(r['request_item.number'] || ''),
          s(r.request_item || '')
        ];
        for (var j = 0; j < cands.length; j++) {
          var n = cands[j].toUpperCase();
          if (!/^RITM\d{6,8}$/.test(n)) continue;
          if (seen[n]) continue;
          seen[n] = true;
          out.push(n);
        }
      }
      if (rows.length < limit) break;
    }
    return JSON.stringify({ok:true, items:out});
  } catch(e){
    return JSON.stringify({ok:false, items:[]});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 12000)
  if (-not $o -or $o.ok -ne $true -or -not $o.PSObject.Properties["items"]) { return @() }
  return @($o.items)
}

function Get-StateLabelCacheFromSnow {
  param($wv)
  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)
  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function pickRows(o){
      if (o && Array.isArray(o.records)) return o.records;
      if (o && Array.isArray(o.result)) return o.result;
      return [];
    }
    var q = encodeURIComponent('nameINsc_req_item,incident,sc_task^element=state');
    var p = '/sys_choice.do?JSONv2&sysparm_limit=500&sysparm_fields=name,element,value,label&sysparm_query=' + q;
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) return JSON.stringify({ok:false, cache:{}});
    var o = {};
    try { o = JSON.parse(x.responseText || '{}'); } catch(e){ return JSON.stringify({ok:false, cache:{}}); }
    var rows = pickRows(o);
    var cache = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      var name = s(r.name || '').toLowerCase();
      var val = s(r.value || '');
      var label = s(r.label || '');
      if (!name || !val || !label) continue;
      if (!cache[name]) cache[name] = {};
      if (!cache[name][val]) cache[name][val] = label;
    }
    return JSON.stringify({ok:true, cache:cache});
  } catch(e){
    return JSON.stringify({ok:false, cache:{}});
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $wv $js 10000)
  if (-not $o -or $o.ok -ne $true -or -not $o.PSObject.Properties["cache"]) { return @{} }
  $out = @{}
  foreach ($p in $o.cache.PSObject.Properties) {
    $inner = @{}
    if ($p.Value -and $p.Value.PSObject -and $p.Value.PSObject.Properties) {
      foreach ($q in $p.Value.PSObject.Properties) {
        $inner[("" + $q.Name)] = ("" + $q.Value)
      }
    }
    $out[("" + $p.Name)] = $inner
  }
  return $out
}

# =============================================================================
# EXTRACTION: JSONv2 via JavaScript inside authenticated WebView2
# =============================================================================
function Extract-Ticket_JSONv2 {
  param(
    $wv,
    [string]$Ticket,
    [hashtable]$StateLabelCache = @{}
  )

  # Determine which table we query for this ticket.
  $table = Ticket-ToTable $Ticket
  $closedStatesJson = ($ClosedTaskStates | ConvertTo-Json -Compress)
  $enableActivitySearchJs = if ($EnableActivityStreamSearch) { "true" } else { "false" }
  $stateLabelCacheJson = if ($StateLabelCache -and $StateLabelCache.Count -gt 0) { $StateLabelCache | ConvertTo-Json -Compress } else { "{}" }
  $useSysChoiceLookupJs = if ($DisableSysChoiceLookup) { "false" } else { "true" }

  # Make sure the browser context is ready before running heavy JS.
  [void](Ensure-SnowReady -wv $wv -MaxWaitMs 6000)

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
      var tn = s(table || '').toLowerCase();
      if (STATE_LABEL_CACHE[tn] && STATE_LABEL_CACHE[tn][s(value)]) {
        return s(STATE_LABEL_CACHE[tn][s(value)]);
      }
      if (!USE_SYS_CHOICE_LOOKUP) return "";
      var u = '/sys_choice.do?JSONv2&sysparm_limit=1&sysparm_query=' +
              encodeURIComponent('name=' + table + '^element=state^value=' + value);
      var obj = httpGetJsonV2(u);
      var rec = pickRec(obj);
      if (!rec) return "";
      var lbl = s(rec.label || rec.value || "");
      if (lbl) {
        if (!STATE_LABEL_CACHE[tn]) STATE_LABEL_CACHE[tn] = {};
        STATE_LABEL_CACHE[tn][s(value)] = lbl;
      }
      return lbl;
    }

    var STATE_LABEL_CACHE = $stateLabelCacheJson;
    var USE_SYS_CHOICE_LOOKUP = $useSysChoiceLookupJs;

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
        var stLabel = "";
        // Fast path: sc_task.state is usually numeric, so avoid sys_choice calls per task.
        if (!/^\d+$/.test(stVal)) {
          stLabel = resolveStateLabel('sc_task', stVal);
        }
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
      var p = path + '?JSONv2&sysparm_limit=200&sysparm_display_value=true&sysparm_query=' + encodeURIComponent(query);
      var o = httpGetJsonV2(p);
      var rows = (o && o.records) ? o.records : ((o && o.result) ? o.result : []);
      return Array.isArray(rows) ? rows : [];
    }

    function splitIntoChunks(arr, size){
      var out = [];
      if (!Array.isArray(arr) || size <= 0) return out;
      for (var i = 0; i < arr.length; i += size) {
        out.push(arr.slice(i, i + size));
      }
      return out;
    }

    function collectAllTextFromRow(row, out){
      try {
        if (!row || !out) return;
        for (var k in row) {
          if (!Object.prototype.hasOwnProperty.call(row, k)) continue;
          var v = row[k];
          if (v === null || v === undefined) continue;
          var t = s(v);
          if (!t) continue;
          // Skip noisy pure sys_ids
          if (/^[0-9a-f]{32}$/i.test(t)) continue;
          out.push(t);
        }
      } catch(e) {}
    }

    function isNewEpUserName(v){
      return /new[\s\W_]*ep[\s\W_]*user|new[\s\W_]*user/i.test(s(v));
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

    function getRitmTasksEvidenceText(reqItemSysId, ritmNumber){
      var out = [];
      var rowsAll = [];

      try {
        if (looksSysId(reqItemSysId)) {
          rowsAll = rowsAll.concat(getRows('/sc_task.do', 'request_item=' + reqItemSysId));
        }
        if (s(ritmNumber)) {
          rowsAll = rowsAll.concat(getRows('/sc_task.do', 'request_item.number=' + ritmNumber));
        }
      } catch(et) {
        activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";task_rows") : "task_rows";
      }

      var seenTask = {};
      var rows = [];
      for (var i0 = 0; i0 < rowsAll.length; i0++) {
        var k = s(rowsAll[i0].sys_id || rowsAll[i0].number || ("idx_" + i0));
        if (!seenTask[k]) {
          seenTask[k] = true;
          rows.push(rowsAll[i0]);
        }
      }

      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var desc = s(r.short_description || "");
        if (desc) out.push(desc);
        // Include known fields and all textual fields defensively.
        var fields = [r.number, r.item, r.description, r.comments, r.work_notes, r.close_notes, r.u_comments];
        for (var fi = 0; fi < fields.length; fi++) { var v = s(fields[fi] || ""); if (v) out.push(v); }
        collectAllTextFromRow(r, out);

        if (/prepare\s*device\s*for\s*new\s*user/i.test(desc)) {
          out.push("TASK_HINT:PREPARE_DEVICE_FOR_NEW_USER");
        }
      }

      // Add journal entries from task(s) in batches for speed.
      try {
        var taskIds = [];
        for (var j = 0; j < rows.length; j++) {
          var ts = s(rows[j].sys_id || "");
          if (looksSysId(ts)) taskIds.push(ts);
        }
        var chunks = splitIntoChunks(taskIds, 20);
        for (var ci = 0; ci < chunks.length; ci++) {
          var inClause = chunks[ci].join(',');
          var q = 'name=sc_task^element_idIN' + inClause + '^elementINcomments,work_notes';
          var jAll = getRows('/sys_journal_field.do', q);
          for (var k2 = 0; k2 < jAll.length; k2++) {
            var jv = s(jAll[k2].value || jAll[k2].message || jAll[k2].comments || jAll[k2].work_notes || "");
            if (jv) out.push(jv);
            collectAllTextFromRow(jAll[k2], out);
          }
        }
      } catch(ej) {
        activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";task_journal") : "task_journal";
      }

      return out.join(' ');
    }

    function getPrepareDevicePiFromTasks(reqItemSysId, ritmNumber){
      var rowsAll = [];
      var selected = null;
      var selectedText = [];
      var selectedTaskNumber = "";

      try {
        if (looksSysId(reqItemSysId)) {
          rowsAll = rowsAll.concat(getRows('/sc_task.do', 'request_item=' + reqItemSysId));
        }
        if (s(ritmNumber)) {
          rowsAll = rowsAll.concat(getRows('/sc_task.do', 'request_item.number=' + ritmNumber));
        }
      } catch(e1) {
        activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";prepare_task_rows") : "prepare_task_rows";
      }

      var seen = {};
      var rows = [];
      for (var i0 = 0; i0 < rowsAll.length; i0++) {
        var k = s(rowsAll[i0].sys_id || rowsAll[i0].number || ("idx_" + i0));
        if (!seen[k]) { seen[k] = true; rows.push(rowsAll[i0]); }
      }

      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        var blob = [
          s(r.short_description || ""),
          s(r.description || ""),
          s(r.item || ""),
          s(r.comments || ""),
          s(r.work_notes || ""),
          s(r.close_notes || ""),
          s(r.u_comments || "")
        ].join(" ");
        if (/prepare[\s\W_]*device[\s\W_]*for[\s\W_]*new[\s\W_]*user/i.test(blob)) {
          selected = r;
          selectedTaskNumber = s(r.number || "");
          collectAllTextFromRow(r, selectedText);
          if (blob) selectedText.push(blob);
          break;
        }
      }

      if (!selected) {
        return { pi:"", task_number:"", text_len:0 };
      }

      try {
        var sid = s(selected.sys_id || "");
        if (looksSysId(sid)) {
          var jAll = [];
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'name=sc_task^element_id=' + sid));
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'element_id=' + sid + '^elementINcomments,work_notes'));
          for (var j = 0; j < jAll.length; j++) {
            collectAllTextFromRow(jAll[j], selectedText);
            var jv = s(jAll[j].value || jAll[j].message || jAll[j].comments || jAll[j].work_notes || "");
            if (jv) selectedText.push(jv);
          }
        }
      } catch(e2) {
        activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";prepare_task_journal") : "prepare_task_journal";
      }

      var txt = selectedText.join(" ");
      var pi = extractMachineFromActivityText(txt);
      return { pi:pi, task_number:selectedTaskNumber, text_len:s(txt).length };
    }

    function getIncidentActivitiesText(incSysId, incNumber){
      var out = [];
      try {
        if (looksSysId(incSysId)) {
          var jAll = [];
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'name=incident^element_id=' + incSysId));
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'element_id=' + incSysId));
          jAll = jAll.concat(getRows('/sys_journal_field.do', 'element_id=' + incSysId + '^elementINcomments,work_notes'));
          for (var i1 = 0; i1 < jAll.length; i1++) {
            var v1 = s(jAll[i1].value || jAll[i1].message || jAll[i1].comments || jAll[i1].work_notes || "");
            if (v1) out.push(v1);
          }
        }
      } catch(ej) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";inc_journal") : "inc_journal"; }

      try {
        if (s(incNumber)) {
          var iRows = getRows('/incident.do', 'number=' + incNumber);
          for (var i2 = 0; i2 < iRows.length; i2++) {
            var v2 = s(iRows[i2].comments || iRows[i2].work_notes || iRows[i2].description || iRows[i2].short_description || "");
            if (v2) out.push(v2);
          }
        }
      } catch(ei) { activityRetrievalError = activityRetrievalError ? (activityRetrievalError + ";inc_rows") : "inc_rows"; }
      return out.join(' ');
    }

    function extractMachineFromActivityText(activityText){
      var txt = s(activityText);
      if (!txt) return "";

      var pats = [
        /\b(?:02PI20[A-Z0-9_-]*|ITEC(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*|MUST(?:BRUN)?[A-Z0-9_-]*\d[A-Z0-9_-]*|EDPSBRUN[A-Z0-9_-]*\d[A-Z0-9_-]*|PRESBRUN[A-Z0-9_-]*\d[A-Z0-9_-]*|[A-Z]{3,}BRUN[A-Z0-9_-]*\d[A-Z0-9_-]*)\b/i,
        /\b(?:MUST|ITEC|EDPS|PRES)\s*[-_ ]?\s*BRUN\s*[-_ ]?\s*\d{6,}\b/i,
        /\b02\s*PI\s*20\s*\d{6,}\b/i
      ];
      for (var i = 0; i < pats.length; i++) {
        var m = txt.match(pats[i]);
        if (!m || !m[0]) continue;
        var v = s(m[0]).toUpperCase();
        if (/\s|-/.test(v)) v = v.replace(/[\s-]+/g, '');
        if (/\d/.test(v)) return v;
      }
      return "";
    }

    function extractLegalNameFromText(txt){
      var t = s(txt);
      if (!t) return "";
      var m0 = t.match(/LEGAL_NAME:\s*([^\r\n]+)/i);
      if (m0 && s(m0[1])) return s(m0[1]);
      var m = t.match(/Legal\s*name[\s:\-]*([A-Za-zÃ€-Ã–Ã˜-Ã¶Ã¸-Ã¿' \-]{3,})/i);
      var cand = m ? s(m[1]) : "";
      if (!cand) return "";
      if (/^explanation$/i.test(cand)) return "";
      return cand;
    }

    function extractLegalNameFromRecord(rec){
      try {
        if (!rec) return "";
        var direct = s(rec.legal_name || rec.u_legal_name || rec.u_legalname || "");
        if (direct) return direct;
        for (var k in rec) {
          if (!Object.prototype.hasOwnProperty.call(rec, k)) continue;
          var nk = s(k).toLowerCase();
          if (nk.indexOf('legal') >= 0 && nk.indexOf('name') >= 0) {
            var v = s(rec[k]);
            if (v) return v;
          }
        }
      } catch(e) {}
      return "";
    }

    // --- Main record fetch ---
    var q1 = 'number=' + '$Ticket';
    var p1 = '/$table.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_query=' + encodeURIComponent(q1);
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

    // --- Ticket-specific enrichments ---
    var openTaskCount = 0;
    var openTasks = [];
    var acts = "";
    var legalName = "";
    var taskEvidenceLength = 0;
    var piSource = "";
    if ('$table' === 'sc_req_item') {
      openTasks = getOpenCatalogTasks(s(r1.sys_id || ""), '$Ticket');
      openTaskCount = openTasks.length;
      legalName = extractLegalNameFromRecord(r1);
      var isNewEpUser = isNewEpUserName(userName) || isNewEpUserName(userDisplay);

      // If activities contain a PI machine id, prefer that value for CI output.
      if ($enableActivitySearchJs) {
        acts = getRitmActivitiesText(s(r1.sys_id || ""), '$Ticket');
        var piFromActivity = extractMachineFromActivityText(acts);
        if (piFromActivity) piSource = "ritm_activity";

        if (!piFromActivity) {
          var prep = getPrepareDevicePiFromTasks(s(r1.sys_id || ""), '$Ticket');
          if (prep && prep.pi) {
            piFromActivity = s(prep.pi);
            taskEvidenceLength = Math.max(taskEvidenceLength, parseInt(prep.text_len || 0, 10) || 0);
            piSource = prep.task_number ? ('prepare_device_task_backend:' + prep.task_number) : 'prepare_device_task_backend';
          }
        }

        if (!piFromActivity) {
          var taskActs = getRitmTasksEvidenceText(s(r1.sys_id || ""), '$Ticket');
          if (taskActs) {
            taskEvidenceLength = s(taskActs).length;
            acts = acts ? (acts + ' ' + taskActs) : taskActs;
            piFromActivity = extractMachineFromActivityText(acts);
            if (piFromActivity) piSource = "sctask_evidence";
          }
        }
        if (piFromActivity) {
          ciVal = piFromActivity;
        }
        if (!legalName) {
          legalName = extractLegalNameFromText(acts);
        }
      }
    }
    else if ('$table' === 'incident') {
      if ($enableActivitySearchJs) {
        acts = getIncidentActivitiesText(s(r1.sys_id || ""), '$Ticket');
        var piFromIncActivity = extractMachineFromActivityText(acts);
        if (piFromIncActivity) {
          ciVal = piFromIncActivity;
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
      legal_name:legalName,
      task_evidence_length:taskEvidenceLength,
      pi_source:piSource,
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
  $o = Parse-WV2Json (ExecJS $wv $js $ExtractJsTimeoutMs)
  if ($o) { return $o }

  # Fallback: minimal extractor to avoid total empty results when complex JS fails.
  $jsMin = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? "" : (""+x).trim(); }
    function pickRec(obj){
      return (obj && obj.records && obj.records[0]) ? obj.records[0] :
             (obj && obj.result && obj.result[0]) ? obj.result[0] :
             (Array.isArray(obj) && obj[0]) ? obj[0] : null;
    }
    var q = 'number=' + '$Ticket';
    var p = '/$table.do?JSONv2&sysparm_limit=1&sysparm_display_value=true&sysparm_query=' + encodeURIComponent(q);
    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);
    if (!(x.status>=200 && x.status<300)) {
      return JSON.stringify({ ok:false, reason:'min_http_'+x.status, ticket:'$Ticket', table:'$table', query:p });
    }
    var obj = {};
    try { obj = JSON.parse(x.responseText || "{}"); } catch(e) { return JSON.stringify({ ok:false, reason:'min_json_parse', ticket:'$Ticket', table:'$table', query:p }); }
    var r = pickRec(obj);
    if (!r) { return JSON.stringify({ ok:false, reason:'not_found', ticket:'$Ticket', table:'$table', query:p }); }
    var user = s(r.requested_for || r.caller_id || "");
    return JSON.stringify({
      ok:true,
      ticket:'$Ticket',
      table:'$table',
      sys_id:s(r.sys_id || ""),
      affected_user:user,
      configuration_item:s(r.configuration_item || r.cmdb_ci || ""),
      status:s(r.state || ""),
      status_value:s(r.state || ""),
      status_label:"",
      open_tasks:0,
      open_task_items:[],
      legal_name:"",
      task_evidence_length:0,
      pi_source:"",
      activity_text:"",
      activity_error:"",
      query:p
    });
  } catch(e) {
    return JSON.stringify({ ok:false, reason:'min_exception', error:''+e, ticket:'$Ticket', table:'$table' });
  }
})();
"@
  $oMin = Parse-WV2Json (ExecJS $wv $jsMin 7000)
  if ($oMin) { return $oMin }

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
    legal_name         = ""
    task_evidence_length = 0
    pi_source         = ""
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
  $swRun = Start-PerfTimer
  # 1) Load WebView2 DLLs from Teams add-in
  $swStage = Start-PerfTimer
  Load-WebView2FromTeams
  [void](Stop-PerfTimer -Timer $swStage -Label "load_webview2")

  # 2) Select Excel file
  $swStage = Start-PerfTimer
  $ExcelPath = Pick-ExcelFile -ExcelPath $ExcelPath -DefaultStartDir $DefaultStartDir -DefaultExcelName $DefaultExcelName
  Log "INFO" "Excel selected: $ExcelPath"
  [void](Stop-PerfTimer -Timer $swStage -Label "pick_excel")

  # 3) Interactive SSO login, then keep WebView2 session
  $swStage = Start-PerfTimer
  $session = Connect-ServiceNowSSO -StartUrl $LoginUrl
  $wv = $session.Wv
  if (-not (Ensure-SnowReady -wv $wv -MaxWaitMs 12000)) {
    Log "ERROR" "SNOW session not ready after SSO. Interactive login is required."
    throw "ServiceNow SSO session is required. Please complete login and try again."
  }
  [void](Stop-PerfTimer -Timer $swStage -Label "connect_sso")

  # Dashboard mode is isolated and does not run export logic.
  if ($DashboardMode) {
    Log "INFO" "Dashboard mode enabled. Launching Check-in / Check-out dashboard."
    Show-CheckInOutDashboard -wv $wv -ExcelPath $ExcelPath -SheetName $SheetName
    return
  }

  # 4) Read tickets list from Excel
  $swStage = Start-PerfTimer
  $tickets = Read-TicketsFromExcel -ExcelPath $ExcelPath -TicketHeader $TicketHeader -SheetName $SheetName -TicketColumn $TicketColumn
  Log "INFO" "Tickets found: $($tickets.Count)"
  [void](Stop-PerfTimer -Timer $swStage -Label "read_tickets_excel")

  if ($tickets.Count -eq 0) {
    throw "No valid tickets found in Excel (INC/RITM/SCTASK + 6-8 digits)."
  }

  switch ($ProcessingScope) {
    "RitmOnly" {
      $tickets = @($tickets | Where-Object { $_ -like "RITM*" })
      Log "INFO" "Processing scope: RITM only (forced). Count=$($tickets.Count)"
    }
    "IncAndRitm" {
      $tickets = @($tickets | Where-Object { ($_ -like "INC*") -or ($_ -like "RITM*") })
      Log "INFO" "Processing scope: INC + RITM (forced). Count=$($tickets.Count)"
    }
    "All" {
      $tickets = @($tickets | Where-Object { ($_ -like "INC*") -or ($_ -like "RITM*") -or ($_ -like "SCTASK*") })
      Log "INFO" "Processing scope: ALL (INC + RITM + SCTASK). Count=$($tickets.Count)"
    }
    default {
      # Auto mode:
      # - If there is at least one INC, process INC + RITM.
      # - If there is no INC, process only RITM.
      $hasInc = @($tickets | Where-Object { $_ -like "INC*" }).Count -gt 0
      if ($hasInc) {
        $tickets = @($tickets | Where-Object { ($_ -like "INC*") -or ($_ -like "RITM*") })
        Log "INFO" "Processing scope: INC + RITM (INC detected). Count=$($tickets.Count)"
      }
      else {
        $tickets = @($tickets | Where-Object { $_ -like "RITM*" })
        Log "INFO" "Processing scope: RITM only (no INC detected). Count=$($tickets.Count)"
      }
    }
  }

  if ($MaxTickets -gt 0 -and $tickets.Count -gt $MaxTickets) {
    $tickets = @($tickets | Select-Object -First $MaxTickets)
    Log "INFO" "Speed mode: limiting tickets to first $MaxTickets."
  }

  if ($tickets.Count -eq 0) {
    throw "No tickets to process after scope filtering."
  }

  # 5) For each ticket: extract + export JSON
  $results = New-Object System.Collections.Generic.List[object]
  $i = 0
  $userDisplayCache = @{}
  $baseEnableActivitySearch = $EnableActivityStreamSearch
  $baseEnableUiFallbackActivitySearch = $EnableUiFallbackActivitySearch
  $ticketLoopMsTotal = [int64]0
  $ticketLoopMaxMs = [int64]0
  $ticketLoopMaxId = ""
  $ticketsSkippedIncremental = 0
  $ticketsProcessedFull = 0

  $swStage = Start-PerfTimer
  $ritmScanDb = Load-RitmScanDatabase -Path $RitmScanDbPath
  Log "INFO" "Incremental scan DB loaded entries=$($ritmScanDb.Count) last_successful_scan_utc='$script:RitmScanDbLastSuccessfulUtc'"
  [void](Stop-PerfTimer -Timer $swStage -Label "load_incremental_db")

  $stateLabelCache = @{}
  if ($TurboMode) {
    $swStage = Start-PerfTimer
    $stateLabelCache = Get-StateLabelCacheFromSnow -wv $wv
    Log "INFO" "Turbo: state label cache loaded tables=$($stateLabelCache.Keys.Count)"
    [void](Stop-PerfTimer -Timer $swStage -Label "load_state_cache")
  }

  if ($TurboMode -and -not [string]::IsNullOrWhiteSpace($script:RitmScanDbLastSuccessfulUtc)) {
    $swStage = Start-PerfTimer
    $changed = @{}
    $ritmChangedDirect = @(Get-RitmNumbersUpdatedSince -wv $wv -SinceUtc $script:RitmScanDbLastSuccessfulUtc)
    foreach ($x in $ritmChangedDirect) { $changed[(("" + $x).Trim().ToUpperInvariant())] = $true }
    $ritmChangedFromTasks = @(Get-RitmNumbersFromUpdatedTasksSince -wv $wv -SinceUtc $script:RitmScanDbLastSuccessfulUtc)
    foreach ($x in $ritmChangedFromTasks) { $changed[(("" + $x).Trim().ToUpperInvariant())] = $true }

    $filtered = New-Object System.Collections.Generic.List[string]
    $prefilterSkipped = 0
    foreach ($tk in $tickets) {
      if ($tk -notlike "RITM*") {
        $filtered.Add($tk) | Out-Null
        continue
      }
      if ($changed.ContainsKey($tk)) {
        $filtered.Add($tk) | Out-Null
        continue
      }
      if (-not $ritmScanDb.ContainsKey($tk)) {
        $filtered.Add($tk) | Out-Null
        continue
      }
      $prefilterSkipped++
      Log "INFO" "Turbo prefilter skip ritm='$tk' reason='not_updated_since_last_successful_scan'"
    }
    $tickets = @($filtered.ToArray())
    Log "INFO" "Turbo prefilter applied since='$script:RitmScanDbLastSuccessfulUtc' direct_changes=$($ritmChangedDirect.Count) task_changes=$($ritmChangedFromTasks.Count) skipped=$prefilterSkipped remaining=$($tickets.Count)"
    [void](Stop-PerfTimer -Timer $swStage -Label "turbo_prefilter")
  }

  foreach ($t in $tickets) {
    $swTicket = Start-PerfTimer
    $i++
    Log "INFO" "[$i/$($tickets.Count)] Open + extract: $t"

    $isIncTicket = ($t -like "INC*")
    $skipActivityParseForTicket = ($SmartMode -and $isIncTicket)
    if ($SmartMode -and $isIncTicket) {
      $EnableActivityStreamSearch = $false
      $EnableUiFallbackActivitySearch = $false
    }
    else {
      $EnableActivityStreamSearch = $baseEnableActivitySearch
      $EnableUiFallbackActivitySearch = $baseEnableUiFallbackActivitySearch
    }

    $incrementalSnapshot = $null
    $incrementalDecision = $null
    if ($t -like "RITM*") {
      $previousScan = if ($ritmScanDb.ContainsKey($t)) { $ritmScanDb[$t] } else { $null }
      $incrementalSnapshot = Get-RitmIncrementalSnapshot -wv $wv -RitmNumber $t
      $incrementalDecision = Get-RitmIncrementalDecision -Ritm $t -Snapshot $incrementalSnapshot -Previous $previousScan
      $decisionAction = if ($incrementalDecision.process -eq $true) { "PROCESS" } else { "SKIP" }
      $decisionWhy = if ($incrementalDecision.process -eq $true) { ("" + $incrementalDecision.process_reason).Trim() } else { ("" + $incrementalDecision.skip_reason).Trim() }
      Log "INFO" "RITM decision ritm='$t' action='$decisionAction' reason='$decisionWhy' ritm_updated_on='$($incrementalSnapshot.ritm_updated_on)' tasks_updated_max='$($incrementalSnapshot.tasks_updated_on_max)' open_tasks=$($incrementalSnapshot.open_task_count)"

      if ($incrementalDecision.process -ne $true) {
        $statusFromSnapshot = Get-RitmStatusFromSnapshot -Snapshot $incrementalSnapshot
        $r = [pscustomobject]@{
          ok                   = $true
          ticket               = $t
          table                = "sc_req_item"
          sys_id               = ""
          affected_user        = ""
          configuration_item   = ""
          status               = $statusFromSnapshot
          status_value         = ""
          status_label         = ""
          open_tasks           = [int]$incrementalSnapshot.open_task_count
          open_task_items      = @()
          legal_name           = ""
          task_evidence_length = 0
          pi_source            = ""
          activity_text        = ""
          activity_error       = ""
          reason               = ("" + $incrementalDecision.skip_reason).Trim()
          query                = "incremental_skip"
        }

        Update-RitmScanDbEntry -Db $ritmScanDb -Ritm $t -Snapshot $incrementalSnapshot -Status $statusFromSnapshot -SkipReason ("" + $incrementalDecision.skip_reason)

        if ($WritePerTicketJson) {
          $perPath = Join-Path $OutDir ("ticket_" + $t + ".json")
          $jsonPer = ($r | ConvertTo-Json -Depth 6) -replace '\\u0027', "'"
          Set-Content -Path $perPath -Value $jsonPer -Encoding UTF8
        }
        $ticketsSkippedIncremental++
        $ticketMs = [int64](Stop-PerfTimer -Timer $swTicket -Label "ticket_$t")
        $ticketLoopMsTotal += $ticketMs
        if ($ticketMs -gt $ticketLoopMaxMs) { $ticketLoopMaxMs = $ticketMs; $ticketLoopMaxId = $t }
        $results.Add($r) | Out-Null
        continue
      }
    }

    # Extract fields via JSONv2 in authenticated session
    $r = Extract-Ticket_JSONv2 -wv $wv -Ticket $t -StateLabelCache $stateLabelCache
    if ($r.ok -ne $true) {
      for ($attempt = 2; $attempt -le $ExtractRetryCount; $attempt++) {
        $reasonTry = if ($r.PSObject.Properties["reason"]) { "" + $r.reason } else { "" }
        Log "INFO" "$t retry $attempt/$ExtractRetryCount after failure reason='$reasonTry'"
        Start-Sleep -Milliseconds $ExtractRetryDelayMs
        $r = Extract-Ticket_JSONv2 -wv $wv -Ticket $t -StateLabelCache $stateLabelCache
        if ($r.ok -eq $true) { break }
      }
    }

    if ($r.ok -eq $true) {
      $uNow = if ($r.PSObject.Properties["affected_user"]) { ("" + $r.affected_user).Trim() } else { "" }
      if ($uNow -match '^[0-9a-fA-F]{32}$') {
        $uResolved = ""
        if ($userDisplayCache.ContainsKey($uNow)) {
          $uResolved = "" + $userDisplayCache[$uNow]
        }
        else {
          $uResolved = Resolve-UserDisplayNameFromSysId -wv $wv -UserSysId $uNow
          $userDisplayCache[$uNow] = "" + $uResolved
        }
        if (-not [string]::IsNullOrWhiteSpace($uResolved)) {
          $r | Add-Member -NotePropertyName affected_user -NotePropertyValue $uResolved -Force
          Log "INFO" "$t user resolved from sys_id => '$uResolved'"
        }
      }
    }

    if (($r.ok -eq $true) -and ($t -like "RITM*")) {
      try {
        $openCountCurrent = 0
        if ($r.PSObject.Properties["open_tasks"]) {
          try { $openCountCurrent = [int]$r.open_tasks } catch { $openCountCurrent = 0 }
        }
        if ($r.PSObject.Properties["open_task_items"] -and $r.open_task_items) {
          try {
            $itemsNow = @($r.open_task_items).Count
            if ($itemsNow -gt $openCountCurrent) { $openCountCurrent = $itemsNow }
          } catch {}
        }

        if ($openCountCurrent -eq 0) {
          $openFallback = @(Get-OpenSCTasksForRitmFallback -wv $wv -RitmNumber $t)
          if ($openFallback.Count -gt 0) {
            $r | Add-Member -NotePropertyName open_task_items -NotePropertyValue @($openFallback) -Force
            $r | Add-Member -NotePropertyName open_tasks -NotePropertyValue ([int]$openFallback.Count) -Force
            $r | Add-Member -NotePropertyName status -NotePropertyValue ("Open:" + $openFallback.Count) -Force
            $nums = @($openFallback | ForEach-Object { if ($_.PSObject.Properties["number"]) { ("" + $_.number).Trim() } else { "" } } | Where-Object { $_ })
            Log "INFO" "$t open SCTASK fallback recovered count=$($openFallback.Count) tasks='$($nums -join ", ")'"
          }
        }
      }
      catch {
        $errMsg = $_.Exception.Message
        $errPos = $_.InvocationInfo.PositionMessage
        Log "ERROR" "$t open SCTASK fallback failed: $errMsg | $errPos"
      }
    }

    if (($r.ok -eq $true) -and (-not $skipActivityParseForTicket) -and (($t -like "RITM*") -or ($t -like "INC*"))) {
      try {
        $activityText = if ($r.PSObject.Properties["activity_text"]) { "" + $r.activity_text } else { "" }
        $uiActivityText = ""
        $activityError = if ($r.PSObject.Properties["activity_error"]) { "" + $r.activity_error } else { "" }
        $legalName = if ($r.PSObject.Properties["legal_name"]) { ("" + $r.legal_name).Trim() } else { "" }
        if ($VerboseTicketLogging) {
          Log "INFO" "$t activity backend text length: $($activityText.Length)"
        }
        if ($activityError) {
          Log "ERROR" "$t activity retrieval issue: $activityError"
        }
        $detectedPi = Get-DetectedPiFromActivityText -ActivityText $activityText
        $currentUserSnapshot = if ($r.PSObject.Properties["affected_user"]) { ("" + $r.affected_user).Trim() } else { "" }
        $isNewEpUserContext = ($currentUserSnapshot -match '(?i)^new\b.*\buser$')
        $needsLegalNameFallback = (
          ($t -like "RITM*") -and
          (-not $legalName) -and
          (
            [string]::IsNullOrWhiteSpace($currentUserSnapshot) -or
            ($currentUserSnapshot -match '^[0-9a-fA-F]{32}$') -or
            ($currentUserSnapshot -match '(?i)^new\b.*\buser$')
          )
        )
        if (
          (
            ([string]::IsNullOrWhiteSpace($detectedPi) -and $EnableActivityStreamSearch) -or
            $needsLegalNameFallback
          ) -and
          $EnableUiFallbackActivitySearch -and
          (
            ($activityText.Length -lt $UiFallbackMinBackendChars) -or
            $needsLegalNameFallback -or
            ($t -like "RITM*")
          )
        ) {
          $recordSysId = if ($r.PSObject.Properties["sys_id"]) { "" + $r.sys_id } else { "" }
          $tableName = if ($r.PSObject.Properties["table"]) { "" + $r.table } else { Ticket-ToTable $t }
          $uiActivityText = Get-RecordActivityTextFromUiPage -wv $wv -RecordSysId $recordSysId -Table $tableName
          if ($VerboseTicketLogging) {
            Log "INFO" "$t activity UI text length: $($uiActivityText.Length)"
          }
          if (-not [string]::IsNullOrWhiteSpace($uiActivityText)) {
            if ($VerboseTicketLogging) {
              Log "INFO" "$t activity UI fallback text collected (len=$($uiActivityText.Length))"
            }
            $detectedPi = Get-DetectedPiFromActivityText -ActivityText $uiActivityText
            if (-not $legalName) { $legalName = Get-LegalNameFromText -Text $uiActivityText }
          }
        }
        if (-not $legalName) { $legalName = Get-LegalNameFromText -Text $activityText }
        if (-not $legalName -and ($t -like "RITM*")) {
          $recordSysId = if ($r.PSObject.Properties["sys_id"]) { "" + $r.sys_id } else { "" }
          $legalFromForm = Get-LegalNameFromUiForm -wv $wv -RecordSysId $recordSysId -Table "sc_req_item"
          if (-not [string]::IsNullOrWhiteSpace($legalFromForm)) {
            $legalName = $legalFromForm
            Log "INFO" "$t Legal name extracted from form => '$legalName'"
          }
        }
        if ([string]::IsNullOrWhiteSpace($detectedPi) -and ($t -like "RITM*") -and ($isNewEpUserContext -or ($t -eq $DebugActivityTicket))) {
          $taskUiText = Get-RitmTaskListTextFromUiPage -wv $wv -RitmNumber $t
          $taskUiLen = if ($taskUiText) { $taskUiText.Length } else { 0 }
          if ($taskUiLen -gt 0) {
            $piFromTaskUi = Get-DetectedPiFromActivityText -ActivityText $taskUiText
            if (-not [string]::IsNullOrWhiteSpace($piFromTaskUi)) {
              $detectedPi = $piFromTaskUi
              $r | Add-Member -NotePropertyName pi_source -NotePropertyValue "sctask_ui_list" -Force
              Log "INFO" "$t PI found from SCTASK UI list => '$detectedPi' (len=$taskUiLen)"
            }
            else {
              Log "INFO" "$t SCTASK UI list scanned, PI not found (len=$taskUiLen)"
              $taskNums = Get-TaskNumbersFromText -Text $taskUiText
              if ($taskNums.Count -eq 0) {
                $taskNums = @(Get-SCTaskNumbersFromBackendByRitm -wv $wv -RitmNumber $t)
                if ($taskNums.Count -gt 0) {
                  Log "INFO" "$t task numbers loaded from backend fallback: $($taskNums.Count)"
                }
              }
              if ($taskNums.Count -gt 0) {
                $maxTaskDeepScan = if ($isNewEpUserContext) { 12 } else { 4 }
                $scanCount = [Math]::Min($taskNums.Count, $maxTaskDeepScan)
                Log "INFO" "$t deep SCTASK scan starting. tasks=$($taskNums.Count) limit=$scanCount new_ep_user=$isNewEpUserContext"
                $matchedPrepareTask = $false
                for ($ti = 0; $ti -lt $scanCount; $ti++) {
                  $tn = $taskNums[$ti]
                  $taskActivityText = Get-SCTaskActivityTextFromUiPage -wv $wv -TaskNumber $tn
                  if (-not [string]::IsNullOrWhiteSpace($taskActivityText)) {
                    $piFromTaskActivity = Get-DetectedPiFromActivityText -ActivityText $taskActivityText
                    if ($isNewEpUserContext -and -not [string]::IsNullOrWhiteSpace($piFromTaskActivity)) {
                      $piDecision = Resolve-ConfidentPiFromSource -PiListText $piFromTaskActivity -SourceText $taskActivityText
                      if ($piDecision -and $piDecision.selected) {
                        if ($piDecision.ambiguous -eq $true) {
                          Log "INFO" "$t multiple PI detected in SCTASK activity $tn; ambiguous decision='$($piDecision.selected)' reason='$($piDecision.reason)'"
                        }
                        elseif ($piDecision.selected -ne $piFromTaskActivity) {
                          Log "INFO" "$t multiple PI detected in SCTASK activity $tn; selected='$($piDecision.selected)' reason='$($piDecision.reason)'"
                        }
                        $piFromTaskActivity = "" + $piDecision.selected
                      }
                    }
                    if (-not [string]::IsNullOrWhiteSpace($piFromTaskActivity)) {
                      $detectedPi = $piFromTaskActivity
                      $src = if ($isNewEpUserContext) { "sctask_activity_record:" + $tn } else { "sctask_activity_record:" + $tn }
                      $r | Add-Member -NotePropertyName pi_source -NotePropertyValue $src -Force
                      Log "INFO" "$t PI found from SCTASK activity $tn => '$detectedPi' source='$src'"
                      break
                    }
                  }

                  $taskRecordText = Get-SCTaskRecordTextFromUiPage -wv $wv -TaskNumber $tn
                  if ([string]::IsNullOrWhiteSpace($taskRecordText)) { continue }
                  if ($isNewEpUserContext) {
                    if ($taskRecordText -match '(?is)prepare[\s\W_]*device[\s\W_]*for[\s\W_]*new[\s\W_]*user') {
                      $matchedPrepareTask = $true
                    }
                    else {
                      continue
                    }
                  }
                  $piFromTaskRecord = Get-DetectedPiFromActivityText -ActivityText $taskRecordText
                  if ([string]::IsNullOrWhiteSpace($piFromTaskRecord) -and $isNewEpUserContext) {
                    $piFromTaskRecord = Get-DetectedMachineHintFromText -Text $taskRecordText
                    if (-not [string]::IsNullOrWhiteSpace($piFromTaskRecord)) {
                      Log "INFO" "$t machine hint found in prepare-device task $tn => '$piFromTaskRecord'"
                    }
                  }
                  if (-not [string]::IsNullOrWhiteSpace($piFromTaskRecord)) {
                    $detectedPi = $piFromTaskRecord
                    $src = if ($isNewEpUserContext) { "sctask_prepare_device_record:" + $tn } else { "sctask_ui_record:" + $tn }
                    $r | Add-Member -NotePropertyName pi_source -NotePropertyValue $src -Force
                    Log "INFO" "$t PI found from SCTASK record $tn => '$detectedPi' source='$src'"
                    break
                  }
                }
                if ($isNewEpUserContext -and (-not $matchedPrepareTask)) {
                  Log "INFO" "$t prepare-device task not found in scanned SCTASK records"
                }
              }
            }
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
        $piSource = if ($r.PSObject.Properties["pi_source"]) { ("" + $r.pi_source).Trim() } else { "" }
        $taskEvidenceLen = 0
        if ($r.PSObject.Properties["task_evidence_length"]) {
          try { $taskEvidenceLen = [int]$r.task_evidence_length } catch { $taskEvidenceLen = 0 }
        }
        if ($taskEvidenceLen -gt 0 -or $piSource) {
          Log "INFO" "$t PI scan source='$piSource' task_evidence_len=$taskEvidenceLen"
        }
        $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue $detectedPi -Force
        if ($legalName) {
          $r | Add-Member -NotePropertyName legal_name -NotePropertyValue $legalName -Force
          $currentUser = if ($r.PSObject.Properties["affected_user"]) { ("" + $r.affected_user).Trim() } else { "" }
          if (Is-InvalidUserDisplay $currentUser) {
            $r | Add-Member -NotePropertyName affected_user -NotePropertyValue $legalName -Force
            Log "INFO" "$t Name updated from Legal name => '$legalName'"
          }
        }
      }
      catch {
        Log "ERROR" "$t activity parsing failed: $($_.Exception.Message)"
        $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue $(if ($WriteNotFoundText) { "Not found" } else { "" }) -Force
      }

      $openTaskItems = @()
      if ($r.PSObject.Properties["open_task_items"] -and $r.open_task_items) {
        $openTaskItems = @($r.open_task_items)
      }
      if ($VerboseTicketLogging) {
        Log "INFO" "$t open SCTASK count: $($openTaskItems.Count)"
        foreach ($ot in $openTaskItems) {
          $taskNo = if ($ot.PSObject.Properties["number"]) { "" + $ot.number } else { "" }
          $taskSys = if ($ot.PSObject.Properties["sys_id"]) { "" + $ot.sys_id } else { "" }
          $taskUrl = Build-SCTaskBestUrl -SysId $taskSys -TaskNumber $taskNo
          Log "INFO" "$t open SCTASK number='$taskNo' sys_id='$taskSys' url='$taskUrl'"
        }
      }
    }
    elseif ($r.ok -eq $true) {
      $r | Add-Member -NotePropertyName detected_pi_machine -NotePropertyValue "" -Force
    }

    if ($t -like "RITM*") {
      $statusForDb = if ($r -and $r.PSObject.Properties["status"]) { ("" + $r.status).Trim() } else { "" }
      $skipReasonForDb = ""
      if ($r -and $r.PSObject.Properties["reason"]) {
        $reasonNow = ("" + $r.reason).Trim()
        if ($reasonNow -match '^unchanged') { $skipReasonForDb = $reasonNow }
      }
      Update-RitmScanDbEntry -Db $ritmScanDb -Ritm $t -Snapshot $incrementalSnapshot -Status $statusForDb -SkipReason $skipReasonForDb
    }

    # Log quick summary line
    $status = if ($r.ok -eq $true) { "OK" } else { "FAIL" }
    $reason = if ($r -and $r.PSObject.Properties["reason"]) { "" + $r.reason } else { "" }
    $userOut = if ($r -and $r.PSObject.Properties["affected_user"]) { "" + $r.affected_user } else { "" }
    $ciOut = if ($r -and $r.PSObject.Properties["configuration_item"]) { "" + $r.configuration_item } else { "" }
    $urlOut = if ($r -and $r.PSObject.Properties["query"]) { "" + $r.query } elseif ($r -and $r.PSObject.Properties["href"]) { "" + $r.href } else { "" }
    Log "INFO" "$t => $status reason=$reason user='$userOut' ci='$ciOut' url='$urlOut'"

    # Save per-ticket JSON file
    if ($WritePerTicketJson) {
      $perPath = Join-Path $OutDir ("ticket_" + $t + ".json")
      $jsonPer = ($r | ConvertTo-Json -Depth 6) -replace '\\u0027', "'"
      Set-Content -Path $perPath -Value $jsonPer -Encoding UTF8
    }

    # Add to in-memory list for combined export + write-back map
    $ticketsProcessedFull++
    $ticketMs = [int64](Stop-PerfTimer -Timer $swTicket -Label "ticket_$t")
    $ticketLoopMsTotal += $ticketMs
    if ($ticketMs -gt $ticketLoopMaxMs) { $ticketLoopMaxMs = $ticketMs; $ticketLoopMaxId = $t }
    $results.Add($r) | Out-Null
  }
  $EnableActivityStreamSearch = $baseEnableActivitySearch
  $EnableUiFallbackActivitySearch = $baseEnableUiFallbackActivitySearch

  # 6) Save combined JSON
  $swStage = Start-PerfTimer
  $jsonAll = ($results | ConvertTo-Json -Depth 6) -replace '\\u0027', "'"
  Set-Content -Path $AllJson -Value $jsonAll -Encoding UTF8
  Log "INFO" "ALL JSON: $AllJson"
  Log "INFO" "DONE. Logs: $LogPath"
  [void](Stop-PerfTimer -Timer $swStage -Label "save_combined_json")

  # 7) Optional write-back to Excel
  if ($WriteBackExcel) {
    $swStage = Start-PerfTimer
    # Build a map ticket -> result object
    $map = @{}
    foreach ($r in $results) { $map[$r.ticket] = $r }

    Write-BackToExcel -ExcelPath $ExcelPath -SheetName $SheetName -TicketHeader $TicketHeader -TicketColumn $TicketColumn `
      -NameHeader $NameHeader -PhoneHeader $PhoneHeader -ActionHeader $ActionHeader -SCTasksHeader $SCTasksHeader -ResultMap $map
    [void](Stop-PerfTimer -Timer $swStage -Label "writeback_excel")
  }

  $script:RitmScanRunSuccessful = $true
  $script:RitmScanDbLastSuccessfulUtc = Convert-ToUtcStampText -UtcDate (Get-Date).ToUniversalTime()
  $swStage = Start-PerfTimer
  [void](Save-RitmScanDatabase -Path $RitmScanDbPath -Map $ritmScanDb -LastSuccessfulScanUtc $script:RitmScanDbLastSuccessfulUtc)
  [void](Stop-PerfTimer -Timer $swStage -Label "save_incremental_db")

  $avgTicketMs = if ($tickets.Count -gt 0) { [int64]($ticketLoopMsTotal / $tickets.Count) } else { 0 }
  Log "INFO" "PERF ticket_summary total=$($tickets.Count) full=$ticketsProcessedFull skipped_incremental=$ticketsSkippedIncremental avg_ms=$avgTicketMs max_ms=$ticketLoopMaxMs max_ticket='$ticketLoopMaxId'"
  [void](Stop-PerfTimer -Timer $swRun -Label "run_total")

  # 8) Final success popup
  if (-not $NoPopups) {
    [System.Windows.Forms.MessageBox]::Show(
      "Export complete.`r`nFolder: $OutDir`r`nAll JSON: $AllJson",
      "SNOW Export",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
  }
}
catch {
  # Any exception: log + show popup
  Log "ERROR" $_.Exception.Message
  try {
    if ($ritmScanDb) { [void](Save-RitmScanDatabase -Path $RitmScanDbPath -Map $ritmScanDb -LastSuccessfulScanUtc $script:RitmScanDbLastSuccessfulUtc) }
  } catch {}

  if (-not $NoPopups) {
    [System.Windows.Forms.MessageBox]::Show(
      $_.Exception.Message,
      "SNOW Export ERROR",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
  }
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
