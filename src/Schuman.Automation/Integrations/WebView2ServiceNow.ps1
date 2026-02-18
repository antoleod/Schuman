Set-StrictMode -Version Latest

function Initialize-WebView2Runtime {
  $arch = if ($env:PROCESSOR_ARCHITECTURE -match '64') { 'x64' } else { 'x86' }
  $base = Join-Path $env:LOCALAPPDATA 'Microsoft\TeamsMeetingAdd-in'
  if (-not (Test-Path -LiteralPath $base)) {
    throw "Teams Meeting Add-in not found: $base"
  }

  $dir = Get-ChildItem -LiteralPath $base -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $dir) {
    throw 'No Teams Meeting Add-in version folder found.'
  }

  Add-Type -AssemblyName System.Windows.Forms, System.Drawing
  Add-Type -Path (Join-Path $dir.FullName "$arch\Microsoft.Web.WebView2.WinForms.dll")
  Add-Type -Path (Join-Path $dir.FullName "$arch\Microsoft.Web.WebView2.Core.dll")
}

function New-ServiceNowSession {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [int]$TimeoutSeconds = 240
  )

  Initialize-WebView2Runtime

  $profileRoot = Join-Path $Config.ServiceNow.WebViewProfileRoot $env:USERNAME
  New-Item -ItemType Directory -Force -Path $profileRoot | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'ServiceNow Login - Complete SSO and wait for green status'
  $form.Size = New-Object System.Drawing.Size(1100, 760)
  $form.StartPosition = 'CenterScreen'
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Dock = 'Top'
  $label.Height = 64
  $label.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
  $label.Text = 'Loading ServiceNow login page...'
  $form.Controls.Add($label)

  $panel = New-Object System.Windows.Forms.Panel
  $panel.Dock = 'Fill'
  $form.Controls.Add($panel)

  $wv = New-Object Microsoft.Web.WebView2.WinForms.WebView2
  $props = New-Object Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties
  $props.UserDataFolder = $profileRoot
  $wv.CreationProperties = $props
  $wv.Dock = 'Fill'
  $panel.Controls.Add($wv)

  $form.Show() | Out-Null
  $task = $wv.EnsureCoreWebView2Async()
  while (-not $task.IsCompleted) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 40
  }
  if ($task.IsFaulted) { throw $task.Exception.InnerException }

  $wv.Source = $Config.ServiceNow.LoginUrl

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $authenticated = $false
  while ($form.Visible -and -not $authenticated -and $sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
    $state = Invoke-WebViewScriptJson -WebView $wv -Script @"
(function(){
  try {
    var href = location.href || '';
    var host = '';
    try { host = (new URL(href)).host || ''; } catch(e) {}
    var isSnow = /service-now\.com/i.test(host);
    var hasLogin = !!document.querySelector('form#login,input#user_name,input#username,input[type=password]');
    var hasNow = (typeof window.NOW !== 'undefined') || (typeof window.g_user !== 'undefined');
    var domReady = !!document.querySelector('sn-polaris-layout, now-global-nav, sn-appshell-root');
    var logged = isSnow && !hasLogin && (hasNow || domReady);
    return JSON.stringify({ href:href, title: document.title || '', logged: logged, isSnow:isSnow });
  } catch(e) {
    return JSON.stringify({ logged:false, error:''+e });
  }
})();
"@ -TimeoutMs 4000

    if ($state) {
      $label.Text = "URL: $($state.href)`r`nTITLE: $($state.title)"
      if ($state.logged -eq $true) {
        $label.ForeColor = [System.Drawing.Color]::Green
        $authenticated = $true
      } elseif ($state.isSnow -eq $true) {
        $label.ForeColor = [System.Drawing.Color]::DarkOrange
      } else {
        $label.ForeColor = [System.Drawing.Color]::Red
      }
    }

    Start-Sleep -Milliseconds 200
  }

  if (-not $authenticated) {
    try { $form.Close() } catch {}
    throw 'ServiceNow SSO not confirmed before timeout/close.'
  }

  Write-RunLog -RunContext $RunContext -Level INFO -Message 'ServiceNow SSO confirmed.'
  $null = $form.Hide()

  return @{
    Form = $form
    WebView = $wv
    BaseUrl = $Config.ServiceNow.BaseUrl
    StateCache = @{}
    UserCache = @{}
    CiCache = @{}
    JsonTimeoutMs = [int]$Config.ServiceNow.JsonTimeoutMs
    RetryCount = [int]$Config.ServiceNow.QueryRetryCount
    RetryDelayMs = [int]$Config.ServiceNow.QueryRetryDelayMs
  }
}

function Close-ServiceNowSession {
  param($Session)
  if (-not $Session) { return }

  try { if ($Session.Form) { $Session.Form.Close() } } catch {}
  try { if ($Session.Form) { $Session.Form.Dispose() } } catch {}
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Invoke-WebViewScriptRaw {
  param(
    [Parameter(Mandatory = $true)]$WebView,
    [Parameter(Mandatory = $true)][string]$Script,
    [int]$TimeoutMs = 12000
  )

  $exec = $WebView.ExecuteScriptAsync($Script)
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while (-not $exec.IsCompleted -and $sw.ElapsedMilliseconds -lt $TimeoutMs) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 40
  }

  if (-not $exec.IsCompleted -or $exec.IsFaulted) { return $null }
  return $exec.GetAwaiter().GetResult()
}

function ConvertFrom-WebViewResult {
  param([string]$Raw)
  if ([string]::IsNullOrWhiteSpace($Raw) -or $Raw -eq 'null') { return $null }

  try {
    $o = $Raw | ConvertFrom-Json
    if ($o -is [string]) { $o = $o | ConvertFrom-Json }
    return $o
  } catch {
    return $null
  }
}

function Invoke-WebViewScriptJson {
  param(
    [Parameter(Mandatory = $true)]$WebView,
    [Parameter(Mandatory = $true)][string]$Script,
    [int]$TimeoutMs = 12000
  )

  $raw = Invoke-WebViewScriptRaw -WebView $WebView -Script $Script -TimeoutMs $TimeoutMs
  return ConvertFrom-WebViewResult -Raw $raw
}

function Invoke-ServiceNowJsonv2Query {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$Table,
    [Parameter(Mandatory = $true)][string]$Query,
    [string[]]$Fields = @(),
    [int]$Limit = 1
  )

  $tableName = $Table.Trim()
  $queryText = $Query.Trim()
  $fieldText = if ($Fields.Count -gt 0) { $Fields -join ',' } else { '' }

  $js = @"
(function(){
  try {
    var table = '$tableName';
    var q = '$queryText';
    var fields = '$fieldText';
    var limit = '$Limit';
    var p = '/' + table + '.do?JSONv2&sysparm_display_value=true&sysparm_limit=' + encodeURIComponent(limit) + '&sysparm_query=' + encodeURIComponent(q);
    if (fields) p += '&sysparm_fields=' + encodeURIComponent(fields);

    var x = new XMLHttpRequest();
    x.open('GET', p, false);
    x.withCredentials = true;
    x.send(null);

    if (!(x.status >= 200 && x.status < 300)) {
      return JSON.stringify({ ok:false, status:x.status, records:[] });
    }

    var o = JSON.parse(x.responseText || '{}');
    var rows = [];
    if (o && Array.isArray(o.records)) rows = o.records;
    else if (o && Array.isArray(o.result)) rows = o.result;

    return JSON.stringify({ ok:true, status:x.status, records:rows });
  } catch(e) {
    return JSON.stringify({ ok:false, error:''+e, records:[] });
  }
})();
"@

  for ($attempt = 1; $attempt -le $Session.RetryCount; $attempt++) {
    $obj = Invoke-WebViewScriptJson -WebView $Session.WebView -Script $js -TimeoutMs $Session.JsonTimeoutMs
    if ($obj -and $obj.ok -eq $true) {
      return @($obj.records)
    }

    if ($attempt -lt $Session.RetryCount) {
      Start-Sleep -Milliseconds $Session.RetryDelayMs
    }
  }

  return @()
}

function Resolve-ServiceNowStateLabel {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$Table,
    [string]$StateValue,
    [string]$FallbackLabel
  )

  $sv = ("" + $StateValue).Trim()
  if ([string]::IsNullOrWhiteSpace($sv)) { return ("" + $FallbackLabel).Trim() }

  $cacheKey = "{0}:{1}" -f $Table, $sv
  if ($Session.StateCache.ContainsKey($cacheKey)) {
    return $Session.StateCache[$cacheKey]
  }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sys_choice' -Query ("name={0}^element=state^value={1}" -f $Table, $sv) -Fields @('label','value') -Limit 1
  $label = if ($rows.Count -gt 0) { ("" + $rows[0].label).Trim() } else { ("" + $FallbackLabel).Trim() }
  if (-not $label) { $label = $sv }

  $Session.StateCache[$cacheKey] = $label
  return $label
}

function Resolve-ServiceNowUserDisplay {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [string]$UserValue
  )

  $u = ("" + $UserValue).Trim()
  if (-not $u) { return '' }
  if ($u -notmatch '^[0-9a-fA-F]{32}$') { return $u }

  if ($Session.UserCache.ContainsKey($u)) { return $Session.UserCache[$u] }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sys_user' -Query ("sys_id={0}" -f $u) -Fields @('name') -Limit 1
  $name = if ($rows.Count -gt 0) { ("" + $rows[0].name).Trim() } else { '' }
  if (-not $name) { $name = $u }

  $Session.UserCache[$u] = $name
  return $name
}

function Resolve-ServiceNowCiDisplay {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [string]$CiValue
  )

  $v = ("" + $CiValue).Trim()
  if (-not $v) { return '' }
  if ($v -notmatch '^[0-9a-fA-F]{32}$') { return $v }

  if ($Session.CiCache.ContainsKey($v)) { return $Session.CiCache[$v] }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table 'cmdb_ci' -Query ("sys_id={0}" -f $v) -Fields @('name') -Limit 1
  $name = if ($rows.Count -gt 0) { ("" + $rows[0].name).Trim() } else { '' }
  if (-not $name) { $name = $v }

  $Session.CiCache[$v] = $name
  return $name
}

function Get-ServiceNowOpenTasksByRitm {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmSysId
  )

  if ([string]::IsNullOrWhiteSpace($RitmSysId)) { return @() }

  $fields = @('number','sys_id','state','state_value','short_description')
  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_task' -Query ("request_item={0}" -f $RitmSysId) -Fields $fields -Limit 200
  $out = New-Object System.Collections.Generic.List[object]

  foreach ($row in $rows) {
    $state = ("" + $row.state).Trim()
    $stateValue = ("" + $row.state_value).Trim()
    if (-not (Test-ClosedState -StateLabel $state -StateValue $stateValue)) {
      $out.Add([pscustomobject]@{
        number = ("" + $row.number).Trim().ToUpperInvariant()
        sys_id = ("" + $row.sys_id).Trim()
        state_label = $state
        state_value = $stateValue
        short_description = ("" + $row.short_description).Trim()
      }) | Out-Null
    }
  }

  return @($out)
}

function Get-ServiceNowTicket {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$Ticket
  )

  $ticketId = ("" + $Ticket).Trim().ToUpperInvariant()
  $ticketType = Get-TicketType -Ticket $ticketId
  if ($ticketType -eq 'UNKNOWN') {
    return [pscustomobject]@{ ticket = $ticketId; ok = $false; reason = 'unsupported_ticket_type' }
  }

  switch ($ticketType) {
    'RITM' {
      $table = 'sc_req_item'
      $fields = @('number','sys_id','state','state_value','requested_for','configuration_item','short_description','comments','work_notes')
      $userField = 'requested_for'
      $ciField = 'configuration_item'
    }
    'INC' {
      $table = 'incident'
      $fields = @('number','sys_id','state','state_value','caller_id','cmdb_ci','short_description','comments','work_notes')
      $userField = 'caller_id'
      $ciField = 'cmdb_ci'
    }
    'SCTASK' {
      $table = 'sc_task'
      $fields = @('number','sys_id','state','state_value','assigned_to','cmdb_ci','short_description','comments','work_notes')
      $userField = 'assigned_to'
      $ciField = 'cmdb_ci'
    }
  }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table $table -Query ("number={0}" -f $ticketId) -Fields $fields -Limit 1
  if ($rows.Count -eq 0) {
    return [pscustomobject]@{ ticket = $ticketId; ok = $false; reason = 'not_found'; table = $table }
  }

  $r = $rows[0]
  $sysId = ("" + $r.sys_id).Trim()
  $stateValue = ("" + $r.state_value).Trim()
  $stateLabel = Resolve-ServiceNowStateLabel -Session $Session -Table $table -StateValue $stateValue -FallbackLabel ("" + $r.state)
  $user = Resolve-ServiceNowUserDisplay -Session $Session -UserValue ("" + $r.$userField)
  $ci = Resolve-ServiceNowCiDisplay -Session $Session -CiValue ("" + $r.$ciField)

  $activityText = (("" + $r.comments) + ' ' + ("" + $r.work_notes)).Trim()
  $piDetected = Get-DetectedPiFromText -Text $activityText

  $openTasks = @()
  if ($ticketType -eq 'RITM') {
    $openTasks = Get-ServiceNowOpenTasksByRitm -Session $Session -RitmSysId $sysId
  }

  $completion = Get-CompletionStatus -Ticket $ticketId -StateLabel $stateLabel -StateValue $stateValue -OpenTasks $openTasks.Count

  return [pscustomobject]@{
    ticket = $ticketId
    ok = $true
    table = $table
    sys_id = $sysId
    status = $stateLabel
    status_value = $stateValue
    affected_user = $user
    configuration_item = $ci
    short_description = ("" + $r.short_description).Trim()
    detected_pi_machine = $piDetected
    open_tasks = $openTasks.Count
    open_task_items = @($openTasks)
    open_task_numbers = @($openTasks | ForEach-Object { $_.number })
    completion_status = $completion
    record_url = if ($sysId) { "{0}/nav_to.do?uri=%2F{1}.do%3Fsys_id%3D{2}" -f $Session.BaseUrl, $table, $sysId } else { '' }
  }
}

function Set-ServiceNowTaskState {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$TaskSysId,
    [Parameter(Mandatory = $true)][string]$TargetStateLabel,
    [string]$WorkNote = ''
  )

  if ([string]::IsNullOrWhiteSpace($TaskSysId)) { return $false }

  $recordUrl = "{0}/nav_to.do?uri=%2Fsc_task.do%3Fsys_id%3D{1}" -f $Session.BaseUrl, $TaskSysId
  try { $Session.WebView.CoreWebView2.Navigate($recordUrl) } catch { return $false }

  $readySw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($readySw.ElapsedMilliseconds -lt 15000) {
    $isReady = Invoke-WebViewScriptJson -WebView $Session.WebView -Script "document.readyState==='complete'" -TimeoutMs 2500
    if ($isReady -eq $true) { break }
    Start-Sleep -Milliseconds 250
  }

  $targetJson = $TargetStateLabel | ConvertTo-Json -Compress
  $noteJson = $WorkNote | ConvertTo-Json -Compress
  $js = @"
(function(){
  try {
    var target = $targetJson;
    var note = $noteJson;
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function norm(x){ return s(x).toLowerCase().replace(/[\s_-]+/g,' ').trim(); }

    function getDoc(){
      var f = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
      if (f && f.contentDocument) return f.contentDocument;
      return document;
    }

    var doc = getDoc();
    var stateEl = doc.querySelector('select[name="sc_task.state"],select[name="state"],select[id$=".state"]');
    if (!stateEl) return JSON.stringify({ok:false, reason:'state_control_missing'});

    var best = null;
    for (var i=0; i<stateEl.options.length; i++) {
      var txt = norm(stateEl.options[i].text || '');
      if (txt === norm(target) || txt.indexOf(norm(target)) >= 0) { best = stateEl.options[i].value; break; }
    }
    if (best === null) return JSON.stringify({ok:false, reason:'state_option_missing'});

    stateEl.value = best;
    stateEl.dispatchEvent(new Event('change', {bubbles:true}));

    if (note) {
      var noteEl = doc.querySelector('textarea[name="work_notes"],textarea[id*="work_notes"]');
      if (noteEl) {
        noteEl.value = note;
        noteEl.dispatchEvent(new Event('input', {bubbles:true}));
        noteEl.dispatchEvent(new Event('change', {bubbles:true}));
      }
    }

    var saveBtn = doc.querySelector('#sysverb_update,#sysverb_update_and_stay,button[name="sysverb_update"],input[name="sysverb_update"]');
    if (!saveBtn) return JSON.stringify({ok:false, reason:'save_button_missing'});
    saveBtn.click();
    return JSON.stringify({ok:true});
  } catch(e) {
    return JSON.stringify({ok:false, reason:''+e});
  }
})();
"@

  for ($i = 0; $i -lt 3; $i++) {
    $res = Invoke-WebViewScriptJson -WebView $Session.WebView -Script $js -TimeoutMs 12000
    if ($res -and $res.ok -eq $true) {
      Start-Sleep -Milliseconds 800
      return $true
    }
    Start-Sleep -Milliseconds 600
  }

  return $false
}

function Get-ServiceNowTasksForRitm {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmNumber
  )

  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if (-not $ritm) { return @() }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_task' -Query ("request_item.number={0}" -f $ritm) -Fields @('number','sys_id','state','state_value') -Limit 200
  $out = New-Object System.Collections.Generic.List[object]
  foreach ($r in $rows) {
    $out.Add([pscustomobject]@{
      number = ("" + $r.number).Trim().ToUpperInvariant()
      sys_id = ("" + $r.sys_id).Trim()
      state = ("" + $r.state).Trim()
      state_value = ("" + $r.state_value).Trim()
    }) | Out-Null
  }

  return @($out)
}
