Set-StrictMode -Version Latest

function Get-ObjectStringValue {
  param(
    [Parameter(Mandatory = $true)]$Object,
    [Parameter(Mandatory = $true)][string]$PropertyName
  )

  if ($null -eq $Object) { return '' }
  $prop = $Object.PSObject.Properties[$PropertyName]
  if ($null -eq $prop) { return '' }
  return ("" + $prop.Value).Trim()
}
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

function global:New-ServiceNowSession {
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
    Start-Sleep -Milliseconds 20
  }
  if ($task.IsFaulted) { throw $task.Exception.InnerException }

  $wv.Source = $Config.ServiceNow.LoginUrl

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $authenticated = $false
  while ($form.Visible -and -not $authenticated -and $sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
    $state = Invoke-WebViewScriptJson -WebView $wv -Script @"
(function(){
  try {
    function isVisible(el){
      if (!el) return false;
      try {
        var s = window.getComputedStyle(el);
        if (!s) return true;
        if (s.display === 'none' || s.visibility === 'hidden') return false;
        if (parseFloat(s.opacity || '1') <= 0.01) return false;
      } catch(e) {}
      var r = null;
      try { r = el.getBoundingClientRect(); } catch(e) {}
      if (!r) return true;
      return (r.width >= 40 && r.height >= 16);
    }

    function hasLoginControls(doc){
      if (!doc) return false;
      var nodes = doc.querySelectorAll('form#login,input#user_name,input#username,input[type=password],button[type=submit][id*="sign"],button[id*="signin"],input[type=submit][value*="Sign"]');
      if (!nodes || nodes.length === 0) return false;
      for (var i=0; i<nodes.length; i++) {
        if (isVisible(nodes[i])) return true;
      }
      return false;
    }

    function hasBlockingLoginInIframes(root){
      if (!root) return false;
      var iframes = root.querySelectorAll('iframe');
      for (var i=0; i<iframes.length; i++) {
        try {
          if (!isVisible(iframes[i])) continue;
          var fr = iframes[i].getBoundingClientRect();
          var iframeIsLarge = !!fr && fr.width > 320 && fr.height > 220;
          if (!iframeIsLarge) continue;
          var fdoc = iframes[i].contentDocument;
          if (hasLoginControls(fdoc)) return true;
        } catch(e) {}
      }
      return false;
    }

    var href = location.href || '';
    var host = '';
    try { host = (new URL(href)).host || ''; } catch(e) {}
    var isSnow = /service-now\.com/i.test(host);
    var hasTopLogin = hasLoginControls(document);
    var hasIframeLogin = hasBlockingLoginInIframes(document);
    var hasNow = (typeof window.NOW !== 'undefined') || (typeof window.g_user !== 'undefined');
    var hasShell = !!document.querySelector('sn-polaris-layout, now-global-nav, sn-appshell-root,#filter');
    var domReady = hasShell;
    var title = (document.title || '').toLowerCase();
    var titleLooksLogin = /sign\s*in|login|log\s*in/.test(title);
    var userId = '';
    try {
      if (window.g_user) {
        userId = '' + (window.g_user.userID || window.g_user.userId || window.g_user.user_id || '');
      } else if (window.NOW && window.NOW.user) {
        userId = '' + (window.NOW.user.userID || window.NOW.user.userId || window.NOW.user.user_id || '');
      }
    } catch(e) {}
    userId = (userId || '').trim();
    var hasValidUser = /^[0-9a-f]{32}$/i.test(userId);
    var loginBlocking = hasTopLogin || (hasIframeLogin && !hasShell);
    var logged = isSnow && !loginBlocking && (hasNow || hasShell) && domReady && hasValidUser && !titleLooksLogin;
    return JSON.stringify({
      href:href,
      title: document.title || '',
      logged: logged,
      isSnow:isSnow,
      hasLogin:loginBlocking,
      hasValidUser:hasValidUser
    });
  } catch(e) {
    return JSON.stringify({ logged:false, error:''+e });
  }
})();
"@ -TimeoutMs 4000

    if ($state) {
      $label.Text = "URL: $($state.href)`r`nTITLE: $($state.title)`r`nLOGIN_FORM: $($state.hasLogin) | USER_OK: $($state.hasValidUser)"
      if ($state.logged -eq $true) {
        $label.ForeColor = [System.Drawing.Color]::Green
        $authenticated = $true
      } elseif ($state.isSnow -eq $true) {
        $label.ForeColor = [System.Drawing.Color]::DarkOrange
      } else {
        $label.ForeColor = [System.Drawing.Color]::Red
      }
    }

    Start-Sleep -Milliseconds 120
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
  $fieldList = @($Fields)
  $fieldText = if ($fieldList.Count -gt 0) { $fieldList -join ',' } else { '' }

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
  $rowList = @($rows)
  $label = if ($rowList.Count -gt 0) { ("" + $rowList[0].label).Trim() } else { ("" + $FallbackLabel).Trim() }
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
  $rowList = @($rows)
  $name = if ($rowList.Count -gt 0) { ("" + $rowList[0].name).Trim() } else { '' }
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
  $rowList = @($rows)
  $name = if ($rowList.Count -gt 0) { ("" + $rowList[0].name).Trim() } else { '' }
  if (-not $name) { $name = $v }

  $Session.CiCache[$v] = $name
  return $name
}

function Test-InvalidUserDisplay {
  param([string]$Name)

  $n = ("" + $Name).Trim()
  if (-not $n) { return $true }
  if ($n -match '^[0-9a-fA-F]{32}$') { return $true }
  if ($n -match '(?i)\bnew\b.*\bep\b.*\busers?\b') { return $true }
  if ($n -match '(?i)^new\b.*\busers?\b') { return $true }
  if ($n -match '(?i)^unknown$|^n/?a$|^null$') { return $true }
  return $false
}

function Get-LegalNameFromText {
  param([string]$Text)

  $t = ("" + $Text)
  if ([string]::IsNullOrWhiteSpace($t)) { return '' }

  $patterns = @(
    '(?im)\bLEGAL_NAME\s*:\s*([^\r\n]+)',
    "(?im)\bLegal\s*name[\s:\-]*([A-Za-z''\- ]{3,})",
    '(?is)\bLegal\s*name\b[\s\S]{0,200}?\bvalue\s*=\s*["'']([^"'']{3,})["'']',
    '(?im)\bLegal\s*name\b[\r\n\t ]+([A-Z][A-Za-z''\- ]{3,})'
  )

  foreach ($pat in $patterns) {
    $m = [System.Text.RegularExpressions.Regex]::Match($t, $pat)
    if (-not $m.Success) { continue }
    $v = ("" + $m.Groups[1].Value).Trim()
    if (-not $v) { continue }
    if ($v -match '(?i)^explanation$') { continue }
    if (Test-InvalidUserDisplay -Name $v) { continue }
    return $v
  }

  return ''
}

function Get-LegalNameFromUiForm {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RecordSysId,
    [string]$Table = 'sc_req_item'
  )

  $sid = ("" + $RecordSysId).Trim()
  if (-not $sid) { return '' }
  $tbl = ("" + $Table).Trim()
  if (-not $tbl) { $tbl = 'sc_req_item' }

  $recordUrl = "{0}/nav_to.do?uri=%2F{1}.do%3Fsys_id%3D{2}" -f $Session.BaseUrl, $tbl, $sid
  try { $Session.WebView.CoreWebView2.Navigate($recordUrl) } catch { return '' }

  $readySw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($readySw.ElapsedMilliseconds -lt 12000) {
    $isReady = Invoke-WebViewScriptJson -WebView $Session.WebView -Script "document.readyState==='complete'" -TimeoutMs 2000
    if ($isReady -eq $true) { break }
    Start-Sleep -Milliseconds 250
  }

  $js = @"
(function(){
  try {
    function s(x){ return (x===null||x===undefined) ? '' : (''+x).trim(); }
    function valid(v){
      var t = s(v);
      if (!t) return '';
      if (/^explanation$/i.test(t)) return '';
      if (/^new\s+.*\s+user$/i.test(t)) return '';
      return t;
    }
    function extractFromDoc(doc){
      if (!doc) return '';

      var exact = doc.querySelector('input[placeholder="Joe Willy Smith"],textarea[placeholder="Joe Willy Smith"]');
      var ev = valid(exact && (exact.value || (exact.getAttribute && exact.getAttribute('value')) || ''));
      if (ev) return ev;

      var labels = doc.querySelectorAll('label, span, div');
      for (var i=0; i<labels.length; i++) {
        var lt = s(labels[i].innerText || labels[i].textContent || '');
        if (!/legal\s*name/i.test(lt)) continue;
        var lid = s(labels[i].id || '');
        var fr = s(labels[i].getAttribute && labels[i].getAttribute('for'));
        var cands = [];
        if (lid) cands = cands.concat([].slice.call(doc.querySelectorAll('[aria-labelledby="'+lid+'"], [aria-describedby="'+lid+'"]')));
        if (fr) cands = cands.concat([].slice.call(doc.querySelectorAll('#'+fr)));
        for (var j=0; j<cands.length; j++) {
          var v = valid(cands[j].value || (cands[j].getAttribute && cands[j].getAttribute('value')) || cands[j].innerText || '');
          if (v) return v;
        }
      }

      var vars = doc.querySelectorAll('input[id^="ni."],textarea[id^="ni."]');
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

  for ($attempt = 1; $attempt -le 6; $attempt++) {
    $o = Invoke-WebViewScriptJson -WebView $Session.WebView -Script $js -TimeoutMs 8000
    if ($o -and $o.ok -eq $true -and $o.PSObject.Properties['legal_name']) {
      $ln = ("" + $o.legal_name).Trim()
      if ($ln -and -not (Test-InvalidUserDisplay -Name $ln)) { return $ln }
    }
    Start-Sleep -Milliseconds 300
  }

  return ''
}

function Get-RitmCatalogFallbackUser {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmSysId
  )

  $details = Get-RitmCatalogFallbackDetails -Session $Session -RitmSysId $RitmSysId
  if ($details -and $details.PSObject.Properties['legal_name']) {
    $legal = ("" + $details.legal_name).Trim()
    if ($legal) { return $legal }
  }
  return ''
}

function Get-RitmCatalogFallbackDetails {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmSysId
  )

  $sid = ("" + $RitmSysId).Trim()
  if (-not $sid) { return [pscustomobject]@{ legal_name = ''; pi_machine = '' } }

  $mtomRows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_item_option_mtom' -Query ("request_item={0}" -f $sid) -Fields @('sc_item_option') -Limit 150)
  if ($mtomRows.Count -eq 0) { return [pscustomobject]@{ legal_name = ''; pi_machine = '' } }

  $bestLegal = ''
  $bestRequestedFor = ''
  $bestRequestedBy = ''
  $piCandidates = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

  foreach ($m in $mtomRows) {
    $optId = (Get-ObjectStringValue -Object $m -PropertyName 'sc_item_option')
    if (-not $optId) { continue }

    $optRows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_item_option' -Query ("sys_id={0}" -f $optId) -Fields @('value','item_option_new') -Limit 1)
    if ($optRows.Count -eq 0) { continue }

    $valueRaw = (Get-ObjectStringValue -Object $optRows[0] -PropertyName 'value')
    if (-not $valueRaw) { continue }

    $questionText = ''
    $ionId = (Get-ObjectStringValue -Object $optRows[0] -PropertyName 'item_option_new')
    if ($ionId) {
      $ionRows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'item_option_new' -Query ("sys_id={0}" -f $ionId) -Fields @('question_text','name') -Limit 1)
      if ($ionRows.Count -gt 0) {
        $questionText = (Get-ObjectStringValue -Object $ionRows[0] -PropertyName 'question_text')
        if (-not $questionText) { $questionText = (Get-ObjectStringValue -Object $ionRows[0] -PropertyName 'name') }
      }
    }
    $q = ("" + $questionText).Trim().ToLowerInvariant()

    $resolved = ("" + $valueRaw).Trim()
    if ($resolved -match '^[0-9a-fA-F]{32}$') {
      $resolvedCi = Resolve-ServiceNowCiDisplay -Session $Session -CiValue $resolved
      if ($resolvedCi -and ($resolvedCi -ne $resolved)) {
        $resolved = ("" + $resolvedCi).Trim()
      }
      else {
        $resolvedUser = Resolve-ServiceNowUserDisplay -Session $Session -UserValue $resolved
        if ($resolvedUser -and ($resolvedUser -ne $resolved)) {
          $resolved = ("" + $resolvedUser).Trim()
        }
      }
    }

    $piFromText = Get-DetectedPiFromText -Text ("{0} {1}" -f $q, $resolved)
    if ($piFromText) {
      foreach ($token in @($piFromText -split ',')) {
        $candidate = ("" + $token).Trim()
        if ($candidate) { [void]$piCandidates.Add($candidate) }
      }
    }

    if ($resolved -and -not (Test-InvalidUserDisplay -Name $resolved)) {
      if ($q -match 'legal\s*name') {
        $bestLegal = $resolved
        continue
      }
      if ($q -match '^requested\s*for$') {
        $bestRequestedFor = $resolved
        continue
      }
      if ($q -match '^requested\s*by$') {
        $bestRequestedBy = $resolved
        continue
      }
    }
  }

  $legalOut = ''
  if ($bestLegal) { $legalOut = $bestLegal }
  elseif ($bestRequestedFor) { $legalOut = $bestRequestedFor }
  elseif ($bestRequestedBy) { $legalOut = $bestRequestedBy }

  $piOut = ''
  $piArray = @($piCandidates)
  if ($piArray.Count -gt 0) {
    $piOut = (($piArray | ForEach-Object { ("" + $_).Trim() } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
  }

  return [pscustomobject]@{
    legal_name = $legalOut
    pi_machine = $piOut
  }
}

function Get-ServiceNowOpenTasksByRitm {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmSysId
  )

  if ([string]::IsNullOrWhiteSpace($RitmSysId)) { return @() }

  $fields = @('number','sys_id','state','state_value','short_description','description','comments','work_notes')
  $rows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_task' -Query ("request_item={0}" -f $RitmSysId) -Fields $fields -Limit 200)
  $out = New-Object System.Collections.Generic.List[object]

  foreach ($row in $rows) {
    $state = Get-ObjectStringValue -Object $row -PropertyName 'state'
    $stateValue = Get-ObjectStringValue -Object $row -PropertyName 'state_value'
    if (-not $stateValue) { $stateValue = $state }
    $stateLabel = Resolve-ServiceNowStateLabel -Session $Session -Table 'sc_task' -StateValue $stateValue -FallbackLabel $state
    if (-not (Test-ClosedState -StateLabel $stateLabel -StateValue $stateValue)) {
      $out.Add([pscustomobject]@{
        number = (Get-ObjectStringValue -Object $row -PropertyName 'number').ToUpperInvariant()
        sys_id = Get-ObjectStringValue -Object $row -PropertyName 'sys_id'
        state_label = $stateLabel
        state_value = $stateValue
        short_description = Get-ObjectStringValue -Object $row -PropertyName 'short_description'
        description = Get-ObjectStringValue -Object $row -PropertyName 'description'
        comments = Get-ObjectStringValue -Object $row -PropertyName 'comments'
        work_notes = Get-ObjectStringValue -Object $row -PropertyName 'work_notes'
      }) | Out-Null
    }
  }

  return @($out.ToArray())
}

function Get-RitmPiEvidenceText {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmSysId
  )

  $sid = ("" + $RitmSysId).Trim()
  if (-not $sid) { return '' }

  $parts = New-Object System.Collections.Generic.List[string]
  $taskIds = New-Object System.Collections.Generic.List[string]

  $taskFields = @('number','sys_id','short_description','description','comments','work_notes','item','close_notes','u_comments')
  $taskRows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_task' -Query ("request_item={0}" -f $sid) -Fields $taskFields -Limit 200)
  foreach ($tr in $taskRows) {
    foreach ($prop in @('number','short_description','description','comments','work_notes','item','close_notes','u_comments')) {
      $v = Get-ObjectStringValue -Object $tr -PropertyName $prop
      if ($v) { [void]$parts.Add($v) }
    }
    $taskId = Get-ObjectStringValue -Object $tr -PropertyName 'sys_id'
    if ($taskId) { [void]$taskIds.Add($taskId) }
  }

  $ritmJournalQueries = @(
    "name=sc_req_item^element_id={0}" -f $sid,
    "element_id={0}^elementINcomments,work_notes" -f $sid
  )
  foreach ($jq in $ritmJournalQueries) {
    $jr = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sys_journal_field' -Query $jq -Fields @('value','message','comments','work_notes') -Limit 200)
    foreach ($row in $jr) {
      foreach ($prop in @('value','message','comments','work_notes')) {
        $v = Get-ObjectStringValue -Object $row -PropertyName $prop
        if ($v) { [void]$parts.Add($v) }
      }
    }
  }

  $taskIdList = @($taskIds | Where-Object { $_ } | Select-Object -Unique)
  if ($taskIdList.Count -gt 0) {
    $chunkSize = 20
    for ($i = 0; $i -lt $taskIdList.Count; $i += $chunkSize) {
      $take = [Math]::Min($chunkSize, $taskIdList.Count - $i)
      $chunk = @($taskIdList | Select-Object -Skip $i -First $take)
      if ($chunk.Count -eq 0) { continue }
      $idClause = ($chunk -join ',')
      $q = "name=sc_task^element_idIN{0}^elementINcomments,work_notes" -f $idClause
      $jr = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sys_journal_field' -Query $q -Fields @('value','message','comments','work_notes') -Limit 500)
      foreach ($row in $jr) {
        foreach ($prop in @('value','message','comments','work_notes')) {
          $v = Get-ObjectStringValue -Object $row -PropertyName $prop
          if ($v) { [void]$parts.Add($v) }
        }
      }
    }
  }

  return (($parts | Where-Object { $_ } | Select-Object -Unique) -join ' ')
}

function ExecJS {
  param($wv, [string]$Js, [int]$TimeoutMs = 12000)
  return Invoke-WebViewScriptRaw -WebView $wv -Script $Js -TimeoutMs $TimeoutMs
}

function Parse-WV2Json {
  param([string]$Raw)
  return ConvertFrom-WebViewResult -Raw $Raw
}

function Build-RitmRecordUrl([string]$SysId) {
  if ([string]::IsNullOrWhiteSpace($SysId)) { return '' }
  return ("{0}/nav_to.do?uri=%2Fsc_req_item.do%3Fsys_id%3D{1}%26sysparm_view%3D" -f $script:ServiceNowBaseForLegacyPi, $SysId.Trim())
}

function Build-IncidentRecordUrl([string]$SysId) {
  if ([string]::IsNullOrWhiteSpace($SysId)) { return '' }
  return ("{0}/nav_to.do?uri=%2Fincident.do%3Fsys_id%3D{1}%26sysparm_view%3D" -f $script:ServiceNowBaseForLegacyPi, $SysId.Trim())
}

function Build-SCTaskFallbackUrl([string]$TaskNumber) {
  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return '' }
  $safeNumber = [System.Uri]::EscapeDataString($TaskNumber.Trim())
  return ("{0}/nav_to.do?uri=%2Fsc_task_list.do%3Fsysparm_query%3Dnumber%3D{1}" -f $script:ServiceNowBaseForLegacyPi, $safeNumber)
}

function Build-SCTaskListByRitmUrl {
  param([string]$RitmNumber)
  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return '' }
  $query = "request_item.number=" + $RitmNumber.Trim()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return ("{0}/nav_to.do?uri=%2Fsc_task_list.do%3Fsysparm_query%3D{1}" -f $script:ServiceNowBaseForLegacyPi, $safeQuery)
}

function Build-SCTaskRecordByNumberUrl {
  param([string]$TaskNumber)
  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return '' }
  $query = "number=" + $TaskNumber.Trim().ToUpperInvariant()
  $safeQuery = [System.Uri]::EscapeDataString($query)
  return ("{0}/nav_to.do?uri=%2Fsc_task.do%3Fsysparm_query%3D{1}" -f $script:ServiceNowBaseForLegacyPi, $safeQuery)
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

function Get-RecordActivityTextFromUiPage {
  param(
    $Session,
    [string]$RecordSysId,
    [string]$Table
  )

  if ([string]::IsNullOrWhiteSpace($RecordSysId)) { return '' }
  $recordUrl = if ($Table -eq 'incident') { Build-IncidentRecordUrl -SysId $RecordSysId } else { Build-RitmRecordUrl -SysId $RecordSysId }
  if ([string]::IsNullOrWhiteSpace($recordUrl)) { return '' }
  try { $Session.WebView.CoreWebView2.Navigate($recordUrl) } catch { return '' }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 12000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $Session.WebView "document.readyState==='complete'" 2000)
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
      return out.join(' ');
    }
    var shellText = collectFromDoc(document);
    var frame = document.querySelector('iframe#gsft_main') || document.querySelector('iframe[name=gsft_main]');
    var fdoc = (frame && frame.contentDocument) ? frame.contentDocument : null;
    var frameText = collectFromDoc(fdoc);
    return JSON.stringify({ ok:true, text: frameText ? frameText : shellText });
  } catch(e) {
    return JSON.stringify({ ok:false, text:'' });
  }
})();
"@
  $o = $null
  for ($attempt = 1; $attempt -le 8; $attempt++) {
    $o = Parse-WV2Json (ExecJS $Session.WebView $js 7000)
    $txt = if ($o -and $o.PSObject.Properties['text']) { ("" + $o.text) } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($txt)) { break }
    Start-Sleep -Milliseconds 350
  }
  if (-not $o) { return '' }
  if ($o.PSObject.Properties['text']) { return ("" + $o.text) }
  return ''
}

function Get-RitmTaskListTextFromUiPage {
  param(
    $Session,
    [string]$RitmNumber
  )

  if ([string]::IsNullOrWhiteSpace($RitmNumber)) { return '' }
  $taskListUrl = Build-SCTaskListByRitmUrl -RitmNumber $RitmNumber
  if ([string]::IsNullOrWhiteSpace($taskListUrl)) { return '' }
  try { $Session.WebView.CoreWebView2.Navigate($taskListUrl) } catch { return '' }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 10000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $Session.WebView "document.readyState==='complete'" 2000)
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
      var sels = ['table','.list2_body','.list2_body tr','.list2 td','.vt','.list_decoration','.linked','a'];
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
    return JSON.stringify({ ok:true, text: ft ? ft : shell });
  } catch(e){
    return JSON.stringify({ ok:false, text:'' });
  }
})();
"@
  $o = Parse-WV2Json (ExecJS $Session.WebView $js 7000)
  if (-not $o) { return '' }
  if ($o.PSObject.Properties['text']) { return ("" + $o.text) }
  return ''
}

function Get-SCTaskRecordTextFromUiPage {
  param(
    $Session,
    [string]$TaskNumber
  )

  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return '' }
  $taskUrl = Build-SCTaskFallbackUrl -TaskNumber $TaskNumber
  if ([string]::IsNullOrWhiteSpace($taskUrl)) { return '' }
  try { $Session.WebView.CoreWebView2.Navigate($taskUrl) } catch { return '' }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 10000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $Session.WebView "document.readyState==='complete'" 2000)
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
      var sels = ['.sn-widget-textblock-body','.sn-widget-textblock-body_formatted','.sn-card-component_accent-bar','.sn-card-component_accent-bar_dark','.activities-form','activities-form','table','input','textarea','label'];
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
  $o = Parse-WV2Json (ExecJS $Session.WebView $js 7000)
  if (-not $o) { return '' }
  if ($o.PSObject.Properties['text']) { return ("" + $o.text) }
  return ''
}

function Get-SCTaskActivityTextFromUiPage {
  param(
    $Session,
    [string]$TaskNumber
  )

  if ([string]::IsNullOrWhiteSpace($TaskNumber)) { return '' }
  $taskUrl = Build-SCTaskRecordByNumberUrl -TaskNumber $TaskNumber
  if ([string]::IsNullOrWhiteSpace($taskUrl)) { return '' }
  try { $Session.WebView.CoreWebView2.Navigate($taskUrl) } catch { return '' }

  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  while ($sw.ElapsedMilliseconds -lt 12000) {
    Start-Sleep -Milliseconds 250
    $isReady = Parse-WV2Json (ExecJS $Session.WebView "document.readyState==='complete'" 2000)
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
      var sels = ['h-card-wrapper activities-form','h-card-wrapper .activities-form','.activities-form','activities-form','.sn-widget-textblock-body','.sn-widget-textblock-body_formatted','.sn-card-component_accent-bar','.sn-card-component_accent-bar_dark','.activity-stream-text','.activity-stream','.journal','.journal_field','[data-stream-entry]'];
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
  $o = Parse-WV2Json (ExecJS $Session.WebView $js 7000)
  if (-not $o) { return '' }
  if ($o.PSObject.Properties['text']) { return ("" + $o.text) }
  return ''
}

function Get-SCTaskNumbersFromBackendByRitm {
  param(
    $Session,
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
  $o = Parse-WV2Json (ExecJS $Session.WebView $js 7000)
  if (-not $o) { return @() }
  if (-not $o.PSObject.Properties['tasks']) { return @() }
  return @($o.tasks)
}

function Get-ServiceNowTicket {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$Ticket,
    [ValidateSet('Auto','ConfigurationItemOnly','CommentsOnly','CommentsAndCI')][string]$PiSearchMode = 'Auto',
    [switch]$SkipLegalNameFallback
  )

  $ticketId = ("" + $Ticket).Trim().ToUpperInvariant()
  $ticketType = Get-TicketType -Ticket $ticketId
  if ($ticketType -eq 'UNKNOWN') {
    return [pscustomobject]@{ ticket = $ticketId; ok = $false; reason = 'unsupported_ticket_type' }
  }

  switch ($ticketType) {
    'RITM' {
      $table = 'sc_req_item'
      $fields = @('number','sys_id','state','state_value','requested_for','configuration_item','cmdb_ci','short_description','description','legal_name','u_legal_name','u_legalname','request')
      $userField = 'requested_for'
      $ciFields = @('configuration_item','cmdb_ci')
    }
    'INC' {
      $table = 'incident'
      $fields = @('number','sys_id','state','state_value','caller_id','cmdb_ci','short_description','description')
      $userField = 'caller_id'
      $ciFields = @('cmdb_ci')
    }
    'SCTASK' {
      $table = 'sc_task'
      $fields = @('number','sys_id','state','state_value','assigned_to','cmdb_ci','short_description','description')
      $userField = 'assigned_to'
      $ciFields = @('cmdb_ci')
    }
  }

  $mode = ("" + $PiSearchMode).Trim()
  if (-not $mode) { $mode = 'Auto' }
  $needComments = $mode -in @('Auto','CommentsOnly','CommentsAndCI')
  if ($needComments) {
    $fields = @($fields + @('comments', 'work_notes') | Select-Object -Unique)
  }

  $rows = Invoke-ServiceNowJsonv2Query -Session $Session -Table $table -Query ("number={0}" -f $ticketId) -Fields $fields -Limit 1
  $rowList = @($rows)
  if ($rowList.Count -eq 0) {
    return [pscustomobject]@{ ticket = $ticketId; ok = $false; reason = 'not_found'; table = $table }
  }

  $r = $rowList[0]
  $sysId = Get-ObjectStringValue -Object $r -PropertyName 'sys_id'
  $stateValue = Get-ObjectStringValue -Object $r -PropertyName 'state_value'
  $stateFallback = Get-ObjectStringValue -Object $r -PropertyName 'state'
  if (-not $stateValue) { $stateValue = $stateFallback }
  $stateLabel = Resolve-ServiceNowStateLabel -Session $Session -Table $table -StateValue $stateValue -FallbackLabel $stateFallback
  $user = Resolve-ServiceNowUserDisplay -Session $Session -UserValue (Get-ObjectStringValue -Object $r -PropertyName $userField)
  $ciRaw = ''
  foreach ($fieldName in @($ciFields)) {
    $candidate = Get-ObjectStringValue -Object $r -PropertyName $fieldName
    if ($candidate) { $ciRaw = $candidate; break }
  }
  $ci = Resolve-ServiceNowCiDisplay -Session $Session -CiValue $ciRaw

  $activityText = if ($needComments) {
    ((Get-ObjectStringValue -Object $r -PropertyName 'comments') + ' ' + (Get-ObjectStringValue -Object $r -PropertyName 'work_notes')).Trim()
  } else { '' }
  $recordTextForPi = @(
    (Get-ObjectStringValue -Object $r -PropertyName 'short_description'),
    (Get-ObjectStringValue -Object $r -PropertyName 'description'),
    $activityText
  ) -join ' '
  $piDetected = ''
  $piSource = 'none'
  $ciPiCandidate = ("" + $ci).Trim()
  $ciPiLooksValid = $ciPiCandidate -and ($ciPiCandidate -match '(?i)[A-Z0-9_-]{6,}')
  switch ($mode) {
    'ConfigurationItemOnly' {
      if ($ciPiLooksValid) {
        $piDetected = $ciPiCandidate
        $piSource = 'ci'
      }
    }
    'CommentsOnly' {
      $piDetected = Get-DetectedPiFromText -Text $recordTextForPi
      if ($piDetected) { $piSource = 'comments' }
    }
    'CommentsAndCI' {
      $piDetected = Get-DetectedPiFromText -Text $recordTextForPi
      if ($piDetected) {
        $piSource = 'comments'
      }
      elseif ($ciPiLooksValid) {
        $piDetected = $ciPiCandidate
        $piSource = 'ci'
      }
    }
    default {
      $piDetected = Get-DetectedPiFromText -Text $recordTextForPi
      if ($piDetected) {
        $piSource = 'comments'
      }
      elseif ($ciPiLooksValid) {
        $piDetected = $ciPiCandidate
        $piSource = 'ci'
      }
    }
  }
  $legalName = ''

  $openTasks = @()
  if ($ticketType -eq 'RITM') {
    $openTasks = Get-ServiceNowOpenTasksByRitm -Session $Session -RitmSysId $sysId
    $catalogFallback = $null

    $legalCandidates = @(
      (Get-ObjectStringValue -Object $r -PropertyName 'legal_name'),
      (Get-ObjectStringValue -Object $r -PropertyName 'u_legal_name'),
      (Get-ObjectStringValue -Object $r -PropertyName 'u_legalname')
    )
    foreach ($cand in $legalCandidates) {
      $v = ("" + $cand).Trim()
      if ($v -and -not (Test-InvalidUserDisplay -Name $v)) { $legalName = $v; break }
    }
    if (-not $legalName) {
      $textHints = @(
        (Get-ObjectStringValue -Object $r -PropertyName 'short_description'),
        (Get-ObjectStringValue -Object $r -PropertyName 'description'),
        $activityText
      ) -join ' '
      $legalName = Get-LegalNameFromText -Text $textHints
    }

    $needsLegalFallback = (Test-InvalidUserDisplay -Name $user)
    if ((-not $SkipLegalNameFallback -and $needsLegalFallback -and -not $legalName -and $sysId) -or (-not $piDetected -and $sysId)) {
      $catalogFallback = Get-RitmCatalogFallbackDetails -Session $Session -RitmSysId $sysId
    }
    if (-not $piDetected -and $catalogFallback -and $catalogFallback.PSObject.Properties['pi_machine']) {
      $piFromCatalog = ("" + $catalogFallback.pi_machine).Trim()
      if ($piFromCatalog) {
        $piDetected = $piFromCatalog
        $piSource = 'catalog'
      }
    }
    if (-not $piDetected -and @($openTasks).Count -gt 0) {
      $taskEvidenceText = @(
        @($openTasks | ForEach-Object { Get-ObjectStringValue -Object $_ -PropertyName 'short_description' }),
        @($openTasks | ForEach-Object { Get-ObjectStringValue -Object $_ -PropertyName 'description' }),
        @($openTasks | ForEach-Object { Get-ObjectStringValue -Object $_ -PropertyName 'comments' }),
        @($openTasks | ForEach-Object { Get-ObjectStringValue -Object $_ -PropertyName 'work_notes' })
      ) -join ' '
      if ($taskEvidenceText) {
        $piFromTasks = Get-DetectedPiFromText -Text $taskEvidenceText
        if ($piFromTasks) {
          $piDetected = $piFromTasks
          $piSource = 'task_evidence'
        }
      }
    }
    if (-not $piDetected -and $sysId) {
      $piEvidenceText = Get-RitmPiEvidenceText -Session $Session -RitmSysId $sysId
      if ($piEvidenceText) {
        $piFromJournal = Get-DetectedPiFromText -Text $piEvidenceText
        if ($piFromJournal) {
          $piDetected = $piFromJournal
          $piSource = 'task_journal'
        }
      }
    }
    if (-not $SkipLegalNameFallback -and $needsLegalFallback -and -not $legalName -and $sysId) {
      if ($catalogFallback -and $catalogFallback.PSObject.Properties['legal_name']) {
        $catalogUser = ("" + $catalogFallback.legal_name).Trim()
        if ($catalogUser) { $legalName = $catalogUser }
      }
    }
    if (-not $SkipLegalNameFallback -and -not $legalName -and $needsLegalFallback -and $sysId) {
      $legalFromForm = Get-LegalNameFromUiForm -Session $Session -RecordSysId $sysId -Table 'sc_req_item'
      if ($legalFromForm) { $legalName = $legalFromForm }
    }
    if ($legalName -and $needsLegalFallback) {
      $user = $legalName
    }
  }

  if ($ticketType -in @('RITM', 'INC')) {
    $EnableActivityStreamSearch = $true
    $EnableUiFallbackActivitySearch = $true
    $UiFallbackMinBackendChars = 120
    $DebugActivityTicket = ''
    $VerboseTicketLogging = $false

    $uiActivityText = ''
    $currentUserSnapshot = ("" + $user).Trim()
    $isNewEpUserContext = ($currentUserSnapshot -match '(?i)^new\b.*\buser$')
    $needsLegalNameFallback = (
      ($ticketType -eq 'RITM') -and
      (-not $legalName) -and
      (
        [string]::IsNullOrWhiteSpace($currentUserSnapshot) -or
        ($currentUserSnapshot -match '^[0-9a-fA-F]{32}$') -or
        ($currentUserSnapshot -match '(?i)^new\b.*\buser$')
      )
    )

    $script:ServiceNowBaseForLegacyPi = $Session.BaseUrl
    if (
      (
        ([string]::IsNullOrWhiteSpace($piDetected) -and $EnableActivityStreamSearch) -or
        $needsLegalNameFallback
      ) -and
      $EnableUiFallbackActivitySearch -and
      (
        ($activityText.Length -lt $UiFallbackMinBackendChars) -or
        $needsLegalNameFallback -or
        ($ticketType -eq 'RITM')
      )
    ) {
      $uiActivityText = Get-RecordActivityTextFromUiPage -Session $Session -RecordSysId $sysId -Table $table
      if (-not [string]::IsNullOrWhiteSpace($uiActivityText)) {
        $piFromUi = Get-DetectedPiFromActivityText -ActivityText $uiActivityText
        if ($piFromUi) {
          $piDetected = $piFromUi
          $piSource = 'ui_activity'
        }
        if (-not $legalName) { $legalName = Get-LegalNameFromText -Text $uiActivityText }
      }
    }

    if (-not $legalName) { $legalName = Get-LegalNameFromText -Text $activityText }
    if (-not $legalName -and ($ticketType -eq 'RITM')) {
      $legalFromForm = Get-LegalNameFromUiForm -Session $Session -RecordSysId $sysId -Table 'sc_req_item'
      if (-not [string]::IsNullOrWhiteSpace($legalFromForm)) { $legalName = $legalFromForm }
    }

    if ([string]::IsNullOrWhiteSpace($piDetected) -and ($ticketType -eq 'RITM') -and ($isNewEpUserContext -or ($ticketId -eq $DebugActivityTicket))) {
      $taskUiText = Get-RitmTaskListTextFromUiPage -Session $Session -RitmNumber $ticketId
      $taskUiLen = if ($taskUiText) { $taskUiText.Length } else { 0 }
      if ($taskUiLen -gt 0) {
        $piFromTaskUi = Get-DetectedPiFromActivityText -ActivityText $taskUiText
        if (-not [string]::IsNullOrWhiteSpace($piFromTaskUi)) {
          $piDetected = $piFromTaskUi
          $piSource = 'sctask_ui_list'
        }
        else {
          $taskNums = Get-TaskNumbersFromText -Text $taskUiText
          if ($taskNums.Count -eq 0) {
            $taskNums = @(Get-SCTaskNumbersFromBackendByRitm -Session $Session -RitmNumber $ticketId)
          }
          if ($taskNums.Count -gt 0) {
            $maxTaskDeepScan = if ($isNewEpUserContext) { 12 } else { 4 }
            $scanCount = [Math]::Min($taskNums.Count, $maxTaskDeepScan)
            $matchedPrepareTask = $false
            for ($ti = 0; $ti -lt $scanCount; $ti++) {
              $tn = $taskNums[$ti]
              $taskActivityText = Get-SCTaskActivityTextFromUiPage -Session $Session -TaskNumber $tn
              if (-not [string]::IsNullOrWhiteSpace($taskActivityText)) {
                $piFromTaskActivity = Get-DetectedPiFromActivityText -ActivityText $taskActivityText
                if ($isNewEpUserContext -and -not [string]::IsNullOrWhiteSpace($piFromTaskActivity)) {
                  $piDecision = Resolve-ConfidentPiFromSource -PiListText $piFromTaskActivity -SourceText $taskActivityText
                  if ($piDecision -and $piDecision.selected) {
                    $piFromTaskActivity = "" + $piDecision.selected
                  }
                }
                if (-not [string]::IsNullOrWhiteSpace($piFromTaskActivity)) {
                  $piDetected = $piFromTaskActivity
                  $piSource = ("sctask_activity_record:" + $tn)
                  break
                }
              }

              $taskRecordText = Get-SCTaskRecordTextFromUiPage -Session $Session -TaskNumber $tn
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
              }
              if (-not [string]::IsNullOrWhiteSpace($piFromTaskRecord)) {
                $piDetected = $piFromTaskRecord
                $piSource = if ($isNewEpUserContext) { ("sctask_prepare_device_record:" + $tn) } else { ("sctask_ui_record:" + $tn) }
                break
              }
            }
          }
        }
      }
    }

    if ($legalName -and (Test-InvalidUserDisplay -Name $user)) {
      $user = $legalName
    }
  }

  $openTaskList = @($openTasks)
  $completion = Get-CompletionStatus -Ticket $ticketId -StateLabel $stateLabel -StateValue $stateValue -OpenTasks $openTaskList.Count

  return [pscustomobject]@{
    ticket = $ticketId
    ok = $true
    table = $table
    sys_id = $sysId
    status = $stateLabel
    status_value = $stateValue
    affected_user = $user
    configuration_item = $ci
    short_description = Get-ObjectStringValue -Object $r -PropertyName 'short_description'
    detected_pi_machine = $piDetected
    pi_source = $piSource
    open_tasks = $openTaskList.Count
    open_task_items = @($openTaskList)
    open_task_numbers = @($openTaskList | ForEach-Object { $_.number })
    completion_status = $completion
    legal_name = $legalName
    record_url = if ($sysId) { "{0}/nav_to.do?uri=%2F{1}.do%3Fsys_id%3D{2}" -f $Session.BaseUrl, $table, $sysId } else { '' }
  }
}

function global:Set-ServiceNowTaskState {
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

function global:Get-ServiceNowTasksForRitm {
  param(
    [Parameter(Mandatory = $true)]$Session,
    [Parameter(Mandatory = $true)][string]$RitmNumber
  )

  $ritm = ("" + $RitmNumber).Trim().ToUpperInvariant()
  if (-not $ritm) { return @() }

  $rows = @(Invoke-ServiceNowJsonv2Query -Session $Session -Table 'sc_task' -Query ("request_item.number={0}" -f $ritm) -Fields @('number','sys_id','state','state_value') -Limit 200)
  $out = New-Object System.Collections.Generic.List[object]
  foreach ($r in $rows) {
    $out.Add([pscustomobject]@{
      number = (Get-ObjectStringValue -Object $r -PropertyName 'number').ToUpperInvariant()
      sys_id = Get-ObjectStringValue -Object $r -PropertyName 'sys_id'
      state = Get-ObjectStringValue -Object $r -PropertyName 'state'
      state_value = Get-ObjectStringValue -Object $r -PropertyName 'state_value'
    }) | Out-Null
  }

  return @($out.ToArray())
}



