Set-StrictMode -Version Latest

function Invoke-DashboardSearchWorkflow {
  param(
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [string]$SearchText = ''
  )

  $rows = Search-DashboardRows -ExcelPath $ExcelPath -SheetName $SheetName -SearchText $SearchText
  Write-RunLog -RunContext $RunContext -Level INFO -Message ("Dashboard search query='{0}' => {1} row(s)" -f $SearchText, $rows.Count)
  return $rows
}

function Invoke-DashboardCheckInWorkflow {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][int]$Row,
    [string]$WorkNote = 'Deliver all credentials to the new user'
  )

  Invoke-DashboardActionCore -Config $Config -RunContext $RunContext -ExcelPath $ExcelPath -SheetName $SheetName -Row $Row `
    -DesiredStatus 'Checked-In' -TargetTaskState 'Work in Progress' -WorkNote $WorkNote
}

function Invoke-DashboardCheckOutWorkflow {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][int]$Row,
    [string]$WorkNote = "Laptop has been delivered.`r`nFirst login made with the user.`r`nOutlook, Teams and Jabber successfully tested."
  )

  Invoke-DashboardActionCore -Config $Config -RunContext $RunContext -ExcelPath $ExcelPath -SheetName $SheetName -Row $Row `
    -DesiredStatus 'Checked-Out' -TargetTaskState 'Closed Complete' -WorkNote $WorkNote
}

function Invoke-DashboardActionCore {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][int]$Row,
    [Parameter(Mandatory = $true)][string]$DesiredStatus,
    [Parameter(Mandatory = $true)][string]$TargetTaskState,
    [string]$WorkNote
  )

  $rows = Search-DashboardRows -ExcelPath $ExcelPath -SheetName $SheetName -SearchText ''
  $match = $rows | Where-Object { $_.Row -eq $Row } | Select-Object -First 1
  if (-not $match) {
    throw "Dashboard row $Row not found."
  }

  $ritm = ("" + $match.RITM).Trim().ToUpperInvariant()
  if (-not ($ritm -match '^RITM\d{6,8}$')) {
    throw "Row $Row does not contain a valid RITM number."
  }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Dashboard action: row=$Row ritm=$ritm desiredStatus=$DesiredStatus"

  $session = $null
  try {
    $session = New-ServiceNowSession -Config $Config -RunContext $RunContext
    $tasks = Get-ServiceNowTasksForRitm -Session $session -RitmNumber $ritm
    if ($tasks.Count -eq 0) {
      throw "No SCTASK found for $ritm"
    }

    $candidate = $null
    if ($DesiredStatus -eq 'Checked-In') {
      $candidate = $tasks | Where-Object { ("" + $_.state_value).Trim() -eq '1' -or (("" + $_.state) -match '(?i)^open|^new$') } | Select-Object -First 1
    } else {
      $candidate = $tasks | Where-Object { ("" + $_.state_value).Trim() -in @('1','2') -or (("" + $_.state) -match '(?i)open|new|in\s*progress|work\s*in\s*progress') } | Select-Object -First 1
    }

    if (-not $candidate) {
      throw "No eligible SCTASK found for $ritm"
    }

    $ok = Set-ServiceNowTaskState -Session $session -TaskSysId $candidate.sys_id -TargetStateLabel $TargetTaskState -WorkNote $WorkNote
    if (-not $ok) {
      throw "ServiceNow update failed for task $($candidate.number)"
    }

    Update-DashboardRow -ExcelPath $ExcelPath -SheetName $SheetName -Row $Row -Status $DesiredStatus -SCTaskNumber $candidate.number
    Write-RunLog -RunContext $RunContext -Level INFO -Message "Dashboard action succeeded: row=$Row ritm=$ritm task=$($candidate.number)"

    return [pscustomobject]@{
      ok = $true
      row = $Row
      ritm = $ritm
      task = $candidate.number
      status = $DesiredStatus
    }
  }
  finally {
    Close-ServiceNowSession -Session $session
  }
}
