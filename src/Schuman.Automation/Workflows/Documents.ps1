Set-StrictMode -Version Latest

function Invoke-DocumentGenerationWorkflow {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Config,
    [Parameter(Mandatory = $true)][hashtable]$RunContext,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [string]$SheetName,
    [string]$TemplatePath,
    [string]$OutputDirectory,
    [switch]$ExportPdf,
    [int[]]$RowNumbers = @()
  )

  if (-not $SheetName) { $SheetName = $Config.Excel.DefaultSheet }
  if (-not $TemplatePath) { $TemplatePath = Join-Path (Split-Path -Parent $Config.Output.SystemRoot) $Config.Documents.TemplateFile }
  if (-not $OutputDirectory) { $OutputDirectory = Join-Path (Split-Path -Parent $Config.Output.SystemRoot) $Config.Documents.OutputFolder }
  $TemplatePath = [System.IO.Path]::GetFullPath($TemplatePath)
  $OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)

  if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Word template not found: $TemplatePath"
  }

  Ensure-Directory -Path $OutputDirectory | Out-Null
  Write-RunLog -RunContext $RunContext -Level INFO -Message "Generating documents from '$TemplatePath'"

  $targets = @(Read-DocumentRowsFromExcel -ExcelPath $ExcelPath -SheetName $SheetName -SelectedRowNumbers $RowNumbers)

  if ($targets.Count -eq 0) {
    Write-RunLog -RunContext $RunContext -Level WARN -Message 'No eligible rows found for document generation.'
    return @()
  }

  $word = $null
  $generated = New-Object System.Collections.Generic.List[object]
  try {
    $word = New-Object -ComObject Word.Application
    if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
      Register-SchumanOwnedComResource -Kind 'WordApplication' -Object $word -Tag 'documents-generate' | Out-Null
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0

    foreach ($row in $targets) {
      $ticketValue = ''
      try {
      $ticketValue = ("" + $row.RITM).Trim()
      $nameValue = ("" + $row.RequestedFor).Trim()
      $piValue = ("" + $row.PI).Trim()
      $equipmentValue = ("" + $row.ITEquipment).Trim()
      if (-not $ticketValue) { $ticketValue = 'UNKNOWN_TICKET' }
      if (-not $nameValue) { $nameValue = 'UNKNOWN_NAME' }
      if (-not $equipmentValue) { $equipmentValue = 'Laptop' }

      $nameSource = if ($nameValue) { $nameValue } else { $ticketValue }
      $safeName = Get-SafeFileName -Text $nameSource
      $baseFile = "{0}_{1}" -f $ticketValue, $safeName
      $docxPath = [System.IO.Path]::GetFullPath((Join-Path $OutputDirectory ("$baseFile.docx")))
      $expectedPdfPath = [System.IO.Path]::GetFullPath((Join-Path $OutputDirectory ("$baseFile.pdf")))

      $docxExists = $false
      $pdfExists = $false
      try { $docxExists = Test-Path -LiteralPath $docxPath } catch { $docxExists = $false }
      try { $pdfExists = Test-Path -LiteralPath $expectedPdfPath } catch { $pdfExists = $false }
      if ($docxExists -and ((-not $ExportPdf) -or $pdfExists)) {
        Write-RunLog -RunContext $RunContext -Level INFO -Message ("Skipping row {0}: output already exists ({1})." -f $row.Row, $baseFile)
        $generated.Add([pscustomobject]@{
            Row = $row.Row
            RITM = $ticketValue
            DocxPath = $docxPath
            PdfPath = if ($ExportPdf) { $expectedPdfPath } else { '' }
            Error = $null
          }) | Out-Null
        continue
      }

      $doc = $null
      try {
        Copy-Item -LiteralPath $TemplatePath -Destination $docxPath -Force
        $doc = $word.Documents.Open($docxPath)
        if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
          Register-SchumanOwnedComResource -Kind 'WordDocument' -Object $doc -Tag 'documents-generate' | Out-Null
        }

        # Legacy compatibility: some templates are protected and require explicit unprotect before replacement.
        try {
          $prot = [int]$doc.ProtectionType
          if ($prot -ne 0) { [void]$doc.Unprotect() }
        }
        catch {}

        Set-WordPlaceholderValue -Document $doc -Key 'FieldDisplayName' -Value $nameValue
        Set-WordPlaceholderValue -Document $doc -Key 'FieldTicketNumber' -Value $ticketValue
        Set-WordPlaceholderValue -Document $doc -Key 'FieldPINumber' -Value $piValue
        Set-WordPlaceholderValue -Document $doc -Key 'FieldITEquipment' -Value $equipmentValue
        Set-WordPlaceholderValue -Document $doc -Key 'FieldTelephoneNumber' -Value $piValue
        Set-WordPlaceholderValue -Document $doc -Key 'FieldUniqueIdentifier' -Value $ticketValue

        Set-WordPlaceholder -Document $doc -Placeholder '{{RITM}}' -Value $ticketValue
        Set-WordPlaceholder -Document $doc -Placeholder '{{NAME}}' -Value $nameValue
        Set-WordPlaceholder -Document $doc -Placeholder '{{SCTASK}}' -Value ("" + $row.SCTASK).Trim()
        Set-WordPlaceholder -Document $doc -Placeholder '{{DATE}}' -Value (Get-Date -Format 'yyyy-MM-dd')

        $doc.Save()

        $pdfPath = ''
        if ($ExportPdf) {
          $pdfPath = $expectedPdfPath
          try {
            $wdExportFormatPDF = 17
            $doc.ExportAsFixedFormat($pdfPath, $wdExportFormatPDF)
          }
          catch {
            $wdFormatPDF = 17
            $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
          }
        }

        $generated.Add([pscustomobject]@{
          Row = $row.Row
          RITM = $ticketValue
          DocxPath = $docxPath
          PdfPath = $pdfPath
          Error = $null
        }) | Out-Null
      }
      finally {
        try { if ($doc) { $doc.Close($true) | Out-Null } } catch {}
        try { if ($doc -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $doc } } catch {}
        try { if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } } catch {}
      }
      }
      catch {
        Write-RunLog -RunContext $RunContext -Level ERROR -Message ("Row {0} failed: {1}" -f $row.Row, $_.Exception.Message)
        $generated.Add([pscustomobject]@{
          Row = $row.Row
          RITM = $ticketValue
          DocxPath = ''
          PdfPath = ''
          Error = $_.Exception.Message
        }) | Out-Null
      }
    }
  }
  finally {
    try { if ($word) { $word.Quit() | Out-Null } } catch {}
    try { if ($word -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $word } } catch {}
    try { if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) } } catch {}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  Write-RunLog -RunContext $RunContext -Level INFO -Message "Document generation completed. Files=$($generated.Count)"
  return @($generated.ToArray())
}

function Read-DocumentRowsFromExcel {
  param(
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [int[]]$SelectedRowNumbers = @()
  )

  $selected = @{}
  foreach ($n in @($SelectedRowNumbers)) {
    try { $selected[[int]$n] = $true } catch {}
  }

  $out = New-Object System.Collections.Generic.List[object]
  $excel = $null
  $wb = $null
  $ws = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, $null, $true)
    $ws = $wb.Worksheets.Item($SheetName)

    $map = Get-ExcelHeaderMap -Worksheet $ws
    $nameCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Name', 'Requested for', 'Requested For', 'User')
    $ticketCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Ticket', 'Number', 'RITM', 'Request Item')
    $piCol = Resolve-HeaderColumn -HeaderMap $map -Names @('PI', 'Phone', 'Configuration Item', 'CI')
    $equipCol = Resolve-HeaderColumn -HeaderMap $map -Names @('Receive ID Equipment', 'IT Equipment', 'Equipment')
    $taskCol = Resolve-HeaderColumn -HeaderMap $map -Names @('SCTasks', 'SCTask', 'SC Task')

    if (-not $ticketCol -and -not $nameCol -and -not $piCol) {
      throw 'Excel does not contain usable columns (Ticket/Number, Name or PI).'
    }

    $lastRow = [int]($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1)
    if ($lastRow -lt 2) { return @() }

    for ($r = 2; $r -le $lastRow; $r++) {
      if ($selected.Count -gt 0 -and -not $selected.ContainsKey($r)) { continue }

      $ticket = if ($ticketCol) { ("" + $ws.Cells.Item($r, $ticketCol).Text).Trim().ToUpperInvariant() } else { '' }
      $name = if ($nameCol) { ("" + $ws.Cells.Item($r, $nameCol).Text).Trim() } else { '' }
      $pi = if ($piCol) { ("" + $ws.Cells.Item($r, $piCol).Text).Trim() } else { '' }
      $equipment = if ($equipCol) { ("" + $ws.Cells.Item($r, $equipCol).Text).Trim() } else { '' }
      $task = if ($taskCol) { ("" + $ws.Cells.Item($r, $taskCol).Text).Trim() } else { '' }

      if (-not $ticket -and -not $name -and -not $pi) { continue }

      $out.Add([pscustomobject]@{
          Row = $r
          RITM = $ticket
          RequestedFor = $name
          PI = $pi
          ITEquipment = $equipment
          SCTASK = $task
        }) | Out-Null
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

  return @($out.ToArray())
}

function Set-WordFieldValue {
  param(
    [Parameter(Mandatory = $true)]$Document,
    [Parameter(Mandatory = $true)][string]$Key,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value
  )

  $k = ("" + $Key).Trim()
  if (-not $k) { return }
  $v = if ($null -eq $Value) { '' } else { ("" + $Value).Trim() }

  $setByCustomProp = $false
  try {
    $customProps = $Document.CustomDocumentProperties
    $existing = $null
    try { $existing = $customProps.Item($k) } catch { $existing = $null }
    if ($existing) {
      try { $existing.Value = $v; $setByCustomProp = $true } catch {}
    }
    else {
      try {
        # 4 = msoPropertyTypeString
        [void]$customProps.Add($k, $false, 4, $v)
        $setByCustomProp = $true
      }
      catch {}
    }
  }
  catch {}

  if (-not $setByCustomProp) {
    try {
      if ($Document.Bookmarks.Exists($k)) {
        $bm = $Document.Bookmarks.Item($k)
        $bm.Range.Text = $v
        [void]$Document.Bookmarks.Add($k, $bm.Range)
      }
    }
    catch {}
    try {
      $ff = $Document.FormFields.Item($k)
      if ($ff) { $ff.Result = $v }
    }
    catch {}
    try {
      foreach ($cc in $Document.ContentControls) {
        if ($cc.Title -eq $k -or $cc.Tag -eq $k) { $cc.Range.Text = $v }
      }
    }
    catch {}
  }

  Set-WordPlaceholder -Document $Document -Placeholder $k -Value $v
  Set-WordPlaceholder -Document $Document -Placeholder ("{{{0}}}" -f $k) -Value $v
  Set-WordPlaceholder -Document $Document -Placeholder ("[{0}]" -f $k) -Value $v
  Set-WordPlaceholder -Document $Document -Placeholder ("<<{0}>>" -f $k) -Value $v
}

function Set-WordPlaceholderValue {
  param(
    [Parameter(Mandatory = $true)]$Document,
    [Parameter(Mandatory = $true)][string]$Key,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value
  )

  $k = ("" + $Key).Trim()
  if (-not $k) { return }
  $v = if ($null -eq $Value) { '' } else { ("" + $Value).Trim() }

  $changed = $false
  try {
    $hits = 0
    foreach ($cc in $Document.ContentControls) {
      if ($cc.Title -eq $k -or $cc.Tag -eq $k) { $cc.Range.Text = $v; $hits++ }
    }
    if ($hits -gt 0) { $changed = $true }
  }
  catch {}

  if (-not $changed) {
    try {
      if ($Document.Bookmarks.Exists($k)) {
        $bm = $Document.Bookmarks.Item($k)
        $bm.Range.Text = $v
        [void]$Document.Bookmarks.Add($k, $bm.Range)
        $changed = $true
      }
    }
    catch {}
  }

  if (-not $changed) {
    try {
      $ff = $Document.FormFields.Item($k)
      if ($ff) { $ff.Result = $v; $changed = $true }
    }
    catch {}
  }

  if (-not $changed) {
    try {
      $customProps = $Document.CustomDocumentProperties
      $existing = $null
      try { $existing = $customProps.Item($k) } catch { $existing = $null }
      if ($existing) {
        try { $existing.Value = $v; $changed = $true } catch {}
      }
      else {
        try { [void]$customProps.Add($k, $false, 4, $v); $changed = $true } catch {}
      }
    }
    catch {}
  }

  $tokens = @($k, "{{${k}}}", "{${k}}", "<<${k}>>", "[${k}]", "[[${k}]]")
  $count = 0
  foreach ($t in $tokens) {
    $count += (Replace-WordToken -Range $Document.Content -Pattern $t -ReplaceValue $v -WholeWord $true)
    $count += (Replace-WordTokenInHeadersFooters -Document $Document -Pattern $t -ReplaceValue $v -WholeWord $true)
  }
  if ($count -eq 0) {
    foreach ($t in $tokens) {
      $count += (Replace-WordToken -Range $Document.Content -Pattern $t -ReplaceValue $v -WholeWord $false)
      $count += (Replace-WordTokenInHeadersFooters -Document $Document -Pattern $t -ReplaceValue $v -WholeWord $false)
    }
  }
}

function Replace-WordToken {
  param(
    [Parameter(Mandatory = $true)]$Range,
    [Parameter(Mandatory = $true)][string]$Pattern,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ReplaceValue,
    [bool]$WholeWord = $true
  )

  try {
    $text = ("" + $Range.Text)
    if ([string]::IsNullOrWhiteSpace($text)) { return 0 }
    $escaped = [regex]::Escape($Pattern)
    $regex = if ($WholeWord) { "(?<!\w)$escaped(?!\w)" } else { $escaped }
    $matches = [regex]::Matches($text, $regex).Count
    if ($matches -le 0) { return 0 }

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
    [void]$find.Execute($Pattern, $false, $false, $false, $false, $false, $true, 1, $false, $ReplaceValue, 2)
    return [int]$matches
  }
  catch {
    return 0
  }
}

function Replace-WordTokenInHeadersFooters {
  param(
    [Parameter(Mandatory = $true)]$Document,
    [Parameter(Mandatory = $true)][string]$Pattern,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ReplaceValue,
    [bool]$WholeWord = $true
  )

  $count = 0
  try {
    foreach ($sec in $Document.Sections) {
      foreach ($hfGroup in @($sec.Headers, $sec.Footers)) {
        foreach ($item in @($hfGroup.Item(1), $hfGroup.Item(2), $hfGroup.Item(3))) {
          try { $count += (Replace-WordToken -Range $item.Range -Pattern $Pattern -ReplaceValue $ReplaceValue -WholeWord $WholeWord) } catch {}
        }
      }
    }
  }
  catch {}
  return [int]$count
}

function Update-WordDocumentFields {
  param([Parameter(Mandatory = $true)]$Document)

  try { [void]$Document.Fields.Update() } catch {}
  try {
    foreach ($section in $Document.Sections) {
      foreach ($hfGroup in @($section.Headers, $section.Footers)) {
        foreach ($hf in @($hfGroup.Item(1), $hfGroup.Item(2), $hfGroup.Item(3))) {
          try { [void]$hf.Range.Fields.Update() } catch {}
        }
      }
    }
  }
  catch {}
}

function Set-WordPlaceholder {
  param(
    [Parameter(Mandatory = $true)]$Document,
    [Parameter(Mandatory = $true)][string]$Placeholder,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value
  )

  $find = $Document.Content.Find
  $find.ClearFormatting()
  $find.Replacement.ClearFormatting()
  $find.Text = $Placeholder
  $find.Replacement.Text = $Value

  $wdReplaceAll = 2
  $null = $find.Execute($find.Text, $false, $false, $false, $false, $false, $true, 1, $false, $find.Replacement.Text, $wdReplaceAll)
}

function Get-SafeFileName {
  param([string]$Text)

  $name = ("" + $Text).Trim()
  if (-not $name) { $name = 'Unknown' }

  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($ch in $invalid) {
    $name = $name.Replace($ch, '_')
  }

  $name = [System.Text.RegularExpressions.Regex]::Replace($name, '\s+', ' ').Trim()
  if ($name.Length -gt 80) {
    $name = $name.Substring(0, 80).Trim()
  }
  if (-not $name) { $name = 'Unknown' }

  return $name
}
