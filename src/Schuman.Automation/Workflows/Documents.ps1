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

  $rows = Search-DashboardRows -ExcelPath $ExcelPath -SheetName $SheetName -SearchText ''
  $targets = @($rows | Where-Object { $_.RITM -match '^RITM\d{6,8}$' })
  if ($RowNumbers -and @($RowNumbers).Count -gt 0) {
    $selected = @{}
    foreach ($n in @($RowNumbers)) {
      try { $selected[[int]$n] = $true } catch {}
    }
    $targets = @($targets | Where-Object {
        $rnum = 0
        if (-not [int]::TryParse(("" + $_.Row).Trim(), [ref]$rnum)) { return $false }
        return $selected.ContainsKey($rnum)
      })
  }

  if ($targets.Count -eq 0) {
    Write-RunLog -RunContext $RunContext -Level WARN -Message 'No eligible RITM rows found for document generation.'
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
      $nameSource = if ($row.RequestedFor) { $row.RequestedFor } else { $row.RITM }
      $safeName = Get-SafeFileName -Text $nameSource
      $baseFile = "{0}_{1}" -f $row.RITM, $safeName
      $docxPath = [System.IO.Path]::GetFullPath((Join-Path $OutputDirectory ("$baseFile.docx")))

      $doc = $null
      try {
        Copy-Item -LiteralPath $TemplatePath -Destination $docxPath -Force
        $doc = $word.Documents.Open($docxPath)
        if (Get-Command -Name Register-SchumanOwnedComResource -ErrorAction SilentlyContinue) {
          Register-SchumanOwnedComResource -Kind 'WordDocument' -Object $doc -Tag 'documents-generate' | Out-Null
        }

        Set-WordPlaceholder -Document $doc -Placeholder '{{RITM}}' -Value $row.RITM
        Set-WordPlaceholder -Document $doc -Placeholder '{{NAME}}' -Value $row.RequestedFor
        Set-WordPlaceholder -Document $doc -Placeholder '{{SCTASK}}' -Value $row.SCTASK
        Set-WordPlaceholder -Document $doc -Placeholder '{{DATE}}' -Value (Get-Date -Format 'yyyy-MM-dd')

        $doc.Save()

        $pdfPath = ''
        if ($ExportPdf) {
          $pdfPath = [System.IO.Path]::GetFullPath((Join-Path $OutputDirectory ("$baseFile.pdf")))
          $wdFormatPDF = 17
          $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
        }

        $generated.Add([pscustomobject]@{
          Row = $row.Row
          RITM = $row.RITM
          DocxPath = $docxPath
          PdfPath = $pdfPath
        }) | Out-Null
      }
      finally {
        try { if ($doc) { $doc.Close($true) | Out-Null } } catch {}
        try { if ($doc -and (Get-Command -Name Unregister-SchumanOwnedComResource -ErrorAction SilentlyContinue)) { Unregister-SchumanOwnedComResource -Object $doc } } catch {}
        try { if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } } catch {}
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
