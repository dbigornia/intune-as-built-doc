# PDF export for IntuneAsBuilt.
#   Word        - best result: builds the table of contents, adds heading bookmarks. Needs desktop Word on Windows.
#   LibreOffice - no Office needed (servers, scheduled runs). The table of contents is not built.

function Find-Soffice {
    $cmd = Get-Command soffice -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $candidates = @(
        (Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe'),
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        '/usr/bin/soffice'
    )
    foreach ($c in $candidates) { if ($c -and (Test-Path -LiteralPath $c)) { return $c } }
    return $null
}

function Convert-DocxToPdfWord {
    <# Also saves the .docx after updating fields, so the Word file opens with its table of contents built. #>
    param([Parameter(Mandatory)][string]$DocxPath, [Parameter(Mandatory)][string]$PdfPath)
    $word = $null; $wdoc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0                                   # wdAlertsNone
        $wdoc = $word.Documents.Open($DocxPath, $false, $false, $false)
        [void]$wdoc.Fields.Update()
        foreach ($toc in $wdoc.TablesOfContents) { $toc.Update() }
        $wdoc.Save()
        # OutputFileName, ExportFormat=PDF(17), OpenAfterExport, OptimizeFor=Print(0), Range=All(0), From, To,
        # Item=Content(0), IncludeDocProps, KeepIRM, CreateBookmarks=Headings(1), DocStructureTags, BitmapMissingFonts, UseISO19005_1
        $wdoc.ExportAsFixedFormat($PdfPath, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1, $true, $true, $false)
    }
    finally {
        if ($wdoc) { $wdoc.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wdoc) }
        if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Convert-DocxToPdfLibreOffice {
    param(
        [Parameter(Mandatory)][string]$DocxPath, [Parameter(Mandatory)][string]$PdfPath, [Parameter(Mandatory)][string]$Soffice,
        [int]$TimeoutSeconds = 600
    )
    $outDir = Join-Path ([IO.Path]::GetTempPath()) "IntuneAsBuilt-pdf-$([guid]::NewGuid())"
    $loProfile = Join-Path ([IO.Path]::GetTempPath()) "IntuneAsBuilt-lo-$([guid]::NewGuid())"
    New-Item -ItemType Directory -Path $outDir | Out-Null
    try {
        # Separate profile so it works even if LibreOffice is already open
        $profileUri = ([Uri](New-Item -ItemType Directory -Path $loProfile).FullName).AbsoluteUri
        $psi = [System.Diagnostics.ProcessStartInfo]::new($Soffice)
        foreach ($a in @('--headless', '--norestore', '--nolockcheck', "-env:UserInstallation=$profileUri",
                         '--convert-to', 'pdf', '--outdir', $outDir, $DocxPath)) { $psi.ArgumentList.Add($a) }
        $psi.UseShellExecute = $false
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $proc = [System.Diagnostics.Process]::Start($psi)
        while (-not $proc.WaitForExit(15000)) {
            if ($sw.Elapsed.TotalSeconds -ge $TimeoutSeconds) {
                try { $proc.Kill($true) } catch { }
                throw "LibreOffice did not finish within $TimeoutSeconds s and was stopped. Try a larger -PdfTimeoutSeconds, or convert manually: soffice --headless --convert-to pdf `"$DocxPath`""
            }
            Write-Host ("  still converting... {0:N0} s" -f $sw.Elapsed.TotalSeconds)
        }
        $produced = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($DocxPath) + '.pdf')
        if (-not (Test-Path -LiteralPath $produced)) { throw "LibreOffice did not produce a PDF (exit code $LASTEXITCODE)." }
        Move-Item -LiteralPath $produced -Destination $PdfPath -Force
    }
    finally {
        Remove-Item -LiteralPath $outDir, $loProfile -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Export-AsBuiltPdf {
    <# Returns the engine used. Engine Auto = Word if available, otherwise LibreOffice. #>
    param(
        [Parameter(Mandatory)][string]$DocxPath,
        [Parameter(Mandatory)][string]$PdfPath,
        [ValidateSet('Auto', 'Word', 'LibreOffice')][string]$Engine = 'Auto',
        [int]$TimeoutSeconds = 600
    )
    $DocxPath = (Resolve-Path -LiteralPath $DocxPath).Path
    $PdfPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($PdfPath)

    if ($Engine -in 'Auto', 'Word') {
        if ($IsWindows) {
            try { Convert-DocxToPdfWord -DocxPath $DocxPath -PdfPath $PdfPath; return 'Word' }
            catch {
                if ($Engine -eq 'Word') { throw "Word PDF export failed: $($_.Exception.Message)" }
                Write-Verbose "Word not available ($($_.Exception.Message)); trying LibreOffice"
            }
        }
        elseif ($Engine -eq 'Word') { throw 'The Word engine needs Windows with desktop Word installed.' }
    }

    $soffice = Find-Soffice
    if (-not $soffice) { throw 'No PDF engine found. Install Microsoft Word (Windows) or LibreOffice.' }
    Write-Host "  Converting with LibreOffice ($soffice) - large documents can take a few minutes"
    Convert-DocxToPdfLibreOffice -DocxPath $DocxPath -PdfPath $PdfPath -Soffice $soffice -TimeoutSeconds $TimeoutSeconds
    return 'LibreOffice'
}
