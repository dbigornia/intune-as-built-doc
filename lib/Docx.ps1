# Minimal, dependency-free .docx writer for IntuneAsBuilt (WordprocessingML via System.IO.Compression).
# No Word install, PSWriteOffice or OpenXML SDK needed.

$script:DocxPartsPath   = Join-Path $PSScriptRoot 'docx-parts'
$script:InvalidXmlChars = [regex]'[^\u0009\u000A\u000D -퟿-�]'
$script:WNs = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
$script:DefaultSectPr = '<w:sectPr><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

#region Zip helpers
function Read-DocxZipEntry {
    param($Zip, [string]$Name)
    $e = $Zip.GetEntry($Name)
    if (-not $e) { return $null }
    $r = [System.IO.StreamReader]::new($e.Open())
    try { return $r.ReadToEnd() } finally { $r.Dispose() }
}

function Write-DocxZipEntry {
    param($Zip, [string]$Name, [string]$Content)
    if ($Zip.Mode -eq [System.IO.Compression.ZipArchiveMode]::Update) {
        $existing = $Zip.GetEntry($Name)
        if ($existing) { $existing.Delete() }
    }
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $w = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
    try { $w.Write($Content) } finally { $w.Dispose() }
}
#endregion

function New-AsBuiltDocument {
    <#
      Creates an in-memory document. If a template is given, the usable text width is read from its
      final section so table column widths fit the template's page size and margins.
    #>
    param([string]$TemplatePath)

    $textWidth = 10080   # Letter, 0.75" margins (default package)
    if ($TemplatePath) {
        $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $TemplatePath).Path)
        try {
            $xml = Read-DocxZipEntry -Zip $zip -Name 'word/document.xml'
            $sect = $xml.Substring([math]::Max(0, $xml.LastIndexOf('<w:sectPr')))
            $pgW = [regex]::Match($sect, '<w:pgSz[^>]*\sw:w="(\d+)"')
            $left = [regex]::Match($sect, '<w:pgMar[^>]*\sw:left="(\d+)"')
            $right = [regex]::Match($sect, '<w:pgMar[^>]*\sw:right="(\d+)"')
            if ($pgW.Success -and $left.Success -and $right.Success) {
                $textWidth = [int]$pgW.Groups[1].Value - [int]$left.Groups[1].Value - [int]$right.Groups[1].Value
            }
            $styles = Read-DocxZipEntry -Zip $zip -Name 'word/styles.xml'
            foreach ($s in 'Heading1', 'Heading2', 'Heading3', 'Title') {
                if ($styles -notmatch "w:styleId=`"$s`"") {
                    Write-Warning "Template has no '$s' style defined. Use it once in the template (then delete the text) so headings render with your styles."
                }
            }
        }
        finally { $zip.Dispose() }
    }
    [pscustomobject]@{
        Body      = [System.Text.StringBuilder]::new()
        TextWidth = $textWidth
    }
}

function ConvertTo-DocxText {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    [System.Security.SecurityElement]::Escape($script:InvalidXmlChars.Replace($Text, ''))
}

function Get-DocxRunXml {
    param(
        [AllowNull()][string]$Text,
        [switch]$Bold, [switch]$Italic, [switch]$Mono,
        [int]$Size = 0,
        [string]$Color
    )
    $rPr = ''
    if ($Mono)   { $rPr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>' }
    if ($Bold)   { $rPr += '<w:b/>' }
    if ($Italic) { $rPr += '<w:i/>' }
    if ($Color)  { $rPr += "<w:color w:val=`"$Color`"/>" }
    if ($Size -gt 0) { $rPr += "<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/>" }

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<w:r>')
    if ($rPr) { [void]$sb.Append("<w:rPr>$rPr</w:rPr>") }
    $lines = ([string]$Text -replace "`t", '    ') -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($i -gt 0) { [void]$sb.Append('<w:br/>') }
        [void]$sb.Append('<w:t xml:space="preserve">').Append((ConvertTo-DocxText $lines[$i])).Append('</w:t>')
    }
    [void]$sb.Append('</w:r>')
    $sb.ToString()
}

function Add-DocxParagraph {
    param(
        [Parameter(Mandatory)]$Doc,
        [AllowEmptyString()][string]$Text = '',
        [string]$Style,
        [switch]$Bold, [switch]$Italic, [switch]$Mono,
        [int]$Size = 0,
        [string]$Color,
        [switch]$KeepNext
    )
    $pPr = ''
    if ($Style)    { $pPr += "<w:pStyle w:val=`"$Style`"/>" }
    if ($KeepNext) { $pPr += '<w:keepNext/>' }
    $pPrXml = if ($pPr) { "<w:pPr>$pPr</w:pPr>" } else { '' }
    $run = Get-DocxRunXml -Text $Text -Bold:$Bold -Italic:$Italic -Mono:$Mono -Size $Size -Color $Color
    [void]$Doc.Body.Append("<w:p>$pPrXml$run</w:p>")
}

function Add-DocxHeading {
    param([Parameter(Mandatory)]$Doc, [Parameter(Mandatory)][string]$Text, [ValidateRange(1, 3)][int]$Level = 1)
    Add-DocxParagraph -Doc $Doc -Text $Text -Style "Heading$Level"
}

function Add-DocxPageBreak {
    param([Parameter(Mandatory)]$Doc)
    [void]$Doc.Body.Append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
}

function Add-DocxToc {
    <# TOC field; Word rebuilds it on open (it asks to update fields). #>
    param([Parameter(Mandatory)]$Doc, [string]$Levels = '1-2')
    [void]$Doc.Body.Append(
        '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
        "<w:r><w:instrText xml:space=`"preserve`"> TOC \o `"$Levels`" \h \z \u </w:instrText></w:r>" +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        '<w:r><w:t>Right-click here and choose Update Field to build the table of contents.</w:t></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
}

function Add-DocxTable {
    <#
      Rows: array of arrays (one inner array per row, one string per column).
      Widths: relative column widths, e.g. 0.3,0.7.
    #>
    param(
        [Parameter(Mandatory)]$Doc,
        [Parameter(Mandatory)][string[]]$Headers,
        [object[]]$Rows,
        [double[]]$Widths,
        [int]$FontSize = 18,
        [int]$MaxCellChars = 3000
    )
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $n = $Headers.Count
    if (-not $Widths -or $Widths.Count -ne $n) { $Widths = @(1.0) * $n }
    $total = ($Widths | Measure-Object -Sum).Sum
    $tw = [int]$Doc.TextWidth
    $cols = @($Widths | ForEach-Object { [int][math]::Floor($tw * $_ / $total) })

    $border = 'w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"'
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<w:tbl><w:tblPr><w:tblW w:w=`"$tw`" w:type=`"dxa`"/>")
    [void]$sb.Append("<w:tblBorders><w:top $border/><w:left $border/><w:bottom $border/><w:right $border/><w:insideH $border/><w:insideV $border/></w:tblBorders>")
    [void]$sb.Append('<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="29" w:type="dxa"/><w:left w:w="72" w:type="dxa"/><w:bottom w:w="29" w:type="dxa"/><w:right w:w="72" w:type="dxa"/></w:tblCellMar>')
    [void]$sb.Append('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>')
    foreach ($c in $cols) { [void]$sb.Append("<w:gridCol w:w=`"$c`"/>") }
    [void]$sb.Append('</w:tblGrid>')

    # Header row (repeats on each page)
    [void]$sb.Append('<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>')
    for ($i = 0; $i -lt $n; $i++) {
        [void]$sb.Append((Get-DocxCellXml -Text $Headers[$i] -Width $cols[$i] -Fill '1F3864' -Color 'FFFFFF' -Bold -FontSize $FontSize -MaxChars $MaxCellChars))
    }
    [void]$sb.Append('</w:tr>')

    $r = 0
    foreach ($row in $Rows) {
        $r++
        $fill = if ($r % 2 -eq 0) { 'F2F2F2' } else { '' }
        $cells = @($row)
        [void]$sb.Append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for ($i = 0; $i -lt $n; $i++) {
            $text = if ($i -lt $cells.Count -and $null -ne $cells[$i]) { [string]$cells[$i] } else { '' }
            [void]$sb.Append((Get-DocxCellXml -Text $text -Width $cols[$i] -Fill $fill -FontSize $FontSize -MaxChars $MaxCellChars))
        }
        [void]$sb.Append('</w:tr>')
    }
    [void]$sb.Append('</w:tbl>')
    [void]$Doc.Body.Append($sb.ToString())
    # Spacer so consecutive tables don't merge
    [void]$Doc.Body.Append('<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>')
}

function Get-DocxCellXml {
    param([string]$Text, [int]$Width, [string]$Fill, [string]$Color, [switch]$Bold, [int]$FontSize, [int]$MaxChars)
    if ($Text.Length -gt $MaxChars) { $Text = $Text.Substring(0, $MaxChars) + "`n... (truncated - see snapshot JSON for the full value)" }
    $shd = if ($Fill) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$Fill`"/>" } else { '' }
    $run = Get-DocxRunXml -Text $Text -Bold:$Bold -Size $FontSize -Color $Color
    "<w:tc><w:tcPr><w:tcW w:w=`"$Width`" w:type=`"dxa`"/>$shd</w:tcPr><w:p><w:pPr><w:spacing w:before=`"0`" w:after=`"0`" w:line=`"240`" w:lineRule=`"auto`"/></w:pPr>$run</w:p></w:tc>"
}

function Save-AsBuiltDocument {
    <#
      Writes the .docx. Without a template, a complete package is generated (styles, footer with page numbers).
      With a template (.docx or .dotx), the template is copied and its body replaced - headers, footers, logos,
      page setup and styles are kept. If the template contains a paragraph with the text {{ASBUILT}}, only
      that paragraph is replaced, so a cover page or intro in the template is preserved.
    #>
    param(
        [Parameter(Mandatory)]$Doc,
        [Parameter(Mandatory)][string]$Path,
        [string]$TemplatePath,
        [string]$Title = 'Intune As-Built',
        [string]$FooterText = 'Intune As-Built'
    )
    $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $body = $Doc.Body.ToString()

    if (-not $TemplatePath) {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $part = { param($f) Get-Content -Raw -LiteralPath (Join-Path $script:DocxPartsPath $f) }
            $created = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            Write-DocxZipEntry $zip '[Content_Types].xml'          (& $part 'content_types.xml')
            Write-DocxZipEntry $zip '_rels/.rels'                  (& $part 'rels.xml')
            Write-DocxZipEntry $zip 'word/_rels/document.xml.rels' (& $part 'document.xml.rels')
            Write-DocxZipEntry $zip 'word/styles.xml'              (& $part 'styles.xml')
            Write-DocxZipEntry $zip 'word/settings.xml'            (& $part 'settings.xml')
            Write-DocxZipEntry $zip 'word/footer1.xml'             ((& $part 'footer1.xml').Replace('{{FOOTER}}', (ConvertTo-DocxText $FooterText)))
            Write-DocxZipEntry $zip 'docProps/core.xml'            ((& $part 'core.xml').Replace('{{TITLE}}', (ConvertTo-DocxText $Title)).Replace('{{CREATED}}', $created))
            $docXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:document $script:WNs><w:body>$body$script:DefaultSectPr</w:body></w:document>"
            Write-DocxZipEntry $zip 'word/document.xml' $docXml
        }
        finally { $zip.Dispose(); $fs.Dispose() }
        return $Path
    }

    $ext = [System.IO.Path]::GetExtension($TemplatePath).ToLowerInvariant()
    if ($ext -notin '.docx', '.dotx') { throw "Template must be .docx or .dotx (macro-enabled templates are not supported): $TemplatePath" }
    Copy-Item -LiteralPath $TemplatePath -Destination $Path
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $docXml = Read-DocxZipEntry $zip 'word/document.xml'
        $bStart = $docXml.IndexOf('<w:body>')
        $bEnd = $docXml.LastIndexOf('</w:body>')
        if ($bStart -lt 0 -or $bEnd -lt 0) { throw 'Could not find <w:body> in the template document.' }
        $inner = $docXml.Substring($bStart + 8, $bEnd - $bStart - 8)
        $sIdx = $inner.LastIndexOf('<w:sectPr')
        $sectPr = if ($sIdx -ge 0) { $inner.Substring($sIdx) } else { '' }
        $content = if ($sIdx -ge 0) { $inner.Substring(0, $sIdx) } else { $inner }

        $newContent = $body
        foreach ($m in [regex]::Matches($content, '<w:p[ >].*?</w:p>', [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            $innerP = $m.Value.Substring(4)
            if ($innerP -match '<w:p[ >]') { continue }          # skip non-innermost matches (text boxes)
            $plain = [regex]::Replace($m.Value, '<[^>]+>', '')
            if ($plain.Contains('{{ASBUILT}}')) {
                $newContent = $content.Substring(0, $m.Index) + $body + $content.Substring($m.Index + $m.Length)
                break
            }
        }
        $docXml = $docXml.Substring(0, $bStart + 8) + $newContent + $sectPr + $docXml.Substring($bEnd)
        Write-DocxZipEntry $zip 'word/document.xml' $docXml

        if ($ext -eq '.dotx') {
            $ct = Read-DocxZipEntry $zip '[Content_Types].xml'
            Write-DocxZipEntry $zip '[Content_Types].xml' ($ct.Replace('wordprocessingml.template.main+xml', 'wordprocessingml.document.main+xml'))
        }
    }
    finally { $zip.Dispose() }
    return $Path
}
