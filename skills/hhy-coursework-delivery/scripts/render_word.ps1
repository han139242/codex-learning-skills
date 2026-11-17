[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [switch]$UpdateFields,
    [switch]$CheckToc,
    [string]$PowerShellPath
)
$ErrorActionPreference = 'Stop'
# This machine's Word PDF export stalls under Windows PowerShell 5.1, while
# the same document exports normally under installed PowerShell 7. Reuse that
# tested backend without installing anything or changing the user's shell.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $backendPath = $null
    if ($PowerShellPath) { $backendPath = (Resolve-Path -LiteralPath $PowerShellPath).ProviderPath }
    else {
        $backend = Get-Command pwsh -ErrorAction SilentlyContinue
        if ($null -ne $backend) { $backendPath = $backend.Source }
    }
    if (-not $backendPath) {
        throw 'Pass -PowerShellPath with an existing PowerShell 7 runtime, or use another native renderer. Nothing has been installed.'
    }
    $forward = @('-NoProfile', '-NonInteractive', '-File', $PSCommandPath,
                 '-InputPath', $InputPath, '-OutputDirectory', $OutputDirectory)
    if ($UpdateFields) { $forward += '-UpdateFields' }
    if ($CheckToc) { $forward += '-CheckToc' }
    & $backendPath @forward
    if ($LASTEXITCODE -ne 0) { throw 'Word export backend failed; original input was not promoted.' }
    return
}
$inputFile = (Resolve-Path -LiteralPath $InputPath).ProviderPath
if ([IO.Path]::GetExtension($inputFile) -ne '.docx') { throw 'Only DOCX input is supported.' }
$outDir = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outDir) { throw 'Output directory already exists; choose a new temporary directory.' }
$parent = [IO.Path]::GetDirectoryName($outDir)
if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw 'Output parent directory must already exist.' }
$beforeHash = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash

# Reject external content before opening Word; ordinary clickable hyperlinks are not fetched.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [IO.Compression.ZipFile]::OpenRead($inputFile)
try {
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName -match 'vbaProject\.bin$') { throw 'Macro-bearing input requires separate review.' }
        if ($entry.FullName -eq 'docProps/core.xml') {
            $reader = New-Object IO.StreamReader($entry.Open())
            try { [xml]$core = $reader.ReadToEnd() } finally { $reader.Dispose() }
            foreach ($element in $core.SelectNodes('//*')) {
                $type = $element.GetAttribute('type', 'http://www.w3.org/2001/XMLSchema-instance')
                if ($type -match '^([^:]+):' -and -not $element.GetNamespaceOfPrefix($Matches[1])) {
                    throw 'Unbound metadata type namespace; repair the working copy before opening Word.'
                }
            }
        }
        if ($entry.FullName -like 'word/*.xml') {
            $reader = New-Object IO.StreamReader($entry.Open())
            try { [xml]$part = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $instructionNodes = $part.SelectNodes("//*[local-name()='instrText'] | //*[local-name()='fldSimple']/@*[local-name()='instr']")
            # PowerShell's XML adapter exposes an attribute's Value, not InnerText.
            $instructions = ($instructionNodes | ForEach-Object {
                if ($_.NodeType -eq [System.Xml.XmlNodeType]::Attribute) { $_.Value }
                else { $_.InnerText }
            }) -join ''
            if ($instructions -match '(?i)\b(DDEAUTO|DDE|INCLUDETEXT|INCLUDEPICTURE|LINK|DATABASE)\b') {
                throw 'External-content field requires separate review before Word export.'
            }
        }
        if ($entry.FullName -like '*.rels') {
            $reader = New-Object IO.StreamReader($entry.Open())
            try { [xml]$relationships = $reader.ReadToEnd() } finally { $reader.Dispose() }
            foreach ($rel in $relationships.DocumentElement.ChildNodes) {
                if ($rel.TargetMode -eq 'External' -and $rel.Type -notmatch '/hyperlink$') {
                    throw 'External content relationship requires separate review before Word export.'
                }
            }
        }
    }
} finally { $archive.Dispose() }

New-Item -ItemType Directory -Path $outDir | Out-Null
$workingFile = Join-Path $outDir 'review.docx'
$pdfFile = Join-Path $outDir 'review.pdf'
Copy-Item -LiteralPath $inputFile -Destination $workingFile
$word = $null
$document = $null
$pages = $null
$tocRows = New-Object 'System.Collections.Generic.List[object]'

function Convert-PageLabel([string]$Label) {
    if ($Label -match '^\d+$') { return [int]$Label }
    if ($Label -notmatch '^[ivxlcdm]+$') { throw "Custom page label needs manual verification: $Label" }
    $values = @{I=1;V=5;X=10;L=50;C=100;D=500;M=1000}
    $total = 0; $previous = 0
    $characters = $Label.ToUpperInvariant().ToCharArray()
    [array]::Reverse($characters)
    foreach ($character in $characters) {
        $value = $values[[string]$character]
        if ($value -lt $previous) { $total -= $value } else { $total += $value }
        $previous = $value
    }
    return $total
}
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($workingFile, $false, $false, $false)
    if ($UpdateFields) {
        # REF, SEQ, TOC, NUMPAGES, PAGE, PAGEREF, SECTIONPAGES only.
        # Never run an indiscriminate Fields.Update against external/include/DDE fields.
        $safeTypes = @(3, 12, 26, 33, 37, 66)
        # Snapshot fields before updates; handle the TOC separately below.
        $fieldSnapshot = New-Object 'System.Collections.Generic.List[object]'
        $fieldCount = $document.Fields.Count
        for ($i = 1; $i -le $fieldCount; $i++) {
            $field = $document.Fields.Item($i)
            if ($safeTypes -contains [int]$field.Type) { $fieldSnapshot.Add($field) }
        }
        foreach ($field in $fieldSnapshot) { [void]$field.Update() }
        for ($i = 1; $i -le $document.TablesOfContents.Count; $i++) {
            $document.TablesOfContents.Item($i).Update()
        }
    }
    $document.Repaginate()
    if ($UpdateFields) {
        for ($i = 1; $i -le $document.TablesOfContents.Count; $i++) {
            $document.TablesOfContents.Item($i).UpdatePageNumbers()
        }
        $document.Save()
    }
    $pages = $document.ComputeStatistics(2)
    $document.ExportAsFixedFormat($pdfFile, 17)
    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    $document = $null
    # Reopen the exact on-disk deliverable; a successful PDF alone is insufficient.
    $document = $word.Documents.Open($workingFile, $false, $true, $false)
    if ($document.ComputeStatistics(2) -ne $pages) { throw 'Page count changed after reopening the saved Word copy.' }
    if ($CheckToc) {
        if ($document.TablesOfContents.Count -eq 0) { throw 'No TOC to check; omit -CheckToc for documents without a TOC.' }
        for ($tocIndex = 1; $tocIndex -le $document.TablesOfContents.Count; $tocIndex++) {
            $entryCount = 0
            foreach ($paragraph in $document.TablesOfContents.Item($tocIndex).Range.Paragraphs) {
                $text = $paragraph.Range.Text.Trim()
                if (-not $text) { continue }
                if ($text -notmatch "`t([^`t]+)$") { throw "TOC page label missing; check the entry manually: $text" }
                $label = $Matches[1].Trim()
                $printedPage = Convert-PageLabel $label
                $anchor = $null
                if ($paragraph.Range.Hyperlinks.Count -gt 0) { $anchor = $paragraph.Range.Hyperlinks.Item(1).SubAddress }
                if (-not $anchor) {
                    foreach ($field in $paragraph.Range.Fields) {
                        if ($field.Type -eq 37 -and $field.Code.Text -match 'PAGEREF\s+"?([^\s"]+)') { $anchor = $Matches[1]; break }
                    }
                }
                if (-not $anchor -or -not $document.Bookmarks.Exists($anchor)) { throw "TOC target missing; check the entry manually: $text" }
                $target = $document.Bookmarks.Item($anchor).Range.Duplicate
                $target.Collapse(1)
                $actualPage = $target.Information(1)
                if ($printedPage -ne $actualPage) { throw "TOC page mismatch: $text; target=$actualPage" }
                $tocRows.Add([ordered]@{toc=$tocIndex;entry=$text;anchor=$anchor;page=$actualPage})
                $entryCount++
            }
            if ($entryCount -eq 0) { throw 'TOC has no verifiable entries.' }
        }
    }
} finally {
    if ($null -ne $document) {
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    if ($null -ne $word) {
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ((Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash -ne $beforeHash) {
    throw 'Input changed while exporting; do not promote this result.'
}
if (-not (Test-Path -LiteralPath $pdfFile) -or (Get-Item -LiteralPath $pdfFile).Length -eq 0) {
    throw 'No PDF was produced.'
}
$manifest = [ordered]@{
    input = $inputFile
    input_sha256 = $beforeHash
    working_copy = $workingFile
    working_sha256 = (Get-FileHash -LiteralPath $workingFile -Algorithm SHA256).Hash
    pdf = $pdfFile
    pages = $pages
    local_fields_updated = [bool]$UpdateFields
    word_reopened = $true
    toc_checked = [bool]$CheckToc
    toc_entries_checked = $tocRows.Count
    toc_entries = @($tocRows.ToArray())
    input_unchanged = $true
    visual_reviewed = $false
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $outDir 'render.json') -Encoding UTF8
$manifest | ConvertTo-Json -Depth 4 -Compress
