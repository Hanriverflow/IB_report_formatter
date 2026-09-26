<#
Render every page of generated DOCX files with Microsoft Word on Windows.
Uses Word's Page.EnhMetaFileBits, not PDF export or a desktop screenshot.
Run with Windows PowerShell 5.1. No default printer or add-in settings change.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputDirectory,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [int]$LongestEdge = 1754,
    [int]$ExpectedPages = 0
)

$ErrorActionPreference = 'Stop'
if ($LongestEdge -lt 600) { throw 'LongestEdge must be at least 600 pixels.' }
if ($ExpectedPages -lt 0) { throw 'ExpectedPages must be zero (unspecified) or positive.' }
Add-Type -AssemblyName System.Drawing
$inputRoot = (Resolve-Path -LiteralPath $InputDirectory).Path
$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
$updatedRoot = Join-Path $outputRoot 'updated-docx'
if ([string]::Equals($inputRoot.TrimEnd('\'), $updatedRoot.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'InputDirectory cannot be the updated-docx output directory.'
}
if (Test-Path -LiteralPath $outputRoot) {
    if (-not (Test-Path -LiteralPath $outputRoot -PathType Container) -or
        @(Get-ChildItem -LiteralPath $outputRoot -Force).Count -ne 0) {
        throw 'OutputDirectory must be new or empty; existing QA evidence will not be overwritten.'
    }
}
$null = New-Item -ItemType Directory -Path $outputRoot -Force
$files = @(Get-ChildItem -LiteralPath $inputRoot -Filter '*.docx' -File | Where-Object { -not $_.Name.StartsWith('~$') } | Sort-Object Name)
if ($files.Count -eq 0) { throw 'No DOCX files found.' }

function Get-SharedFileHash([string]$Path) {
    # Read without locking or closing a document already open in the user's Word.
    $stream = $algorithm = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $algorithm = [Security.Cryptography.SHA256]::Create()
        return [BitConverter]::ToString($algorithm.ComputeHash($stream)).Replace('-', '')
    }
    finally {
        if ($null -ne $algorithm) { $algorithm.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

$word = $null
$doc = $null
$results = @()
$failures = @()
[object]$noSave = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    foreach ($file in $files) {
        $startedAt = [DateTime]::UtcNow.ToString('o')
        $sourceHash = Get-SharedFileHash $file.FullName
        Write-Output ('Opening: ' + $file.Name)
        $doc = $word.Documents.Open($file.FullName, $false, $true, $false)
        $doc.ActiveWindow.View.Type = 3
        $null = $doc.Fields.Update()
        $doc.Repaginate()
        foreach ($toc in $doc.TablesOfContents) { $toc.UpdatePageNumbers() }
        # Word paginates lazily. Materialize all pages before refreshing footer totals.
        Write-Output ('Document pages: ' + $doc.ComputeStatistics(2) + '; sections: ' + $doc.Sections.Count)
        foreach ($section in $doc.Sections) {
            foreach ($header in $section.Headers) { if ($header.Exists) { $null = $header.Range.Fields.Update() } }
            foreach ($footer in $section.Footers) { if ($footer.Exists) { $null = $footer.Range.Fields.Update() } }
        }
        $doc.Repaginate()
        # A persisted QA-only copy clears Word's stale page/header rendering cache.
        $null = New-Item -ItemType Directory -Path $updatedRoot -Force
        $updatedPath = Join-Path $updatedRoot $file.Name
        $doc.SaveAs2([string]$updatedPath, 16)
        $doc.Close([ref]$noSave)
        $doc = $word.Documents.Open($updatedPath, $false, $true, $false)
        $doc.ActiveWindow.View.Type = 3
        $doc.Repaginate()
        $warmPages = $doc.ActiveWindow.Panes.Item(1).Pages
        # Repaginate alone does not refresh EMF's cached NUMPAGES value.
        for ($warmIndex = 1; $warmIndex -le $warmPages.Count; $warmIndex++) {
            $null = $warmPages.Item($warmIndex).EnhMetaFileBits
        }
        foreach ($section in $doc.Sections) {
            foreach ($footer in $section.Footers) { if ($footer.Exists) { $null = $footer.Range.Fields.Update() } }
        }
        Write-Output ('Fields updated: ' + $file.Name)
        $pages = $doc.ActiveWindow.Panes.Item(1).Pages
        $pageCount = $pages.Count
        $pageFiles = @()
        for ($index = 1; $index -le $pageCount; $index++) {
            $stream = $source = $bitmap = $graphics = $null
            try {
                $pageRange = $doc.GoTo(1, 1, $index)
                $pageRange.Select()
                $doc.ActiveWindow.ScrollIntoView($pageRange, $true)
                # Materialize repeated header artwork for this page before capture.
                $doc.ActiveWindow.View.SeekView = 9
                $doc.ActiveWindow.View.SeekView = 0
                $word.ScreenRefresh()
                $bits = [byte[]]$pages.Item($index).EnhMetaFileBits
                $stream = New-Object IO.MemoryStream(,$bits)
                $source = [Drawing.Image]::FromStream($stream)
                $scale = $LongestEdge / [Math]::Max($source.Width, $source.Height)
                $width = [int][Math]::Round($source.Width * $scale)
                $height = [int][Math]::Round($source.Height * $scale)
                $bitmap = New-Object Drawing.Bitmap($width, $height)
                $graphics = [Drawing.Graphics]::FromImage($bitmap)
                $graphics.Clear([Drawing.Color]::White)
                $graphics.DrawImage($source, 0, 0, $width, $height)
                $target = Join-Path $outputRoot ('{0}-page-{1:D3}.png' -f $file.BaseName, $index)
                $bitmap.Save($target, [Drawing.Imaging.ImageFormat]::Png)
                $pageFiles += @{ page = $index; path = $target; width = $width; height = $height; sha256 = (Get-SharedFileHash $target) }
                Write-Output ('Rendered: ' + [IO.Path]::GetFileName($target))
            }
            finally {
                foreach ($resource in @($graphics, $bitmap, $source, $stream)) {
                    if ($null -ne $resource) { $resource.Dispose() }
                }
            }
        }
        $doc.Close([ref]$noSave)
        $doc = $null
        $sourceHashAfter = Get-SharedFileHash $file.FullName
        $unchanged = $sourceHash -eq $sourceHashAfter
        $pageCountMatches = ($ExpectedPages -eq 0) -or ($ExpectedPages -eq $pageCount)
        $results += @{
            schemaVersion = 2
            source = $file.FullName; sourceSha256 = $sourceHash; sourceSha256After = $sourceHashAfter
            sourceUnchanged = $unchanged
            updatedDocument = $updatedPath; updatedSha256 = (Get-SharedFileHash $updatedPath)
            wordVersion = [string]$word.Version; wordBuild = [string]$word.Build
            startedAtUtc = $startedAt; completedAtUtc = [DateTime]::UtcNow.ToString('o')
            pageCount = $pageCount; expectedPages = $ExpectedPages; pageCountMatches = $pageCountMatches
            pages = $pageFiles; visualReview = 'pending'
        }
        # Persist completed evidence even when a later document fails.
        ConvertTo-Json -InputObject @($results) -Depth 6 | Set-Content -LiteralPath (Join-Path $outputRoot 'manifest.json') -Encoding UTF8
        if (-not $unchanged) { $failures += ('Source changed during QA: ' + $file.Name) }
        if (-not $pageCountMatches) { $failures += ('ExpectedPages mismatch: ' + $file.Name + ' expected ' + $ExpectedPages + ', actual ' + $pageCount) }
    }
    if ($failures.Count -gt 0) { throw ($failures -join '; ') }
}
finally {
    try {
        if ($null -ne $doc) { $doc.Close([ref]$noSave) }
    }
    finally {
        if ($null -ne $word) {
            $word.Quit([ref]$noSave)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
        }
    }
}
