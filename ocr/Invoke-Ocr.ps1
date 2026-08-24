<#
.SYNOPSIS
    Read scanned pages and write out every word with its position on the page.

.DESCRIPTION
    Point it at a folder of scans. It renders anything that is not already a
    page image, prepares each page, recognises it, works out the lines and
    columns, and writes one .ocr.xml per page into a new run folder.

    The result is not a wall of text. It is every word with its box, grouped
    into lines and columns, in points from the top-left of the page - the same
    frame and the same shape the PDF text extractor produces, so a scanned
    page and a text page can be read by the same code afterwards.

    Nothing is downloaded at run time. Nothing outside the run folder is
    written to. The originals are never modified.

.PARAMETER Path
    A file or a folder of scans. Defaults to the _Scans folder beside this
    script.

.PARAMETER RunFolder
    Where run folders are created. Defaults to _OcrRuns beside this script.

.PARAMETER Engine
    tesseract  the bundled engine (default, most accurate on documents)
    windows    the OCR engine built into Windows 10+ (no install, no network)
    both       run both and report where they disagree

.PARAMETER Language
    Recognition language. 'eng' for Tesseract, a tag like 'en-GB' for Windows.

.PARAMETER PageSegmentation
    How the engine should carve up the page. 3 = automatic (default),
    4 = a single column of text, 6 = one uniform block, 11 = sparse text.

.PARAMETER SourceDpi
    What resolution the images actually are. Read from the file when it says;
    only set this when the file is wrong or silent.

.PARAMETER TargetDpi
    What to scale pages to before recognition. 300 is the sweet spot.

.PARAMETER NoPreprocess
    Hand images to the engine exactly as they are.

.EXAMPLE
    .\Invoke-Ocr.ps1

    Reads everything in _Scans and writes a new run folder under _OcrRuns.

.EXAMPLE
    .\Invoke-Ocr.ps1 -Path .\Invoices -Layout -Csv

    Also writes a plain-text picture of each page and a CSV of every word.

.EXAMPLE
    .\Invoke-Ocr.ps1 -Path .\fax.png -SourceDpi 204 -TargetDpi 400

    A fax is 204 dpi. Saying so lets it be scaled by the right amount.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $Path,

    [Parameter(Mandatory = $false)]
    [string] $RunFolder,

    [Parameter(Mandatory = $false)]
    [ValidateSet('tesseract', 'windows', 'both')]
    [string] $Engine = 'tesseract',

    [Parameter(Mandatory = $false)]
    [string] $Language,

    [Parameter(Mandatory = $false)]
    [ValidateSet('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13')]
    [string] $PageSegmentation = '3',

    [Parameter(Mandatory = $false)]
    [int] $SourceDpi = 0,

    [Parameter(Mandatory = $false)]
    [int] $TargetDpi = 300,

    [Parameter(Mandatory = $false)]
    [ValidateSet('sauvola', 'otsu')]
    [string] $Threshold = 'sauvola',

    [Parameter(Mandatory = $false)]
    [int] $Workers = 0,

    [Parameter(Mandatory = $false)]
    [int] $RefineBelow = 70,

    [Parameter(Mandatory = $false)]
    [switch] $NoPreprocess,

    [Parameter(Mandatory = $false)]
    [switch] $NoDeskew,

    [Parameter(Mandatory = $false)]
    [switch] $NoRefine,

    [Parameter(Mandatory = $false)]
    [switch] $Crop,

    [Parameter(Mandatory = $false)]
    [switch] $Layout,

    [Parameter(Mandatory = $false)]
    [switch] $Csv,

    [Parameter(Mandatory = $false)]
    [switch] $KeepImages,

    [Parameter(Mandatory = $false)]
    [switch] $Recurse,

    [Parameter(Mandatory = $false)]
    [switch] $NoInvoice,

    [Parameter(Mandatory = $false)]
    [switch] $NonInteractive,

    [Parameter(Mandatory = $false)]
    [switch] $Install
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Dot-sourced rather than imported as a module so the whole thing stays a
# folder of files that can be copied anywhere and run - no manifest, no
# install, nothing to register.
$here = $PSScriptRoot
if (-not $here) { $here = Split-Path -Parent $MyInvocation.MyCommand.Path }

. (Join-Path $here 'OcrCommon.ps1')
. (Join-Path $here 'OcrDocument.ps1')
. (Join-Path $here 'OcrRaster.ps1')
. (Join-Path $here 'OcrWindows.ps1')
. (Join-Path $here 'OcrNode.ps1')

# --------------------------------------------------------------------- setup

Write-OcrBanner 'OCR - scanned pages to positioned words'

if ($Install) {
    $null = Install-OcrNodeKit -Root $here
    Write-Host ''
    return
}

<#
    Ask for the folder.

    Running the script with no arguments should not be an error - it should
    ask. The one thing a person always has to hand is the folder their PDFs
    are in, and the natural way to give it is to paste the path or drag the
    folder onto the window.

    Both of those arrive slightly mangled, and cleaning them up is the whole
    job of Get-OcrFolderFromUser:

      - Windows Explorer's "Copy as path" wraps the path in double quotes;
      - dragging a folder onto a console window wraps it in quotes too, but
        only when it contains a space;
      - either can arrive with trailing whitespace or a trailing backslash.

    Left alone, every one of those produces "path not found" for a path that
    is plainly right there on screen.
#>
function Get-OcrFolderFromUser {
    [CmdletBinding()]
    param()

    for ($attempt = 0; $attempt -lt 5; $attempt++) {
        Write-Host ''
        Write-Host '  Paste the folder holding the PDFs to read' -ForegroundColor White
        Write-Host '  (drag the folder onto this window, or paste the path, then press Enter)' -ForegroundColor DarkGray
        Write-Host '  Press Enter on its own to quit.' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  folder> ' -ForegroundColor Cyan -NoNewline

        $entered = Read-Host

        if ([string]::IsNullOrWhiteSpace($entered)) { return $null }

        $cleaned = $entered.Trim()
        # Strip a matched pair of quotes of either kind.
        if ($cleaned.Length -ge 2) {
            $first = $cleaned.Substring(0, 1)
            $last = $cleaned.Substring($cleaned.Length - 1, 1)
            if (($first -eq '"' -and $last -eq '"') -or ($first -eq "'" -and $last -eq "'")) {
                $cleaned = $cleaned.Substring(1, $cleaned.Length - 2)
            }
        }
        $cleaned = $cleaned.Trim()
        # A trailing separator is harmless to Explorer and fatal to Test-Path
        # when the path is a drive-relative one.
        if ($cleaned.Length -gt 3) { $cleaned = $cleaned.TrimEnd('\', '/') }

        if ([string]::IsNullOrWhiteSpace($cleaned)) { continue }

        if (-not (Test-Path -LiteralPath $cleaned)) {
            Write-Host ''
            Write-OcrLog ('That path does not exist: ' + $cleaned) -Level Bad
            Write-OcrLog 'Check it and try again.' -Level Detail
            continue
        }

        $files = @()
        try {
            $files = Get-OcrInputFiles -Path $cleaned
        } catch {
            Write-OcrLog $_.Exception.Message -Level Bad
            continue
        }

        if ($files.Count -eq 0) {
            Write-Host ''
            Write-OcrLog 'There are no PDFs or page images in that folder.' -Level Warn
            Write-OcrLog 'Accepted: .pdf .png .jpg .jpeg .tif .tiff .bmp .gif .webp' -Level Detail
            continue
        }

        # Show what was found before spending minutes on it.
        $pdfCount = @($files | Where-Object { Test-OcrIsPdf $_.FullName }).Count
        $imageCount = $files.Count - $pdfCount
        $bytes = ($files | Measure-Object -Property Length -Sum).Sum

        Write-Host ''
        Write-OcrLog ('Found ' + $files.Count + ' file(s), ' + (Format-OcrSize $bytes) + ':') -Level Good
        if ($pdfCount -gt 0) { Write-OcrLog ($pdfCount.ToString() + ' PDF(s)') -Level Detail }
        if ($imageCount -gt 0) { Write-OcrLog ($imageCount.ToString() + ' page image(s)') -Level Detail }
        foreach ($file in ($files | Select-Object -First 8)) {
            Write-OcrLog ('  ' + $file.Name) -Level Detail
        }
        if ($files.Count -gt 8) {
            Write-OcrLog ('  ... and ' + ($files.Count - 8) + ' more') -Level Detail
        }

        Write-Host ''
        Write-Host '  Read these? [Y/n] ' -ForegroundColor Cyan -NoNewline
        $answer = Read-Host
        if ([string]::IsNullOrWhiteSpace($answer) -or $answer.Trim().ToLowerInvariant().StartsWith('y')) {
            return $cleaned
        }
        # Anything else means "no, let me pick a different folder".
    }

    return $null
}

if (-not $Path) {
    if ($NonInteractive) {
        throw 'No -Path was given and -NonInteractive was set, so there is nothing to read.'
    }
    $Path = Get-OcrFolderFromUser
    if (-not $Path) {
        Write-Host ''
        Write-OcrLog 'Nothing to do.' -Level Detail
        Write-Host ''
        return
    }
}

if (-not (Test-Path -LiteralPath $Path)) {
    throw ('Not found: ' + $Path)
}

$run = New-OcrRunFolder -Parent $RunFolder -Label ([System.IO.Path]::GetFileName($Path))
Set-OcrLogPath $run.LogPath

Write-OcrLog ('Reading   ' + (Resolve-OcrPath $Path))
Write-OcrLog ('Writing   ' + $run.Path)
Write-OcrLog ('Engine    ' + $Engine)

$files = Get-OcrInputFiles -Path $Path -Recurse:$Recurse
if ($files.Count -eq 0) {
    Write-OcrLog 'Nothing to read - no images or PDFs found.' -Level Warn
    Write-Host ''
    return
}

$totalBytes = ($files | Measure-Object -Property Length -Sum).Sum
Write-OcrLog ('Found     ' + $files.Count + ' file(s), ' + (Format-OcrSize $totalBytes))

# ------------------------------------------------------- documents to images

<#
    Staging: get everything into one folder of things the recogniser can open.

    THE ORDER MATTERS, and it is the opposite of the obvious one.

    A scanned PDF is a wrapper around one photograph per page. The Node side
    can open it and lift that photograph straight out - the scanner's own
    pixels, at the scanner's own resolution, with nothing resampled. Rendering
    the PDF instead means decoding the photograph, painting it at some chosen
    size and re-encoding it, and every one of those steps softens the letter
    edges a little.

    So PDFs are handed to Node FIRST, and the Windows renderer is kept as the
    fallback for the ones it cannot lift: fax-compressed (CCITT) and JBIG2
    scans, and PDFs that are drawn rather than scanned. That way the common
    case gets the best possible pixels and the awkward case still works.

    TIFF is the exception - Tesseract.js genuinely cannot read it - so a TIFF
    is split into PNG frames up front, on Windows.
#>
$tiffs = @($files | Where-Object {
    $_.Extension.ToLowerInvariant() -eq '.tif' -or $_.Extension.ToLowerInvariant() -eq '.tiff'
})
$direct = @($files | Where-Object {
    $_.Extension.ToLowerInvariant() -ne '.tif' -and $_.Extension.ToLowerInvariant() -ne '.tiff'
})

$pageCount = 0

foreach ($tiff in $tiffs) {
    try {
        Write-OcrLog ('Splitting ' + $tiff.Name) -Level Step
        $rendered = Convert-OcrTiffToImages -Path $tiff.FullName -DestinationFolder $run.PagesPath
        $pageCount = $pageCount + $rendered.Count
        Write-OcrLog ($rendered.Count.ToString() + ' page(s)') -Level Good
    } catch {
        Write-OcrLog ($tiff.Name + ': ' + $_.Exception.Message) -Level Bad
    }
}

foreach ($item in $direct) {
    # Copied, never moved: the originals are not ours to disturb.
    Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $run.PagesPath $item.Name) -Force
    $pageCount = $pageCount + 1
}

if ($pageCount -eq 0) {
    Write-OcrLog 'Nothing could be staged, so there is nothing to recognise.' -Level Bad
    Write-Host ''
    return
}

Write-OcrLog ($pageCount.ToString() + ' page image(s) ready')

# ------------------------------------------------------------- recognition

$started = Get-Date
$summary = $null

if ($Engine -eq 'tesseract' -or $Engine -eq 'both') {
    Write-OcrBanner 'Recognising (tesseract.js)'

    $language = $Language
    if (-not $language) { $language = 'eng' }

    $summary = Invoke-OcrNode `
        -InputPath $run.PagesPath `
        -OutputPath $run.Path `
        -Root $here `
        -Language $language `
        -PageSegmentation $PageSegmentation `
        -SourceDpi $SourceDpi `
        -TargetDpi $TargetDpi `
        -Threshold $Threshold `
        -Workers $Workers `
        -RefineBelow $RefineBelow `
        -NoPreprocess:$NoPreprocess `
        -NoDeskew:$NoDeskew `
        -NoRefine:$NoRefine `
        -Crop:$Crop `
        -NoInvoice:$NoInvoice `
        -Layout:$Layout `
        -KeepImages:$KeepImages `
        -Force
}

<#
    The fallback.

    Node reports, per file, whether it got anything. A PDF it could not lift
    an image out of is either drawn rather than scanned - in which case it has
    a text layer and needs no OCR at all - or compressed with CCITT or JBIG2,
    which the Windows renderer handles and Node does not. Either way the
    answer is to render it and try again, and only PDFs that actually failed
    are put through that.
#>
if ($summary -and $Engine -ne 'windows') {
    $stuck = @($summary.Pages | Where-Object {
        (-not $_.ok) -and ([System.IO.Path]::GetExtension([string]$_.file).ToLowerInvariant() -eq '.pdf')
    })

    if ($stuck.Count -gt 0) {
        $support = Test-OcrRasterSupport
        if (-not $support.Pdf) {
            Write-Host ''
            Write-OcrLog ($stuck.Count.ToString() + ' PDF(s) could not be read directly, and this machine cannot render PDFs.') -Level Warn
            Write-OcrLog $support.Note -Level Detail
            foreach ($item in $stuck) {
                Write-OcrLog ([System.IO.Path]::GetFileName([string]$item.file) + ': ' + $item.note) -Level Detail
            }
        } else {
            Write-OcrBanner ('Rendering ' + $stuck.Count + ' PDF(s) that could not be read directly')

            $retryFolder = Join-Path $run.Path 'rendered'
            $null = New-Item -ItemType Directory -Path $retryFolder -Force
            $renderedAny = 0

            foreach ($item in $stuck) {
                $pdfPath = [string]$item.file
                try {
                    Write-OcrLog ([System.IO.Path]::GetFileName($pdfPath)) -Level Step
                    $rendered = Convert-OcrPdfToImages -Path $pdfPath -DestinationFolder $retryFolder -Dpi $TargetDpi
                    $renderedAny = $renderedAny + $rendered.Count
                    Write-OcrLog ($rendered.Count.ToString() + ' page(s) at ' + $TargetDpi + ' dpi') -Level Good
                } catch {
                    Write-OcrLog ([System.IO.Path]::GetFileName($pdfPath) + ': ' + $_.Exception.Message) -Level Bad
                }
            }

            if ($renderedAny -gt 0) {
                $retry = Invoke-OcrNode `
                    -InputPath $retryFolder `
                    -OutputPath $run.Path `
                    -Root $here `
                    -Language $(if ($Language) { $Language } else { 'eng' }) `
                    -PageSegmentation $PageSegmentation `
                    -SourceDpi $TargetDpi `
                    -TargetDpi $TargetDpi `
                    -Threshold $Threshold `
                    -Workers $Workers `
                    -RefineBelow $RefineBelow `
                    -NoPreprocess:$NoPreprocess `
                    -NoDeskew:$NoDeskew `
                    -NoRefine:$NoRefine `
                    -NoInvoice:$NoInvoice `
                    -Crop:$Crop `
                    -Layout:$Layout `
                    -KeepImages:$KeepImages `
                    -Force
                if (-not $retry.Ok) {
                    Write-OcrLog 'Some rendered pages still could not be read.' -Level Warn
                }
            }
        }
    }
}

if ($Engine -eq 'windows' -or $Engine -eq 'both') {
    Write-OcrBanner 'Recognising (Windows OCR)'
    $check = Test-OcrWindowsEngine
    if (-not $check.Available) {
        Write-OcrLog $check.Note -Level Warn
    } else {
        $target = $run.Path
        if ($Engine -eq 'both') {
            # Two engines cannot write the same filenames into the same folder.
            $target = Join-Path $run.Path 'windows'
            $null = New-Item -ItemType Directory -Path $target -Force
        }
        $null = Invoke-OcrWindows `
            -InputPath $run.PagesPath `
            -OutputPath $target `
            -Root $here `
            -Language $Language `
            -Dpi $TargetDpi `
            -Layout:$Layout
    }
}

$elapsed = (Get-Date) - $started

# ------------------------------------------------------------------ results

Write-OcrBanner 'Results'

$documents = @(Get-ChildItem -LiteralPath $run.OcrPath -Filter '*.ocr.xml' -ErrorAction SilentlyContinue |
    Sort-Object Name)

if ($documents.Count -eq 0) {
    Write-OcrLog 'No documents were produced.' -Level Bad
    if ($summary -and $summary.Note) { Write-OcrLog $summary.Note -Level Detail }
    Write-Host ''
    return
}

$totalWords = 0
$confidences = New-Object System.Collections.ArrayList
$lowConfidence = New-Object System.Collections.ArrayList

foreach ($document in $documents) {
    $parsed = Import-OcrDocument $document.FullName
    foreach ($page in $parsed.Pages) {
        $totalWords = $totalWords + $page.WordCount
        $null = $confidences.Add($page.Confidence)

        $tables = @($page.Regions | Where-Object { $_.Kind -eq 'table' })
        $shape = $page.Regions.Count.ToString() + ' region(s)'
        if ($tables.Count -gt 0) {
            $shape = $shape + ', ' + $tables.Count + ' table(s)'
        }

        Write-OcrLog ($document.Name.PadRight(34) +
                      $page.WordCount.ToString().PadLeft(5) + ' words   ' +
                      $page.Confidence.ToString('0.0').PadLeft(5) + '%   ' + $shape)

        # A page nobody looks at is a page nobody checks. Naming the weak ones
        # is the difference between a review queue and a pile.
        if ($page.Confidence -lt 75) {
            $null = $lowConfidence.Add($document.Name)
        }
    }

    if ($Csv) {
        # GetFileNameWithoutExtension strips only the LAST extension, so
        # "statement.ocr.xml" comes back as "statement.ocr" and the CSV would
        # be named "statement.ocr.words.csv". Take off the .ocr as well.
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($document.Name)
        if ($stem.ToLowerInvariant().EndsWith('.ocr')) {
            $stem = $stem.Substring(0, $stem.Length - 4)
        }
        $csvPath = Join-Path $run.OcrPath ($stem + '.words.csv')
        $null = Export-OcrWords -Document $parsed -Path $csvPath
    }
}

$meanConfidence = 0
if ($confidences.Count -gt 0) {
    $meanConfidence = ($confidences | Measure-Object -Average).Average
}

Write-Host ''
Write-OcrLog ($documents.Count.ToString() + ' page(s), ' + $totalWords + ' words, mean confidence ' +
              $meanConfidence.ToString('0.0') + '%, in ' + (Format-OcrDuration $elapsed)) -Level Good

if ($lowConfidence.Count -gt 0) {
    Write-Host ''
    Write-OcrLog ($lowConfidence.Count.ToString() + ' page(s) came back under 75% and are worth a look:') -Level Warn
    foreach ($name in ($lowConfidence | Select-Object -Unique)) {
        Write-OcrLog $name -Level Detail
    }
}

Write-Host ''
Write-OcrLog ('Results   ' + $run.OcrPath)
Write-OcrLog ('Log       ' + $run.LogPath)
Write-Host ''
Write-Host '  To look at a page the way the parser sees it:' -ForegroundColor DarkGray
Write-Host ('    . .\OcrDocument.ps1; Show-OcrPage (Import-OcrDocument "' +
            (Join-Path $run.OcrPath ($documents[0].Name)) + '")') -ForegroundColor DarkGray
Write-Host ''
