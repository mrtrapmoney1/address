<#
    Test-EndToEnd.ps1 - the whole pipeline, for real.

    Runs Invoke-Ocr.ps1 over the sample pages and checks what comes out the
    other end. This is the suite that fails when the two halves stop agreeing
    with each other - which is the failure mode that unit tests, by
    construction, cannot see.

    If Node or the OCR engine is not installed the suite says so and skips,
    because that is a setup step rather than a defect.
#>

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

. (Join-Path $here 'OcrTest.ps1')
. (Join-Path $root 'OcrCommon.ps1')
. (Join-Path $root 'OcrDocument.ps1')

$node = Get-OcrNode
$kit = Test-OcrNodeKit -Root $root

if (-not $node.Ok -or -not $kit.Ok) {
    Write-Host ''
    Write-Host '  Test-EndToEnd: SKIPPED' -ForegroundColor Yellow
    if (-not $node.Ok) {
        Write-Host '    Node.js 18 or newer is not available.' -ForegroundColor DarkGray
    }
    foreach ($problem in $kit.Problems) {
        Write-Host ('    ' + $problem) -ForegroundColor DarkGray
    }
    Write-Host '    Run  .\Invoke-Ocr.ps1 -Install  to set it up.' -ForegroundColor DarkGray
    Write-Host ''
    exit 0
}

$samples = Join-Path $root 'samples'
$clean = Join-Path $samples 'page_clean.png'
$scan = Join-Path $samples 'page_scan.png'

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ocre2e-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temp -Force

try {
    Start-OcrSection 'The sample pages are present'
    Assert-OcrTrue (Test-Path -LiteralPath $clean) 'page_clean.png'
    Assert-OcrTrue (Test-Path -LiteralPath $scan) 'page_scan.png'

    Start-OcrSection 'A full run over the samples'

    $runs = Join-Path $temp 'runs'
    & (Join-Path $root 'Invoke-Ocr.ps1') -Path $samples -RunFolder $runs -Csv -Layout | Out-Null

    $runFolders = @(Get-ChildItem -LiteralPath $runs -Directory)
    Assert-OcrEqual $runFolders.Count 1 'exactly one run folder was created'

    $run = $runFolders[0].FullName
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $run 'run.log')) 'the run wrote a log'
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $run 'ocr')) 'and an ocr folder'

    $documents = @(Get-ChildItem -LiteralPath (Join-Path $run 'ocr') -Filter '*.ocr.xml')
    Assert-OcrEqual $documents.Count 2 'one document per sample page'

    <#
        The originals must be untouched. A tool that modifies the scans it was
        given is a tool nobody can safely run twice.
    #>
    Assert-OcrTrue (Test-Path -LiteralPath $clean) 'the original clean page is still there'
    Assert-OcrTrue (Test-Path -LiteralPath $scan) 'the original scan is still there'

    Start-OcrSection 'What the clean page produced'

    $cleanDoc = Import-OcrDocument (Join-Path $run 'ocr\page_clean.ocr.xml')
    $cleanPage = $cleanDoc.Pages[0]

    Assert-OcrTrue ($cleanPage.WordCount -gt 40) ('a page of words, got ' + $cleanPage.WordCount)
    Assert-OcrTrue ($cleanPage.Confidence -gt 80) ('high confidence, got ' + $cleanPage.Confidence)
    Assert-OcrTrue ($cleanPage.Text -match '12 point text') 'the text is recognisably right'
    Assert-OcrTrue ($cleanPage.Preprocess -match 'sauvola') ('the preprocessing history is recorded: ' + $cleanPage.Preprocess)

    Start-OcrSection 'Every word has a usable position'

    $words = Get-OcrWords -Document $cleanDoc
    Assert-OcrTrue ($words.Count -gt 40) 'the word stream is populated'

    $bad = 0
    foreach ($word in $words) {
        if ($word.W -le 0 -or $word.H -le 0) { $bad++ }
        if ($word.X -lt 0 -or $word.Y -lt 0) { $bad++ }
        # The page is 640x480 at 96 dpi, so 480x360 points. Nothing may sit
        # outside it - a word off the page means the coordinate conversion is
        # wrong, and everything downstream would inherit the error.
        if ($word.X -gt $cleanPage.Width + 1) { $bad++ }
        if ($word.Y -gt $cleanPage.Height + 1) { $bad++ }
    }
    Assert-OcrEqual $bad 0 'no word has an impossible position or size'

    Assert-OcrTrue ($cleanPage.Width -gt 0 -and $cleanPage.Height -gt 0) 'the page has a size in points'

    Start-OcrSection 'The tilted scan was straightened'

    $scanDoc = Import-OcrDocument (Join-Path $run 'ocr\page_scan.ocr.xml')
    $scanPage = $scanDoc.Pages[0]

    # The fixture is tilted by 1.8 degrees on purpose.
    Assert-OcrTrue ([Math]::Abs($scanPage.Skew) -gt 1) ('the tilt was detected, got ' + $scanPage.Skew)
    Assert-OcrTrue ($scanPage.Preprocess -match 'deskew') ('and corrected: ' + $scanPage.Preprocess)
    Assert-OcrTrue ($scanPage.WordCount -gt 30) ('the bad scan still read, got ' + $scanPage.WordCount + ' words')

    Start-OcrSection 'The extra outputs'

    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $run 'ocr\page_clean.words.csv')) 'the CSV is written'
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $run 'ocr\page_clean.layout.txt')) 'the layout view is written'

    $csv = Import-Csv -LiteralPath (Join-Path $run 'ocr\page_clean.words.csv')
    Assert-OcrEqual $csv.Count $words.Count 'the CSV has one row per word'

    $layout = Get-Content -LiteralPath (Join-Path $run 'ocr\page_clean.layout.txt') -Raw
    Assert-OcrTrue ($layout -match '12\s+point\s+text') 'the layout view shows the text where it sat'

    Start-OcrSection 'A second run does not disturb the first'

    & (Join-Path $root 'Invoke-Ocr.ps1') -Path $samples -RunFolder $runs | Out-Null
    $after = @(Get-ChildItem -LiteralPath $runs -Directory)
    Assert-OcrEqual $after.Count 2 'the second run got its own folder'
    Assert-OcrTrue (Test-Path -LiteralPath (Join-Path $run 'ocr\page_clean.words.csv')) 'the first run is intact'

    Start-OcrSection 'An empty input folder is handled'

    $emptyIn = Join-Path $temp 'empty'
    $null = New-Item -ItemType Directory -Path $emptyIn -Force
    $threw = $false
    try {
        & (Join-Path $root 'Invoke-Ocr.ps1') -Path $emptyIn -RunFolder (Join-Path $temp 'runs2') | Out-Null
    } catch {
        $threw = $true
    }
    Assert-OcrTrue (-not $threw) 'nothing to do is reported, not thrown'

} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

Complete-OcrSuite 'Test-EndToEnd'
