<#
    Test-Common.ps1 - run folders, logging, input listing and the small
    helpers. Nothing here touches an OCR engine.
#>

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

. (Join-Path $here 'OcrTest.ps1')
. (Join-Path $root 'OcrCommon.ps1')

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ocrcommon-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temp -Force

try {
    Start-OcrSection 'Run folders'

    $run = New-OcrRunFolder -Parent $temp
    Assert-OcrTrue (Test-Path -LiteralPath $run.Path) 'the run folder is created'
    Assert-OcrTrue (Test-Path -LiteralPath $run.PagesPath) 'with a pages folder'
    Assert-OcrTrue (Test-Path -LiteralPath $run.OcrPath) 'with an ocr folder'
    Assert-OcrTrue (Test-Path -LiteralPath $run.LogPath) 'and a log file'

    # The name has to sort as text in the order the runs happened, or a folder
    # with a hundred runs in it is unusable.
    $name = Split-Path -Leaf $run.Path
    Assert-OcrTrue ($name -match '^\d{4}-\d{2}-\d{2}_\d{6}') ('the run folder is time-stamped: ' + $name)

    <#
        Two runs started inside the same second must not collide. Without the
        uniquifying suffix the second run would write its results into the
        first one's folder and quietly mix the two together.
    #>
    $second = New-OcrRunFolder -Parent $temp
    $third = New-OcrRunFolder -Parent $temp
    Assert-OcrTrue ($run.Path -ne $second.Path) 'a second run gets its own folder'
    Assert-OcrTrue ($second.Path -ne $third.Path) 'and so does a third'

    $labelled = New-OcrRunFolder -Parent $temp -Label 'Invoices Q3/2026'
    $labelledName = Split-Path -Leaf $labelled.Path
    Assert-OcrTrue ($labelledName -notmatch '[/\\:*?"<>|]') ('the label is made filesystem-safe: ' + $labelledName)
    Assert-OcrTrue ($labelledName -match 'Invoices-Q3-2026') 'and it is still recognisable'

    $longLabel = New-OcrRunFolder -Parent $temp -Label ('x' * 200)
    Assert-OcrTrue ((Split-Path -Leaf $longLabel.Path).Length -lt 80) 'a very long label is trimmed'

    Start-OcrSection 'Logging'

    Set-OcrLogPath $run.LogPath
    Write-OcrLog 'a plain message' -NoConsole
    Write-OcrLog 'something went wrong' -Level Bad -NoConsole
    $log = Get-Content -LiteralPath $run.LogPath -Raw
    Assert-OcrTrue ($log -match 'a plain message') 'the message reaches the log'
    Assert-OcrTrue ($log -match 'something went wrong') 'so does the failure'
    Assert-OcrTrue ($log -match 'Bad') 'and the level is recorded'

    # A log that cannot be written must never stop a run that is producing
    # real results.
    Set-OcrLogPath (Join-Path $temp 'no\such\folder\run.log')
    $threw = $false
    try { Write-OcrLog 'into the void' -NoConsole } catch { $threw = $true }
    Assert-OcrTrue (-not $threw) 'an unwritable log is survived, not fatal'
    Set-OcrLogPath $run.LogPath

    Start-OcrSection 'Listing input'

    $scans = Join-Path $temp 'scans'
    $null = New-Item -ItemType Directory -Path $scans -Force
    foreach ($name in @('a.png', 'b.JPG', 'c.pdf', 'd.tiff', 'notes.txt', 'sheet.xlsx')) {
        Set-Content -LiteralPath (Join-Path $scans $name) -Value 'x' -Encoding ASCII
    }

    $found = Get-OcrInputFiles -Path $scans
    Assert-OcrEqual $found.Count 4 'only the images and PDFs are picked up'
    # Extension matching must not care about case: scanners produce .JPG as
    # often as .jpg.
    Assert-OcrTrue (@($found | Where-Object { $_.Name -eq 'b.JPG' }).Count -eq 1) 'an upper-case extension is matched'
    Assert-OcrTrue (@($found | Where-Object { $_.Name -eq 'notes.txt' }).Count -eq 0) 'a text file is ignored'

    $single = Get-OcrInputFiles -Path (Join-Path $scans 'a.png')
    Assert-OcrTrue ($single -is [array]) 'a single file still comes back as an array'
    Assert-OcrEqual $single.Count 1 'with Count 1'

    $empty = Join-Path $temp 'empty'
    $null = New-Item -ItemType Directory -Path $empty -Force
    $none = Get-OcrInputFiles -Path $empty
    Assert-OcrTrue ($none -is [array]) 'an empty folder gives an array'
    Assert-OcrEqual $none.Count 0 'with Count 0, not $null'

    Assert-OcrThrows { Get-OcrInputFiles -Path (Join-Path $temp 'nowhere') } 'a missing path throws'

    # Recursion has to be asked for: silently walking a whole drive because
    # someone pointed at C:\ is not a helpful default.
    $nested = Join-Path $scans 'more'
    $null = New-Item -ItemType Directory -Path $nested -Force
    Set-Content -LiteralPath (Join-Path $nested 'deep.png') -Value 'x' -Encoding ASCII
    Assert-OcrEqual (Get-OcrInputFiles -Path $scans).Count 4 'subfolders are skipped by default'
    Assert-OcrEqual (Get-OcrInputFiles -Path $scans -Recurse).Count 5 '-Recurse reaches them'

    Start-OcrSection 'PDF detection'

    Assert-OcrTrue (Test-OcrIsPdf 'C:\a\b\invoice.pdf') 'a .pdf is a PDF'
    Assert-OcrTrue (Test-OcrIsPdf 'invoice.PDF') 'case does not matter'
    Assert-OcrTrue (-not (Test-OcrIsPdf 'scan.png')) 'a .png is not'

    Start-OcrSection 'Formatting'

    Assert-OcrEqual (Format-OcrSize 512) '512 B' 'bytes'
    Assert-OcrEqual (Format-OcrSize 2048) '2.0 KB' 'kilobytes'
    Assert-OcrEqual (Format-OcrSize 5242880) '5.0 MB' 'megabytes'
    Assert-OcrTrue ((Format-OcrSize 3221225472) -match 'GB') 'gigabytes'

    Assert-OcrEqual (Format-OcrDuration (New-TimeSpan -Seconds 5)) '5.0s' 'seconds'
    Assert-OcrTrue ((Format-OcrDuration (New-TimeSpan -Minutes 3)) -match '^3m') 'minutes'
    Assert-OcrTrue ((Format-OcrDuration (New-TimeSpan -Hours 2)) -match '^2h') 'hours'

    Start-OcrSection 'Paths'

    $unresolved = Resolve-OcrPath (Join-Path $temp 'not-created-yet.txt')
    Assert-OcrTrue ($unresolved -match 'not-created-yet') 'a path that does not exist yet still resolves'
    Assert-OcrTrue ([System.IO.Path]::IsPathRooted($unresolved)) 'and comes back absolute'

    Start-OcrSection 'Node detection'

    # Whether Node is present depends on the machine; what must always hold is
    # that the answer is a complete object rather than a crash.
    $node = Get-OcrNode
    foreach ($field in @('Path', 'Version', 'MajorVersion', 'Ok')) {
        Assert-OcrTrue ($null -ne $node.PSObject.Properties[$field]) ('Get-OcrNode reports ' + $field)
    }
    if ($node.Path) {
        Assert-OcrTrue ($node.MajorVersion -gt 0) 'a found Node has a parsed major version'
    }

    $kit = Test-OcrNodeKit -Root $root
    foreach ($field in @('Ok', 'JsPath', 'CliPath', 'Problems')) {
        Assert-OcrTrue ($null -ne $kit.PSObject.Properties[$field]) ('Test-OcrNodeKit reports ' + $field)
    }
    Assert-OcrTrue ($kit.Problems -is [array]) 'Problems is always an array'

} finally {
    Set-OcrLogPath $null
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

Complete-OcrSuite 'Test-Common'
