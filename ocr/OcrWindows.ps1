<#
    OcrWindows.ps1 - the second engine, built into Windows.

    Windows 10 and later ship an OCR engine in Windows.Media.Ocr. It is
    already on the machine, it needs no install and no network, and it is
    fast. That makes it worth having even though Tesseract is usually more
    accurate on documents: it is the engine that works on a locked-down
    machine at nine in the morning when npm install is not an option, and it
    is a genuinely independent second opinion when accuracy matters.

    The approach to driving WinRT from Windows PowerShell 5.1 - loading the
    projected types and blocking on IAsyncOperation through
    WindowsRuntimeSystemExtensions.GetAwaiter - follows JoelGMSec's
    Invoke-Transfer, which is the clearest published example of it.

    WHAT IS DIFFERENT HERE, AND WHY IT MATTERS

    Invoke-Transfer reads a screenshot and takes OcrResult.Text - the whole
    page as one string. For its purpose (recovering a base64 blob from a
    screenshot) that is exactly right.

    For reading a document it throws away the most valuable thing the engine
    produced. OcrResult also carries Lines, each with Words, each with a
    BoundingRect - so the engine knows where every word sat, and .Text
    discards it. Once "4987 36.07 99,914.15" is one string, nothing can
    recover which column the 36.07 came from.

    So this walks the Lines/Words tree and keeps every box. The words then go
    to the same segmentation code the Tesseract path uses, and both engines
    produce byte-identical document structure.
#>

Set-StrictMode -Version 2.0

<#
    Is the Windows OCR engine usable, and for which languages?

    TryCreateFromUserProfileLanguages returns null when no OCR language pack
    is installed for any of the user's languages. That is a real and common
    state on a stripped server build, and it must be reported as "install a
    language pack", not as a null-reference error three calls later.
#>
function Test-OcrWindowsEngine {
    [CmdletBinding()]
    param()

    if (-not (Test-OcrOnWindows)) {
        return [pscustomobject]@{
            Available = $false
            Languages = @()
            Note      = 'The Windows OCR engine needs Windows 10 or later.'
        }
    }

    if (-not (Initialize-OcrWinRt)) {
        return [pscustomobject]@{
            Available = $false
            Languages = @()
            Note      = 'The Windows runtime types could not be loaded on this machine.'
        }
    }

    try {
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
        $null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]

        $languages = @()
        foreach ($language in [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages) {
            $languages = $languages + $language.LanguageTag
        }

        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if ($null -eq $engine) {
            return [pscustomobject]@{
                Available = $false
                Languages = @($languages)
                Note      = ('No OCR language is installed for your display language. ' +
                             'Settings > Time & Language > Language > (your language) > Options > ' +
                             'Optical character recognition. Installed recognisers: ' +
                             $(if ($languages.Count) { $languages -join ', ' } else { 'none' }))
            }
        }

        return [pscustomobject]@{
            Available = $true
            Languages = @($languages)
            Note      = ''
        }
    } catch {
        return [pscustomobject]@{
            Available = $false
            Languages = @()
            Note      = ('The Windows OCR engine could not be reached: ' + $_.Exception.Message)
        }
    }
}

<#
    Recognise one image with the Windows engine, keeping every word's box.

    Returns a page object in the shape ocr-segment.js expects, so the result
    goes through exactly the same line grouping, column detection and XML
    writing as the Tesseract path.
#>
function Invoke-OcrWindowsPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $false)]
        [string] $Language,

        [Parameter(Mandatory = $false)]
        [int] $Dpi = 300
    )

    if (-not (Initialize-OcrWinRt)) {
        throw 'The Windows OCR engine is not available on this machine.'
    }

    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]

    $engine = $null
    if ($Language) {
        $tag = New-Object Windows.Globalization.Language($Language)
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($tag)
        if ($null -eq $engine) {
            throw ('No Windows OCR recogniser is installed for "' + $Language + '".')
        }
    } else {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        if ($null -eq $engine) {
            throw ('No Windows OCR language is installed. ' +
                   'Add one under Settings > Time & Language > Language > Options.')
        }
    }

    $full = Resolve-OcrPath $Path
    if (-not (Test-Path -LiteralPath $full)) {
        throw ('Image not found: ' + $Path)
    }

    $storageFile = Wait-OcrAsync ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $stream = Wait-OcrAsync ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])

    try {
        $decoder = Wait-OcrAsync ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-OcrAsync ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

        try {
            $pixelWidth = [int] $bitmap.PixelWidth
            $pixelHeight = [int] $bitmap.PixelHeight

            # The engine refuses an image outside its size limits rather than
            # scaling it, so the limit is checked here where a clear message
            # can be given.
            $maximum = [int] [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
            if ($pixelWidth -gt $maximum -or $pixelHeight -gt $maximum) {
                throw ('The image is ' + $pixelWidth + 'x' + $pixelHeight +
                       ', larger than the Windows OCR limit of ' + $maximum +
                       ' pixels on a side. Render the page at a lower resolution, ' +
                       'or use the Tesseract engine, which has no such limit.')
            }

            $result = Wait-OcrAsync ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

            $words = New-Object System.Collections.ArrayList
            foreach ($line in $result.Lines) {
                foreach ($word in $line.Words) {
                    $rect = $word.BoundingRect
                    $null = $words.Add([pscustomobject]@{
                        text = [string] $word.Text
                        x    = [double] $rect.X
                        y    = [double] $rect.Y
                        w    = [double] $rect.Width
                        h    = [double] $rect.Height
                        # The Windows engine reports no per-word confidence.
                        # Inventing one would be worse than admitting it.
                        conf = $null
                    })
                }
            }

            # TextAngle is the tilt the engine detected, in degrees, or null
            # when it could not tell. It is recorded rather than corrected:
            # the boxes are already in the source image's own pixels.
            $skew = 0.0
            if ($null -ne $result.TextAngle) { $skew = [double] $result.TextAngle }

            $effectiveDpi = $Dpi
            if ($effectiveDpi -le 0) { $effectiveDpi = 300 }

            return [pscustomobject]@{
                number      = 1
                pixelWidth  = $pixelWidth
                pixelHeight = $pixelHeight
                dpi         = $effectiveDpi
                widthPt     = ($pixelWidth * 72.0) / $effectiveDpi
                heightPt    = ($pixelHeight * 72.0) / $effectiveDpi
                skew        = $skew
                text        = [string] $result.Text
                words       = @($words)
                note        = 'recognised by the Windows OCR engine (no per-word confidence)'
            }
        } finally {
            $bitmap.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

<#
    Recognise a folder of page images with the Windows engine and write the
    same OcrDocument XML the Tesseract path produces.

    Node is still used - but only to segment and to write the XML, which is
    a few milliseconds of work and keeps one implementation of the layout
    rules. If Node is genuinely unavailable, the words are still written out
    as JSON so nothing is lost.
#>
function Invoke-OcrWindows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $InputPath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath,

        [Parameter(Mandatory = $false)]
        [string] $Root,

        [Parameter(Mandatory = $false)]
        [string] $Language,

        [Parameter(Mandatory = $false)]
        [int] $Dpi = 300,

        [Parameter(Mandatory = $false)]
        [switch] $Layout
    )

    if (-not $Root) { $Root = Get-OcrRoot }

    $check = Test-OcrWindowsEngine
    if (-not $check.Available) {
        throw $check.Note
    }

    $files = Get-OcrInputFiles -Path $InputPath
    $images = @($files | Where-Object { -not (Test-OcrIsPdf $_.FullName) })
    if ($images.Count -eq 0) {
        throw ('No page images found in ' + $InputPath)
    }

    $ocrFolder = Join-Path $OutputPath 'ocr'
    if (-not (Test-Path -LiteralPath $ocrFolder)) {
        $null = New-Item -ItemType Directory -Path $ocrFolder -Force
    }

    $node = Get-OcrNode
    $segmenter = Join-Path (Join-Path $Root 'js') 'ocr-segment.js'
    $results = New-Object System.Collections.ArrayList

    foreach ($image in $images) {
        Write-OcrLog ($image.Name) -Level Step
        $page = $null
        try {
            $page = Invoke-OcrWindowsPage -Path $image.FullName -Language $Language -Dpi $Dpi
        } catch {
            # One unreadable page must not cost the rest of the batch.
            Write-OcrLog ($image.Name + ': ' + $_.Exception.Message) -Level Bad
            $null = $results.Add([pscustomobject]@{
                File = $image.FullName; Ok = $false; Words = 0; Note = $_.Exception.Message
            })
            continue
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($image.Name)
        $jobPath = Join-Path $ocrFolder ($baseName + '.words.json')
        $xmlPath = Join-Path $ocrFolder ($baseName + '.ocr.xml')

        $job = [pscustomobject]@{
            source = [pscustomobject]@{ path = $image.FullName; kind = 'image'; pages = 1 }
            engine = [pscustomobject]@{
                name = 'Windows.Media.Ocr'
                lang = $(if ($Language) { $Language } else { 'user profile' })
                dpi  = $Dpi
            }
            pages  = @($page)
        }

        # Depth matters: the default of 2 would flatten the word list into
        # type names and produce a file full of "System.Object[]".
        $job | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jobPath -Encoding UTF8

        if ($node.Ok -and (Test-Path -LiteralPath $segmenter)) {
            $arguments = New-Object System.Collections.ArrayList
            $null = $arguments.Add($segmenter)
            $null = $arguments.Add('--in');  $null = $arguments.Add($jobPath)
            $null = $arguments.Add('--out'); $null = $arguments.Add($xmlPath)
            if ($Layout) { $null = $arguments.Add('--layout') }

            & $node.Path $arguments.ToArray()
            if ($LASTEXITCODE -ne 0) {
                Write-OcrLog ('Segmentation failed for ' + $image.Name +
                              '; the words are still in ' + [System.IO.Path]::GetFileName($jobPath)) -Level Warn
            } else {
                Remove-Item -LiteralPath $jobPath -Force -ErrorAction SilentlyContinue
            }
        } else {
            Write-OcrLog ('Node is not available, so no XML was written. The words are in ' +
                          [System.IO.Path]::GetFileName($jobPath) + '.') -Level Warn
        }

        Write-OcrLog ($image.Name + ': ' + $page.words.Count + ' words') -Level Good
        $null = $results.Add([pscustomobject]@{
            File = $image.FullName; Ok = $true; Words = $page.words.Count; Note = ''
        })
    }

    return ,@($results)
}
