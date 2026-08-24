<#
    OcrRaster.ps1 - turning documents into page images.

    Tesseract.js reads png, jpg, bmp, gif, webp and pbm. It does NOT read PDF
    and it does NOT read TIFF - which between them are what almost every
    scanner and every accounts inbox actually produces. Something has to
    render them first, and that something must not be a download.

    Windows 10 and later ship two APIs that do exactly this, and nothing has
    to be installed to use them:

        Windows.Data.Pdf      renders a PDF page to a bitmap
        System.Drawing        reads the frames of a multi-page TIFF

    Both are Windows-only. On any other host these functions say so plainly
    and the caller is told to supply page images instead - which is much
    better than a type-not-found exception from deep inside a loop.

    The WinRT plumbing (the [Type, Assembly, ContentType = WindowsRuntime]
    loading and the Await helper) follows the approach used by
    JoelGMSec's Invoke-Transfer, which is the clearest published example of
    driving these APIs from Windows PowerShell 5.1.
#>

Set-StrictMode -Version 2.0

$script:OcrWinRtReady = $false

<#
    Load the WinRT types and build the awaiter.

    WinRT methods return IAsyncOperation<T>, which PowerShell has no syntax
    for. The trick - and it is the whole reason this is fiddly - is to reach
    into WindowsRuntimeSystemExtensions for the generic GetAwaiter, close it
    over the right result type by reflection, and call GetResult() to block.
#>
function Initialize-OcrWinRt {
    [CmdletBinding()]
    param()

    if ($script:OcrWinRtReady) { return $true }

    if (-not (Test-OcrOnWindows)) {
        return $false
    }

    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop

        # Referencing each type with its contract is what actually loads the
        # WinRT projection; the assignment to $null is only to keep the type
        # object off the pipeline.
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Foundation.IAsyncOperation`1, Windows.Foundation, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapEncoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]

        $script:OcrAwaiterMethod = [WindowsRuntimeSystemExtensions].GetMember('GetAwaiter') |
            Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } |
            Select-Object -First 1

        if (-not $script:OcrAwaiterMethod) {
            return $false
        }

        $script:OcrWinRtReady = $true
        return $true
    } catch {
        return $false
    }
}

<#
    Block on a WinRT IAsyncOperation<T> and return its result.
#>
function Wait-OcrAsync {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Operation,

        [Parameter(Mandatory = $true, Position = 1)]
        [Type] $ResultType
    )

    if (-not $script:OcrWinRtReady) {
        throw 'The Windows runtime types are not loaded. Call Initialize-OcrWinRt first.'
    }

    return $script:OcrAwaiterMethod.
        MakeGenericMethod($ResultType).
        Invoke($null, @($Operation)).
        GetResult()
}

<#
    Can this machine turn a PDF into page images by itself?
#>
function Test-OcrRasterSupport {
    [CmdletBinding()]
    param()

    $windows = Test-OcrOnWindows
    $winrt = $false
    if ($windows) { $winrt = Initialize-OcrWinRt }

    $drawing = $false
    if ($windows) {
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $drawing = $true
        } catch {
            $drawing = $false
        }
    }

    return [pscustomobject]@{
        OnWindows = $windows
        Pdf       = $winrt
        Tiff      = $drawing
        Note      = $(if ($windows) { '' } else { 'PDF and TIFF conversion need Windows 10 or later. Supply PNG or JPEG page images instead.' })
    }
}

<#
    Render every page of a PDF to a PNG.

    Resolution is the point of this function. Windows.Data.Pdf renders at
    whatever size it is asked for, and asking for the page's natural size
    gives 96 dpi - at which eight-point type on an invoice is about eight
    pixels tall and recognises badly. So the render is scaled to the target
    resolution up front, which is far better than rendering small and
    enlarging afterwards: this way the glyph edges are drawn sharp rather
    than interpolated from a blur.
#>
function Convert-OcrPdfToImages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $DestinationFolder,

        [Parameter(Mandatory = $false)]
        [int] $Dpi = 300,

        [Parameter(Mandatory = $false)]
        [int] $FirstPage = 1,

        [Parameter(Mandatory = $false)]
        [int] $LastPage = 0,

        [Parameter(Mandatory = $false)]
        [int] $MaxDimension = 6000
    )

    if (-not (Initialize-OcrWinRt)) {
        throw ('This machine cannot render PDFs: it needs Windows 10 or later. ' +
               'Convert ' + [System.IO.Path]::GetFileName($Path) + ' to PNG or JPEG page images and run those instead.')
    }

    $full = Resolve-OcrPath $Path
    if (-not (Test-Path -LiteralPath $full)) {
        throw ('PDF not found: ' + $Path)
    }
    if (-not (Test-Path -LiteralPath $DestinationFolder)) {
        $null = New-Item -ItemType Directory -Path $DestinationFolder -Force
    }

    $storageFile = Wait-OcrAsync ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $document = Wait-OcrAsync ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])

    $pageCount = [int] $document.PageCount
    $from = $FirstPage
    if ($from -lt 1) { $from = 1 }
    $to = $LastPage
    if ($to -lt 1 -or $to -gt $pageCount) { $to = $pageCount }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($full)
    $written = New-Object System.Collections.ArrayList

    for ($number = $from; $number -le $to; $number++) {
        $page = $document.GetPage([uint32]($number - 1))
        try {
            # Size is in points; 72 points to the inch, so this is the exact
            # pixel count for the requested resolution.
            $widthPt = [double] $page.Size.Width
            $heightPt = [double] $page.Size.Height
            $scale = $Dpi / 72.0

            $pixelWidth = [int][Math]::Round($widthPt * $scale)
            $pixelHeight = [int][Math]::Round($heightPt * $scale)

            # A very large page at a high resolution can exhaust memory before
            # a single word is read; cap it and say so rather than dying.
            $longest = [Math]::Max($pixelWidth, $pixelHeight)
            if ($longest -gt $MaxDimension) {
                $shrink = $MaxDimension / $longest
                $pixelWidth = [int][Math]::Round($pixelWidth * $shrink)
                $pixelHeight = [int][Math]::Round($pixelHeight * $shrink)
            }

            $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $options.DestinationWidth = [uint32] $pixelWidth
            $options.DestinationHeight = [uint32] $pixelHeight

            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            try {
                $null = Wait-OcrAsync ($page.RenderToStreamAsync($stream, $options)) ([Object])

                $target = Join-Path $DestinationFolder ($baseName + '-p' + $number.ToString('0000') + '.png')

                # RenderToStreamAsync produces a bitmap stream; copy it to disk
                # through a .NET stream, which is the least surprising route
                # from a WinRT buffer to a file.
                $reader = New-Object Windows.Storage.Streams.DataReader($stream.GetInputStreamAt(0))
                try {
                    $size = [uint32] $stream.Size
                    $null = Wait-OcrAsync ($reader.LoadAsync($size)) ([uint32])
                    $bytes = New-Object 'byte[]' $size
                    $reader.ReadBytes($bytes)
                    [System.IO.File]::WriteAllBytes($target, $bytes)
                } finally {
                    $reader.Dispose()
                }

                $null = $written.Add([pscustomobject]@{
                    Page        = $number
                    Path        = $target
                    PixelWidth  = $pixelWidth
                    PixelHeight = $pixelHeight
                    WidthPt     = $widthPt
                    HeightPt    = $heightPt
                    Dpi         = [int][Math]::Round(($pixelWidth * 72.0) / $widthPt)
                })
            } finally {
                $stream.Dispose()
            }
        } finally {
            $page.Close()
        }
    }

    return ,@($written)
}

<#
    Split a TIFF into one PNG per frame.

    Multi-page TIFF is what a departmental scanner produces when it is asked
    for "one file per batch", so a single file is routinely a whole day's
    invoices. System.Drawing reads the frames through the page dimension,
    which is the part people miss - without SelectActiveFrame only the first
    page is ever seen, and the rest of the batch silently disappears.
#>
function Convert-OcrTiffToImages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $DestinationFolder
    )

    if (-not (Test-OcrOnWindows)) {
        throw ('TIFF conversion needs Windows. Convert ' +
               [System.IO.Path]::GetFileName($Path) + ' to PNG page images and run those instead.')
    }

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    } catch {
        throw ('System.Drawing is not available, so TIFF files cannot be split here: ' + $_.Exception.Message)
    }

    $full = Resolve-OcrPath $Path
    if (-not (Test-Path -LiteralPath $full)) {
        throw ('TIFF not found: ' + $Path)
    }
    if (-not (Test-Path -LiteralPath $DestinationFolder)) {
        $null = New-Item -ItemType Directory -Path $DestinationFolder -Force
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($full)
    $written = New-Object System.Collections.ArrayList
    $image = [System.Drawing.Image]::FromFile($full)

    try {
        $dimension = New-Object System.Drawing.Imaging.FrameDimension($image.FrameDimensionsList[0])
        $frames = $image.GetFrameCount($dimension)

        for ($index = 0; $index -lt $frames; $index++) {
            $null = $image.SelectActiveFrame($dimension, $index)
            $target = Join-Path $DestinationFolder ($baseName + '-p' + ($index + 1).ToString('0000') + '.png')

            # Copy the frame into a fresh bitmap before saving: saving the
            # source Image directly writes the whole TIFF again, not the frame.
            $frame = New-Object System.Drawing.Bitmap($image.Width, $image.Height)
            try {
                $frame.SetResolution($image.HorizontalResolution, $image.VerticalResolution)
                $graphics = [System.Drawing.Graphics]::FromImage($frame)
                try {
                    $graphics.DrawImage($image, 0, 0, $image.Width, $image.Height)
                } finally {
                    $graphics.Dispose()
                }
                $frame.Save($target, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $frame.Dispose()
            }

            $dpi = [int][Math]::Round($image.HorizontalResolution)
            if ($dpi -le 0) { $dpi = 200 }

            $null = $written.Add([pscustomobject]@{
                Page        = $index + 1
                Path        = $target
                PixelWidth  = $image.Width
                PixelHeight = $image.Height
                WidthPt     = ($image.Width * 72.0) / $dpi
                HeightPt    = ($image.Height * 72.0) / $dpi
                Dpi         = $dpi
            })
        }
    } finally {
        $image.Dispose()
    }

    return ,@($written)
}

<#
    One entry point: hand it anything, get page images back.

    An image that is already in a format the engine reads is passed through
    untouched rather than re-encoded - every re-encode of a JPEG loses a
    little more of the text.
#>
function ConvertTo-OcrPageImages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [string] $DestinationFolder,

        [Parameter(Mandatory = $false)]
        [int] $Dpi = 300
    )

    $full = Resolve-OcrPath $Path
    $extension = [System.IO.Path]::GetExtension($full).ToLowerInvariant()

    if ($extension -eq '.pdf') {
        return Convert-OcrPdfToImages -Path $full -DestinationFolder $DestinationFolder -Dpi $Dpi
    }

    if ($extension -eq '.tif' -or $extension -eq '.tiff') {
        return Convert-OcrTiffToImages -Path $full -DestinationFolder $DestinationFolder
    }

    return ,@([pscustomobject]@{
        Page        = 1
        Path        = $full
        PixelWidth  = 0
        PixelHeight = 0
        WidthPt     = 0
        HeightPt    = 0
        Dpi         = 0
    })
}
