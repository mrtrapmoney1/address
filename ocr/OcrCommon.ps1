<#
    OcrCommon.ps1 - the shared floor everything else stands on.

    Run folders, logging, locating Node, and the small helpers that keep the
    rest of the scripts short.

    Windows PowerShell 5.1 and PowerShell 7 both run this. That constraint is
    real and it costs: no ternary operator, no ?? , no ?. , and every file
    pure ASCII, because 5.1 reads a file with no BOM as the machine's ANSI
    code page and a stray smart quote in a comment will stop the whole script
    from parsing. tests\Test-Syntax.ps1 enforces both.
#>

Set-StrictMode -Version 2.0

# --------------------------------------------------------------------- paths

<#
    Where the toolkit lives. Everything is resolved from here rather than from
    the caller's current directory, so the tool works the same whether it is
    run from its own folder, from a scheduled task, or by double-clicking.
#>
function Get-OcrRoot {
    [CmdletBinding()]
    param()

    if ($PSScriptRoot) { return $PSScriptRoot }
    # $PSScriptRoot is empty when the file is dot-sourced from the console.
    if ($MyInvocation.MyCommand.Path) {
        return (Split-Path -Parent $MyInvocation.MyCommand.Path)
    }
    return (Get-Location).Path
}

<#
    Resolve a path that may not exist yet. Resolve-Path throws for a missing
    file, which is unhelpful when the whole point is to create it.
#>
function Resolve-OcrPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    return $resolved
}

# ----------------------------------------------------------------- run folder

<#
    Make the folder this run writes into.

    Every run gets its own, stamped with the time it started, so a second run
    can never quietly overwrite the first one's results and two runs can be
    compared side by side. The name sorts correctly as text, which matters
    when there are a hundred of them.

    Layout:
        _OcrRuns\2026-08-24_143012\
            pages\      the page images, including any rasterised from PDFs
            ocr\        one <name>.ocr.xml per page, plus optional layout dumps
            run.log     everything the run printed
            summary.xml what happened, machine-readable
#>
function New-OcrRunFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $Parent,

        [Parameter(Mandatory = $false)]
        [string] $Label
    )

    if (-not $Parent) { $Parent = Join-Path (Get-OcrRoot) '_OcrRuns' }
    $Parent = Resolve-OcrPath $Parent

    $stamp = (Get-Date).ToString('yyyy-MM-dd_HHmmss')
    if ($Label) {
        # Keep the folder name to characters every filesystem and every backup
        # tool is happy with.
        $safe = ($Label -replace '[^A-Za-z0-9._-]', '-')
        if ($safe.Length -gt 40) { $safe = $safe.Substring(0, 40) }
        $stamp = $stamp + '_' + $safe
    }

    $runPath = Join-Path $Parent $stamp

    # Two runs started in the same second must not collide.
    $suffix = 1
    while (Test-Path -LiteralPath $runPath) {
        $runPath = Join-Path $Parent ($stamp + '-' + $suffix)
        $suffix = $suffix + 1
    }

    $null = New-Item -ItemType Directory -Path $runPath -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $runPath 'pages') -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $runPath 'ocr') -Force

    $run = [pscustomobject]@{
        Path      = $runPath
        PagesPath = Join-Path $runPath 'pages'
        OcrPath   = Join-Path $runPath 'ocr'
        LogPath   = Join-Path $runPath 'run.log'
        Started   = Get-Date
    }

    Set-Content -LiteralPath $run.LogPath -Value ('OCR run started ' + $run.Started.ToString('s')) -Encoding ASCII
    return $run
}

# -------------------------------------------------------------------- logging

$script:OcrLogPath = $null

function Set-OcrLogPath {
    [CmdletBinding()]
    param(
        # AllowEmptyString as well as AllowNull: the [string] cast turns $null
        # into '' before validation sees it, so AllowNull on its own is not
        # enough and "Set-OcrLogPath $null" - the way logging is turned off -
        # fails on a mandatory parameter.
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string] $Path
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { $script:OcrLogPath = $null }
    else { $script:OcrLogPath = $Path }
}

<#
    One place that writes to the screen and to the run log at once.

    A run that only prints to the console leaves nothing behind to look at
    when someone asks why page 47 came out blank, and a run that only writes a
    log leaves the operator staring at a still cursor for ten minutes.
#>
function Write-OcrLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [AllowEmptyString()]
        [string] $Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Good', 'Warn', 'Bad', 'Step', 'Detail')]
        [string] $Level = 'Info',

        [Parameter(Mandatory = $false)]
        [switch] $NoConsole
    )

    $prefix = ''
    $colour = 'Gray'
    switch ($Level) {
        'Good'   { $prefix = '  [ok]   '; $colour = 'Green' }
        'Warn'   { $prefix = '  [warn] '; $colour = 'Yellow' }
        'Bad'    { $prefix = '  [FAIL] '; $colour = 'Red' }
        'Step'   { $prefix = '';          $colour = 'Cyan' }
        'Detail' { $prefix = '         '; $colour = 'DarkGray' }
        default  { $prefix = '  ';        $colour = 'Gray' }
    }

    if (-not $NoConsole) {
        Write-Host ($prefix + $Message) -ForegroundColor $colour
    }

    if ($script:OcrLogPath) {
        $stamp = (Get-Date).ToString('HH:mm:ss')
        $line = $stamp + '  ' + $Level.PadRight(6) + '  ' + $Message
        try {
            Add-Content -LiteralPath $script:OcrLogPath -Value $line -Encoding ASCII
        } catch {
            # A log that cannot be written must never stop the run that is
            # producing the actual results.
        }
    }
}

function Write-OcrBanner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Title
    )

    $line = '-' * 74
    Write-Host ''
    Write-Host ('  ' + $Title) -ForegroundColor White
    Write-Host ('  ' + $line) -ForegroundColor DarkGray
}

# ----------------------------------------------------------------------- Node

<#
    Find Node, and be specific when it is missing.

    "node is not recognised" sends people to a search engine. Saying which
    versions work and where to get it does not.
#>
function Get-OcrNode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $NodePath
    )

    $candidates = New-Object System.Collections.ArrayList

    if ($NodePath) { $null = $candidates.Add($NodePath) }

    $command = Get-Command -Name 'node' -CommandType Application -ErrorAction SilentlyContinue
    if ($command) {
        foreach ($item in @($command)) { $null = $candidates.Add($item.Source) }
    }

    # The usual install locations, for the case where PATH was not refreshed
    # after installing Node in the same session. Every one of these variables
    # is unset somewhere - on a non-Windows host, or in a stripped service
    # account's environment - and Join-Path throws on a null path rather than
    # returning nothing, so each is checked before it is used.
    foreach ($root in @($env:ProgramFiles, $env:ProgramW6432)) {
        if ($root) { $null = $candidates.Add((Join-Path $root 'nodejs\node.exe')) }
    }
    if ($env:LOCALAPPDATA) {
        $null = $candidates.Add((Join-Path $env:LOCALAPPDATA 'Programs\nodejs\node.exe'))
    }

    foreach ($candidate in $candidates) {
        if (-not $candidate) { continue }
        if (Test-Path -LiteralPath $candidate) {
            $version = $null
            try {
                $version = (& $candidate '--version' 2>$null | Select-Object -First 1)
            } catch {
                $version = $null
            }
            if ($version) {
                $numeric = 0
                if ($version -match '^v(\d+)') { $numeric = [int]$Matches[1] }
                return [pscustomobject]@{
                    Path         = $candidate
                    Version      = $version.Trim()
                    MajorVersion = $numeric
                    Ok           = ($numeric -ge 18)
                }
            }
        }
    }

    return [pscustomobject]@{
        Path         = $null
        Version      = $null
        MajorVersion = 0
        Ok           = $false
    }
}

<#
    Is the Node half ready to run? Checks the script, the engine and the
    language data separately, because each has a different fix.
#>
function Test-OcrNodeKit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $Root
    )

    if (-not $Root) { $Root = Get-OcrRoot }
    $jsPath = Join-Path $Root 'js'
    $cli = Join-Path $jsPath 'ocr-cli.js'
    $modules = Join-Path $jsPath 'node_modules'
    $engine = Join-Path $modules 'tesseract.js'
    $language = Join-Path $modules '@tesseract.js-data'

    $problems = New-Object System.Collections.ArrayList

    if (-not (Test-Path -LiteralPath $cli)) {
        $null = $problems.Add('ocr-cli.js is missing from ' + $jsPath)
    }
    if (-not (Test-Path -LiteralPath $engine)) {
        $null = $problems.Add('the OCR engine is not installed - run "npm install" in ' + $jsPath)
    }
    if (-not (Test-Path -LiteralPath $language)) {
        $null = $problems.Add('the language data is not installed - run "npm install" in ' + $jsPath)
    }

    return [pscustomobject]@{
        Ok         = ($problems.Count -eq 0)
        JsPath     = $jsPath
        CliPath    = $cli
        ModulePath = $modules
        Problems   = @($problems)
    }
}

# ------------------------------------------------------------------- listing

$script:OcrImageExtensions = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.pbm', '.pnm', '.tif', '.tiff')
$script:OcrDocumentExtensions = @('.pdf')

function Get-OcrInputFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,

        [Parameter(Mandatory = $false)]
        [switch] $Recurse
    )

    $full = Resolve-OcrPath $Path
    if (-not (Test-Path -LiteralPath $full)) {
        throw ('Input not found: ' + $Path)
    }

    $wanted = $script:OcrImageExtensions + $script:OcrDocumentExtensions

    if (Test-Path -LiteralPath $full -PathType Leaf) {
        $item = Get-Item -LiteralPath $full
        return ,@($item)
    }

    $items = Get-ChildItem -LiteralPath $full -File -Recurse:$Recurse |
        Where-Object { $wanted -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName

    # The comma is what actually keeps this an array. Without it PowerShell
    # unrolls the result: no matches returns $null and one match returns a
    # bare FileInfo, and .Count throws on both under Set-StrictMode.
    return ,@($items)
}

function Test-OcrIsPdf {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )
    return ([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -eq '.pdf')
}

# --------------------------------------------------------------------- misc

<#
    Human-readable size. A run that says it processed "48327104 bytes" makes
    the reader do arithmetic; "46.1 MB" does not.
#>
function Format-OcrSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [double] $Bytes
    )

    if ($Bytes -lt 1024) { return ([int]$Bytes).ToString() + ' B' }
    if ($Bytes -lt 1048576) { return ($Bytes / 1024).ToString('0.0') + ' KB' }
    if ($Bytes -lt 1073741824) { return ($Bytes / 1048576).ToString('0.0') + ' MB' }
    return ($Bytes / 1073741824).ToString('0.00') + ' GB'
}

function Format-OcrDuration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [TimeSpan] $Span
    )

    if ($Span.TotalSeconds -lt 60) { return $Span.TotalSeconds.ToString('0.0') + 's' }

    # The casts to string are not decoration. PowerShell decides what "+"
    # means from its LEFT operand, so "[int]$x + 'm '" is integer addition and
    # throws trying to convert 'm ' to a number. Making the left side a string
    # is what makes it concatenation.
    if ($Span.TotalMinutes -lt 60) {
        return ([int]$Span.TotalMinutes).ToString() + 'm ' + $Span.Seconds.ToString() + 's'
    }
    return ([int]$Span.TotalHours).ToString() + 'h ' + $Span.Minutes.ToString() + 'm'
}

<#
    Is this Windows? Several parts of this toolkit use Windows-only APIs, and
    they must say so plainly rather than failing with a type-not-found error.
    $IsWindows exists only in PowerShell 6+, so 5.1 needs the fallback.
#>
function Test-OcrOnWindows {
    [CmdletBinding()]
    param()

    if (Get-Variable -Name 'IsWindows' -Scope Global -ErrorAction SilentlyContinue) {
        return [bool] (Get-Variable -Name 'IsWindows' -Scope Global -ValueOnly)
    }
    # Windows PowerShell 5.1 only ever runs on Windows.
    return $true
}
