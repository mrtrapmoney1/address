<#
    OcrNode.ps1 - running the Node half from PowerShell.

    PowerShell decides WHAT to recognise and where the results go; Node does
    the recognition. This file is the join between them, and it is deliberately
    thin: the more logic that lives on this side of the boundary, the more of
    the tool cannot be tested without a machine that has both halves working.

    The summary comes back through a FILE, not through stdout. Capturing a
    child process's standard output while also letting its progress reach the
    operator's console means merging the two streams, and a single warning
    from a dependency then lands in the middle of the JSON and breaks the
    parse. A file cannot be corrupted that way.
#>

Set-StrictMode -Version 2.0

<#
    Install the Node dependencies. Run once, on a machine with a network; the
    engine and the language data then live in ocr\js\node_modules and nothing
    is ever downloaded again - which is the whole point, because the machines
    this runs on are usually the ones that cannot reach a CDN.
#>
function Install-OcrNodeKit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $Root
    )

    if (-not $Root) { $Root = Get-OcrRoot }
    $jsPath = Join-Path $Root 'js'

    if (-not (Test-Path -LiteralPath (Join-Path $jsPath 'package.json'))) {
        throw ('No package.json in ' + $jsPath + ' - the js folder is incomplete.')
    }

    <#
        Select-Object -First 1 is not tidiness, it is a bug fix.

        Get-Command returns EVERY match on PATH, and a machine with more than
        one Node install has more than one npm. Without this, $npm.Source is
        an array, and using it as a command produces the wonderful error

            The term '/opt/node22/bin/npm /usr/local/bin/npm' is not recognized

        because PowerShell joined the two paths with a space and tried to run
        the result. The first match is the one PATH would have chosen anyway.
    #>
    $npm = Get-Command -Name 'npm' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if (-not $npm) {
        throw ('npm was not found. Install Node.js (which includes npm) from https://nodejs.org, ' +
               'then open a new PowerShell window so PATH is picked up, and run this again.')
    }

    Write-OcrLog ('Installing the OCR engine into ' + $jsPath) -Level Step
    Write-OcrLog 'This needs the network, and only has to be done once.' -Level Detail

    $code = 0
    Push-Location -LiteralPath $jsPath
    try {
        # npm writes progress to stderr even on success, so the exit code is
        # the only trustworthy verdict - never the presence of stderr output.
        & $npm.Source 'install' '--no-audit' '--no-fund'
        $code = $LASTEXITCODE
        if ($null -eq $code) { $code = 0 }
    } finally {
        Pop-Location
    }

    if ($code -ne 0) {
        throw ('npm install failed with exit code ' + $code + '. Look at the output above.')
    }

    $check = Test-OcrNodeKit -Root $Root
    if (-not $check.Ok) {
        throw ('npm install finished but the kit is still incomplete: ' + ($check.Problems -join '; '))
    }

    Write-OcrLog 'The OCR engine and language data are installed.' -Level Good
    return $check
}

<#
    Run the recogniser over a folder of page images.

    Every switch the CLI understands is exposed here as a parameter, because
    the alternative - a single string of arguments - puts the burden of
    quoting paths with spaces on the caller, and that is a bug waiting in
    every "C:\Users\Firstname Lastname\..." on earth.
#>
function Invoke-OcrNode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $InputPath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath,

        [Parameter(Mandatory = $false)]
        [string] $Root,

        [Parameter(Mandatory = $false)]
        [string] $NodePath,

        [Parameter(Mandatory = $false)]
        [string] $Language = 'eng',

        [Parameter(Mandatory = $false)]
        [string] $LanguagePath,

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
        [string] $Whitelist,

        [Parameter(Mandatory = $false)]
        [switch] $NoPreprocess,

        [Parameter(Mandatory = $false)]
        [switch] $NoDeskew,

        [Parameter(Mandatory = $false)]
        [switch] $NoRefine,

        [Parameter(Mandatory = $false)]
        [switch] $NoInvoice,

        [Parameter(Mandatory = $false)]
        [switch] $Crop,

        [Parameter(Mandatory = $false)]
        [switch] $Layout,

        [Parameter(Mandatory = $false)]
        [switch] $Json,

        [Parameter(Mandatory = $false)]
        [switch] $KeepImages,

        [Parameter(Mandatory = $false)]
        [switch] $Force,

        [Parameter(Mandatory = $false)]
        [switch] $Quiet
    )

    if (-not $Root) { $Root = Get-OcrRoot }

    $kit = Test-OcrNodeKit -Root $Root
    if (-not $kit.Ok) {
        throw ('The Node half is not ready: ' + ($kit.Problems -join '; '))
    }

    $node = Get-OcrNode -NodePath $NodePath
    if (-not $node.Ok) {
        if ($node.Path) {
            throw ('Node ' + $node.Version + ' was found at ' + $node.Path +
                   ' but version 18 or newer is required.')
        }
        throw ('Node.js was not found. Install it from https://nodejs.org (the LTS build), ' +
               'then open a new PowerShell window so PATH is picked up.')
    }

    $summaryFile = Join-Path ([System.IO.Path]::GetTempPath()) ('ocr-summary-' + [Guid]::NewGuid().ToString('N') + '.json')

    # Built as a list rather than a string: each element is passed to the
    # process as one argument, so a path with a space in it needs no quoting
    # and cannot be split.
    $arguments = New-Object System.Collections.ArrayList
    $null = $arguments.Add($kit.CliPath)
    $null = $arguments.Add('--in');  $null = $arguments.Add((Resolve-OcrPath $InputPath))
    $null = $arguments.Add('--out'); $null = $arguments.Add((Resolve-OcrPath $OutputPath))
    $null = $arguments.Add('--lang'); $null = $arguments.Add($Language)
    $null = $arguments.Add('--psm');  $null = $arguments.Add($PageSegmentation)
    $null = $arguments.Add('--target-dpi'); $null = $arguments.Add([string] $TargetDpi)
    $null = $arguments.Add('--method'); $null = $arguments.Add($Threshold)
    $null = $arguments.Add('--refine-below'); $null = $arguments.Add([string] $RefineBelow)
    $null = $arguments.Add('--summary-file'); $null = $arguments.Add($summaryFile)

    if ($LanguagePath) { $null = $arguments.Add('--lang-path'); $null = $arguments.Add((Resolve-OcrPath $LanguagePath)) }
    if ($SourceDpi -gt 0) { $null = $arguments.Add('--dpi'); $null = $arguments.Add([string] $SourceDpi) }
    if ($Workers -gt 0) { $null = $arguments.Add('--workers'); $null = $arguments.Add([string] $Workers) }
    if ($Whitelist) { $null = $arguments.Add('--whitelist'); $null = $arguments.Add($Whitelist) }
    if ($NoPreprocess) { $null = $arguments.Add('--no-preprocess') }
    if ($NoDeskew) { $null = $arguments.Add('--no-deskew') }
    if ($NoRefine) { $null = $arguments.Add('--no-refine') }
    if ($NoInvoice) { $null = $arguments.Add('--no-invoice') }
    if ($Crop) { $null = $arguments.Add('--crop') }
    if ($Layout) { $null = $arguments.Add('--layout') }
    if ($Json) { $null = $arguments.Add('--json') }
    if ($KeepImages) { $null = $arguments.Add('--keep-images') }
    if ($Force) { $null = $arguments.Add('--force') }
    if ($Quiet) { $null = $arguments.Add('--quiet') }

    Write-OcrLog ('node ' + [System.IO.Path]::GetFileName($kit.CliPath) +
                  ' (' + $node.Version + ')') -Level Detail

    $exitCode = 0
    try {
        # The call operator lets Node's progress reach the console as it
        # happens, which is what a person watching a long batch needs.
        & $node.Path $arguments.ToArray()
        $exitCode = $LASTEXITCODE
    } catch {
        throw ('Could not run Node: ' + $_.Exception.Message)
    }

    $summary = $null
    if (Test-Path -LiteralPath $summaryFile) {
        try {
            $summary = Get-Content -LiteralPath $summaryFile -Raw | ConvertFrom-Json
        } catch {
            Write-OcrLog ('The run summary could not be read: ' + $_.Exception.Message) -Level Warn
        }
        Remove-Item -LiteralPath $summaryFile -Force -ErrorAction SilentlyContinue
    }

    if ($null -eq $summary) {
        # No summary means Node did not get far enough to write one. The exit
        # code is all there is to report, and pretending otherwise would hide
        # a real failure.
        return [pscustomobject]@{
            Ok       = $false
            ExitCode = $exitCode
            Pages    = @()
            Elapsed  = 0
            Note     = 'Node produced no run summary; the recognition did not complete.'
        }
    }

    return [pscustomobject]@{
        Ok       = ($exitCode -eq 0 -and $summary.ok)
        ExitCode = $exitCode
        Pages    = @($summary.pages)
        Elapsed  = $summary.elapsedSeconds
        Engine   = $summary.engine
        OutPath  = $summary.out
        Note     = ''
    }
}
