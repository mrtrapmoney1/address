<#
    OcrTest.ps1 - the shared assertion helpers.

    Dot-sourced by each suite. Deliberately tiny: a test framework that has to
    be installed is a test framework that does not get run.
#>

$script:OcrPass = 0
$script:OcrFail = 0
$script:OcrSection = ''

function Start-OcrSection {
    param([string] $Name)
    $script:OcrSection = $Name
    Write-Host ''
    Write-Host ('  ' + $Name) -ForegroundColor Cyan
}

function Assert-OcrTrue {
    param(
        [Parameter(Position = 0)] $Condition,
        [Parameter(Position = 1)] [string] $Message
    )
    if ($Condition) {
        $script:OcrPass++
    } else {
        $script:OcrFail++
        Write-Host ('    FAIL  ' + $Message) -ForegroundColor Red
    }
}

function Assert-OcrEqual {
    param(
        [Parameter(Position = 0)] $Actual,
        [Parameter(Position = 1)] $Expected,
        [Parameter(Position = 2)] [string] $Message
    )
    if ("$Actual" -eq "$Expected") {
        $script:OcrPass++
    } else {
        $script:OcrFail++
        Write-Host ('    FAIL  ' + $Message) -ForegroundColor Red
        Write-Host ('          got      "' + $Actual + '"') -ForegroundColor DarkGray
        Write-Host ('          expected "' + $Expected + '"') -ForegroundColor DarkGray
    }
}

function Assert-OcrNear {
    param(
        [Parameter(Position = 0)] [double] $Actual,
        [Parameter(Position = 1)] [double] $Expected,
        [Parameter(Position = 2)] [double] $Tolerance,
        [Parameter(Position = 3)] [string] $Message
    )
    if ([Math]::Abs($Actual - $Expected) -le $Tolerance) {
        $script:OcrPass++
    } else {
        $script:OcrFail++
        Write-Host ('    FAIL  ' + $Message) -ForegroundColor Red
        Write-Host ('          got ' + $Actual + ', expected ' + $Expected + ' +/- ' + $Tolerance) -ForegroundColor DarkGray
    }
}

function Assert-OcrThrows {
    param(
        [Parameter(Position = 0)] [scriptblock] $Action,
        [Parameter(Position = 1)] [string] $Message
    )
    $threw = $false
    try { & $Action | Out-Null } catch { $threw = $true }
    Assert-OcrTrue $threw $Message
}

function Complete-OcrSuite {
    param([string] $Name)
    Write-Host ''
    $total = $script:OcrPass + $script:OcrFail
    if ($script:OcrFail -eq 0) {
        Write-Host ('  ' + $Name + ': ' + $script:OcrPass + '/' + $total + ' passed') -ForegroundColor Green
        exit 0
    }
    Write-Host ('  ' + $Name + ': ' + $script:OcrFail + ' of ' + $total + ' FAILED') -ForegroundColor Red
    exit 1
}
