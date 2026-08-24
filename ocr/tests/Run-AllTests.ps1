<#
    Run-AllTests.ps1 - every suite, both halves, one answer at the end.

        .\tests\Run-AllTests.ps1

    Works in Windows PowerShell 5.1 and in PowerShell 7.
#>

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

# Each suite runs in its own process using THIS host's executable, so a
# suite's `exit` cannot take the runner down with it and $LASTEXITCODE is
# always about the suite that just ran.
$exe = (Get-Process -Id $PID).Path
if (-not $exe) {
    if ($PSVersionTable.PSEdition -eq 'Core') { $exe = 'pwsh' } else { $exe = 'powershell' }
}

$failed = 0
$suites = @('Test-Syntax.ps1', 'Test-Common.ps1', 'Test-Document.ps1', 'Test-EndToEnd.ps1')

foreach ($suite in $suites) {
    Write-Host ''
    Write-Host ('########## ' + $suite + ' ##########') -ForegroundColor Cyan
    & $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here $suite)
    if ($LASTEXITCODE -ne 0) { $failed++ }
}

# The JavaScript half has its own suite, and a green PowerShell run means
# nothing if the engine underneath it is broken.
Write-Host ''
Write-Host '########## the JavaScript suite ##########' -ForegroundColor Cyan

. (Join-Path $root 'OcrCommon.ps1')
$node = Get-OcrNode
$jsPath = Join-Path $root 'js'

if (-not $node.Ok) {
    Write-Host '  SKIPPED - Node.js 18 or newer was not found.' -ForegroundColor Yellow
} elseif (-not (Test-Path -LiteralPath (Join-Path $jsPath 'node_modules'))) {
    Write-Host '  SKIPPED - run  .\Invoke-Ocr.ps1 -Install  first.' -ForegroundColor Yellow
} else {
    <#
        The test files are DISCOVERED, never listed.

        They were listed once, and the list went stale the moment two new
        suites were added: thirty-three tests sat in the folder and quietly
        did not run, while the runner still reported success. A hard-coded
        list of tests is a list that will be wrong, and wrong in the direction
        that hides failures.
    #>
    $testFiles = @(Get-ChildItem -LiteralPath (Join-Path $jsPath 'test') -Filter '*.test.js' -File |
        Sort-Object Name |
        ForEach-Object { 'test/' + $_.Name })

    if ($testFiles.Count -eq 0) {
        Write-Host '  No JavaScript test files were found.' -ForegroundColor Yellow
        $failed++
    } else {
        Write-Host ('  ' + $testFiles.Count + ' test file(s)') -ForegroundColor DarkGray
        Push-Location -LiteralPath $jsPath
        try {
            $arguments = @('--test') + $testFiles
            & $node.Path $arguments
            if ($LASTEXITCODE -ne 0) { $failed++ }
        } finally {
            Pop-Location
        }
    }
}

Write-Host ''
if ($failed -gt 0) {
    Write-Host ('SUITE FAILED (' + $failed + ' of ' + ($suites.Count + 1) + ')') -ForegroundColor Red
    exit 1
}
Write-Host 'ALL SUITES PASSED' -ForegroundColor Green
exit 0
