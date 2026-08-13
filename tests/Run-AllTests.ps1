<#
 Run-AllTests.ps1 - runs the logic and paste suites and reports a single result.
   pwsh -File tests/Run-AllTests.ps1
#>
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$fail = 0
foreach ($t in @('Test-Logic.ps1','Test-Paste.ps1')) {
    Write-Host "`n########## $t ##########" -ForegroundColor Cyan
    & (Join-Path $here $t)
    if ($LASTEXITCODE -ne 0) { $fail++ }
}
Write-Host ''
if ($fail) { Write-Host "SUITE FAILED ($fail file(s))" -ForegroundColor Red; exit 1 }
Write-Host "ALL SUITES PASSED" -ForegroundColor Green
