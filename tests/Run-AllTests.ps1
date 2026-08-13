<#
 Run-AllTests.ps1 - runs the logic and paste suites and reports a single result.

 Works in Windows PowerShell 5.1 (powershell.exe) and PowerShell 7 (pwsh).
   Windows PowerShell 5.1:  .\tests\Run-AllTests.ps1
   or:                      powershell -ExecutionPolicy Bypass -File tests\Run-AllTests.ps1
#>
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

# run each child in its own process with THIS host's executable, so a child's
# `exit` can never affect us and $LASTEXITCODE is always reliable
$exe = (Get-Process -Id $PID).Path
if (-not $exe) { $exe = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh' } else { 'powershell' } }

$fail = 0
foreach ($t in @('Test-Logic.ps1', 'Test-Match.ps1', 'Test-Paste.ps1', 'Test-Pipeline.ps1')) {
    Write-Host "`n########## $t ##########" -ForegroundColor Cyan
    & $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here $t)
    if ($LASTEXITCODE -ne 0) { $fail++ }
}
Write-Host ''
if ($fail) { Write-Host "SUITE FAILED ($fail file(s))" -ForegroundColor Red; exit 1 }
Write-Host "ALL SUITES PASSED" -ForegroundColor Green
exit 0
