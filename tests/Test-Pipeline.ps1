<#
 Test-Pipeline.ps1
 End-to-end run of the WHOLE Get-CityCode.ps1 pipeline against a mock Excel:
 open workpaper -> detect columns -> group by quarter -> read the quarterly
 file -> match (incl. fuzzy) -> write the code+flag columns -> build the Report
 sheet -> save. If any line of the script errors at runtime, this fails.

 It injects the mock via New-NeExcelForTest (a hook the script checks for).
   .\tests\Test-Pipeline.ps1
#>
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path $here -Parent
. (Join-Path $here 'ExcelMock.ps1')

$pass = 0; $fail = 0
function Eq($got, $want, $msg) { if ("$got" -eq "$want") { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg (got '$got' want '$want')" -ForegroundColor Red } }
function Ok($c, $msg) { if ($c) { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg" -ForegroundColor Red } }

# ---- mock WORKPAPER: headers row 1, data rows 2-4 (Loc/Address/City/State/Zip/Period) ----
$wpGrid = @(
    @('Loc', 'Address', 'City', 'State', 'Zip', 'Period'),
    @(1, '150 Main St',    'Aurora',       'NE', 68818, 202108),   # -> OK 94
    @(2, '999 Nowhere Rd', 'Aurora',       'NE', 68818, 202108),   # -> NO MATCH
    @(3, 'N/A',            'Nowheresville', 'KS', 0,     202108)    # -> OOS
)
$wpWB = New-MockDataWorkbook $wpGrid

# ---- mock QUARTERLY file: header row 1 (col7 Street Name, col25 City Code (final)),
#      one data row: MAIN 68818, 100-199, both sides, ST, code 94 ----
function Col([hashtable]$byIndex, [int]$width) { $a = New-Object 'object[]' $width; foreach ($k in $byIndex.Keys) { $a[$k] = $byIndex[$k] }; return , $a }
$qHeader = Col @{ 6 = 'Street Name'; 23 = '_key (sorted A-Z)'; 24 = 'City Code (final)' } 25
$qRow1   = Col @{ 2 = 100; 3 = 199; 4 = 'B'; 6 = 'MAIN'; 7 = 'ST'; 13 = 'Aurora'; 14 = 68818; 23 = 'MAIN 68818'; 24 = 94 } 25
$qWB = New-MockDataWorkbook @($qHeader, $qRow1)

# ---- mock Excel with both workbooks registered by the paths the script will open ----
$excel = New-MockExcel
$qFile = Join-Path (Join-Path $root '_Address') '2021 Q3 Address Data.xlsx'
$wpFile = Join-Path ([IO.Path]::GetTempPath()) ('wp_' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
Set-Content -LiteralPath $qFile -Value 'placeholder' -NoNewline
Set-Content -LiteralPath $wpFile -Value 'placeholder' -NoNewline
$excel.Workbooks._openMap[(Get-Item -LiteralPath $wpFile).FullName] = $wpWB
$excel.Workbooks._openMap[(Get-Item -LiteralPath $qFile).FullName] = $qWB
function New-NeExcelForTest { $excel }

Write-Host "`n== running the full Get-CityCode pipeline against the mock ==" -ForegroundColor Cyan
$ran = $false
try {
    . (Join-Path $root 'Get-CityCode.ps1') -WorkingPaper $wpFile -HeaderRow 1 -PasteAt END -IncludeFlag -NonInteractive
    $ran = $true
} catch {
    $fail++; Write-Host "  FAIL: pipeline threw: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    Remove-Item -LiteralPath $qFile -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $wpFile -ErrorAction SilentlyContinue
}

Ok $ran 'pipeline ran to completion without a runtime error'
if ($ran) {
    $ws = $wpWB.Worksheets.Item(1)
    # 6 input columns -> results appended at col 7 (code) and 8 (flag), header row 1
    Eq $ws.Cells(1, 7).Value2 'NE City Code' 'code header appended at col 7'
    Eq $ws.Cells(1, 8).Value2 'Match Flag'   'flag header appended at col 8'
    Eq $ws.Cells(2, 7).Value2 94    'row 2 (150 Main St) -> code 94'
    Eq $ws.Cells(2, 8).Value2 'OK'  'row 2 -> flag OK'
    Ok ("$($ws.Cells(3, 8).Value2)" -match 'NO MATCH') 'row 3 (999 Nowhere Rd) -> NO MATCH'
    Eq $ws.Cells(4, 8).Value2 'OOS' 'row 4 (Kansas) -> OOS'
    Ok ($null -eq $ws.Cells(4, 7).Value2) 'row 4 OOS -> code left blank'
    # code column formatted 000
    Eq $ws.Range($ws.Cells(2, 7), $ws.Cells(4, 7)).NumberFormat '000' 'code column formatted 000'
    # Report sheet built, workbook saved
    $rep = $null; foreach ($s in $wpWB.Worksheets) { if ($s.Name -eq 'Report') { $rep = $s } }
    Ok ($null -ne $rep) 'Report sheet created in the workpaper'
    if ($rep) { Eq $rep.Cells(1, 1).Value2 'City Code Report' 'Report title present' }
    Ok ($wpWB.Saved) 'workpaper was saved'
}

Write-Host ('-' * 50) -ForegroundColor DarkGray
Write-Host ("  PASS: {0}   FAIL: {1}" -f $pass, $fail) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
if ($fail) { exit 1 }
exit 0
