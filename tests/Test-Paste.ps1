<#
================================================================================
 Test-Paste.ps1
 Verifies the workbook PASTE - the part that scrambled before - by running the
 real NeTaxPaste.ps1 functions against a throwaway workbook and reading every
 written cell back.

 It uses real Excel COM if this machine has it; otherwise it falls back to the
 mock in ExcelMock.ps1 so it can still run on a build box. The assertions are
 identical either way.

 Run:   powershell -ExecutionPolicy Bypass -File .\tests\Test-Paste.ps1
        pwsh -File ./tests/Test-Paste.ps1        (uses the mock off-Windows)
================================================================================
#>
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path (Split-Path $here -Parent) 'NeTaxPaste.ps1')

$script:pass = 0; $script:fail = 0
function Ok($cond, $msg) { if ($cond) { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg" -ForegroundColor Red } }
function Eq($got, $want, $msg) { if ("$got" -eq "$want") { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg  (got '$got' want '$want')" -ForegroundColor Red } }

# ---- pick a backend: real Excel if available, else the mock ----
$excel = $null; $usingReal = $false; $realApp = $null
try { $realApp = New-Object -ComObject Excel.Application; $usingReal = $true } catch { $usingReal = $false }
if ($usingReal) {
    $realApp.Visible = $false; $realApp.DisplayAlerts = $false
    $excel = $realApp
    Write-Host "Backend: REAL Excel COM" -ForegroundColor Cyan
} else {
    . (Join-Path $here 'ExcelMock.ps1')
    $excel = New-MockExcel
    Write-Host "Backend: MOCK (no Excel on this machine)" -ForegroundColor Cyan
}

# helper: make a fresh data sheet with a header row at row 3 and $n data rows.
# We build it with Range + (N x 1) arrays - the SAME way production writes - so
# it also exercises that path and dodges a Windows-PowerShell quirk where a
# scalar Int32 assigned to a single cell's .Value2 throws an InvalidCastException.
$HDR = 3; $FIRST = 4
function Set-Col($ws, [int]$col, [int]$r1, $values) {
    $n = $values.Count
    $blk = [Array]::CreateInstance([object], $n, 1)
    for ($i = 0; $i -lt $n; $i++) { $blk.SetValue($values[$i], $i, 0) }
    $ws.Range($ws.Cells($r1, $col), $ws.Cells($r1 + $n - 1, $col)).Value2 = $blk
}
function New-Fixture([int]$n) {
    if ($usingReal) { $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1) }
    else { $ws = New-MockSheet 'Data' }   # ExcelMock is already dot-sourced at top level
    # header row (A..E) as a 1 x 5 array.
    # NOTE: do not name this $hdr - PowerShell variables are case-insensitive, so
    # $hdr would clobber $HDR (the header row number) and wreck every Cells() call.
    $hdrNames = @('Loc#','Address','City','State','Zip')
    $hb = [Array]::CreateInstance([object], 1, 5)
    for ($c = 0; $c -lt 5; $c++) { $hb.SetValue($hdrNames[$c], 0, $c) }
    $ws.Range($ws.Cells($HDR, 1), $ws.Cells($HDR, 5)).Value2 = $hb
    # data columns, each as an (N x 1) array
    $loc = @(); $adr = @(); $cty = @(); $sta = @(); $zip = @()
    for ($i = 0; $i -lt $n; $i++) { $loc += $i; $adr += ("ADDR{0} " -f $i); $cty += ("City{0}" -f $i); $sta += 'NE'; $zip += (68000 + $i) }
    Set-Col $ws 1 $FIRST $loc
    Set-Col $ws 2 $FIRST $adr
    Set-Col $ws 3 $FIRST $cty
    Set-Col $ws 4 $FIRST $sta
    Set-Col $ws 5 $FIRST $zip
    return $ws
}

# known result arrays of length n (distinct per row so a swap is obvious)
function New-Results([int]$n) {
    $code = New-Object 'object[]' $n; $flag = New-Object 'string[]' $n
    $row  = New-Object 'object[]' $n; $src  = New-Object 'string[]' $n
    for ($i = 0; $i -lt $n; $i++) {
        switch ($i % 5) {
            0 { $code[$i] = 0;   $flag[$i] = 'OK';       $row[$i] = 100 + $i }
            1 { $code[$i] = 94;  $flag[$i] = 'OK';       $row[$i] = 200 + $i }
            2 { $code[$i] = 305; $flag[$i] = 'FALLBACK'; $row[$i] = "FALLBACK - row $((900+$i)) (unvalidated)" }
            3 { $code[$i] = $null; $flag[$i] = 'NO MATCH'; $row[$i] = $null }
            4 { $code[$i] = $null; $flag[$i] = 'OOS';      $row[$i] = $null }
        }
        $src[$i] = '2021 Q3 Address Data.xlsx'
    }
    return @{ Code = $code; Flag = $flag; Row = $row; Src = $src }
}

$N = 20                       # 20 data rows
$readChunk = 7                # force 3 chunks so chunk-boundary bugs show up

# ============================================================================
Write-Host "`n== A. Append, CODE ONLY, 000 format ==" -ForegroundColor Cyan
# ============================================================================
$ws = New-Fixture $N
$res = New-Results $N
$lastRow = $FIRST + $N - 1
$startCol = Write-ResultBlock $ws $excel $HDR $FIRST $lastRow 5 $N 0 @('NE City Code') (,$res.Code) $readChunk '000' 0
Eq $startCol 6 'append lands at column 6 (after A:E)'
Eq $ws.Cells($HDR, 6).Value2 'NE City Code' 'header written'
$allCodeOk = $true; $scramble = $false
for ($i = 0; $i -lt $N; $i++) {
    $r = $FIRST + $i
    if ("$($ws.Cells($r,6).Value2)" -ne "$($res.Code[$i])") { $allCodeOk = $false }
    if ($null -ne $ws.Cells($r,7).Value2) { $scramble = $true }   # nothing must land beyond the single column
}
Ok $allCodeOk  'every code cell matches its row (no misalignment across chunks)'
Ok (-not $scramble) 'code-only wrote exactly one column (nothing spilled into col 7)'
Eq $ws.Range($ws.Cells($FIRST,6),$ws.Cells($lastRow,6)).NumberFormat '000' 'code column formatted 000'
# originals untouched (incl. trailing space and ints)
Eq $ws.Cells($FIRST,2).Value2 'ADDR0 ' 'address col A:E untouched (trailing space kept)'
Eq $ws.Cells($FIRST,1).Value2 0        'Loc# untouched'
Eq $ws.Cells($FIRST,5).Value2 68000    'Zip untouched'

# ============================================================================
Write-Host "`n== B. Append, CODE + FLAG ==" -ForegroundColor Cyan
# ============================================================================
$ws = New-Fixture $N
$res = New-Results $N
$startCol = Write-ResultBlock $ws $excel $HDR $FIRST $lastRow 5 $N 0 @('NE City Code','Match Flag') ($res.Code, $res.Flag) $readChunk '000' 0
Eq $ws.Cells($HDR,6).Value2 'NE City Code' 'code header'
Eq $ws.Cells($HDR,7).Value2 'Match Flag'   'flag header'
$codeOk = $true; $flagOk = $true
for ($i = 0; $i -lt $N; $i++) {
    $r = $FIRST + $i
    if ("$($ws.Cells($r,6).Value2)" -ne "$($res.Code[$i])") { $codeOk = $false }
    if ("$($ws.Cells($r,7).Value2)" -ne "$($res.Flag[$i])") { $flagOk = $false }
}
Ok $codeOk 'code column correct, row-aligned'
Ok $flagOk 'flag column correct, row-aligned (no "OK 0" jamming)'
# the classic scramble signature must be ABSENT: a flag cell holding "<x> <number>"
$sig = $false
for ($i = 0; $i -lt $N; $i++) { $f = "$($ws.Cells($FIRST+$i,7).Value2)"; if ($f -match '\s-?\d+$') { $sig = $true } }
Ok (-not $sig) 'no flag cell contains a trailing jammed number'

# ============================================================================
Write-Host "`n== C. INSERT at an existing column (no overwrite) ==" -ForegroundColor Cyan
# ============================================================================
$ws = New-Fixture $N
$res = New-Results $N
# capture originals that must survive the insert
$origB = $ws.Cells($FIRST,2).Value2      # 'ADDR0 '
$origE = $ws.Cells($FIRST,5).Value2      # 68000
$startCol = Write-ResultBlock $ws $excel $HDR $FIRST $lastRow 5 $N 2 @('NE City Code') (,$res.Code) $readChunk '000' 0
Eq $startCol 2 'insert lands exactly at column 2'
Eq $ws.Cells($HDR,2).Value2 'NE City Code' 'inserted header at col 2'
Eq $ws.Cells($FIRST,2).Value2 $res.Code[0] 'inserted code at col 2'
# original column B has shifted to C, original E to F - values intact, nothing overwritten
Eq $ws.Cells($HDR,3).Value2 'Address' 'original Address header shifted to col 3'
Eq $ws.Cells($FIRST,3).Value2 $origB 'original Address value shifted to col 3, intact'
Eq $ws.Cells($FIRST,6).Value2 $origE 'original Zip value shifted to col 6, intact'
Eq $ws.Cells($FIRST,1).Value2 0 'Loc# (left of insert) unchanged'

# ============================================================================
Write-Host "`n== D. Engine backup: one value-pasted sheet per quarter ==" -ForegroundColor Cyan
# ============================================================================
# two quarters over 6 rows
$firstDataRow = 4
$aAddr = @('A0','A1','A2','A3','A4','A5'); $aCity = @('C0','C1','C2','C3','C4','C5'); $aZip = @(1,2,3,4,5,6)
$backCode = @(0,94,$null,305,0,$null); $backRow = @(10,20,$null,40,50,$null)
$backFlag = @('OK','OK','NO MATCH','FALLBACK','OK','NO MATCH'); $outSrc = @('Q3','Q3','Q3','Q4','Q4','Q4')
$groups = @{}
$groups['2021 Q3'] = [System.Collections.Generic.List[int]]::new(); 0,1,2 | ForEach-Object { $groups['2021 Q3'].Add($_) }
$groups['2021 Q4'] = [System.Collections.Generic.List[int]]::new(); 3,4,5 | ForEach-Object { $groups['2021 Q4'].Add($_) }
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("engbak_{0}.xlsx" -f ([Guid]::NewGuid().ToString('N').Substring(0,6)))
$bk = New-EngineBackup $excel $tmp @('2021 Q3','2021 Q4') $groups $firstDataRow $aAddr $aCity $aZip $backCode $backRow $backFlag $outSrc
Eq $bk.Sheets.Count 2 'two quarter sheets created'
Ok ($bk.Sheets -contains '2021Q3 (engine)') 'Q3 sheet named'
Ok ($bk.Sheets -contains '2021Q4 (engine)') 'Q4 sheet named'
if (-not $usingReal) {
    $book = $excel.Workbooks._books[-1]
    $q3 = $book.Worksheets.Items | Where-Object { $_.Name -eq '2021Q3 (engine)' }
    Eq $q3.Cells(1,1).Value2 'SheetRow' 'backup header row'
    Eq $q3.Cells(2,1).Value2 4  'Q3 row1 SheetRow = firstDataRow+0'
    Eq $q3.Cells(2,2).Value2 'A0' 'Q3 row1 address'
    Eq $q3.Cells(2,5).Value2 0    'Q3 row1 engine code'
    Eq $q3.Cells(4,7).Value2 'NO MATCH' 'Q3 row3 engine flag'
    Eq $book.SavedPath $tmp 'engine backup SaveAs path recorded'
    Ok ($book.SavedFormat -eq 51) 'engine backup saved as .xlsx (fmt 51)'
}

# ============================================================================
Write-Host "`n== E. Report sheet (flags + unique addresses, idempotent) ==" -ForegroundColor Cyan
# ============================================================================
if ($usingReal) { $rwb = $excel.Workbooks.Add() } else { $rwb = $excel.Workbooks.Add() }
$flagRows = @([pscustomobject]@{Flag='OK';Count=57;Pct=49.1}, [pscustomobject]@{Flag='NO MATCH';Count=20;Pct=17.2})
$uniqueRows = @(
    [pscustomobject]@{Address='2225 Q St'; City='Aurora'; Zip=68818; Code=0; Flag='OK'; Count=8},
    [pscustomobject]@{Address='1414 Mankin St'; City='Aurora'; Zip=68818; Code=$null; Flag='NO MATCH'; Count=1})
$rep = New-NeReport $excel $rwb 'Report' $flagRows $uniqueRows '000'
$rsh = $null; foreach ($s in $rwb.Worksheets) { if ($s.Name -eq 'Report') { $rsh = $s } }
Ok ($null -ne $rsh) 'Report sheet created'
Eq $rsh.Cells(1,1).Value2 'City Code Report' 'report title'
Eq $rsh.Cells(6,1).Value2 'OK' 'first flag row label'
Eq $rsh.Cells(6,2).Value2 57 'first flag row count'
$fu = $rep.FirstUnique
Eq $rsh.Cells($fu,1).Value2 '2225 Q St' 'first unique address'
Eq $rsh.Cells($fu,4).Value2 0 'first unique code'
Eq $rsh.Cells($fu,6).Value2 8 'first unique count'
$null = New-NeReport $excel $rwb 'Report' $flagRows $uniqueRows '000'
$rc = 0; foreach ($s in $rwb.Worksheets) { if ($s.Name -eq 'Report') { $rc++ } }
Eq $rc 1 'Report is idempotent (re-run replaces, not duplicates)'

# ---- cleanup real Excel ----
if ($usingReal) { try { $realApp.Quit() } catch {} }
Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue

Write-Host "`n---------------------------------------------" -ForegroundColor DarkGray
Write-Host ("  PASS: {0}   FAIL: {1}   (backend: {2})" -f $script:pass, $script:fail, $(if($usingReal){'Excel'}else{'mock'})) -ForegroundColor $(if($script:fail){'Red'}else{'Green'})
if ($script:fail) { exit 1 }

exit 0
