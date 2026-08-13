<#
================================================================================
 Get-CityCode.ps1   (pure PowerShell - no engine template, no parallel runspaces)
 Address -> Nebraska city tax code, straight onto your working paper.

 HOW IT WORKS
   It ports the Address_Tax_Engine formulas into PowerShell (NeMatch.ps1) and
   matches your addresses directly against the quarterly tax-data files. No
   Excel engine workbook, no second workbook, no concurrency - just read the
   workpaper, match, write the code, save.

   The matcher is validated to reproduce the engine's answer exactly
   (tests\Test-Match.ps1).

 FILES NEEDED IN THE FOLDER
   Get-CityCode.ps1        this script
   NeMatch.ps1             the ported matching logic
   NeDicts.ps1             the city / direction / suffix lists (from the engine)
   NeTaxPaste.ps1          the workbook-writing helpers (tested)
   Your Working Paper.xlsx
   _Address\               the quarterly files (2021 Q3 Address Data.xlsx, ...)

 REQUIREMENTS
   Windows PowerShell 5.1 and Microsoft Excel.
================================================================================
#>

[CmdletBinding()]
param(
    [string]$WorkingPaper,
    [string]$SheetName,
    [int]   $HeaderRow = 0,
    [string]$PasteAt,
    [switch]$IncludeFlag,
    [switch]$NonInteractive
)
$ErrorActionPreference = 'Stop'
$StartTime = Get-Date

$CodeNumberFormat = '000'
$ReadChunk = 50000

# ---- constants ----
$xlCalcManual = -4135; $xlCalcAutomatic = -4105

$folder = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
foreach ($lib in @('NeDicts.ps1','NeMatch.ps1','NeTaxPaste.ps1')) {
    $p = Join-Path $folder $lib
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "Missing required file: $lib" -ForegroundColor Red; if (-not $NonInteractive) { Read-Host '  Press Enter to close' }; exit 1 }
}
. (Join-Path $folder 'NeMatch.ps1')       # also dot-sources NeDicts.ps1
. (Join-Path $folder 'NeTaxPaste.ps1')

function H($t) { Write-Host ''; Write-Host ('=' * 72) -ForegroundColor DarkCyan; Write-Host "  $t" -ForegroundColor Cyan; Write-Host ('=' * 72) -ForegroundColor DarkCyan }
function S($t) { Write-Host "  $t" -ForegroundColor Gray }
function G($t) { Write-Host "  $t" -ForegroundColor Green }
function W2($t){ Write-Host "  $t" -ForegroundColor Yellow }
function Die($t) { Write-Host "  $t" -ForegroundColor Red; if ($script:excel) { try { $script:excel.Quit() } catch {} }; if (-not $NonInteractive) { Read-Host '  Press Enter to close' }; exit 1 }

function Get-Nth($arr, [int]$i) {
    if ($null -eq $arr) { return $null }
    if ($arr -is [Array]) { if ($arr.Rank -eq 2) { return $arr.GetValue($i + $arr.GetLowerBound(0) - 1, $arr.GetLowerBound(1)) } return $arr.GetValue($i + $arr.GetLowerBound(0) - 1) }
    if ($i -eq 1) { return $arr }
    return $null
}

# ============================================================================
H 'Nebraska City Tax Code  (pure PowerShell)'
S "Folder: $folder"
$addrFolder = Join-Path $folder '_Address'
if (-not (Test-Path -LiteralPath $addrFolder)) { Die "The '_Address' folder was not found next to this script." }
$qMap = @{}
foreach ($f in (Get-ChildItem -LiteralPath $addrFolder -File -Filter '*.xlsx' | Where-Object { $_.Name -notlike '~$*' })) {
    if ($f.Name -match '(?<y>(19|20)\d{2}).{0,3}?Q(?<q>[1-4])') { $k = "$($Matches['y']) Q$($Matches['q'])"; if (-not $qMap.ContainsKey($k)) { $qMap[$k] = $f.FullName } }
}
if ($qMap.Count -eq 0) { Die 'No quarterly files found in _Address (need names like "2021 Q3 Address Data.xlsx").' }
G "_Address holds $($qMap.Count) quarterly file(s)."

# ---- working paper ----
if ($WorkingPaper) { if (-not (Test-Path -LiteralPath $WorkingPaper)) { Die "Working paper not found: $WorkingPaper" }; $wbFile = Get-Item -LiteralPath $WorkingPaper }
else {
    $books = Get-ChildItem -LiteralPath $folder -File | Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xlsb)$' -and $_.Name -notlike '~$*' } | Sort-Object Name
    if ($books.Count -eq 0) { Die 'No working paper found in this folder.' }
    if ($books.Count -eq 1 -or $NonInteractive) { $wbFile = $books[0] }
    else { Write-Host ''; for ($i = 0; $i -lt $books.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $books[$i].Name) }; do { $pick = Read-Host "`n  Which workbook is the working paper? (number)" } while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $books.Count); $wbFile = $books[[int]$pick - 1] }
}
G "Working paper: $($wbFile.Name)"

# ============================================================================
H 'Opening the workbook'
$script:excel = New-Object -ComObject Excel.Application
$excel = $script:excel
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false; $excel.EnableEvents = $false; $excel.AskToUpdateLinks = $false
$wb = $null
try { $wb = $excel.Workbooks.Open($wbFile.FullName, 0, $false) } catch { Die "Could not open the workbook: $($_.Exception.Message)" }
$prevCalc = $xlCalcAutomatic; try { $prevCalc = $excel.Calculation; $excel.Calculation = $xlCalcManual } catch {}

$sheets = @(); foreach ($s in $wb.Worksheets) { $sheets += $s }
if ($SheetName) { $ws = $sheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1; if (-not $ws) { Die "No sheet named '$SheetName'." } }
elseif ($sheets.Count -eq 1 -or $NonInteractive) { $ws = $sheets[0] }
else { Write-Host ''; for ($i = 0; $i -lt $sheets.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $sheets[$i].Name) }; do { $pick = Read-Host "`n  Which sheet holds the data? (number)" } while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $sheets.Count); $ws = $sheets[[int]$pick - 1] }
G "Sheet: $($ws.Name)"

# ============================================================================
H 'Which columns'
$hdrRow = 1
if ($HeaderRow -gt 0) { $hdrRow = $HeaderRow }
elseif (-not $NonInteractive) { $x = Read-Host '  Which row holds the headers? (Enter = 1)'; if ($x -and ($x -as [int])) { $hdrRow = [int]$x } }

$used = $ws.UsedRange
$lastRow = $used.Row + $used.Rows.Count - 1
$lastCol = $used.Column + $used.Columns.Count - 1
$firstDataRow = $hdrRow + 1
$nRows = $lastRow - $firstDataRow + 1
if ($nRows -lt 1) { Die 'No data rows below the header row.' }

$headers = New-Object 'string[]' ($lastCol + 1)
for ($c = 1; $c -le $lastCol; $c++) { $v = $ws.Cells($hdrRow, $c).Value2; $headers[$c] = if ($null -eq $v) { '' } else { [string]$v } }

Write-Host ''
Write-Host "  Headers (row $hdrRow), $nRows data row(s):" -ForegroundColor White
for ($c = 1; $c -le $lastCol; $c++) { $nm = $headers[$c]; if (-not $nm) { $nm = '(blank)' }; Write-Host ("   [{0,3}]  {1,-3}  {2}" -f $c, (Get-ColLetter $c), $nm) }

function Ask-Column([string]$what, [bool]$optional, [string[]]$hints) {
    $guess = 0
    for ($c = 1; $c -le $lastCol; $c++) { $h = ($headers[$c] -replace '[^A-Za-z0-9]', '').ToUpper(); if (-not $h) { continue }; foreach ($hint in $hints) { if ($h -eq $hint -or $h -like "*$hint*") { $guess = $c; break } }; if ($guess) { break } }
    if ($NonInteractive) { if ($guess) { return $guess } elseif ($optional) { return 0 } else { Die "Could not auto-detect the $what column." } }
    while ($true) {
        $sfx = if ($optional) { "  (0 = none)" } else { "" }; $def = if ($guess) { " [Enter = $guess : $($headers[$guess])]" } else { "" }
        $ans = Read-Host "  Column for the $what$sfx$def"
        if (-not $ans) { if ($guess) { return $guess }; if ($optional) { return 0 }; W2 '   Required.'; continue }
        if ($optional -and ($ans.Trim() -eq '0' -or $ans.Trim() -match '^(none|no|n/a|na)$')) { return 0 }
        if ($ans -as [int]) { $n = [int]$ans; if ($n -ge 1 -and $n -le $lastCol) { return $n }; W2 '   Out of range.'; continue }
        for ($c = 1; $c -le $lastCol; $c++) { if ($headers[$c] -and $headers[$c].Trim().ToUpper() -eq $ans.Trim().ToUpper()) { return $c } }
        W2 '   No header by that name.'
    }
}
Write-Host ''
$colAddr = Ask-Column 'STREET ADDRESS' $false @('ADDRESSLINE1','ADDRESS1','SHIPTOADDRESS','STREETADDRESS','ADDRESS','STREET')
$colCity = Ask-Column 'CITY'  $false @('SHIPTOCITY','SOURCEDTOCITY','BILLTOCITY','CITY')
$colZip  = Ask-Column 'ZIP'   $false @('ZIPCODE','SHIPTOZIP','POSTALCODE','ZIP','POSTAL')
$colDate = Ask-Column 'DATE (YYYYMM)' $true @('YYYYMM','PERIOD','INVOICEDATE','SALEDATE','TRANDATE','DATE')
$colState= Ask-Column 'STATE' $true @('SHIPTOSTATE','SOURCEDTOSTATE','BILLTOSTATE','STATE')

$fixedQuarter = $null
if ($colDate -eq 0) {
    if ($qMap.Count -eq 1) { $fixedQuarter = @($qMap.Keys)[0]; W2 "No date column; using the only quarter present: $fixedQuarter" }
    elseif ($NonInteractive) { Die 'No date column and multiple quarters; cannot choose in -NonInteractive.' }
    else {
        Write-Host ''; W2 'No date column. Which period does the whole sheet use?'
        S ('Available: ' + (($qMap.Keys | Sort-Object | ForEach-Object { $_ -replace ' ','' }) -join '  '))
        while ($true) { $ans = Read-Host '  Period (YYYYQX, e.g. 2021Q3)'; $k = $null
            if ($ans -match '^\s*(?<y>(19|20)\d{2})\s*[-_ ]?[Qq]?\s*(?<q>[1-4])\s*$') { $k = "$($Matches['y']) Q$($Matches['q'])" }
            if ($k -and $qMap.ContainsKey($k)) { $fixedQuarter = $k; break }; W2 '   Not found. Pick from the list.' }
    }
}
$assumeNE = $false
if ($colState -eq 0) { if ($NonInteractive) { $assumeNE = $true } else { $a = Read-Host '  No state column. Treat EVERY row as Nebraska? (Y/N)'; if ($a -match '^[Yy]') { $assumeNE = $true } else { Die 'A state column is needed.' } } }

# ---- output options ----
Write-Host ''
Write-Host '  Result: press ENTER to append at the end, or type a column (e.g. T) to INSERT there.' -ForegroundColor White
$insertAt = 0
if ($PasteAt) { if ($PasteAt.Trim().ToUpper() -ne 'END') { $insertAt = Resolve-PasteLocation $PasteAt $headers $lastCol; if ($insertAt -lt 1) { Die "Could not understand paste location '$PasteAt'." } } }
elseif (-not $NonInteractive) { $ans = Read-Host "  Result location (Enter = end)"; if ($ans) { $insertAt = Resolve-PasteLocation $ans $headers $lastCol; if ($insertAt -lt 1) { W2 '   Not recognised - appending at end.'; $insertAt = 0 } } }
$writeFlag = [bool]$IncludeFlag
if (-not $IncludeFlag -and -not $NonInteractive) { $af = Read-Host '  Also add a Match Flag column? (y/N)'; if ($af -match '^[Yy]') { $writeFlag = $true } }
if ($insertAt -gt 0) { G "Insert at column $(Get-ColLetter $insertAt)." } else { G 'Append at end.' }
G $(if ($writeFlag) { 'Columns: NE City Code (000) + Match Flag.' } else { 'Column: NE City Code only (000).' })

if (-not $NonInteractive) { $go = Read-Host "`n  Press Enter to run, or N to cancel"; if ($go -match '^[Nn]') { $wb.Close($false); $excel.Quit(); exit 0 } }

# ============================================================================
H 'Backup'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$bakDir = Join-Path $folder '_Backups'; if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir | Out-Null }
$bak = Join-Path $bakDir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension)
Copy-Item -LiteralPath $wbFile.FullName -Destination $bak -Force
G "Backup: _Backups\$(Split-Path $bak -Leaf)"

# ============================================================================
H 'Reading the workpaper'
function Read-Col([int]$col) {
    $out = New-Object 'string[]' $nRows; $done = 0
    while ($done -lt $nRows) { $take = [Math]::Min($ReadChunk, $nRows - $done); $r1 = $firstDataRow + $done; $r2 = $r1 + $take - 1
        $vals = $ws.Range($ws.Cells($r1, $col), $ws.Cells($r2, $col)).Value2
        for ($i = 1; $i -le $take; $i++) { $v = Get-Nth $vals $i; $out[$done + $i - 1] = if ($null -eq $v) { '' } else { [string]$v } }
        $done += $take }
    return , $out
}
$aAddr = Read-Col $colAddr; $aCity = Read-Col $colCity; $aZip = Read-Col $colZip
$aDate = if ($colDate -gt 0) { Read-Col $colDate } else { $null }
$aState = if ($assumeNE) { $null } else { Read-Col $colState }
G "Read $nRows row(s)."

# ============================================================================
H 'Grouping rows into quarters'
function Get-Quarter([string]$v) {
    if (-not $v) { return $null }
    $s = ($v -replace '[^0-9]', '')
    if ($s.Length -ge 6) { $y = [int]$s.Substring(0, 4); $m = [int]$s.Substring(4, 2); if ($y -ge 1990 -and $y -le 2100 -and $m -ge 1 -and $m -le 12) { return ("{0} Q{1}" -f $y, [Math]::Ceiling($m / 3)) } }
    $d = 0.0; if ([double]::TryParse($v, [ref]$d) -and $d -gt 20000 -and $d -lt 80000) { $dt = [DateTime]::FromOADate($d); return ("{0} Q{1}" -f $dt.Year, [Math]::Ceiling($dt.Month / 3)) }
    $dt2 = [DateTime]::MinValue; if ([DateTime]::TryParse($v, [ref]$dt2)) { return ("{0} Q{1}" -f $dt2.Year, [Math]::Ceiling($dt2.Month / 3)) }
    return $null
}
$finalCode = New-Object 'object[]' $nRows
$finalFlag = New-Object 'string[]' $nRows
$groups = @{}
$anyZip9 = $false
for ($i = 0; $i -lt $nRows; $i++) {
    if (-not $assumeNE) { $st = ($aState[$i] -replace '[^A-Za-z]', '').ToUpper(); if ($st -and $st -ne 'NE' -and $st -ne 'NEBRASKA' -and $st -ne 'NEB') { $finalFlag[$i] = 'OOS'; continue } }
    if (-not $aAddr[$i]) { $finalFlag[$i] = 'NO ADDRESS'; continue }
    $q = if ($fixedQuarter) { $fixedQuarter } else { Get-Quarter $aDate[$i] }
    if (-not $q) { $finalFlag[$i] = 'NO DATE'; continue }
    if (-not $qMap.ContainsKey($q)) { $finalFlag[$i] = "NO TAX FILE - $($q -replace ' ','')"; continue }
    if (($aZip[$i] -replace '[^0-9]', '').Length -ge 9) { $anyZip9 = $true }
    if (-not $groups.ContainsKey($q)) { $groups[$q] = New-Object 'System.Collections.Generic.List[int]' }
    $groups[$q].Add($i)
}
if ($groups.Count -eq 0) { W2 'No rows matched a quarterly file.' } else { foreach ($k in ($groups.Keys | Sort-Object)) { S ("{0}  ->  {1,8:N0} row(s)" -f $k, $groups[$k].Count) } }

# ============================================================================
H 'Matching against the tax data'
function Read-QuarterColumns($path) {
    $qwb = $excel.Workbooks.Open($path, 0, $true)
    try {
        $qws = $qwb.Worksheets.Item(1)
        $qHdr = 0
        for ($r = 1; $r -le 5; $r++) { $v = $qws.Cells($r, 7).Value2; if ($v -and ([string]$v).ToUpper().Contains('STREET')) { $qHdr = $r; break } }
        if ($qHdr -eq 0) { throw "no 'Street Name' header in $(Split-Path $path -Leaf)" }
        $yv = [string]$qws.Cells($qHdr, 25).Value2
        if ($yv -and -not $yv.ToUpper().Contains('CITY CODE')) { throw "column Y of $(Split-Path $path -Leaf) is '$yv', not 'City Code (final)'" }
        $qUsed = $qws.UsedRange; $qLast = $qUsed.Row + $qUsed.Rows.Count - 1; $qFirst = $qHdr + 1; $qn = $qLast - $qFirst + 1
        if ($qn -lt 1) { throw "no data rows in $(Split-Path $path -Leaf)" }
        function RC([int]$col) {
            $arr = New-Object 'object[]' $qn; $done = 0
            while ($done -lt $qn) { $take = [Math]::Min($ReadChunk, $qn - $done); $r1 = $qFirst + $done; $r2 = $r1 + $take - 1
                $vals = $qws.Range($qws.Cells($r1, $col), $qws.Cells($r2, $col)).Value2
                for ($i = 1; $i -le $take; $i++) { $arr[$done + $i - 1] = Get-Nth $vals $i }
                $done += $take }
            return , $arr
        }
        $dr = New-Object 'object[]' $qn; for ($i = 0; $i -lt $qn; $i++) { $dr[$i] = $qFirst + $i }
        $res = @{ n = $qn; dr = $dr; key = (RC 24); lo = (RC 3); hi = (RC 4); e = (RC 5); h = (RC 8); code = (RC 25); concat = $(if ($anyZip9) { RC 23 } else { $null }) }
        return $res
    } finally { $qwb.Close($false) }
}

$qi = 0
foreach ($q in ($groups.Keys | Sort-Object)) {
    $qi++
    S ("[{0}/{1}] {2}: loading tax data..." -f $qi, $groups.Count, $q)
    $t0 = Get-Date
    try { $qd = Read-QuarterColumns $qMap[$q] } catch { Die "Quarter $q failed: $($_.Exception.Message)" }
    # normalise code to int
    $codeArr = New-Object 'object[]' $qd.n
    for ($i = 0; $i -lt $qd.n; $i++) { $cy = 0.0; $codeArr[$i] = if ([double]::TryParse("$($qd.code[$i])", [ref]$cy)) { [int]$cy } else { $null } }
    $index = New-NeIndex $qd.dr $qd.key $qd.lo $qd.hi $qd.e $qd.h $codeArr $qd.concat
    S ("      {0:N0} tax rows in {1:N1}s; matching {2:N0} address(es)..." -f $qd.n, ((Get-Date) - $t0).TotalSeconds, $groups[$q].Count)
    $repaired = 0
    foreach ($ri in $groups[$q]) {
        $p = Get-NeParse $aAddr[$ri] $aCity[$ri] $aZip[$ri]
        $m = Find-NeCode $p $index
        # PASS 2 - fuzzy repair only on rows that failed
        if ($m.Flag -eq 'NO MATCH' -or $m.Flag -eq 'FALLBACK') {
            $rep = Repair-NeAddress $aAddr[$ri]
            if ($rep.Note) {
                $p2 = Get-NeParse $rep.Text $aCity[$ri] $aZip[$ri]
                $m2 = Find-NeCode $p2 $index
                $better = ($m2.Flag -eq 'OK' -or $m2.Flag -eq 'ZIP9' -or ($m2.Flag -eq 'FALLBACK' -and $m.Flag -eq 'NO MATCH'))
                if ($better) { $m = @{ Code = $m2.Code; Flag = "$($m2.Flag) - REPAIRED: $($rep.Note)" }; $repaired++ }
            }
        }
        $finalCode[$ri] = $m.Code; $finalFlag[$ri] = $m.Flag
    }
    if ($repaired -gt 0) { S ("      fuzzy repair recovered {0:N0} address(es)" -f $repaired) }
    $index = $null; [System.GC]::Collect()
}
G 'Matching complete.'

# ============================================================================
H 'Writing the code onto the working paper'
if ($writeFlag) { $headerNames = @('NE City Code','Match Flag'); $columns = $finalCode, $finalFlag } else { $headerNames = @('NE City Code'); $columns = , $finalCode }
$excel.ScreenUpdating = $true
$startCol = Write-ResultBlock $ws $excel $hdrRow $firstDataRow $lastRow $lastCol $nRows $insertAt $headerNames $columns $ReadChunk $CodeNumberFormat 0
if ($insertAt -gt 0) { S "Inserted at $(Get-ColLetter $startCol) (existing data shifted right)." } else { S "Appended at $(Get-ColLetter $startCol)." }
S "NE City Code -> column $(Get-ColLetter $startCol), formatted $CodeNumberFormat"
if ($writeFlag) { S "Match Flag   -> column $(Get-ColLetter ($startCol + 1))" }

# ============================================================================
H 'Building the Report sheet'
# flag summary (base flag, i.e. everything before " - REPAIRED")
$tally2 = @{}
for ($i = 0; $i -lt $nRows; $i++) { $f = $finalFlag[$i]; if (-not $f) { $f = '(blank)' }; $b = ($f -split ' - ')[0].Trim(); if (-not $tally2.ContainsKey($b)) { $tally2[$b] = 0 }; $tally2[$b]++ }
$flagRows = foreach ($k in ($tally2.Keys | Sort-Object { -$tally2[$_] })) { [pscustomobject]@{ Flag = $k; Count = $tally2[$k]; Pct = (100 * $tally2[$k] / $nRows) } }
# unique addresses with code/flag/count
$uniq = @{}
for ($i = 0; $i -lt $nRows; $i++) {
    $key = (("" + $aAddr[$i]).Trim().ToUpper()) + '|' + (("" + $aCity[$i]).Trim().ToUpper()) + '|' + (("" + $aZip[$i]).Trim())
    if (-not $uniq.ContainsKey($key)) { $uniq[$key] = [pscustomobject]@{ Address = $aAddr[$i]; City = $aCity[$i]; Zip = $aZip[$i]; Code = $finalCode[$i]; Flag = $finalFlag[$i]; Count = 0 } }
    $uniq[$key].Count++
}
$uniqueRows = $uniq.Values | Sort-Object @{e = { ($_.Flag -split ' - ')[0] } }, @{e = { -$_.Count } }, City, Address
try {
    $rep = New-NeReport $excel $wb 'Report' $flagRows $uniqueRows $CodeNumberFormat
    G "Report sheet written: '$($rep.Sheet)' ($($uniqueRows.Count) unique address(es))."
} catch { W2 "Could not build the Report sheet: $($_.Exception.Message)" }

# ============================================================================
H 'Saving'
try { $excel.Calculation = $prevCalc } catch { $excel.Calculation = $xlCalcAutomatic }
$saved = $false
for ($a = 1; $a -le 4 -and -not $saved; $a++) { try { $wb.Save(); $saved = $true } catch { W2 "Save attempt $a failed: $($_.Exception.Message)"; Start-Sleep -Seconds ([Math]::Pow(2, $a)) } }
if (-not $saved) { $alt = Join-Path $folder ("{0}_RESULTS_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension); try { $wb.SaveAs($alt); $saved = $true; W2 "Saved a copy as $(Split-Path $alt -Leaf)." } catch {} }
if ($saved) { G "Saved: $($wbFile.Name)" } else { Write-Host '  COULD NOT SAVE - do not close Excel until you save by hand.' -ForegroundColor Red }

# ============================================================================
H 'Summary'
$tally = @{}
for ($i = 0; $i -lt $nRows; $i++) { $f = $finalFlag[$i]; if (-not $f) { $f = '(blank)' }; if (-not $tally.ContainsKey($f)) { $tally[$f] = 0 }; $tally[$f]++ }
foreach ($k in ($tally.Keys | Sort-Object { -$tally[$_] })) { Write-Host ("   {0,-16} {1,8:N0}   {2,5:N1}%" -f $k, $tally[$k], (100 * $tally[$k] / $nRows)) }
Write-Host ''
G ("Finished in {0:N1} min." -f ((Get-Date) - $StartTime).TotalMinutes)
$excel.Visible = $true; $excel.EnableEvents = $true
if (-not $NonInteractive) { Read-Host '  Excel is open with your results. Press Enter to close this window' }
