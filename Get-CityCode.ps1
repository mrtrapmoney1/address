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
$ReadChunk = 200000        # bigger chunks = far fewer COM round-trips

# ---- constants ----
$xlCalcManual = -4135; $xlCalcAutomatic = -4105

$folder = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
foreach ($lib in @('NeDicts.ps1','NeMatch.ps1','NeTaxPaste.ps1')) {
    $p = Join-Path $folder $lib
    if (-not (Test-Path -LiteralPath $p)) { Write-Host "Missing required file: $lib" -ForegroundColor Red; if (-not $NonInteractive) { Read-Host '  Press Enter to close' }; exit 1 }
}
. (Join-Path $folder 'NeMatch.ps1')       # also dot-sources NeDicts.ps1
. (Join-Path $folder 'NeTaxPaste.ps1')

$script:Phase = 0
function WrHead($t) {
    $script:Phase++
    Write-Host ''
    Write-Host ('  +' + ('-' * 68) + '+') -ForegroundColor DarkCyan
    Write-Host ('  | ' + ("{0,-2} {1,-64}" -f $script:Phase, $t) + '|') -ForegroundColor Cyan
    Write-Host ('  +' + ('-' * 68) + '+') -ForegroundColor DarkCyan
}
function WrBanner($t, $sub) {
    Write-Host ''
    Write-Host ('  ' + ('=' * 70)) -ForegroundColor DarkCyan
    Write-Host ('  ' + $t) -ForegroundColor White
    if ($sub) { Write-Host ('  ' + $sub) -ForegroundColor DarkGray }
    Write-Host ('  ' + ('=' * 70)) -ForegroundColor DarkCyan
}
function WrStep($t) { Write-Host "  $t" -ForegroundColor Gray }
function WrGood($t) { Write-Host "  $t" -ForegroundColor Green }
function WrWarn($t){ Write-Host "  $t" -ForegroundColor Yellow }
function WrDie($t) { Write-Host "  $t" -ForegroundColor Red; if ($script:excel) { try { $script:excel.Quit() } catch {} }; if (-not $NonInteractive) { Read-Host '  Press Enter to close' }; exit 1 }

function Get-Nth($arr, [int]$i) {
    if ($null -eq $arr) { return $null }
    if ($arr -is [Array]) { if ($arr.Rank -eq 2) { return $arr.GetValue($i + $arr.GetLowerBound(0) - 1, $arr.GetLowerBound(1)) } return $arr.GetValue($i + $arr.GetLowerBound(0) - 1) }
    if ($i -eq 1) { return $arr }
    return $null
}

# ============================================================================
WrBanner 'NEBRASKA CITY TAX CODE' 'address -> city code, written straight onto your working paper'
WrStep "Folder: $folder"
$addrFolder = Join-Path $folder '_Address'
if (-not (Test-Path -LiteralPath $addrFolder)) { WrDie "The '_Address' folder was not found next to this script." }
$qMap = @{}
foreach ($f in (Get-ChildItem -LiteralPath $addrFolder -File -Filter '*.xlsx' | Where-Object { $_.Name -notlike '~$*' })) {
    if ($f.Name -match '(?<y>(19|20)\d{2}).{0,3}?Q(?<q>[1-4])') { $k = "$($Matches['y']) Q$($Matches['q'])"; if (-not $qMap.ContainsKey($k)) { $qMap[$k] = $f.FullName } }
}
if ($qMap.Count -eq 0) { WrDie 'No quarterly files found in _Address (need names like "2021 Q3 Address Data.xlsx").' }
WrGood "_Address holds $($qMap.Count) quarterly file(s)."

# ---- working paper ----
if ($WorkingPaper) { if (-not (Test-Path -LiteralPath $WorkingPaper)) { WrDie "Working paper not found: $WorkingPaper" }; $wbFile = Get-Item -LiteralPath $WorkingPaper }
else {
    $books = Get-ChildItem -LiteralPath $folder -File | Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xlsb)$' -and $_.Name -notlike '~$*' } | Sort-Object Name
    if ($books.Count -eq 0) { WrDie 'No working paper found in this folder.' }
    if ($books.Count -eq 1 -or $NonInteractive) { $wbFile = $books[0] }
    else { Write-Host ''; for ($i = 0; $i -lt $books.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $books[$i].Name) }; do { $pick = Read-Host "`n  Which workbook is the working paper? (number)" } while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $books.Count); $wbFile = $books[[int]$pick - 1] }
}
WrGood "Working paper: $($wbFile.Name)"

# ============================================================================
WrHead 'Opening the workbook'
# Excel is injectable for the integration test (New-NeExcelForTest). In normal
# use that function does not exist, so this is a plain COM Excel.Application.
$script:excel = if (Get-Command -Name New-NeExcelForTest -ErrorAction SilentlyContinue) { New-NeExcelForTest } else { New-Object -ComObject Excel.Application }
$excel = $script:excel
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false; $excel.EnableEvents = $false; $excel.AskToUpdateLinks = $false
$wb = $null
try { $wb = $excel.Workbooks.Open($wbFile.FullName, 0, $false) } catch { WrDie "Could not open the workbook: $($_.Exception.Message)" }
$prevCalc = $xlCalcAutomatic; try { $prevCalc = $excel.Calculation; $excel.Calculation = $xlCalcManual } catch {}

$sheets = @(); foreach ($s in $wb.Worksheets) { $sheets += $s }
if ($SheetName) { $ws = $sheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1; if (-not $ws) { WrDie "No sheet named '$SheetName'." } }
elseif ($sheets.Count -eq 1 -or $NonInteractive) { $ws = $sheets[0] }
else { Write-Host ''; for ($i = 0; $i -lt $sheets.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $sheets[$i].Name) }; do { $pick = Read-Host "`n  Which sheet holds the data? (number)" } while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $sheets.Count); $ws = $sheets[[int]$pick - 1] }
WrGood "Sheet: $($ws.Name)"

# ============================================================================
WrHead 'Which columns'
$hdrRow = 1
if ($HeaderRow -gt 0) { $hdrRow = $HeaderRow }
elseif (-not $NonInteractive) { $x = Read-Host '  Which row holds the headers? (Enter = 1)'; if ($x -and ($x -as [int])) { $hdrRow = [int]$x } }

$used = $ws.UsedRange
$lastRow = $used.Row + $used.Rows.Count - 1
$lastCol = $used.Column + $used.Columns.Count - 1
$firstDataRow = $hdrRow + 1
$nRows = $lastRow - $firstDataRow + 1
if ($nRows -lt 1) { WrDie 'No data rows below the header row.' }

$headers = New-Object 'string[]' ($lastCol + 1)
for ($c = 1; $c -le $lastCol; $c++) { $v = $ws.Cells($hdrRow, $c).Value2; $headers[$c] = if ($null -eq $v) { '' } else { [string]$v } }

Write-Host ''
Write-Host "  Headers (row $hdrRow), $nRows data row(s):" -ForegroundColor White
for ($c = 1; $c -le $lastCol; $c++) { $nm = $headers[$c]; if (-not $nm) { $nm = '(blank)' }; Write-Host ("   [{0,3}]  {1,-3}  {2}" -f $c, (Get-ColLetter $c), $nm) }

function Ask-Column([string]$what, [bool]$optional, [string[]]$hints) {
    $guess = 0
    for ($c = 1; $c -le $lastCol; $c++) { $h = ($headers[$c] -replace '[^A-Za-z0-9]', '').ToUpper(); if (-not $h) { continue }; foreach ($hint in $hints) { if ($h -eq $hint -or $h -like "*$hint*") { $guess = $c; break } }; if ($guess) { break } }
    if ($NonInteractive) { if ($guess) { return $guess } elseif ($optional) { return 0 } else { WrDie "Could not auto-detect the $what column." } }
    while ($true) {
        $sfx = if ($optional) { "  (0 = none)" } else { "" }; $def = if ($guess) { " [Enter = $guess : $($headers[$guess])]" } else { "" }
        $ans = Read-Host "  Column for the $what$sfx$def"
        if (-not $ans) { if ($guess) { return $guess }; if ($optional) { return 0 }; WrWarn '   Required.'; continue }
        if ($optional -and ($ans.Trim() -eq '0' -or $ans.Trim() -match '^(none|no|n/a|na)$')) { return 0 }
        if ($ans -as [int]) { $n = [int]$ans; if ($n -ge 1 -and $n -le $lastCol) { return $n }; WrWarn '   Out of range.'; continue }
        for ($c = 1; $c -le $lastCol; $c++) { if ($headers[$c] -and $headers[$c].Trim().ToUpper() -eq $ans.Trim().ToUpper()) { return $c } }
        WrWarn '   No header by that name.'
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
    if ($qMap.Count -eq 1) { $fixedQuarter = @($qMap.Keys)[0]; WrWarn "No date column; using the only quarter present: $fixedQuarter" }
    elseif ($NonInteractive) { WrDie 'No date column and multiple quarters; cannot choose in -NonInteractive.' }
    else {
        Write-Host ''; WrWarn 'No date column. Which period does the whole sheet use?'
        WrStep ('Available: ' + (($qMap.Keys | Sort-Object | ForEach-Object { $_ -replace ' ','' }) -join '  '))
        while ($true) { $ans = Read-Host '  Period (YYYYQX, e.g. 2021Q3)'; $k = $null
            if ($ans -match '^\s*(?<y>(19|20)\d{2})\s*[-_ ]?[Qq]?\s*(?<q>[1-4])\s*$') { $k = "$($Matches['y']) Q$($Matches['q'])" }
            if ($k -and $qMap.ContainsKey($k)) { $fixedQuarter = $k; break }; WrWarn '   Not found. Pick from the list.' }
    }
}
$assumeNE = $false
if ($colState -eq 0) { if ($NonInteractive) { $assumeNE = $true } else { $a = Read-Host '  No state column. Treat EVERY row as Nebraska? (Y/N)'; if ($a -match '^[Yy]') { $assumeNE = $true } else { WrDie 'A state column is needed.' } } }

# ---- true last data row (guard against UsedRange bloat, common in .xlsb / heavily
#      formatted sheets, where UsedRange can claim ~1,048,576 rows). Use the real
#      bottom of the ADDRESS column instead so we don't read a million blank rows. ----
$xlUp = -4162
try {
    $sheetRowCount = [int]$ws.Rows.Count
    $lastByAddr = [int]$ws.Cells($sheetRowCount, $colAddr).End($xlUp).Row
    if ($lastByAddr -ge $firstDataRow -and $lastByAddr -lt $lastRow) {
        WrWarn ("UsedRange claimed {0:N0} rows; real data ends at row {1:N0}. Using the smaller." -f $lastRow, $lastByAddr)
        $lastRow = $lastByAddr; $nRows = $lastRow - $firstDataRow + 1
    }
} catch { }
WrGood ("Will process {0:N0} data row(s)." -f $nRows)

# ---- output options ----
Write-Host ''
Write-Host '  Result: press ENTER to append at the end, or type a column (e.g. T) to INSERT there.' -ForegroundColor White
$insertAt = 0
if ($PasteAt) { if ($PasteAt.Trim().ToUpper() -ne 'END') { $insertAt = Resolve-PasteLocation $PasteAt $headers $lastCol; if ($insertAt -lt 1) { WrDie "Could not understand paste location '$PasteAt'." } } }
elseif (-not $NonInteractive) { $ans = Read-Host "  Result location (Enter = end)"; if ($ans) { $insertAt = Resolve-PasteLocation $ans $headers $lastCol; if ($insertAt -lt 1) { WrWarn '   Not recognised - appending at end.'; $insertAt = 0 } } }
$writeFlag = [bool]$IncludeFlag
if (-not $IncludeFlag -and -not $NonInteractive) { $af = Read-Host '  Also add a Match Flag column? (y/N)'; if ($af -match '^[Yy]') { $writeFlag = $true } }
if ($insertAt -gt 0) { WrGood "Insert at column $(Get-ColLetter $insertAt)." } else { WrGood 'Append at end.' }
WrGood $(if ($writeFlag) { 'Columns: NE City Code (000) + Match Flag.' } else { 'Column: NE City Code only (000).' })

if (-not $NonInteractive) { $go = Read-Host "`n  Press Enter to run, or N to cancel"; if ($go -match '^[Nn]') { $wb.Close($false); $excel.Quit(); exit 0 } }

# ============================================================================
WrHead 'Backup'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$bakDir = Join-Path $folder '_Backups'; if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir | Out-Null }
$bak = Join-Path $bakDir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension)
Copy-Item -LiteralPath $wbFile.FullName -Destination $bak -Force
WrGood "Backup: _Backups\$(Split-Path $bak -Leaf)"

# ============================================================================
WrHead 'Reading the workpaper'
# Bulk column read. Excel hands back a whole chunk's .Value2 as one 2-D array in
# a single COM call; we flatten it with Array.Copy (native, fast) instead of a
# per-cell reflection loop, which is what made a big .xlsb crawl.
function Read-Col([int]$col, [string]$label) {
    $out = New-Object 'string[]' $nRows; $done = 0
    while ($done -lt $nRows) {
        $take = [Math]::Min($ReadChunk, $nRows - $done); $r1 = $firstDataRow + $done; $r2 = $r1 + $take - 1
        $vals = $ws.Range($ws.Cells($r1, $col), $ws.Cells($r2, $col)).Value2
        if ($take -eq 1) { $out[$done] = if ($null -eq $vals) { '' } else { [string]$vals } }
        else {
            $flat = New-Object 'object[]' $take
            try { [Array]::Copy($vals, $flat, $take) } catch { for ($i = 1; $i -le $take; $i++) { $flat[$i - 1] = Get-Nth $vals $i } }
            for ($i = 0; $i -lt $take; $i++) { $v = $flat[$i]; $out[$done + $i] = if ($null -eq $v) { '' } else { [string]$v } }
        }
        $done += $take
        if ($nRows -ge 20000) { Write-Progress -Activity "Reading $label" -Status ("{0:N0} / {1:N0}" -f $done, $nRows) -PercentComplete (100 * $done / $nRows) }
    }
    Write-Progress -Activity "Reading $label" -Completed
    return , $out
}
$aAddr = Read-Col $colAddr 'address'; $aCity = Read-Col $colCity 'city'; $aZip = Read-Col $colZip 'zip'
$aDate = if ($colDate -gt 0) { Read-Col $colDate 'date' } else { $null }
$aState = if ($assumeNE) { $null } else { Read-Col $colState 'state' }
WrGood "Read $nRows row(s)."

# ============================================================================
WrHead 'Grouping rows into quarters'
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
    if ($nRows -ge 20000 -and ($i % 50000) -eq 0) { Write-Progress -Activity 'Grouping rows into quarters' -PercentComplete (100 * $i / $nRows) -Status ("{0:N0} / {1:N0}" -f $i, $nRows) }
}
Write-Progress -Activity 'Grouping rows into quarters' -Completed
if ($groups.Count -eq 0) { WrWarn 'No rows matched a quarterly file.' } else { foreach ($k in ($groups.Keys | Sort-Object)) { WrStep ("{0}  ->  {1,8:N0} row(s)" -f $k, $groups[$k].Count) } }

# ============================================================================
WrHead 'Matching against the tax data'

# FAST PATH: pulling ~300k x 7 cells out of Excel one range at a time costs
# minutes (COM marshals every cell). Instead we have Excel write the sheet out
# once as a tab-delimited file and read that with .NET - typically 10-50x
# faster - and keep it in _Cache so later runs skip the export completely.
# Returns $null if anything goes wrong, so the caller falls back to COM.
function Read-QuarterFast($path) {
    $cacheDir = Join-Path $folder '_Cache'
    if (-not (Test-Path -LiteralPath $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir | Out-Null }
    $tsv = Join-Path $cacheDir ([IO.Path]::GetFileNameWithoutExtension($path) + '.txt')
    $src = Get-Item -LiteralPath $path
    $fresh = (Test-Path -LiteralPath $tsv) -and ((Get-Item -LiteralPath $tsv).LastWriteTimeUtc -ge $src.LastWriteTimeUtc)
    if (-not $fresh) {
        WrStep '      exporting tax file to a fast cache (one time per file)...'
        Write-Progress -Activity 'Exporting tax data' -Status 'Excel is writing the cache file...' -PercentComplete 30
        $xlUnicodeText = 42
        $qwb = $excel.Workbooks.Open($src.FullName, 0, $true)
        try {
            $qwb.Worksheets.Item(1).Copy()          # new workbook holding just that sheet
            $tmpWb = $excel.ActiveWorkbook
            $tmpWb.SaveAs($tsv, $xlUnicodeText)
            $tmpWb.Close($false)
        } finally { $qwb.Close($false) }
        Write-Progress -Activity 'Exporting tax data' -Completed
    } else { WrStep '      using cached tax data (delete _Cache to force a refresh)' }

    # ---- read the tab file with .NET ----
    $need = @{ lo = 2; hi = 3; e = 4; h = 7; concat = 22; key = 23; code = 24 }   # 0-based
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $sr = New-Object System.IO.StreamReader($tsv, [System.Text.Encoding]::Unicode)
    try { while ($null -ne ($ln = $sr.ReadLine())) { $lines.Add($ln) } } finally { $sr.Dispose() }
    if ($lines.Count -lt 2) { return $null }

    # header row: column G (index 6) says "Street Name"
    $hdrIdx = -1
    for ($i = 0; $i -lt [Math]::Min(5, $lines.Count); $i++) {
        $f = $lines[$i].Split("`t")
        if ($f.Count -gt 6 -and $f[6] -and $f[6].ToUpper().Contains('STREET')) { $hdrIdx = $i; break }
    }
    if ($hdrIdx -lt 0) { return $null }
    $hf = $lines[$hdrIdx].Split("`t")
    if ($hf.Count -gt 24 -and $hf[24] -and -not $hf[24].ToUpper().Contains('CITY CODE')) {
        throw "column Y of $(Split-Path $path -Leaf) is '$($hf[24])', not 'City Code (final)'"
    }

    $qn = $lines.Count - $hdrIdx - 1
    if ($qn -lt 1) { return $null }
    $res = @{ n = $qn; dr = (New-Object 'object[]' $qn); key = (New-Object 'object[]' $qn)
        lo = (New-Object 'object[]' $qn); hi = (New-Object 'object[]' $qn); e = (New-Object 'object[]' $qn)
        h = (New-Object 'object[]' $qn); code = (New-Object 'object[]' $qn); concat = $null }
    if ($anyZip9) { $res.concat = New-Object 'object[]' $qn }
    for ($i = 0; $i -lt $qn; $i++) {
        $f = $lines[$hdrIdx + 1 + $i].Split("`t")
        $res.dr[$i] = $hdrIdx + 2 + $i
        if ($f.Count -gt $need.key)  { $res.key[$i]  = $f[$need.key] }
        if ($f.Count -gt $need.lo)   { $res.lo[$i]   = $f[$need.lo] }
        if ($f.Count -gt $need.hi)   { $res.hi[$i]   = $f[$need.hi] }
        if ($f.Count -gt $need.e)    { $res.e[$i]    = $f[$need.e] }
        if ($f.Count -gt $need.h)    { $res.h[$i]    = $f[$need.h] }
        if ($f.Count -gt $need.code) { $res.code[$i] = $f[$need.code] }
        if ($null -ne $res.concat -and $f.Count -gt $need.concat) { $res.concat[$i] = $f[$need.concat] }
        if (($i % 50000) -eq 0) { Write-Progress -Activity 'Reading tax data' -Status ("{0:N0} / {1:N0}" -f $i, $qn) -PercentComplete (100 * $i / $qn) }
    }
    Write-Progress -Activity 'Reading tax data' -Completed
    return $res
}

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
        function RdCol([int]$col) {
            $arr = New-Object 'object[]' $qn; $done = 0
            while ($done -lt $qn) {
                $take = [Math]::Min($ReadChunk, $qn - $done); $r1 = $qFirst + $done; $r2 = $r1 + $take - 1
                $vals = $qws.Range($qws.Cells($r1, $col), $qws.Cells($r2, $col)).Value2
                if ($take -eq 1) { $arr[$done] = $vals }
                else {
                    $flat = New-Object 'object[]' $take
                    try { [Array]::Copy($vals, $flat, $take) } catch { for ($i = 1; $i -le $take; $i++) { $flat[$i - 1] = Get-Nth $vals $i } }
                    [Array]::Copy($flat, 0, $arr, $done, $take)
                }
                $done += $take
            }
            return , $arr
        }
        $dr = New-Object 'object[]' $qn; for ($i = 0; $i -lt $qn; $i++) { $dr[$i] = $qFirst + $i }
        $cols = @(
            @{ n = 'key';  c = 24 }, @{ n = 'lo'; c = 3 }, @{ n = 'hi'; c = 4 },
            @{ n = 'e';    c = 5 },  @{ n = 'h';  c = 8 }, @{ n = 'code'; c = 25 }
        )
        if ($anyZip9) { $cols += @{ n = 'concat'; c = 23 } }
        $res = @{ n = $qn; dr = $dr; concat = $null }
        $ci = 0
        foreach ($cd in $cols) {
            $ci++
            Write-Progress -Activity "Loading tax data ($qn rows)" -Status ("column {0} of {1}: {2}" -f $ci, $cols.Count, $cd.n) -PercentComplete (100 * $ci / $cols.Count)
            $res[$cd.n] = RdCol $cd.c
        }
        Write-Progress -Activity "Loading tax data ($qn rows)" -Completed
        return $res
    } finally { $qwb.Close($false) }
}

$qi = 0
foreach ($q in ($groups.Keys | Sort-Object)) {
    $qi++
    $rowIdx = $groups[$q]; $qTot = $rowIdx.Count
    Write-Host ''
    WrStep ("[{0}/{1}]  {2}   {3:N0} row(s)" -f $qi, $groups.Count, $q, $qTot)

    # ---- 1. distinct addresses (sales data repeats the same stores endlessly) ----
    $tA = Get-Date
    $uniqKeys = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $uAddr = New-Object 'System.Collections.Generic.List[string]'
    $uCity = New-Object 'System.Collections.Generic.List[string]'
    $uZip  = New-Object 'System.Collections.Generic.List[string]'
    $rowToU = New-Object 'int[]' $qTot
    $seen = 0; $ui = 0
    foreach ($ri in $rowIdx) {
        $ck = $aAddr[$ri] + '|' + $aCity[$ri] + '|' + $aZip[$ri]
        $u = 0
        if (-not $uniqKeys.TryGetValue($ck, [ref]$u)) {
            $u = $uAddr.Count; $uniqKeys[$ck] = $u
            $uAddr.Add($aAddr[$ri]); $uCity.Add($aCity[$ri]); $uZip.Add($aZip[$ri])
        }
        $rowToU[$ui] = $u; $ui++
        $seen++
        if (($seen % 100000) -eq 0) { Write-Progress -Activity "$q  -  finding distinct addresses" -Status ("{0:N0} / {1:N0}   ({2:N0} distinct)" -f $seen, $qTot, $uAddr.Count) -PercentComplete (100 * $seen / $qTot) }
    }
    Write-Progress -Activity "$q  -  finding distinct addresses" -Completed
    $nU = $uAddr.Count
    WrStep ("      {0,10:N0} distinct address(es)   ({1:N1}s)" -f $nU, ((Get-Date) - $tA).TotalSeconds)

    # ---- 2. parse them once, and work out which tax keys we actually need ----
    $tB = Get-Date
    $parses = New-Object 'object[]' $nU
    $parses2 = New-Object 'object[]' $nU      # repaired variant (or $null)
    $notes2 = New-Object 'string[]' $nU
    $needK = New-Object 'System.Collections.Generic.HashSet[string]'
    $needZ = New-Object 'System.Collections.Generic.HashSet[string]'
    for ($u = 0; $u -lt $nU; $u++) {
        $p = Get-NeParse $uAddr[$u] $uCity[$u] $uZip[$u]
        $parses[$u] = $p
        if ($p.W -ne '') { [void]$needK.Add($p.W + ' ' + $p.V) }
        if ($p.Plus4 -ne '') { [void]$needZ.Add($p.V + $p.Plus4 + '_') }
        $rep = Repair-NeAddress $uAddr[$u]
        if ($rep.Note) {
            $p2 = Get-NeParse $rep.Text $uCity[$u] $uZip[$u]
            $parses2[$u] = $p2; $notes2[$u] = $rep.Note
            if ($p2.W -ne '') { [void]$needK.Add($p2.W + ' ' + $p2.V) }
        }
        if (($u % 2000) -eq 0 -and $nU -gt 2000) { Write-Progress -Activity "$q  -  parsing addresses" -Status ("{0:N0} / {1:N0}" -f $u, $nU) -PercentComplete (100 * $u / $nU) }
    }
    Write-Progress -Activity "$q  -  parsing addresses" -Completed
    WrStep ("      {0,10:N0} street key(s) needed     ({1:N1}s)" -f $needK.Count, ((Get-Date) - $tB).TotalSeconds)

    # ---- 3. load the quarterly file ----
    $tC = Get-Date
    $qd = $null
    try { $qd = Read-QuarterFast $qMap[$q] }
    catch { WrWarn "      fast read unavailable ($($_.Exception.Message)); using the slower COM read"; $qd = $null }
    if ($null -eq $qd) {
        try { $qd = Read-QuarterColumns $qMap[$q] } catch { WrDie "Quarter $q failed: $($_.Exception.Message)" }
    }
    WrStep ("      {0,10:N0} tax row(s) read          ({1:N1}s)" -f $qd.n, ((Get-Date) - $tC).TotalSeconds)

    # ---- 4. index ONLY the keys this batch needs ----
    $tD = Get-Date
    Write-Progress -Activity "$q  -  indexing tax data" -Status "building lookup..." -PercentComplete 50
    $codeArr = New-Object 'object[]' $qd.n
    $cyd = 0.0
    for ($i = 0; $i -lt $qd.n; $i++) { $codeArr[$i] = if ([double]::TryParse("$($qd.code[$i])", [ref]$cyd)) { [int]$cyd } else { $null } }
    $index = New-NeIndex $qd.dr $qd.key $qd.lo $qd.hi $qd.e $qd.h $codeArr $qd.concat $needK $needZ
    Write-Progress -Activity "$q  -  indexing tax data" -Completed
    $qd = $null
    WrStep ("      {0,10:N0} street(s) indexed        ({1:N1}s)" -f $index.Map.Count, ((Get-Date) - $tD).TotalSeconds)

    # ---- 5. match each distinct address once (with the fuzzy retry) ----
    $tE = Get-Date
    $uCode = New-Object 'object[]' $nU
    $uFlag = New-Object 'string[]' $nU
    $repaired = 0
    for ($u = 0; $u -lt $nU; $u++) {
        $m = Find-NeCode $parses[$u] $index
        if (($m.Flag -eq 'NO MATCH' -or $m.Flag -eq 'FALLBACK') -and $null -ne $parses2[$u]) {
            $m2 = Find-NeCode $parses2[$u] $index
            if ($m2.Flag -eq 'OK' -or $m2.Flag -eq 'ZIP9' -or ($m2.Flag -eq 'FALLBACK' -and $m.Flag -eq 'NO MATCH')) {
                $m = @{ Code = $m2.Code; Flag = "$($m2.Flag) - REPAIRED: $($notes2[$u])" }; $repaired++
            }
        }
        $uCode[$u] = $m.Code; $uFlag[$u] = $m.Flag
        if (($u % 2000) -eq 0 -and $nU -gt 2000) { Write-Progress -Activity "$q  -  matching" -Status ("{0:N0} / {1:N0}" -f $u, $nU) -PercentComplete (100 * $u / $nU) }
    }
    Write-Progress -Activity "$q  -  matching" -Completed
    WrStep ("      {0,10:N0} matched{1}   ({2:N1}s)" -f $nU, $(if ($repaired) { " ($repaired fuzzy-repaired)" } else { '' }), ((Get-Date) - $tE).TotalSeconds)

    # ---- 6. fan the answers back out to every row ----
    $tF = Get-Date
    for ($j = 0; $j -lt $qTot; $j++) {
        $ri = $rowIdx[$j]; $u = $rowToU[$j]
        $finalCode[$ri] = $uCode[$u]; $finalFlag[$ri] = $uFlag[$u]
        if (($j % 100000) -eq 0 -and $qTot -gt 100000) { Write-Progress -Activity "$q  -  applying to rows" -Status ("{0:N0} / {1:N0}" -f $j, $qTot) -PercentComplete (100 * $j / $qTot) }
    }
    Write-Progress -Activity "$q  -  applying to rows" -Completed
    WrGood ("      {0,10:N0} row(s) done              ({1:N1}s total for {2})" -f $qTot, ((Get-Date) - $tA).TotalSeconds, $q)
    $index = $null; $parses = $null; $parses2 = $null; $uniqKeys = $null
    [System.GC]::Collect()
}
Write-Host ''
WrGood 'Matching complete.'

# ============================================================================
WrHead 'Writing the code onto the working paper'
if ($writeFlag) { $headerNames = @('NE City Code','Match Flag'); $columns = $finalCode, $finalFlag } else { $headerNames = @('NE City Code'); $columns = , $finalCode }
$excel.ScreenUpdating = $true
$startCol = Write-ResultBlock $ws $excel $hdrRow $firstDataRow $lastRow $lastCol $nRows $insertAt $headerNames $columns $ReadChunk $CodeNumberFormat 0
if ($insertAt -gt 0) { WrStep "Inserted at $(Get-ColLetter $startCol) (existing data shifted right)." } else { WrStep "Appended at $(Get-ColLetter $startCol)." }
WrStep "NE City Code -> column $(Get-ColLetter $startCol), formatted $CodeNumberFormat"
if ($writeFlag) { WrStep "Match Flag   -> column $(Get-ColLetter ($startCol + 1))" }

# ============================================================================
WrHead 'Building the Report sheet'
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
    WrGood "Report sheet written: '$($rep.Sheet)' ($($uniqueRows.Count) unique address(es))."
} catch { WrWarn "Could not build the Report sheet: $($_.Exception.Message)" }

# ============================================================================
WrHead 'Saving'
try { $excel.Calculation = $prevCalc } catch { $excel.Calculation = $xlCalcAutomatic }
$saved = $false
for ($a = 1; $a -le 4 -and -not $saved; $a++) { try { $wb.Save(); $saved = $true } catch { WrWarn "Save attempt $a failed: $($_.Exception.Message)"; Start-Sleep -Seconds ([Math]::Pow(2, $a)) } }
if (-not $saved) { $alt = Join-Path $folder ("{0}_RESULTS_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension); try { $wb.SaveAs($alt); $saved = $true; WrWarn "Saved a copy as $(Split-Path $alt -Leaf)." } catch {} }
if ($saved) { WrGood "Saved: $($wbFile.Name)" } else { Write-Host '  COULD NOT SAVE - do not close Excel until you save by hand.' -ForegroundColor Red }

# ============================================================================
WrHead 'Summary'
$tally = @{}
for ($i = 0; $i -lt $nRows; $i++) { $f = $finalFlag[$i]; if (-not $f) { $f = '(blank)' }; $b = ($f -split ' - ')[0].Trim(); if (-not $tally.ContainsKey($b)) { $tally[$b] = 0 }; $tally[$b]++ }
Write-Host ''
Write-Host ("   {0,-16} {1,10} {2,7}   {3}" -f 'FLAG', 'ROWS', 'SHARE', '') -ForegroundColor White
Write-Host ('   ' + ('-' * 62)) -ForegroundColor DarkGray
foreach ($k in ($tally.Keys | Sort-Object { -$tally[$_] })) {
    $pct = 100 * $tally[$k] / $nRows
    $bar = '#' * [Math]::Max(0, [Math]::Round($pct / 4))
    $col = switch -Wildcard ($k) { 'OK' { 'Green' } 'ZIP9' { 'Green' } 'NO CODE-CITY' { 'Gray' } 'OOS' { 'Gray' } 'NO DATE' { 'Gray' } 'NO TAX FILE*' { 'Gray' } default { 'Yellow' } }
    Write-Host ("   {0,-16} {1,10:N0} {2,6:N1}%   {3}" -f $k, $tally[$k], $pct, $bar) -ForegroundColor $col
}
Write-Host ''
WrGood ("Finished in {0:N1} min." -f ((Get-Date) - $StartTime).TotalMinutes)
$excel.Visible = $true; $excel.EnableEvents = $true
if (-not $NonInteractive) { Read-Host '  Excel is open with your results. Press Enter to close this window' }
