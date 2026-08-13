<#
================================================================================
 NeTaxPaste.ps1
 The workbook-writing half of the pipeline, split out on its own so it can be
 tested against a throwaway workbook WITHOUT the engine or a real run.

 This is the code that scrambled before (code+row jammed into the flag cell).
 Everything here is built to make that impossible:
   * every result column is written as its own (N x 1) object array - a single
     column can never trade places with another;
   * INSERT shifts existing data right, it never writes over it;
   * the engine backup is values only, one sheet per quarter, no formulas.

 Dot-source it:   . .\NeTaxPaste.ps1
 Both Get-NeTaxCode.ps1 (production) and tests\Test-Paste.ps1 (verification)
 call these exact functions, so the test proves the real thing.
================================================================================
#>

# ---- Excel constants used by the paste ----
$script:NP_xlShiftToRight = -4161
$script:NP_xlOpenXMLWorkbook = 51

function Get-ColLetter([int]$n) {
    $s = ''
    while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int](($n - $m) / 26) }
    return $s
}
function Get-ColNumber([string]$letters) {
    $letters = ($letters -replace '[^A-Za-z]', '').ToUpper()
    if (-not $letters) { return 0 }
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

# Turn whatever the user typed for a paste location into a column number.
# '' or 'END' -> 0 (append at the end). A letter, a number, or a header name
# -> that column. Returns -1 if it can't be resolved.
function Resolve-PasteLocation {
    param([string]$text, [string[]]$headers, [int]$lastCol)
    if (-not $text -or $text.Trim().ToUpper() -eq 'END') { return 0 }
    $t = $text.Trim()
    if ($t -as [int]) { return [int]$t }
    $asCol = Get-ColNumber $t
    if ($asCol -ge 1) {
        # only treat it as a column letter if it's plausibly a letter token
        if ($t -match '^[A-Za-z]{1,3}$') { return $asCol }
    }
    for ($c = 1; $c -le $lastCol; $c++) {
        if ($headers[$c] -and $headers[$c].Trim().ToUpper() -eq $t.ToUpper()) { return $c }
    }
    return -1
}

# Decide the first result column and lay down the header cells.
#   $insertAt = 0  -> append after the last used column (skipping any stray
#                     non-empty columns so we never land on top of something)
#   $insertAt > 0  -> INSERT that many brand-new columns there; existing data
#                     slides right and is never overwritten.
# Returns the starting column number. Writes ONLY the header row here.
function Add-ResultColumns {
    param($ws, $excel, [int]$hdrRow, [int]$lastRow, [int]$lastCol, [int]$insertAt, [string[]]$headerNames)
    $nCols = $headerNames.Count
    if ($insertAt -gt 0) {
        $startCol = $insertAt
        for ($k = 0; $k -lt $nCols; $k++) { $ws.Columns($startCol).Insert($script:NP_xlShiftToRight) | Out-Null }
    } else {
        $startCol = $lastCol + 1
        $safe = $false
        while (-not $safe) {
            $safe = $true
            for ($c = $startCol; $c -lt $startCol + $nCols; $c++) {
                $probe = $ws.Range($ws.Cells(1, $c), $ws.Cells([Math]::Min($lastRow, 1048576), $c))
                if ($excel.WorksheetFunction.CountA($probe) -gt 0) { $startCol = $c + 1; $safe = $false; break }
            }
        }
    }
    for ($k = 0; $k -lt $nCols; $k++) { $ws.Cells($hdrRow, $startCol + $k).Value2 = $headerNames[$k] }
    return $startCol
}

# Write ONE column of values, chunked. The block is always an (N x 1) 2-D
# object array assigned to an N-row-by-1-column range: an unambiguous shape
# that Excel cannot re-interpret. This is the anti-scramble guarantee.
function Write-ColumnValues {
    param($ws, [int]$col, [int]$firstDataRow, [int]$nRows, $values, [int]$readChunk = 50000)
    if ($readChunk -lt 1) { $readChunk = 50000 }
    $done = 0
    while ($done -lt $nRows) {
        $take = [Math]::Min($readChunk, $nRows - $done)
        $blk = [Array]::CreateInstance([object], $take, 1)
        for ($i = 0; $i -lt $take; $i++) { $blk.SetValue($values[$done + $i], $i, 0) }
        $r1 = $firstDataRow + $done; $r2 = $r1 + $take - 1
        $ws.Range($ws.Cells($r1, $col), $ws.Cells($r2, $col)).Value2 = $blk
        $done += $take
    }
}

# Write the whole result block: header + one column per array in $columns.
# $columns is an array of arrays, in the same order as $headerNames.
# Returns the starting column so the caller can report / format it.
function Write-ResultBlock {
    param(
        $ws, $excel, [int]$hdrRow, [int]$firstDataRow, [int]$lastRow, [int]$lastCol,
        [int]$nRows, [int]$insertAt, [string[]]$headerNames, $columns,
        [int]$readChunk = 50000, [string]$codeNumberFormat = 'General', [int]$codeColumnOffset = 0
    )
    # Guard against the @()-flatten footgun: @($oneArray) unwraps to that array's
    # ELEMENTS, so a single column silently arrives as N scalars instead of one
    # column. Pass one column as (,$array). Fail loudly if the counts disagree.
    if ($null -eq $columns) { throw "Write-ResultBlock: `$columns is null." }
    if ($columns.Count -ne $headerNames.Count) {
        throw ("Write-ResultBlock: got {0} value-column(s) for {1} header(s). A single column must be passed with the unary comma, e.g. (,`$array)." -f $columns.Count, $headerNames.Count)
    }
    $startCol = Add-ResultColumns $ws $excel $hdrRow $lastRow $lastCol $insertAt $headerNames
    for ($k = 0; $k -lt $headerNames.Count; $k++) {
        Write-ColumnValues $ws ($startCol + $k) $firstDataRow $nRows $columns[$k] $readChunk
    }
    if ($codeColumnOffset -ge 0) {
        try { $ws.Range($ws.Cells($firstDataRow, $startCol + $codeColumnOffset), $ws.Cells($lastRow, $startCol + $codeColumnOffset)).NumberFormat = $codeNumberFormat } catch { }
    }
    return $startCol
}

# Build the engine (pass-1) backup workbook: one VALUE-PASTED sheet per quarter.
# Nothing here is a formula, so a later re-sort or repair can never disturb it.
# Returns @{ Path; Sheets } .
function New-EngineBackup {
    param(
        $excel, [string]$path, [string[]]$quarterOrder, [hashtable]$groups,
        [int]$firstDataRow, $aAddr, $aCity, $aZip, $backCode, $backRow, $backFlag, $outSrc
    )
    $bwb = $excel.Workbooks.Add()
    $created = @()
    foreach ($q in $quarterOrder) {
        if (-not $groups.ContainsKey($q)) { continue }
        $idx = $groups[$q].ToArray()
        $sheetName = ($q -replace ' ', '') + ' (engine)'
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        # guard against an accidental duplicate 31-char truncation
        $suffix = 1
        while ($created -contains $sheetName) { $sheetName = (($q -replace ' ', '') + "_$suffix"); if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0,31) }; $suffix++ }
        $sh = $bwb.Worksheets.Add()
        $sh.Name = $sheetName
        $created += $sheetName
        $hdr = @('SheetRow','Address','City','Zip','Engine City Code','Engine Match Row #','Engine Flag','Source File')
        for ($k = 0; $k -lt $hdr.Count; $k++) { $sh.Cells(1, $k + 1).Value2 = $hdr[$k] }
        $m = $idx.Count
        if ($m -gt 0) {
            $cSheetRow = New-Object 'object[]' $m; $cAddr = New-Object 'object[]' $m; $cCity = New-Object 'object[]' $m; $cZip = New-Object 'object[]' $m
            $cCode = New-Object 'object[]' $m; $cRow = New-Object 'object[]' $m; $cFlag = New-Object 'object[]' $m; $cSrc = New-Object 'object[]' $m
            for ($i = 0; $i -lt $m; $i++) {
                $g = $idx[$i]
                $cSheetRow[$i] = $firstDataRow + $g; $cAddr[$i] = $aAddr[$g]; $cCity[$i] = $aCity[$g]; $cZip[$i] = $aZip[$g]
                $cCode[$i] = $backCode[$g]; $cRow[$i] = $backRow[$g]; $cFlag[$i] = $backFlag[$g]; $cSrc[$i] = $outSrc[$g]
            }
            $cols = @($cSheetRow,$cAddr,$cCity,$cZip,$cCode,$cRow,$cFlag,$cSrc)
            for ($k = 0; $k -lt $cols.Count; $k++) {
                $blk = [Array]::CreateInstance([object], $m, 1)
                for ($i = 0; $i -lt $m; $i++) { $blk.SetValue($cols[$k][$i], $i, 0) }
                $sh.Range($sh.Cells(2, $k + 1), $sh.Cells(1 + $m, $k + 1)).Value2 = $blk
            }
        }
    }
    $toRemove = @()
    foreach ($s in $bwb.Worksheets) { if ($created -notcontains $s.Name) { $toRemove += $s } }
    foreach ($s in $toRemove) { if ($bwb.Worksheets.Count -gt 1) { $s.Delete() | Out-Null } }
    $bwb.SaveAs($path, $script:NP_xlOpenXMLWorkbook)
    $bwb.Close($false)
    return @{ Path = $path; Sheets = $created }
}

# Build/refresh a "Report" analysis sheet in the SAME workbook. Since there is
# no engine sheet to eyeball anymore, this gives the flag breakdown and a table
# of UNIQUE addresses with their code/flag/count. It replaces any existing sheet
# of the same name (it is our own output, never customer data) and writes the
# whole grid in ONE value assignment so nothing can scramble.
#   $flagRows   : array of [pscustomobject]@{ Flag; Count; Pct }
#   $uniqueRows : array of [pscustomobject]@{ Address; City; Zip; Code; Flag; Count }
function New-NeReport {
    param($excel, $wb, [string]$sheetName, $flagRows, $uniqueRows, [string]$codeFormat = '000')
    $existing = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $sheetName) { $existing = $s } }
    if ($existing) { $existing.Delete() | Out-Null }
    $sh = $wb.Worksheets.Add()
    $sh.Name = $sheetName

    $rows = New-Object 'System.Collections.Generic.List[object]'
    $rows.Add(@('City Code Report', '', '', '', '', ''))
    $rows.Add(@('Generated', (Get-Date).ToString('yyyy-MM-dd HH:mm'), '', '', '', ''))
    $rows.Add(@('', '', '', '', '', ''))
    $rows.Add(@('FLAG SUMMARY (all rows)', '', '', '', '', ''))
    $rows.Add(@('Flag', 'Count', 'Percent', '', '', ''))
    foreach ($f in $flagRows) { $rows.Add(@($f.Flag, $f.Count, ('{0:N1}%' -f $f.Pct), '', '', '')) }
    $rows.Add(@('', '', '', '', '', ''))
    $rows.Add(@("UNIQUE ADDRESSES ($($uniqueRows.Count))", '', '', '', '', ''))
    $rows.Add(@('Address', 'City', 'Zip', 'NE City Code', 'Flag', 'Count'))
    $firstUnique = $rows.Count + 1   # 1-based sheet row of the first unique-address data row
    foreach ($u in $uniqueRows) { $rows.Add(@($u.Address, $u.City, $u.Zip, $u.Code, $u.Flag, $u.Count)) }

    # NOTE: PowerShell variables are case-insensitive - do NOT use $R and $r
    # together (a loop counter $r would clobber a row-count $R). Distinct names.
    $nrows = $rows.Count
    $arr = [Array]::CreateInstance([object], $nrows, 6)
    for ($ir = 0; $ir -lt $nrows; $ir++) { $line = $rows[$ir]; for ($ic = 0; $ic -lt 6; $ic++) { $arr.SetValue($line[$ic], $ir, $ic) } }
    $sh.Range($sh.Cells(1, 1), $sh.Cells($nrows, 6)).Value2 = $arr

    # 000 format on the code column of the unique table
    if ($nrows -ge $firstUnique) { try { $sh.Range($sh.Cells($firstUnique, 4), $sh.Cells($nrows, 4)).NumberFormat = $codeFormat } catch { } }
    # widen the address column a bit if the host supports it (real Excel; mock ignores)
    try { $sh.Columns(1).ColumnWidth = 34; $sh.Columns(2).ColumnWidth = 16 } catch { }
    return @{ Sheet = $sheetName; Rows = $R; FirstUnique = $firstUnique }
}
