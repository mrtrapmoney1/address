<#
================================================================================
 Get-NeTaxCode.ps1   (v4)
 Nebraska city tax-jurisdiction code lookup.

 WHAT THIS IS
   It does NOT re-implement the matching logic. It feeds your addresses into the
   proven Excel template (Address_Tax_Engine.xlsx), lets the template's own
   formulas work them out, and carries the answers back to your working sheet.
   Excel does the thinking; this script does the moving.

 WHAT'S NEW IN v4  (see docs/CHANGES.md for the full write-up)
   1. CONCURRENCY. Quarters are processed in parallel, each in its own Excel
      process with its own copy of the engine. One quarter no longer waits on
      another. Set $MaxParallel = 1 for the old, strictly-sequential behaviour.
   2. THE ENGINE RUN IS BACKED UP. Pass 1 is the raw template answer; pass 2 is
      the repaired retry. The old script let pass 2 overwrite pass 1 in memory,
      so the "main bulk of the work" was lost. Now pass 1 is kept and written,
      value-pasted, to per-quarter backup sheets BEFORE the repair touches it.
   3. CLEAN OUTPUT, YOUR CHOICE OF LOCATION. On the working paper you either
      press Enter to append the result columns at the end of the sheet, or name
      a spot and the script INSERTS brand-new columns there (existing data is
      shifted right, never overwritten). Every column is written on its own so
      two columns can never swap places - the scramble you saw can't happen.
   4. IT ALWAYS SAVES. The working paper and the backup workbook are both saved
      before the script hands control back, with the path printed.

 HARD RULE
   Nothing existing is ever changed. Results go into brand-new columns and a
   timestamped backup of the whole workbook is taken first.

 REQUIREMENTS
   Windows PowerShell 5.1 and Microsoft Excel (32-bit or 64-bit).
================================================================================
#>

[CmdletBinding()]
param(
    # All optional - the script prompts for anything you don't pass. Passing
    # them lets you script the run unattended (e.g. from Task Scheduler).
    [string]$WorkingPaper,        # path to the working paper .xlsx
    [string]$SheetName,           # sheet inside it that holds the data
    [int]   $HeaderRow = 0,       # header row (0 = ask / default 1)
    [string]$PasteAt,             # 'END' to append, or a column letter/header to insert at
    [int]   $MaxParallel = 0,     # 0 = auto (CPU-based). 1 = old sequential behaviour.
    [switch]$NonInteractive       # never prompt; use params + sensible defaults
)

$ErrorActionPreference = 'Stop'
$script:StartTime = Get-Date

# ============================================================================
#  SETTINGS
# ============================================================================
$EngineFileName   = 'Address_Tax_Engine.xlsx'  # the clean copy of your template
$BatchSize        = 5000    # rows pushed through the template at a time.
                            # 32-bit Excel: 5000. 64-bit: 20000 is fine.
$ReadChunk        = 50000   # rows read from / written to the working paper at a time
$TryRepairs       = $true   # retry failed rows with a cleaned-up address (pass 2)
$BackupEngineRun  = $true   # write the pass-1 engine answers to per-quarter backup sheets
$KeepEngine       = $true   # save one engine copy per quarter so you can open and look
$ShowSample       = 8       # print this many finished rows to the screen as a sanity check
$CodeNumberFormat = 'General'   # how the code column LOOKS. '000' shows 094 instead of 94.
$DumpCsv          = 200     # also dump this many finished rows to _Results_check.csv
$NoCodeCityZero   = $false  # $true = write 0 instead of blank when a city levies no local tax
$MaxParallelCap   = 6       # never launch more than this many Excel processes at once

# ---------------------------------------------------------------- constants --
$xlCalcManual    = -4135
$xlCalcAutomatic = -4105
$xlPasteValues   = -4163
$xlDone          = 0
$xlShiftToRight  = -4161

# ------------------------------------------------------------------ helpers --
function Write-Head($t) { Write-Host ''; Write-Host ('=' * 74) -ForegroundColor DarkCyan; Write-Host "  $t" -ForegroundColor Cyan; Write-Host ('=' * 74) -ForegroundColor DarkCyan }
function Write-Step($t) { Write-Host "  $t" -ForegroundColor Gray }
function Write-Good($t) { Write-Host "  $t" -ForegroundColor Green }
function Write-Warn2($t){ Write-Host "  $t" -ForegroundColor Yellow }
function Write-Bad($t)  { Write-Host "  $t" -ForegroundColor Red }
function Write-Rule()   { Write-Host ('  ' + ('-' * 70)) -ForegroundColor DarkGray }

function Get-ScriptFolder {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    return (Get-Location).Path
}

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

$script:excel = $null
function Stop-Everything([string]$msg) {
    if ($msg) { Write-Bad $msg }
    if ($script:excel) {
        try { $script:excel.DisplayAlerts = $false } catch { }
        try { $script:excel.Quit() } catch { }
    }
    if (-not $NonInteractive) { Read-Host '  Press Enter to close' }
    exit 1
}

# ============================================================================
#  WORKER  -  everything needed to grind one quarter, packaged so it can run
#  either right here (sequential) or inside its own runspace (parallel). A
#  runspace cannot see the functions defined above, so the worker carries its
#  own copies of every helper it needs. Yes, that duplicates a little code -
#  it is the price of isolation, and isolation is what lets quarters run at
#  the same time without tripping over one shared Excel instance.
#
#  INPUT  (one hashtable): EnginePath, QuarterKey, QuarterFile, Indexes,
#         Addr, City, Zip, BatchSize, TryRepairs, KeepEngine, KeepEnginePath
#  OUTPUT (one hashtable): per-position arrays, aligned to $Indexes -
#         EngineCode/EngineRow/EngineFlag  (pass 1, the raw engine answer)
#         FinalCode /FinalRow /FinalFlag   (pass 2, after repair; == pass 1
#                                           when nothing was repaired)
#         Source, Error
# ============================================================================
$WorkerScript = {
    param($job)

    $xlCalcManual  = -4135
    $xlPasteValues = -4163
    $xlDone        = 0

    # ---- helpers (worker-local copies) -----------------------------------
    function Get-Nth($arr, [int]$i) {
        if ($null -eq $arr) { return $null }
        if ($arr -is [Array]) {
            if ($arr.Rank -eq 2) { return $arr.GetValue($i + $arr.GetLowerBound(0) - 1, $arr.GetLowerBound(1)) }
            return $arr.GetValue($i + $arr.GetLowerBound(0) - 1)
        }
        if ($i -eq 1) { return $arr }
        return $null
    }
    function Reset-CopyMode($xl) {
        try { $xl.CutCopyMode = $false; return } catch { }
        try { $xl.CutCopyMode = 0;      return } catch { }
        try { $xl.ActiveSheet.Cells(1048576, 16384).Copy() | Out-Null } catch { }
    }

    # ---- address repair (only used on rows the engine could not match) ----
    $SuffixTypos = @{
        'AT'='ST';'SR'='ST';'SST'='ST';'STR'='ST';'STRT'='ST';'STRET'='ST';'STEET'='ST'
        'STREE'='ST';'STREETT'='ST';'SREET'='ST';'STREER'='ST';'STEERT'='ST'
        'AVENEU'='AVE';'AVENU'='AVE';'AVENUW'='AVE';'AVNUE'='AVE';'AV'='AVE';'AVN'='AVE'
        'AVEN'='AVE';'AVEUE'='AVE';'DIRVE'='DR';'DRIV'='DR';'DRV'='DR';'DRIVW'='DR'
        'DRE'='DR';'DRIEV'='DR';'RAOD'='RD';'RODA'='RD';'ROAND'='RD';'BLV'='BLVD'
        'BVLD'='BLVD';'BOULEVARDE'='BLVD';'BOLEVARD'='BLVD';'BULEVARD'='BLVD'
        'CIRLE'='CIR';'CIRC'='CIR';'CRICLE'='CIR';'CIRCE'='CIR';'CRT'='CT';'COUT'='CT'
        'CORT'='CT';'COURTT'='CT';'COURTE'='CT';'LNE'='LN';'LAN'='LN';'PKY'='PKWY'
        'PARKWY'='PKWY';'PRKWY'='PKWY';'TERR'='TER';'TRAL'='TRL';'TRIAL'='TRL'
        'HIGHWY'='HWY';'HWAY'='HWY';'HIWAY'='HWY';'PLC'='PL';'PLAC'='PL';'PLZA'='PLZ'
    }
    $RealSuffixes = @('ST','AVE','BLVD','RD','DR','CT','LN','PL','PLZ','PKWY','TER','CIR',
                      'HWY','WAY','TRL','RDG','LK','HTS','CV','PT','LOOP','PATH','SQ','XING',
                      'VW','BND','CRK','PARK','CTR','HL','ROW','VLG','TRCE','MALL','SPUR','RUN',
                      'STREET','AVENUE','BOULEVARD','ROAD','DRIVE','COURT','LANE','PLACE','PLAZA',
                      'PARKWAY','TERRACE','CIRCLE','HIGHWAY','TRAIL','RIDGE','LAKE','HEIGHTS',
                      'COVE','POINT','SQUARE','CROSSING','VIEW','BEND','CREEK','CENTER','HILL',
                      'VILLAGE','TRACE','LANDING','MEADOWS','VISTA','EXPRESSWAY','TURNPIKE')
    $Directionals = @{
        'NORTH'='N';'SOUTH'='S';'EAST'='E';'WEST'='W';'NORTHEAST'='NE';'NORTHWEST'='NW'
        'SOUTHEAST'='SE';'SOUTHWEST'='SW';'N'='N';'S'='S';'E'='E';'W'='W';'NE'='NE'
        'NW'='NW';'SE'='SE';'SW'='SW'
    }
    $CardWords = @{
        'ONE'=1;'TWO'=2;'THREE'=3;'FOUR'=4;'FIVE'=5;'SIX'=6;'SEVEN'=7;'EIGHT'=8;'NINE'=9
        'TEN'=10;'ELEVEN'=11;'TWELVE'=12;'THIRTEEN'=13;'FOURTEEN'=14;'FIFTEEN'=15
        'SIXTEEN'=16;'SEVENTEEN'=17;'EIGHTEEN'=18;'NINETEEN'=19;'TWENTY'=20;'THIRTY'=30
        'FORTY'=40;'FOURTY'=40;'FIFTY'=50;'SIXTY'=60;'SEVENTY'=70;'EIGHTY'=80;'NINETY'=90
    }
    $OrdWords = @{
        'FIRST'=1;'SECOND'=2;'THIRD'=3;'FOURTH'=4;'FIFTH'=5;'SIXTH'=6;'SEVENTH'=7
        'EIGHTH'=8;'NINTH'=9;'TENTH'=10;'ELEVENTH'=11;'TWELFTH'=12;'THIRTEENTH'=13
        'FOURTEENTH'=14;'FIFTEENTH'=15;'SIXTEENTH'=16;'SEVENTEENTH'=17;'EIGHTEENTH'=18
        'NINETEENTH'=19;'TWENTIETH'=20;'THIRTIETH'=30;'FORTIETH'=40;'FIFTIETH'=50
        'SIXTIETH'=60;'SEVENTIETH'=70;'EIGHTIETH'=80;'NINETIETH'=90
    }
    function Get-OrdinalEnding([int]$n) {
        $h = $n % 100
        if ($h -eq 11 -or $h -eq 12 -or $h -eq 13) { return 'TH' }
        switch ($n % 10) { 1 { 'ST' } 2 { 'ND' } 3 { 'RD' } default { 'TH' } }
    }
    function Repair-Address([string]$raw) {
        $result = @{ Text = $raw; Note = '' }
        if (-not $raw) { return $result }
        $s = $raw.ToUpper().Replace([char]160, ' ')
        foreach ($ch in @(',','.','/','\','-','#')) { $s = $s.Replace($ch, ' ') }
        $toks = @($s -split '\s+' | Where-Object { $_ -ne '' })
        if ($toks.Count -eq 0) { return $result }
        $notes = New-Object System.Collections.Generic.List[string]
        if ($toks[0] -eq 'PO' -or ($toks.Count -gt 1 -and $toks[0] -eq 'P' -and $toks[1] -eq 'O')) { return $result }
        $first = $toks[0]
        if ($first -match '^(\d+)([A-Z]+)$') {
            $tail = $Matches[2]
            if ($tail -notin @('ST','ND','RD','TH')) {
                $rest = @(); if ($toks.Count -gt 1) { $rest = $toks[1..($toks.Count - 1)] }
                $toks = @($Matches[1]) + @($tail) + $rest
                $notes.Add('no space after house number')
            }
        }
        if ($toks.Count -gt 2) {
            $last = $toks[$toks.Count - 1]; $prev = $toks[$toks.Count - 2]
            if ($last -match '^\d+$' -and ($RealSuffixes -contains $prev -or $SuffixTypos.ContainsKey($prev))) {
                $toks = $toks[0..($toks.Count - 2)]; $notes.Add('stray unit number on the end')
            }
        }
        if ($toks.Count -gt 1) {
            $last = $toks[$toks.Count - 1]
            if ($SuffixTypos.ContainsKey($last)) { $toks[$toks.Count - 1] = $SuffixTypos[$last]; $notes.Add('misspelled street suffix') }
        }
        if ($toks.Count -gt 2) {
            $last = $toks[$toks.Count - 1]
            if ($Directionals.ContainsKey($last)) {
                $headIsDir = $false
                if ($toks.Count -gt 1 -and $Directionals.ContainsKey($toks[1])) { $headIsDir = $true }
                if (-not $headIsDir) {
                    $d = $Directionals[$last]
                    $toks = $toks[0..($toks.Count - 2)]
                    $toks = @($toks[0]) + @($d) + $toks[1..($toks.Count - 1)]
                    $notes.Add('direction on the wrong end')
                }
            }
        }
        if ($toks.Count -ge 3) {
            $begin = 1
            if ($Directionals.ContainsKey($toks[1])) { $begin = 2 }
            $total = 0; $used = 0; $ok = $false
            for ($i = $begin; $i -lt $toks.Count; $i++) {
                $w = $toks[$i]
                if ($CardWords.ContainsKey($w))    { $total += $CardWords[$w]; $used = $i }
                elseif ($OrdWords.ContainsKey($w)) { $total += $OrdWords[$w];  $used = $i; $ok = $true; break }
                else { break }
            }
            if ($ok -and $total -gt 0) {
                $tail = @(); if ($used -lt $toks.Count - 1) { $tail = @($toks[($used + 1)..($toks.Count - 1)]) }
                $tailOk = ($tail.Count -eq 0) -or ($tail.Count -eq 1 -and $RealSuffixes -contains $tail[0])
                if ($tailOk) {
                    $ordinal = "$total" + (Get-OrdinalEnding $total)
                    $head = @(); if ($begin -eq 2) { $head = @($toks[0], $toks[1]) } else { $head = @($toks[0]) }
                    $toks = $head + @($ordinal) + $tail
                    $notes.Add('street number spelled out')
                }
            }
        }
        $new = ($toks -join ' ').Trim()
        if ($new -ne ($raw.ToUpper().Trim()) -and $notes.Count -gt 0) { $result.Text = $new; $result.Note = ($notes -join '; ') }
        return $result
    }

    # ---- open this worker's own Excel + its own copy of the engine --------
    $err = ''
    $xl = $null; $ewb = $null
    $n = $job.Indexes.Count
    $EngineCode = New-Object 'object[]' $n
    $EngineRow  = New-Object 'object[]' $n
    $EngineFlag = New-Object 'string[]' $n
    $FinalCode  = New-Object 'object[]' $n
    $FinalRow   = New-Object 'object[]' $n
    $FinalFlag  = New-Object 'string[]' $n
    $Source     = New-Object 'string[]' $n

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
        $xl.EnableEvents = $false; $xl.AskToUpdateLinks = $false
        try { $xl.Calculation = $xlCalcManual } catch { }

        $engTemp = Join-Path $env:TEMP ("NeTaxEngine_{0}_{1}.xlsx" -f ($job.QuarterKey -replace '[^0-9A-Za-z]',''), [Guid]::NewGuid().ToString('N').Substring(0,8))
        Copy-Item -LiteralPath $job.EnginePath -Destination $engTemp -Force
        $ewb = $xl.Workbooks.Open($engTemp, 0, $false)

        $eAddr = $null; $eTax = $null
        foreach ($s in $ewb.Worksheets) {
            if ($s.Name -eq 'Addresses') { $eAddr = $s }
            if ($s.Name -eq 'TaxData')   { $eTax  = $s }
        }
        if (-not $eAddr -or -not $eTax) { throw "engine is missing an 'Addresses' or 'TaxData' sheet" }

        $bs = [Math]::Max(1000, [int]$job.BatchSize)
        $engineLastRow = 2 + $bs

        # trim the engine to exactly one batch of formula rows
        $eAddrUsedLast = $eAddr.UsedRange.Row + $eAddr.UsedRange.Rows.Count - 1
        if ($eAddrUsedLast -gt $engineLastRow) { $eAddr.Rows("$($engineLastRow + 1):$eAddrUsedLast").Delete() | Out-Null }
        $eAddr.Range("C3:E$engineLastRow").ClearContents() | Out-Null
        if ($engineLastRow -gt 3) { $eAddr.Range('F3:AL3').Copy($eAddr.Range("F4:AL$engineLastRow")) | Out-Null }
        Reset-CopyMode $xl

        # the parsing lists (AN:BB) must survive - an empty one silently poisons every row
        $cityList = $xl.WorksheetFunction.CountA($eAddr.Range('AN3:AN284'))
        $dirList  = $xl.WorksheetFunction.CountA($eAddr.Range('AX3:AY19'))
        $sufList  = $xl.WorksheetFunction.CountA($eAddr.Range('BA3:BB116'))
        if ($cityList -lt 1) { throw "engine city list AN3:AN284 is empty" }
        if ($dirList  -lt 1) { throw "engine direction list AX3:AY19 is empty" }
        if ($sufList  -lt 1) { throw "engine suffix list BA3:BB116 is empty" }

        # ---- load this quarter's TaxData ---------------------------------
        $qwb = $xl.Workbooks.Open($job.QuarterFile, 0, $true)
        $qws = $qwb.Worksheets.Item(1)
        $qHdr = 0
        for ($r = 1; $r -le 5; $r++) {
            $v = $qws.Cells($r, 7).Value2
            if ($v -and ([string]$v).ToUpper().Contains('STREET')) { $qHdr = $r; break }
        }
        if ($qHdr -eq 0) { $qwb.Close($false); throw "no 'Street Name' header found in $(Split-Path $job.QuarterFile -Leaf)" }
        $yv = [string]$qws.Cells($qHdr, 25).Value2
        if ($yv -and -not $yv.ToUpper().Contains('CITY CODE')) { $qwb.Close($false); throw "column Y of $(Split-Path $job.QuarterFile -Leaf) is '$yv', not 'City Code (final)'" }
        $qUsed  = $qws.UsedRange
        $qLast  = $qUsed.Row + $qUsed.Rows.Count - 1
        $qFirst = $qHdr + 1
        $qCount = $qLast - $qFirst + 1
        if ($qCount -lt 1) { $qwb.Close($false); throw "no data rows in $(Split-Path $job.QuarterFile -Leaf)" }
        if ($qCount -gt 399999) { $qwb.Close($false); throw "$(Split-Path $job.QuarterFile -Leaf) has $qCount rows - beyond the template's 399,999 limit" }
        $eTax.Range('A3:Y400001').ClearContents() | Out-Null
        $slice = 50000; $moved = 0
        while ($moved -lt $qCount) {
            $take = [Math]::Min($slice, $qCount - $moved)
            $r1 = $qFirst + $moved; $r2 = $r1 + $take - 1
            $qws.Range($qws.Cells($r1, 1), $qws.Cells($r2, 25)).Copy() | Out-Null
            $eTax.Range($eTax.Cells(3 + $moved, 1), $eTax.Cells(3 + $moved, 1)).PasteSpecial($xlPasteValues) | Out-Null
            Reset-CopyMode $xl
            $moved += $take
        }
        $qwb.Close($false)
        $lastTax = 2 + $qCount
        if ($xl.WorksheetFunction.CountA($eTax.Range("X3:X$lastTax")) -lt 1) { throw "TaxData key column X empty after load" }
        if ($xl.WorksheetFunction.CountA($eTax.Range("Y3:Y$lastTax")) -lt 1) { throw "TaxData code column Y empty after load" }

        # ---- push one batch through the template -------------------------
        $Invoke = {
            param($addr, $city, $zip)
            $m = $addr.Count
            $block = [Array]::CreateInstance([object], $m, 3)
            for ($i = 0; $i -lt $m; $i++) { $block.SetValue($addr[$i], $i, 0); $block.SetValue($city[$i], $i, 1); $block.SetValue($zip[$i], $i, 2) }
            $eAddr.Range("C3:E$engineLastRow").ClearContents() | Out-Null
            $eAddr.Range($eAddr.Cells(3, 3), $eAddr.Cells(2 + $m, 5)).Value2 = $block
            $eAddr.Calculate() | Out-Null
            $guard = 0
            while ($xl.CalculationState -ne $xlDone -and $guard -lt 6000) { Start-Sleep -Milliseconds 100; $guard++ }
            return @{
                F = $eAddr.Range($eAddr.Cells(3, 6), $eAddr.Cells(2 + $m, 6)).Value2
                G = $eAddr.Range($eAddr.Cells(3, 7), $eAddr.Cells(2 + $m, 7)).Value2
                H = $eAddr.Range($eAddr.Cells(3, 8), $eAddr.Cells(2 + $m, 8)).Value2
            }
        }

        $label = Split-Path $job.QuarterFile -Leaf
        $tot = $n; $done = 0
        $retry = New-Object 'System.Collections.Generic.List[int]'

        # ---- PASS 1: the address exactly as it appears -------------------
        while ($done -lt $tot) {
            $take = [Math]::Min($bs, $tot - $done)
            $a = New-Object 'string[]' $take; $c = New-Object 'string[]' $take; $z = New-Object 'string[]' $take
            for ($i = 0; $i -lt $take; $i++) { $a[$i] = $job.Addr[$done + $i]; $c[$i] = $job.City[$done + $i]; $z[$i] = $job.Zip[$done + $i] }
            $res = & $Invoke $a $c $z
            for ($i = 1; $i -le $take; $i++) {
                $p = $done + $i - 1
                $EngineCode[$p] = Get-Nth $res.F $i
                $EngineRow[$p]  = Get-Nth $res.G $i
                $fh = Get-Nth $res.H $i
                $EngineFlag[$p] = if ($null -eq $fh) { '' } else { [string]$fh }
                # pass 1 IS the starting point for the final answer
                $FinalCode[$p] = $EngineCode[$p]; $FinalRow[$p] = $EngineRow[$p]; $FinalFlag[$p] = $EngineFlag[$p]
                $Source[$p] = $label
                if ($job.TryRepairs -and ($EngineFlag[$p] -eq 'NO MATCH' -or $EngineFlag[$p] -eq 'FALLBACK')) { $retry.Add($p) }
            }
            $done += $take
        }

        # ---- PASS 2: retry the failures with a repaired address ----------
        if ($job.TryRepairs -and $retry.Count -gt 0) {
            $fixable = New-Object 'System.Collections.Generic.List[int]'
            $fixText = @{}; $fixNote = @{}
            foreach ($p in $retry) {
                $rep = Repair-Address $job.Addr[$p]
                if ($rep.Note) { $fixable.Add($p); $fixText[$p] = $rep.Text; $fixNote[$p] = $rep.Note }
            }
            if ($fixable.Count -gt 0) {
                $fa = $fixable.ToArray(); $done2 = 0
                while ($done2 -lt $fa.Count) {
                    $take = [Math]::Min($bs, $fa.Count - $done2)
                    $a = New-Object 'string[]' $take; $c = New-Object 'string[]' $take; $z = New-Object 'string[]' $take
                    for ($i = 0; $i -lt $take; $i++) { $p = $fa[$done2 + $i]; $a[$i] = $fixText[$p]; $c[$i] = $job.City[$p]; $z[$i] = $job.Zip[$p] }
                    $res = & $Invoke $a $c $z
                    for ($i = 1; $i -le $take; $i++) {
                        $p = $fa[$done2 + $i - 1]
                        $nf = Get-Nth $res.H $i
                        $fl = if ($null -eq $nf) { '' } else { [string]$nf }
                        $old = $EngineFlag[$p]
                        $better = ($fl -eq 'OK' -or $fl -eq 'ZIP9' -or ($fl -eq 'FALLBACK' -and $old -eq 'NO MATCH'))
                        if ($better) {
                            # ONLY the FINAL arrays change. The Engine* arrays keep pass 1 -
                            # that is the backup the old script threw away.
                            $FinalCode[$p] = Get-Nth $res.F $i
                            $FinalRow[$p]  = Get-Nth $res.G $i
                            $FinalFlag[$p] = "$fl - REPAIRED: $($fixNote[$p])"
                        }
                    }
                    $done2 += $take
                }
            }
        }

        if ($job.KeepEngine -and $job.KeepEnginePath) {
            try { $ewb.SaveAs($job.KeepEnginePath) } catch { }
        }
        $ewb.Close($false); $ewb = $null
        Remove-Item -LiteralPath $engTemp -Force -ErrorAction SilentlyContinue
    }
    catch { $err = $_.Exception.Message }
    finally {
        try { if ($ewb) { $ewb.Close($false) } } catch { }
        try { if ($xl)  { $xl.Quit() } } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) 2>$null | Out-Null
    }

    return @{
        QuarterKey = $job.QuarterKey
        Indexes    = $job.Indexes
        EngineCode = $EngineCode; EngineRow = $EngineRow; EngineFlag = $EngineFlag
        FinalCode  = $FinalCode;  FinalRow  = $FinalRow;  FinalFlag  = $FinalFlag
        Source     = $Source
        Error      = $err
        QCount     = $n
    }
}

# ============================================================================
#  1. FIND EVERYTHING
# ============================================================================
Write-Head 'Nebraska Tax Jurisdiction Code Lookup  (v4)'
$folder = Get-ScriptFolder
Write-Step "Folder: $folder"

$enginePath = Join-Path $folder $EngineFileName
if (-not (Test-Path -LiteralPath $enginePath)) { Stop-Everything "'$EngineFileName' was not found in this folder. See README.txt." }
Write-Good "Engine: $EngineFileName"

$addrFolder = Join-Path $folder '_Address'
if (-not (Test-Path -LiteralPath $addrFolder)) { Stop-Everything "The '_Address' folder was not found next to this script." }
$quarterFiles = Get-ChildItem -LiteralPath $addrFolder -File -Filter '*.xlsx' | Where-Object { $_.Name -notlike '~$*' }
$qMap = @{}
foreach ($f in $quarterFiles) {
    if ($f.Name -match '(?<y>(19|20)\d{2}).{0,3}?Q(?<q>[1-4])') {
        $k = "$($Matches['y']) Q$($Matches['q'])"
        if (-not $qMap.ContainsKey($k)) { $qMap[$k] = $f.FullName }
    }
}
if ($qMap.Count -eq 0) { Stop-Everything 'No quarterly files found in _Address.' }
Write-Good "_Address holds $($qMap.Count) quarterly file(s)."

# ---- pick the working paper ----
Write-Head '1. Working paper'
if ($WorkingPaper) {
    if (-not (Test-Path -LiteralPath $WorkingPaper)) { Stop-Everything "Working paper not found: $WorkingPaper" }
    $wbFile = Get-Item -LiteralPath $WorkingPaper
    Write-Good "Using: $($wbFile.Name)"
} else {
    $books = Get-ChildItem -LiteralPath $folder -File |
             Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xlsb)$' -and $_.Name -notlike '~$*' -and $_.Name -ne $EngineFileName -and $_.Name -notlike '_Engine_*' } |
             Sort-Object Name
    if ($books.Count -eq 0) { Stop-Everything 'No working paper found in this folder.' }
    if ($books.Count -eq 1) { $wbFile = $books[0]; Write-Good "Found: $($wbFile.Name)" }
    else {
        Write-Host ''
        for ($i = 0; $i -lt $books.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $books[$i].Name) }
        do { $pick = Read-Host "`n  Which workbook is the working paper? (number)" }
        while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $books.Count)
        $wbFile = $books[[int]$pick - 1]
    }
}

# ============================================================================
#  2. OPEN EXCEL  (the parent instance - reads input, writes output)
# ============================================================================
Write-Head '2. Opening the workbook'
$script:excel = New-Object -ComObject Excel.Application
$excel = $script:excel
$excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
$excel.EnableEvents = $false; $excel.AskToUpdateLinks = $false

$wb = $null
try { $wb = $excel.Workbooks.Open($wbFile.FullName, 0, $false) }
catch { Stop-Everything "Could not open the workbook: $($_.Exception.Message)" }

$prevCalc = $xlCalcAutomatic
try { $prevCalc = $excel.Calculation; $excel.Calculation = $xlCalcManual } catch { }

$sheets = @(); foreach ($s in $wb.Worksheets) { $sheets += $s }
if ($SheetName) {
    $ws = $sheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
    if (-not $ws) { Stop-Everything "No sheet named '$SheetName'." }
    Write-Good "Sheet: $($ws.Name)"
} elseif ($sheets.Count -eq 1) { $ws = $sheets[0]; Write-Good "Sheet: $($ws.Name)" }
else {
    Write-Host ''
    for ($i = 0; $i -lt $sheets.Count; $i++) { Write-Host ("   [{0}]  {1}" -f ($i + 1), $sheets[$i].Name) }
    do { $pick = Read-Host "`n  Which sheet holds the data? (number)" }
    while (-not ($pick -as [int]) -or [int]$pick -lt 1 -or [int]$pick -gt $sheets.Count)
    $ws = $sheets[[int]$pick - 1]
}

# ============================================================================
#  3. WHICH COLUMNS
# ============================================================================
Write-Head '3. Tell me which columns to use'
$hdrRow = 1
if ($HeaderRow -gt 0) { $hdrRow = $HeaderRow }
elseif (-not $NonInteractive) {
    $hdrRowIn = Read-Host '  Which row holds the column headers? (Enter = 1)'
    if ($hdrRowIn -and ($hdrRowIn -as [int])) { $hdrRow = [int]$hdrRowIn }
}

$used     = $ws.UsedRange
$lastRow  = $used.Row + $used.Rows.Count - 1
$lastCol  = $used.Column + $used.Columns.Count - 1
$firstDataRow = $hdrRow + 1
$nRows    = $lastRow - $firstDataRow + 1
if ($nRows -lt 1) { Stop-Everything 'No data rows found below the header row.' }

$headers = New-Object 'string[]' ($lastCol + 1)
for ($c = 1; $c -le $lastCol; $c++) {
    $v = $ws.Cells($hdrRow, $c).Value2
    $headers[$c] = if ($null -eq $v) { '' } else { [string]$v }
}
$script:headersRef = $headers
$script:lastColRef = $lastCol

Write-Host ''
Write-Host "  Detected headers (row $hdrRow), $nRows data row(s):" -ForegroundColor White
for ($c = 1; $c -le $lastCol; $c++) {
    $nm = $headers[$c]; if (-not $nm) { $nm = '(blank)' }
    Write-Host ("   [{0,3}]  {1,-3}  {2}" -f $c, (Get-ColLetter $c), $nm)
}

function Ask-Column([string]$what, [bool]$optional, [string[]]$hints) {
    $guess = 0
    for ($c = 1; $c -le $script:lastColRef; $c++) {
        $h = ($script:headersRef[$c] -replace '[^A-Za-z0-9]', '').ToUpper()
        if (-not $h) { continue }
        foreach ($hint in $hints) { if ($h -eq $hint -or $h -like "*$hint*") { $guess = $c; break } }
        if ($guess) { break }
    }
    if ($NonInteractive) { if ($guess) { return $guess } elseif ($optional) { return 0 } else { Stop-Everything "Could not auto-detect the $what column; re-run without -NonInteractive." } }
    while ($true) {
        $sfx = if ($optional) { "  (0 = there isn't one)" } else { "" }
        $def = if ($guess) { " [Enter = $guess : $($script:headersRef[$guess])]" } else { "" }
        $ans = Read-Host "  Column holding the $what$sfx$def"
        if (-not $ans) { if ($guess) { return $guess }; if ($optional) { return 0 }; Write-Warn2 '   That one is required.'; continue }
        if ($optional -and ($ans.Trim() -eq '0' -or $ans.Trim() -match '^(none|no|n/a|na)$')) { return 0 }
        if ($ans -as [int]) { $nn = [int]$ans; if ($nn -ge 1 -and $nn -le $script:lastColRef) { return $nn }; Write-Warn2 '   Out of range.'; continue }
        for ($c = 1; $c -le $script:lastColRef; $c++) { if ($script:headersRef[$c] -and $script:headersRef[$c].Trim().ToUpper() -eq $ans.Trim().ToUpper()) { return $c } }
        Write-Warn2 '   No header by that name. Try the number instead.'
    }
}

Write-Host ''
$colAddr  = Ask-Column 'STREET ADDRESS' $false @('ADDRESSLINE1','ADDRESS1','SHIPTOADDRESS','STREETADDRESS','ADDRESS','STREET')
$colCity  = Ask-Column 'CITY'           $false @('SHIPTOCITY','SOURCEDTOCITY','BILLTOCITY','CITY')
$colZip   = Ask-Column 'ZIP'            $false @('ZIPCODE','SHIPTOZIP','POSTALCODE','ZIP','POSTAL')
$colDate  = Ask-Column 'DATE (YYYYMM)'  $true  @('YYYYMM','PERIOD','INVOICEDATE','SALEDATE','TRANDATE','DATE')
$colState = Ask-Column 'STATE'          $true  @('SHIPTOSTATE','SOURCEDTOSTATE','BILLTOSTATE','STATE')

$fixedQuarter = $null
if ($colDate -eq 0) {
    $qKeys = @($qMap.Keys | Sort-Object)
    Write-Host ''
    Write-Warn2 'No date column, so every row will be treated as one period.'
    Write-Step ('Available in _Address:  ' + (($qKeys | ForEach-Object { $_ -replace ' ', '' }) -join '  '))
    if ($NonInteractive) { Stop-Everything 'No date column and no fixed quarter available in -NonInteractive mode.' }
    while ($true) {
        $ans = Read-Host '  Which period should the whole sheet assume? (YYYYQX, e.g. 2024Q4)'
        $k = $null
        if     ($ans -match '^\s*(?<y>(19|20)\d{2})\s*[-_ ]?[Qq]?\s*(?<q>[1-4])\s*$') { $k = "$($Matches['y']) Q$($Matches['q'])" }
        elseif ($ans -match '^\s*(?<y>(19|20)\d{2})\s*[-_ ]?(?<m>0[1-9]|1[0-2])\s*$') { $k = "{0} Q{1}" -f $Matches['y'], [Math]::Ceiling([int]$Matches['m'] / 3) }
        if (-not $k) { Write-Warn2 '   Format is YYYYQX, for example 2024Q4.'; continue }
        if ($qMap.ContainsKey($k)) { $fixedQuarter = $k; break }
        Write-Warn2 ("   Nothing in _Address for {0}." -f ($k -replace ' ', ''))
    }
}

$assumeNE = $false
if ($colState -eq 0) {
    if ($NonInteractive) { $assumeNE = $true }
    else {
        $a = Read-Host '  No state column given. Treat EVERY row as Nebraska? (Y/N)'
        if ($a -match '^[Yy]') { $assumeNE = $true } else { Stop-Everything 'A state column is needed then. Re-run and pick one.' }
    }
}

Write-Host ''
Write-Good "Address-> $(Get-ColLetter $colAddr)  $($headers[$colAddr])"
Write-Good "City   -> $(Get-ColLetter $colCity)  $($headers[$colCity])"
Write-Good "Zip    -> $(Get-ColLetter $colZip)  $($headers[$colZip])"
if ($fixedQuarter) { Write-Good "Date   -> no column; every row uses $fixedQuarter" } else { Write-Good "Date   -> $(Get-ColLetter $colDate)  $($headers[$colDate])" }
if ($assumeNE)     { Write-Good "State  -> every row treated as NE" } else { Write-Good "State  -> $(Get-ColLetter $colState)  $($headers[$colState])" }

# ============================================================================
#  3b. WHERE SHOULD THE RESULTS GO?
# ============================================================================
Write-Head '3b. Where should the results land?'
Write-Host '  Press ENTER to add the result columns at the END of the sheet,' -ForegroundColor White
Write-Host '  or type a column to INSERT them at (e.g. "T", or a header name).' -ForegroundColor White
Write-Host '  Either way the columns are brand-new - nothing is written over.' -ForegroundColor Gray

$insertAt = 0   # 0 = append at end; >0 = insert new columns starting at this column number
if ($PasteAt) {
    if ($PasteAt.Trim().ToUpper() -ne 'END') {
        $insertAt = if ($PasteAt -as [int]) { [int]$PasteAt } else { Get-ColNumber $PasteAt }
        if ($insertAt -lt 1) {
            for ($c = 1; $c -le $lastCol; $c++) { if ($headers[$c] -and $headers[$c].Trim().ToUpper() -eq $PasteAt.Trim().ToUpper()) { $insertAt = $c; break } }
        }
        if ($insertAt -lt 1) { Stop-Everything "Could not understand paste location '$PasteAt'." }
    }
} elseif (-not $NonInteractive) {
    $ans = Read-Host "`n  Result location (Enter = end of sheet)"
    if ($ans) {
        $insertAt = if ($ans -as [int]) { [int]$ans } else { Get-ColNumber $ans }
        if ($insertAt -lt 1) { for ($c = 1; $c -le $lastCol; $c++) { if ($headers[$c] -and $headers[$c].Trim().ToUpper() -eq $ans.Trim().ToUpper()) { $insertAt = $c; break } } }
        if ($insertAt -lt 1) { Write-Warn2 "   Didn't recognise that - appending at the end instead."; $insertAt = 0 }
    }
}
if ($insertAt -gt 0) { Write-Good "Results will be INSERTED at column $(Get-ColLetter $insertAt) (existing data shifts right)." }
else                 { Write-Good "Results will be APPENDED after the last used column." }

# ---- parallelism ----
if ($MaxParallel -le 0) {
    $cpu = 2; try { $cpu = [int]$env:NUMBER_OF_PROCESSORS } catch { }
    $MaxParallel = [Math]::Max(1, [Math]::Min($MaxParallelCap, [Math]::Floor($cpu / 2)))
}
$MaxParallel = [Math]::Min($MaxParallel, $MaxParallelCap)
Write-Good ("Concurrency -> up to {0} quarter(s) at once." -f $MaxParallel)

if (-not $NonInteractive) {
    $go = Read-Host "`n  Look good? Press Enter to run, or type N to cancel"
    if ($go -match '^[Nn]') { $wb.Close($false); $excel.Quit(); exit 0 }
}

# ============================================================================
#  4. BACKUP THE WHOLE WORKBOOK FIRST
# ============================================================================
Write-Head '4. Backup'
$stamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$bakDir = Join-Path $folder '_Backups'
if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir | Out-Null }
$bak = Join-Path $bakDir ("{0}_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension)
Copy-Item -LiteralPath $wbFile.FullName -Destination $bak -Force
Write-Good "Copy saved to _Backups\$(Split-Path $bak -Leaf)"

# ============================================================================
#  5. READ THE INPUT COLUMNS
# ============================================================================
Write-Head '5. Reading your data'
function Get-Nth($arr, [int]$i) {
    if ($null -eq $arr) { return $null }
    if ($arr -is [Array]) {
        if ($arr.Rank -eq 2) { return $arr.GetValue($i + $arr.GetLowerBound(0) - 1, $arr.GetLowerBound(1)) }
        return $arr.GetValue($i + $arr.GetLowerBound(0) - 1)
    }
    if ($i -eq 1) { return $arr }
    return $null
}
function Read-Column([int]$col) {
    $out = New-Object 'string[]' $nRows
    $done = 0
    while ($done -lt $nRows) {
        $take = [Math]::Min($ReadChunk, $nRows - $done)
        $r1 = $firstDataRow + $done; $r2 = $r1 + $take - 1
        $vals = $ws.Range($ws.Cells($r1, $col), $ws.Cells($r2, $col)).Value2
        for ($i = 1; $i -le $take; $i++) { $v = Get-Nth $vals $i; $out[$done + $i - 1] = if ($null -eq $v) { '' } else { [string]$v } }
        $done += $take
        Write-Progress -Activity "Reading column $(Get-ColLetter $col)" -PercentComplete (100 * $done / $nRows)
    }
    Write-Progress -Activity "Reading column $(Get-ColLetter $col)" -Completed
    return ,$out
}
$aAddr  = Read-Column $colAddr
$aCity  = Read-Column $colCity
$aZip   = Read-Column $colZip
$aDate  = if ($colDate  -gt 0) { Read-Column $colDate }  else { $null }
$aState = if ($assumeNE)       { $null }                 else { Read-Column $colState }
Write-Good "Read $nRows row(s)."

# ============================================================================
#  6. SORT ROWS INTO QUARTERS
# ============================================================================
Write-Head '6. Sorting rows into quarters'
function Get-Quarter([string]$v) {
    if (-not $v) { return $null }
    $s = ($v -replace '[^0-9]', '')
    if ($s.Length -ge 6) {
        $y = [int]$s.Substring(0, 4); $m = [int]$s.Substring(4, 2)
        if ($y -ge 1990 -and $y -le 2100 -and $m -ge 1 -and $m -le 12) { return ("{0} Q{1}" -f $y, [Math]::Ceiling($m / 3)) }
    }
    $d = 0.0
    if ([double]::TryParse($v, [ref]$d) -and $d -gt 20000 -and $d -lt 80000) { $dt = [DateTime]::FromOADate($d); return ("{0} Q{1}" -f $dt.Year, [Math]::Ceiling($dt.Month / 3)) }
    $dt2 = [DateTime]::MinValue
    if ([DateTime]::TryParse($v, [ref]$dt2)) { return ("{0} Q{1}" -f $dt2.Year, [Math]::Ceiling($dt2.Month / 3)) }
    return $null
}

# per-row results, kept as plain 1-D arrays (three separate arrays cannot be
# misread the way one reshaped 2-D array can)
$engCode = New-Object 'object[]' $nRows   # pass 1 - raw engine code
$engRow  = New-Object 'object[]' $nRows
$engFlag = New-Object 'string[]' $nRows
$finCode = New-Object 'object[]' $nRows   # pass 2 - after repair (== pass 1 when nothing changed)
$finRow  = New-Object 'object[]' $nRows
$finFlag = New-Object 'string[]' $nRows
$preFlag = New-Object 'string[]' $nRows   # only set for rows never sent to the engine
$outSrc  = New-Object 'string[]' $nRows

$groups = @{}
for ($i = 0; $i -lt $nRows; $i++) {
    if (-not $assumeNE) {
        $st = ($aState[$i] -replace '[^A-Za-z]', '').ToUpper()
        if ($st -and $st -ne 'NE' -and $st -ne 'NEBRASKA' -and $st -ne 'NEB') { $preFlag[$i] = 'OOS'; continue }
    }
    if (-not $aAddr[$i]) { $preFlag[$i] = 'NO ADDRESS'; continue }
    $q = if ($fixedQuarter) { $fixedQuarter } else { Get-Quarter $aDate[$i] }
    if (-not $q) { $preFlag[$i] = 'NO DATE'; continue }
    if (-not $qMap.ContainsKey($q)) { $preFlag[$i] = "NO TAX FILE - $($q -replace ' ','')"; continue }
    if (-not $groups.ContainsKey($q)) { $groups[$q] = New-Object 'System.Collections.Generic.List[int]' }
    $groups[$q].Add($i)
}
if ($groups.Count -eq 0) { Write-Warn2 'No rows could be matched to a quarterly file. Check the date column.' }
else { foreach ($k in ($groups.Keys | Sort-Object)) { Write-Step ("{0}  ->  {1,10:N0} row(s)" -f $k, $groups[$k].Count) } }

# ============================================================================
#  7. RUN THE QUARTERS  (concurrently)
# ============================================================================
Write-Head "7. Running $($groups.Count) quarter(s) through the engine"
$engineForWorkers = Join-Path $env:TEMP ("NeTaxEngineMaster_{0}.xlsx" -f $stamp)
Copy-Item -LiteralPath $enginePath -Destination $engineForWorkers -Force   # workers copy from here, never the original

# build one job per quarter
$jobs = @()
foreach ($q in ($groups.Keys | Sort-Object)) {
    $idx = $groups[$q].ToArray()
    $a = New-Object 'string[]' $idx.Count; $c = New-Object 'string[]' $idx.Count; $z = New-Object 'string[]' $idx.Count
    for ($i = 0; $i -lt $idx.Count; $i++) { $a[$i] = $aAddr[$idx[$i]]; $c[$i] = $aCity[$idx[$i]]; $z[$i] = $aZip[$idx[$i]] }
    $keepPath = if ($KeepEngine) { Join-Path $folder ("_Engine_last_run_{0}.xlsx" -f ($q -replace '[^0-9A-Za-z]','')) } else { $null }
    $jobs += @{
        EnginePath = $engineForWorkers; QuarterKey = $q; QuarterFile = $qMap[$q]
        Indexes = $idx; Addr = $a; City = $c; Zip = $z
        BatchSize = $BatchSize; TryRepairs = $TryRepairs; KeepEngine = $KeepEngine; KeepEnginePath = $keepPath
    }
}

$results = @()
$runParallel = ($MaxParallel -gt 1 -and $jobs.Count -gt 1)

if (-not $runParallel) {
    # -------- sequential: identical logic, one quarter after another --------
    $qn = 0
    foreach ($job in $jobs) {
        $qn++
        Write-Step ("[{0}/{1}] {2}  ({3:N0} rows)..." -f $qn, $jobs.Count, $job.QuarterKey, $job.Indexes.Count)
        $t0 = Get-Date
        $r = & $WorkerScript $job
        if ($r.Error) { Stop-Everything "Quarter $($job.QuarterKey) failed: $($r.Error)" }
        Write-Good ("      done in {0:N1}s" -f ((Get-Date) - $t0).TotalSeconds)
        $results += $r
    }
} else {
    # -------- parallel: a runspace pool, STA (COM needs it) -----------------
    Write-Step ("Launching up to {0} worker(s)..." -f $MaxParallel)
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxParallel, $iss, $Host)
    $pool.ApartmentState = 'STA'
    $pool.ThreadOptions  = 'ReuseThread'
    $pool.Open()
    $running = @()
    foreach ($job in $jobs) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($WorkerScript.ToString()).AddArgument($job)
        $running += [PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke(); Key = $job.QuarterKey; Count = $job.Indexes.Count }
        Write-Step ("   queued {0}  ({1:N0} rows)" -f $job.QuarterKey, $job.Indexes.Count)
    }
    $finished = 0
    while ($finished -lt $running.Count) {
        Start-Sleep -Milliseconds 250
        foreach ($r in $running) {
            if (-not $r.Done -and $r.Handle.IsCompleted) {
                $out = $r.PS.EndInvoke($r.Handle)
                $r | Add-Member -NotePropertyName Done -NotePropertyValue $true -Force
                $r.PS.Dispose()
                $res = $out | Select-Object -First 1
                if ($res.Error) { $pool.Close(); Stop-Everything "Quarter $($r.Key) failed: $($res.Error)" }
                $results += $res
                $finished++
                Write-Good ("      finished {0}  ({1:N0} rows)   [{2}/{3}]" -f $r.Key, $r.Count, $finished, $running.Count)
            }
        }
    }
    $pool.Close(); $pool.Dispose()
}
Remove-Item -LiteralPath $engineForWorkers -Force -ErrorAction SilentlyContinue

# ---- fold every quarter's results back into the master row arrays --------
foreach ($res in $results) {
    for ($p = 0; $p -lt $res.Indexes.Count; $p++) {
        $g = $res.Indexes[$p]
        $engCode[$g] = $res.EngineCode[$p]; $engRow[$g] = $res.EngineRow[$p]; $engFlag[$g] = $res.EngineFlag[$p]
        $finCode[$g] = $res.FinalCode[$p];  $finRow[$g] = $res.FinalRow[$p];  $finFlag[$g] = $res.FinalFlag[$p]
        $outSrc[$g]  = $res.Source[$p]
    }
}
Write-Good "All quarters complete."

# ============================================================================
#  8. NORMALISE  (turn raw engine output into clean, validated values)
# ============================================================================
Write-Head '8. Cleaning up the answers'
function Test-XlError($v) {
    if ($null -eq $v) { return $false }
    if ($v -is [int] -or $v -is [long] -or $v -is [double]) { $n = [double]$v; return ($n -le -2146826246 -and $n -ge -2146826288) }
    return $false
}

# Normalise ONE (code,row,flag) triple into a validated (code,row,flag). Used
# for both the final answer and the engine-backup answer, so they clean up the
# same way and can never disagree on formatting.
function Resolve-Answer($codeRaw, $rowRaw, $flagRaw, $preSet) {
    $fl = $preSet
    if (-not $fl) { $fl = if ($null -eq $flagRaw -or "$flagRaw" -eq '') { 'NO MATCH' } else { [string]$flagRaw } }
    if ((Test-XlError $codeRaw) -or (Test-XlError $rowRaw) -or (Test-XlError $flagRaw)) { $fl = 'ENGINE ERROR'; $codeRaw = $null; $rowRaw = $null }
    if ($fl -eq 'OOS') { $codeRaw = $null; $rowRaw = $null }
    $num = $null
    if ($null -ne $codeRaw -and "$codeRaw".Trim() -ne '') { $d = 0.0; if ([double]::TryParse("$codeRaw".Trim(), [ref]$d)) { $num = $d } }
    if ($null -eq $num -and $NoCodeCityZero -and $fl -like 'NO CODE-CITY*') { $num = 0 }
    if ($null -ne $num -and ($num -lt 0 -or $num -gt 999 -or $num -ne [Math]::Floor($num))) { $fl = "BAD CODE - engine returned '$codeRaw'"; $num = $null; $rowRaw = $null; $script:badCodes++ }
    return @{ Code = $num; Row = $rowRaw; Flag = $fl }
}

$script:badCodes = 0
$finalCode = New-Object 'object[]' $nRows; $finalRow = New-Object 'object[]' $nRows
$finalFlag = New-Object 'string[]' $nRows; $finalSrc = New-Object 'string[]' $nRows
$backCode  = New-Object 'object[]' $nRows; $backRow  = New-Object 'object[]' $nRows
$backFlag  = New-Object 'string[]' $nRows

for ($j = 0; $j -lt $nRows; $j++) {
    $pre = $preFlag[$j]
    $f = Resolve-Answer $finCode[$j] $finRow[$j] $finFlag[$j] $pre
    $finalCode[$j] = $f.Code; $finalRow[$j] = $f.Row; $finalFlag[$j] = $f.Flag; $finalSrc[$j] = $outSrc[$j]
    $preFlag[$j] = $f.Flag   # so the summary counts what was actually written
    $e = Resolve-Answer $engCode[$j] $engRow[$j] $engFlag[$j] $pre
    $backCode[$j] = $e.Code; $backRow[$j] = $e.Row; $backFlag[$j] = $e.Flag
}

# a CSV straight from memory, before Excel is written - if the sheet ever
# disagrees with this file, the bug is in the writing, not the matching
if ($DumpCsv -gt 0) {
    try {
        $csvPath = Join-Path $folder '_Results_check.csv'
        $n = [Math]::Min($DumpCsv, $nRows)
        $dump = for ($j = 0; $j -lt $n; $j++) {
            [PSCustomObject]@{ SheetRow = $firstDataRow + $j; Address = $aAddr[$j]; City = $aCity[$j]; Zip = $aZip[$j]
                Code = $finalCode[$j]; MatchRow = $finalRow[$j]; Flag = $finalFlag[$j]
                EngineCode = $backCode[$j]; EngineFlag = $backFlag[$j]; Source = $finalSrc[$j] }
        }
        $dump | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Step "First $n row(s) also dumped to _Results_check.csv"
    } catch { Write-Warn2 "Could not write the CSV check file: $($_.Exception.Message)" }
}

# ============================================================================
#  9. WRITE THE RESULT COLUMNS INTO THE WORKING PAPER
# ============================================================================
Write-Head '9. Writing results into the working paper'

# work out the first result column: either INSERT (shift existing right) or APPEND
$headerNames = @('NE City Code','Match Row #','Match Flag','Source File')
$nCols = $headerNames.Count
if ($insertAt -gt 0) {
    $startCol = $insertAt
    # insert brand-new blank columns so nothing is overwritten - existing data slides right
    for ($k = 0; $k -lt $nCols; $k++) { $ws.Columns($startCol).Insert($xlShiftToRight) | Out-Null }
    Write-Step "Inserted $nCols new column(s) at $(Get-ColLetter $startCol); existing data shifted right."
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
    Write-Step "Appending new columns starting at $(Get-ColLetter $startCol)"
}

for ($k = 0; $k -lt $nCols; $k++) { $ws.Cells($hdrRow, $startCol + $k).Value2 = $headerNames[$k] }

# one column at a time - a single column of values can never trade places with another
function Write-OneColumn([int]$col, $values, [string]$what) {
    $done = 0
    while ($done -lt $nRows) {
        $take = [Math]::Min($ReadChunk, $nRows - $done)
        $colBlock = [Array]::CreateInstance([object], $take, 1)
        for ($i = 0; $i -lt $take; $i++) { $colBlock.SetValue($values[$done + $i], $i, 0) }
        $r1 = $firstDataRow + $done; $r2 = $r1 + $take - 1
        $ws.Range($ws.Cells($r1, $col), $ws.Cells($r2, $col)).Value2 = $colBlock
        $done += $take
        Write-Progress -Activity "Writing $what" -PercentComplete (100 * $done / $nRows)
    }
    Write-Progress -Activity "Writing $what" -Completed
    Write-Step ("  {0,-14} -> column {1}" -f $what, (Get-ColLetter $col))
}
Write-OneColumn  $startCol        $finalCode 'NE City Code'
Write-OneColumn ($startCol + 1)   $finalRow  'Match Row #'
Write-OneColumn ($startCol + 2)   $finalFlag 'Match Flag'
Write-OneColumn ($startCol + 3)   $finalSrc  'Source File'
try { $ws.Range($ws.Cells($firstDataRow, $startCol), $ws.Cells($lastRow, $startCol)).NumberFormat = $CodeNumberFormat } catch { }

if ($script:badCodes -gt 0) {
    Write-Host ''
    Write-Bad ("{0:N0} row(s) came back as something other than a valid city code - flagged BAD CODE, left blank." -f $script:badCodes)
}

# ============================================================================
#  10. BACK UP THE ENGINE (PASS-1) RUN  -  one value-pasted sheet per quarter
# ============================================================================
if ($BackupEngineRun) {
    Write-Head '10. Backing up the engine (pass-1) run'
    $engBookPath = Join-Path $bakDir ("{0}_ENGINE_{1}.xlsx" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp)
    $bwb = $excel.Workbooks.Add()
    # remove the default extra sheets later; build one sheet per quarter
    $created = @()
    foreach ($q in ($groups.Keys | Sort-Object)) {
        $idx = $groups[$q].ToArray()
        $sheetName = ($q -replace ' ', '') + ' (engine)'
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $sh = $bwb.Worksheets.Add()
        $sh.Name = $sheetName
        $created += $sheetName
        # header
        $hdr = @('SheetRow','Address','City','Zip','Engine City Code','Engine Match Row #','Engine Flag','Source File')
        for ($k = 0; $k -lt $hdr.Count; $k++) { $sh.Cells(1, $k + 1).Value2 = $hdr[$k] }
        # build the whole block in memory, then value-write it column by column
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
                $sh.Range($sh.Cells(2, $k + 1), $sh.Cells(1 + $m, $k + 1)).Value2 = $blk   # values only - no formulas anywhere
            }
        }
        Write-Step ("  {0,-16} {1,8:N0} row(s)" -f $sheetName, $m)
    }
    # drop the blank default sheet(s) Excel created with the new workbook
    $toRemove = @()
    foreach ($s in $bwb.Worksheets) { if ($created -notcontains $s.Name) { $toRemove += $s } }
    foreach ($s in $toRemove) { if ($bwb.Worksheets.Count -gt 1) { $s.Delete() | Out-Null } }
    try { $bwb.SaveAs($engBookPath); $bwb.Close($true) } catch { Write-Warn2 "Could not save the engine backup: $($_.Exception.Message)" }
    Write-Good "Engine backup saved: _Backups\$(Split-Path $engBookPath -Leaf)"
    Write-Step "  (one value-pasted sheet per quarter; the repair pass cannot touch these)"
}

# ============================================================================
#  11. SAVE THE WORKING PAPER  (always, before we hand control back)
# ============================================================================
Write-Head '11. Saving'
$excel.ScreenUpdating = $true
try { $excel.Calculation = $prevCalc } catch { $excel.Calculation = $xlCalcAutomatic }
$saved = $false
for ($attempt = 1; $attempt -le 4 -and -not $saved; $attempt++) {
    try { $wb.Save(); $saved = $true } catch { Write-Warn2 "Save attempt $attempt failed: $($_.Exception.Message)"; Start-Sleep -Seconds ([Math]::Pow(2, $attempt)) }
}
if (-not $saved) {
    $alt = Join-Path $folder ("{0}_RESULTS_{1}{2}" -f [IO.Path]::GetFileNameWithoutExtension($wbFile.Name), $stamp, $wbFile.Extension)
    try { $wb.SaveAs($alt); $saved = $true; Write-Warn2 "Original was locked; saved a copy as $(Split-Path $alt -Leaf) instead." } catch { }
}
if ($saved) { Write-Good "Saved: $($wbFile.Name)" } else { Write-Bad 'COULD NOT SAVE. Your results are in the workbook in memory - do NOT close Excel until you save it by hand.' }

# ============================================================================
#  12. SUMMARY
# ============================================================================
Write-Head '12. Summary'
$tally = @{}
for ($i = 0; $i -lt $nRows; $i++) {
    $f = $preFlag[$i]; if (-not $f) { $f = '(blank)' }
    $b = ($f -split ' - ')[0].Trim()
    if (-not $tally.ContainsKey($b)) { $tally[$b] = 0 }
    $tally[$b]++
}
foreach ($k in ($tally.Keys | Sort-Object { -$tally[$_] })) { Write-Host ("   {0,-24} {1,10:N0}   {2,5:N1}%" -f $k, $tally[$k], (100 * $tally[$k] / $nRows)) }

Write-Host ''
Write-Host '  WHAT THE FLAGS MEAN' -ForegroundColor White
Write-Host '   OK / ZIP9        trust these - full match.' -ForegroundColor Green
Write-Host '   ... REPAIRED     matched only after cleanup (see the engine backup for the raw answer).' -ForegroundColor Yellow
Write-Host '   FALLBACK / NO MATCH / NO ADDRESS    fix by hand.' -ForegroundColor Red
Write-Host '   NO CODE-CITY / OOS / NO DATE / NO TAX FILE    normal, no action.' -ForegroundColor Gray

Write-Host ''
Write-Good ("Finished in {0:N1} minutes." -f ((Get-Date) - $script:StartTime).TotalMinutes)
Write-Step  "Whole-workbook backup: $bak"

$excel.Visible = $true
$excel.EnableEvents = $true
Write-Host ''
if (-not $NonInteractive) { Read-Host '  Excel is open with your results. Press Enter to close this window' }
