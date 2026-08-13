# Unit tests for the pure (non-COM) logic lifted verbatim from Get-NeTaxCode.ps1
$ErrorActionPreference = 'Stop'
$script:fail = 0; $script:pass = 0
function Assert($cond, $msg) { if ($cond) { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg" -ForegroundColor Red } }
function AssertEq($got, $want, $msg) { if ("$got" -eq "$want") { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg  (got '$got' want '$want')" -ForegroundColor Red } }

# ---------------- functions copied verbatim ----------------
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
function Get-Nth($arr, [int]$i) {
    if ($null -eq $arr) { return $null }
    if ($arr -is [Array]) {
        if ($arr.Rank -eq 2) { return $arr.GetValue($i + $arr.GetLowerBound(0) - 1, $arr.GetLowerBound(1)) }
        return $arr.GetValue($i + $arr.GetLowerBound(0) - 1)
    }
    if ($i -eq 1) { return $arr }
    return $null
}
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
function Test-XlError($v) {
    if ($null -eq $v) { return $false }
    if ($v -is [int] -or $v -is [long] -or $v -is [double]) { $n = [double]$v; return ($n -le -2146826246 -and $n -ge -2146826288) }
    return $false
}

# Repair-Address (verbatim)
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

# Resolve-Answer (verbatim, with script-scope deps)
$NoCodeCityZero = $false
$script:badCodes = 0
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

# ===================== TESTS =====================
Write-Host "`n== Get-ColLetter / Get-ColNumber ==" -ForegroundColor Cyan
AssertEq (Get-ColLetter 1)  'A'   'col 1 = A'
AssertEq (Get-ColLetter 20) 'T'   'col 20 = T'
AssertEq (Get-ColLetter 26) 'Z'   'col 26 = Z'
AssertEq (Get-ColLetter 27) 'AA'  'col 27 = AA'
AssertEq (Get-ColLetter 28) 'AB'  'col 28 = AB'
AssertEq (Get-ColLetter 54) 'BB'  'col 54 = BB'
foreach ($n in 1,20,26,27,28,54,702,703) { AssertEq (Get-ColNumber (Get-ColLetter $n)) $n "round-trip col $n" }
AssertEq (Get-ColNumber 'T')  20 'T = 20'
AssertEq (Get-ColNumber 'aa') 27 'aa = 27 (case-insensitive)'

Write-Host "`n== Get-Quarter ==" -ForegroundColor Cyan
AssertEq (Get-Quarter '202410') '2024 Q4' 'YYYYMM 202410'
AssertEq (Get-Quarter '202101') '2021 Q1' 'YYYYMM 202101'
AssertEq (Get-Quarter '202409') '2024 Q3' 'YYYYMM 202409'
AssertEq (Get-Quarter '2021-08-13') '2021 Q3' 'ISO date'
AssertEq (Get-Quarter '8/13/2021') '2021 Q3' 'US date'
AssertEq (Get-Quarter '') '' 'empty -> null'
AssertEq (Get-Quarter 'garbage') '' 'garbage -> null'
# Excel serial for 2021-08-13 = 44421
AssertEq (Get-Quarter '44421') '2021 Q3' 'Excel serial date'

Write-Host "`n== Get-Nth (array shapes) ==" -ForegroundColor Cyan
$flat = @('a','b','c')
AssertEq (Get-Nth $flat 1) 'a' '1-D first'
AssertEq (Get-Nth $flat 3) 'c' '1-D last'
$two = [Array]::CreateInstance([object], @(3,1), @(1,1))  # 1-based 3x1 like Excel
$two.SetValue('x',1,1); $two.SetValue('y',2,1); $two.SetValue('z',3,1)
AssertEq (Get-Nth $two 1) 'x' '2-D 1-based first'
AssertEq (Get-Nth $two 2) 'y' '2-D 1-based second'
AssertEq (Get-Nth $two 3) 'z' '2-D 1-based third'
AssertEq (Get-Nth 'single' 1) 'single' 'bare single-cell value'
Assert  ($null -eq (Get-Nth $null 1)) 'null array -> null'

Write-Host "`n== Test-XlError ==" -ForegroundColor Cyan
Assert (Test-XlError (-2146826246)) '#N/A code is an error'      # 2042 -> -2146826246
Assert (Test-XlError (-2146826281)) 'DIV/0 code is an error'
Assert (-not (Test-XlError 0))    '0 is not an error'
Assert (-not (Test-XlError 94))   '94 is not an error'
Assert (-not (Test-XlError 'OK')) 'string is not an error'
Assert (-not (Test-XlError $null))'null is not an error'

Write-Host "`n== Repair-Address ==" -ForegroundColor Cyan
AssertEq (Repair-Address '123MAIN ST').Text '123 MAIN ST' 'no space after house number'
AssertEq (Repair-Address '123MAIN ST').Note 'no space after house number' '  note'
AssertEq (Repair-Address '100 MAIN ST N').Text '100 N MAIN ST' 'direction on wrong end (with house #)'
AssertEq (Repair-Address '223 Mill Street South').Text '223 S MILL STREET' 'direction wrong end (real sample, matched engine)'
AssertEq (Repair-Address '147TH STREET 100').Text '147TH STREET' 'stray unit number on end'
AssertEq (Repair-Address '100 FORTY SECOND ST').Text '100 42ND ST' 'spelled-out street number (with house #)'
AssertEq (Repair-Address 'N FORTY SECOND ST').Text 'N 42ND ST' 'spelled-out with leading dir'
AssertEq (Repair-Address '100 MAIN STREE').Text '100 MAIN ST' 'misspelled suffix'
$po = Repair-Address 'PO BOX 100'
AssertEq $po.Note '' 'PO Box left alone'
$second = Repair-Address 'SECOND CHANCE DR'
AssertEq $second.Note '' 'SECOND CHANCE DR not mangled (guard works)'
$clean = Repair-Address '2225 Q ST'
AssertEq $clean.Note '' 'already-clean address unchanged'

Write-Host "`n== Resolve-Answer ==" -ForegroundColor Cyan
$r = Resolve-Answer 0 236652 'OK' $null
AssertEq $r.Code 0 'OK code 0'; AssertEq $r.Flag 'OK' 'OK flag'
$r = Resolve-Answer 94 24820 'OK' $null
AssertEq $r.Code 94 'code 94'
$r = Resolve-Answer $null $null $null 'OOS'
AssertEq $r.Flag 'OOS' 'preset OOS wins'; Assert ($null -eq $r.Code) 'OOS code blank'
$r = Resolve-Answer 12345 999 'OK' $null
AssertEq $r.Flag "BAD CODE - engine returned '12345'" 'bad code flagged'; Assert ($null -eq $r.Code) 'bad code blanked'
$r = Resolve-Answer (-2146826246) 5 'OK' $null
AssertEq $r.Flag 'ENGINE ERROR' 'xl error -> ENGINE ERROR'
$r = Resolve-Answer $null $null 'NO MATCH' $null
AssertEq $r.Flag 'NO MATCH' 'no match passes through'
$r = Resolve-Answer '' '' '' $null
AssertEq $r.Flag 'NO MATCH' 'empty engine flag -> NO MATCH'

Write-Host "`n---------------------------------------------" -ForegroundColor DarkGray
Write-Host ("  PASS: {0}   FAIL: {1}" -f $script:pass, $script:fail) -ForegroundColor $(if($script:fail){'Red'}else{'Green'})
if ($script:fail) { exit 1 }

exit 0
