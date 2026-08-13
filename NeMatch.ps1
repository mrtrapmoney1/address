<#
================================================================================
 NeMatch.ps1
 A faithful, pure-PowerShell port of the Address_Tax_Engine "Addresses" sheet
 formulas (columns J..AB and F/G/H). No Excel engine, no COM - just the same
 logic the template uses, so it produces the same code the sheet would.

 The three lookup lists (taxing cities, directions, street suffixes) are the
 exact ones from the engine, extracted into NeDicts.ps1.

 PUBLIC FUNCTIONS
   Get-NeParse   $addr $city $zip     -> parsed pieces (house, street key, ...)
   Find-NeCode   $parse $index        -> @{ Code; Flag; Row }
   New-NeIndex   (columns)            -> the in-memory tax-data index

 The FLAGS match the engine exactly:
   OK  ZIP9  FALLBACK  NO MATCH  NO CODE-CITY   (OOS/NO ADDRESS handled upstream)
================================================================================
#>

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'NeDicts.ps1')

$script:NeUnitMarkers = @(' APT ',' APARTMENT ',' UNIT ',' STE ',' SUITE ',' LOT ',' TRLR ',' RM ',' FL ',' BLDG ',' SPC ')

# Excel TRIM: collapse internal runs of spaces to one, trim the ends.
function ConvertTo-NeTrim([string]$s) { return (($s -replace ' +', ' ').Trim()) }

# City gate key: UPPER, SAINT<->ST, then strip to A-Z0-9 only. So "Saint Paul",
# "St. Paul", "ST PAUL" all collapse to STPAUL and match the same list entry.
function Get-NeCityKey([string]$city) {
    $u = ("" + $city).Trim().ToUpper()
    $u = $u -replace '\bSAINT\b', 'ST'
    return ($u -replace '[^A-Z0-9]', '')
}

# column J - the cleaned address
function ConvertTo-NeClean([string]$raw) {
    if ($null -eq $raw -or $raw -eq '') { return '' }
    $s = $raw.ToUpper().Replace([char]160, ' ')
    foreach ($ch in @(',', '.', '/', '-', '#')) { $s = $s.Replace($ch, ' ') }
    $s = ConvertTo-NeTrim $s
    $s = $s.Replace('P O BOX', 'PO BOX')
    return $s
}

function Get-NeOrdinal([int]$n) {
    $h = $n % 100
    if ($h -eq 11 -or $h -eq 12 -or $h -eq 13) { return "${n}TH" }
    switch ($n % 10) { 1 { "${n}ST" } 2 { "${n}ND" } 3 { "${n}RD" } default { "${n}TH" } }
}

# Parse one transaction address the way the sheet does. Returns the pieces the
# match needs: House (int or $null), W (street key), R (suffix abbr), U (city
# upper), V (zip5), Plus4, CityOk.
function Get-NeParse([string]$rawAddr, [string]$rawCity, [string]$rawZip) {
    $J = ConvertTo-NeClean $rawAddr
    # K - unit marker position (1-based FIND in J+' ')
    $Jp = $J + ' '; $K = 9999
    foreach ($m in $script:NeUnitMarkers) { $i = $Jp.IndexOf($m); if ($i -ge 0 -and ($i + 1) -lt $K) { $K = $i + 1 } }
    # L - address with the unit/secondary stripped
    if ($K -lt 9999) { $L = ConvertTo-NeTrim ($Jp.Substring(0, $K)) } else { $L = $J }
    # M - everything after the first space (drop the house number)
    $sp = $L.IndexOf(' '); $M = if ($sp -ge 0) { $L.Substring($sp + 1) } else { '' }
    # N - first word of M ; O - leading directional (abbr) if any
    $N = if ($M) { ($M -split ' ', 2)[0] } else { $M }
    $O = if ($script:NeDirs.ContainsKey($N)) { $script:NeDirs[$N] } else { '' }
    # P - street part (drop the leading directional)
    if ($O -eq '') { $P = $M } else { $sp2 = $M.IndexOf(' '); $P = if ($sp2 -ge 0) { $M.Substring($sp2 + 1) } else { $M } }
    # Q - last word of P ; R - suffix (abbr) if any
    $Q = if ($P) { ($P -split ' ')[-1] } else { '' }
    $R = if ($script:NeSufs.ContainsKey($Q)) { $script:NeSufs[$Q] } else { '' }
    # S - house number (or PO box number)
    if ($L.Length -ge 7 -and $L.Substring(0, 7) -eq 'PO BOX ') { $S = ConvertTo-NeTrim ($L.Substring(7, [Math]::Min(20, $L.Length - 7))) }
    elseif ($L.Contains(' ')) { $S = ($L -split ' ', 2)[0] } else { $S = $L }
    # U city, V zip5
    $U = ("" + $rawCity).Trim().ToUpper()
    $zdig = (("" + $rawZip).Trim()) -replace '[ \-]', ''
    $V = if ($zdig.Length -ge 5) { $zdig.Substring(0, 5) } elseif ($zdig) { $zdig } else { '' }
    # T - street name (drop the suffix word) ; W - street key (numeric -> ordinal)
    if ($L.Length -ge 7 -and $L.Substring(0, 7) -eq 'PO BOX ') { $T = 'PO BOX' }
    elseif ($R -eq '') { $T = $P }
    elseif (-not $P.Contains(' ')) { $T = $P }
    else { $T = ConvertTo-NeTrim ($P.Substring(0, $P.Length - $Q.Length - 1)) }
    if ($T -eq '') { $W = '' }
    elseif ($T -match '^\d+$') { $W = Get-NeOrdinal ([int]$T) }
    else { $W = $T }
    $Plus4 = if ($zdig.Length -ge 9) { $zdig.Substring(5, 4) } else { '' }
    # House numbers can exceed Int32 in raw sales data (e.g. 3194158696),
    # so parse as Int64 and simply leave it null if it still won't fit.
    $House = $null
    if ($S -match '^\d+$') { $hv64 = [long]0; if ([long]::TryParse($S, [ref]$hv64)) { $House = $hv64 } }
    return [pscustomobject]@{ House = $House; W = $W; R = $R; U = $U; V = $V; Plus4 = $Plus4; CityOk = $script:NeCities.ContainsKey((Get-NeCityKey $U)) }
}

# ---- the match (columns Z/AA/AB and F/G/H) --------------------------------
# Kept for compatibility / readability; the hot path below inlines these.
function Test-NeRange($lo, $hi, $h) {
    if ($null -eq $h) { return $false }
    $loi = [long]0; $hii = [long]0
    if (-not [long]::TryParse("$lo", [ref]$loi)) { return $false }
    if (-not [long]::TryParse("$hi", [ref]$hii)) { return $false }
    return ($h -ge $loi -and $h -le $hii)
}
function Test-NeOddEven([string]$e, $h) {
    $hh = if ($null -eq $h) { 0 } else { $h }
    if ($e -eq 'B' -or $e -eq '') { return $true }
    if ($e -eq 'O') { return (($hh % 2) -eq 1) }
    if ($e -eq 'E') { return (($hh % 2) -eq 0) }
    return $false
}

# The index is FLAT typed arrays plus a Dictionary of key -> row positions.
# (The old shape allocated one PSCustomObject per tax row - ~300k objects per
# quarter - which is what made loading a quarter feel like a hang.)
#   Map    : Dictionary[string, List[int]]  "STREET ZIP5" -> positions
#   Lo/Hi  : int[]     (-1 when the file value is not a number)
#   E/H    : string[]  (upper-cased once, here, not per comparison)
#   Code   : the City Code (final) column, as read
#   Dr     : the source data-row number, for reporting
#   ByZip9 : Dictionary[string, int]  "zip5plus4_" -> position
function Find-NeCode($p, $index) {
    # ZIP9 (column AC) wins when the customer gave a 9-digit zip
    if ($p.Plus4 -ne '' -and $index.ByZip9.Count -gt 0) {
        $zi = 0
        if ($index.ByZip9.TryGetValue(($p.V + $p.Plus4 + '_'), [ref]$zi)) {
            return @{ Code = $index.Code[$zi]; Flag = 'ZIP9'; Row = $index.Dr[$zi] }
        }
    }
    if (-not $p.CityOk) { return @{ Code = $null; Flag = 'NO CODE-CITY'; Row = $null } }
    if ($p.W -eq '') { return @{ Code = $null; Flag = 'NO MATCH'; Row = $null } }
    $lst = $null
    if (-not $index.Map.TryGetValue(($p.W + ' ' + $p.V), [ref]$lst)) { return @{ Code = $null; Flag = 'NO MATCH'; Row = $null } }

    $h = $p.House
    $ab = -1
    if ($null -ne $h) {
        $odd = ($h % 2)
        $r = $p.R
        # pass 1: range + odd/even + suffix agrees (or either side blank). Keep the LAST.
        foreach ($i in $lst) {
            $lo = $index.Lo[$i]
            if ($lo -lt 0 -or $h -lt $lo -or $h -gt $index.Hi[$i]) { continue }
            $e = $index.E[$i]
            if (-not ($e -eq 'B' -or $e -eq '' -or ($e -eq 'O' -and $odd -eq 1) -or ($e -eq 'E' -and $odd -eq 0))) { continue }
            $sh = $index.H[$i]
            if ($sh -eq $r -or $r -eq '' -or $sh -eq '') { $ab = $i }
        }
        # pass 2 (fallback): same, ignoring the suffix
        if ($ab -lt 0) {
            foreach ($i in $lst) {
                $lo = $index.Lo[$i]
                if ($lo -lt 0 -or $h -lt $lo -or $h -gt $index.Hi[$i]) { continue }
                $e = $index.E[$i]
                if ($e -eq 'B' -or $e -eq '' -or ($e -eq 'O' -and $odd -eq 1) -or ($e -eq 'E' -and $odd -eq 0)) { $ab = $i }
            }
        }
    }
    if ($ab -ge 0) { return @{ Code = $index.Code[$ab]; Flag = 'OK'; Row = $index.Dr[$ab] } }
    return @{ Code = 0; Flag = 'FALLBACK'; Row = $index.Dr[$lst[0]] }
}

# Build the index from parallel arrays (as read from a quarterly file).
# $neededKeys / $neededZip9 (optional HashSet[string]): when supplied, only rows
# whose key is actually wanted by this batch are indexed. A quarter file has
# ~300k rows but a workpaper usually needs a few hundred distinct streets, so
# this is the difference between minutes and seconds.
function New-NeIndex($dr, $key, $lo, $hi, $e, $h, $code, $concat, $neededKeys = $null, $neededZip9 = $null) {
    $n = $key.Count
    $map = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[int]]'
    $z9  = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $loA = New-Object 'long[]' $n
    $hiA = New-Object 'long[]' $n
    $eA  = New-Object 'string[]' $n
    $hA  = New-Object 'string[]' $n
    $filterK = ($null -ne $neededKeys -and $neededKeys.Count -gt 0)
    $filterZ = ($null -ne $neededZip9 -and $neededZip9.Count -gt 0)
    $lv = [long]0; $hv = [long]0
    for ($i = 0; $i -lt $n; $i++) {
        $k = [string]$key[$i]
        if ($k -ne '' -and ((-not $filterK) -or $neededKeys.Contains($k))) {
            if ([long]::TryParse([string]$lo[$i], [ref]$lv)) { $loA[$i] = $lv } else { $loA[$i] = -1 }
            if ([long]::TryParse([string]$hi[$i], [ref]$hv)) { $hiA[$i] = $hv } else { $hiA[$i] = -1 }
            $eA[$i] = ([string]$e[$i]).ToUpper()
            $hA[$i] = ([string]$h[$i]).ToUpper()
            $lst = $null
            if (-not $map.TryGetValue($k, [ref]$lst)) { $lst = New-Object 'System.Collections.Generic.List[int]'; $map[$k] = $lst }
            $lst.Add($i)
        }
        if ($null -ne $concat) {
            $w = [string]$concat[$i]
            if ($w -ne '') {
                $us = $w.IndexOf('_')
                if ($us -ge 0) {
                    $wk = $w.Substring(0, $us + 1)
                    if (((-not $filterZ) -or $neededZip9.Contains($wk)) -and (-not $z9.ContainsKey($wk))) { $z9[$wk] = $i }
                }
            }
        }
    }
    return @{ Map = $map; ByZip9 = $z9; Lo = $loA; Hi = $hiA; E = $eA; H = $hA; Code = $code; Dr = $dr }
}

# ============================================================================
#  FUZZY REPAIR  (pass 2)  -  only used on rows that came back NO MATCH/FALLBACK.
#  Ported from the original Get-NeTaxCode repair pass. Returns @{Text;Note};
#  Note is '' when nothing was changed. Also normalises SAINT -> ST.
# ============================================================================
$script:NeSuffixTypos = @{
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
$script:NeRealSuffixes = @('ST','AVE','BLVD','RD','DR','CT','LN','PL','PLZ','PKWY','TER','CIR',
    'HWY','WAY','TRL','RDG','LK','HTS','CV','PT','LOOP','PATH','SQ','XING','VW','BND','CRK','PARK',
    'CTR','HL','ROW','VLG','TRCE','MALL','SPUR','RUN','STREET','AVENUE','BOULEVARD','ROAD','DRIVE',
    'COURT','LANE','PLACE','PLAZA','PARKWAY','TERRACE','CIRCLE','HIGHWAY','TRAIL','RIDGE','LAKE',
    'HEIGHTS','COVE','POINT','SQUARE','CROSSING','VIEW','BEND','CREEK','CENTER','HILL','VILLAGE',
    'TRACE','LANDING','MEADOWS','VISTA','EXPRESSWAY','TURNPIKE')
$script:NeCardWords = @{ 'ONE'=1;'TWO'=2;'THREE'=3;'FOUR'=4;'FIVE'=5;'SIX'=6;'SEVEN'=7;'EIGHT'=8;'NINE'=9;'TEN'=10;'ELEVEN'=11;'TWELVE'=12;'THIRTEEN'=13;'FOURTEEN'=14;'FIFTEEN'=15;'SIXTEEN'=16;'SEVENTEEN'=17;'EIGHTEEN'=18;'NINETEEN'=19;'TWENTY'=20;'THIRTY'=30;'FORTY'=40;'FOURTY'=40;'FIFTY'=50;'SIXTY'=60;'SEVENTY'=70;'EIGHTY'=80;'NINETY'=90 }
$script:NeOrdWords = @{ 'FIRST'=1;'SECOND'=2;'THIRD'=3;'FOURTH'=4;'FIFTH'=5;'SIXTH'=6;'SEVENTH'=7;'EIGHTH'=8;'NINTH'=9;'TENTH'=10;'ELEVENTH'=11;'TWELFTH'=12;'THIRTEENTH'=13;'FOURTEENTH'=14;'FIFTEENTH'=15;'SIXTEENTH'=16;'SEVENTEENTH'=17;'EIGHTEENTH'=18;'NINETEENTH'=19;'TWENTIETH'=20;'THIRTIETH'=30;'FORTIETH'=40;'FIFTIETH'=50;'SIXTIETH'=60;'SEVENTIETH'=70;'EIGHTIETH'=80;'NINETIETH'=90 }

function Repair-NeAddress([string]$raw) {
    $result = @{ Text = $raw; Note = '' }
    if (-not $raw) { return $result }
    $s = $raw.ToUpper().Replace([char]160, ' ')
    foreach ($ch in @(',', '.', '/', '\', '-', '#')) { $s = $s.Replace($ch, ' ') }
    $toks = @($s -split '\s+' | Where-Object { $_ -ne '' })
    if ($toks.Count -eq 0) { return $result }
    $notes = New-Object System.Collections.Generic.List[string]
    # SAINT -> ST anywhere
    for ($i = 0; $i -lt $toks.Count; $i++) { if ($toks[$i] -eq 'SAINT') { $toks[$i] = 'ST'; $notes.Add('Saint -> St') } }
    if ($toks[0] -eq 'PO' -or ($toks.Count -gt 1 -and $toks[0] -eq 'P' -and $toks[1] -eq 'O')) {
        $new = ($toks -join ' ').Trim(); if ($new -ne ($raw.ToUpper().Trim()) -and $notes.Count -gt 0) { $result.Text = $new; $result.Note = ($notes -join '; ') }
        return $result
    }
    $first = $toks[0]
    if ($first -match '^(\d+)([A-Z]+)$') { $tail = $Matches[2]; if ($tail -notin @('ST','ND','RD','TH')) { $rest = @(); if ($toks.Count -gt 1) { $rest = $toks[1..($toks.Count - 1)] }; $toks = @($Matches[1]) + @($tail) + $rest; $notes.Add('no space after house number') } }
    if ($toks.Count -gt 2) { $last = $toks[$toks.Count - 1]; $prev = $toks[$toks.Count - 2]; if ($last -match '^\d+$' -and ($script:NeRealSuffixes -contains $prev -or $script:NeSuffixTypos.ContainsKey($prev))) { $toks = $toks[0..($toks.Count - 2)]; $notes.Add('stray unit number on the end') } }
    if ($toks.Count -gt 1) { $last = $toks[$toks.Count - 1]; if ($script:NeSuffixTypos.ContainsKey($last)) { $toks[$toks.Count - 1] = $script:NeSuffixTypos[$last]; $notes.Add('misspelled street suffix') } }
    if ($toks.Count -gt 2) { $last = $toks[$toks.Count - 1]; if ($script:NeDirs.ContainsKey($last)) { $headIsDir = $false; if ($toks.Count -gt 1 -and $script:NeDirs.ContainsKey($toks[1])) { $headIsDir = $true }; if (-not $headIsDir) { $d = $script:NeDirs[$last]; $toks = $toks[0..($toks.Count - 2)]; $toks = @($toks[0]) + @($d) + $toks[1..($toks.Count - 1)]; $notes.Add('direction on the wrong end') } } }
    if ($toks.Count -ge 3) {
        $begin = 1; if ($script:NeDirs.ContainsKey($toks[1])) { $begin = 2 }
        $total = 0; $used = 0; $ok = $false
        for ($i = $begin; $i -lt $toks.Count; $i++) { $w = $toks[$i]; if ($script:NeCardWords.ContainsKey($w)) { $total += $script:NeCardWords[$w]; $used = $i } elseif ($script:NeOrdWords.ContainsKey($w)) { $total += $script:NeOrdWords[$w]; $used = $i; $ok = $true; break } else { break } }
        if ($ok -and $total -gt 0) { $tail = @(); if ($used -lt $toks.Count - 1) { $tail = @($toks[($used + 1)..($toks.Count - 1)]) }; $tailOk = ($tail.Count -eq 0) -or ($tail.Count -eq 1 -and $script:NeRealSuffixes -contains $tail[0]); if ($tailOk) { $ordinal = Get-NeOrdinal $total; $head = @(); if ($begin -eq 2) { $head = @($toks[0], $toks[1]) } else { $head = @($toks[0]) }; $toks = $head + @($ordinal) + $tail; $notes.Add('street number spelled out') } }
    }
    $new = ($toks -join ' ').Trim()
    if ($new -ne ($raw.ToUpper().Trim()) -and $notes.Count -gt 0) { $result.Text = $new; $result.Note = ($notes -join '; ') }
    return $result
}
