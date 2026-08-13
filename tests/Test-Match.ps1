<#
 Test-Match.ps1 - exercises the pure matcher (NeMatch.ps1) against a small,
 hand-built tax index, one assertion per matching rule. No Excel needed.
   .\tests\Test-Match.ps1
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) -Parent
. (Join-Path $root 'NeMatch.ps1')

$pass = 0; $fail = 0
function Eq($got, $want, $msg) { if ("$got" -eq "$want") { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg (got '$got' want '$want')" -ForegroundColor Red } }
function Ok($c, $msg) { if ($c) { $script:pass++ } else { $script:fail++; Write-Host "  FAIL: $msg" -ForegroundColor Red } }

# ---- build a synthetic tax index -----------------------------------------
# columns: dr, key "W V", low, high, oddeven, suffix, code
$rows = @(
    @(1,'MAIN 68000',100,199,'B','ST',5),
    @(2,'MAIN 68000',200,298,'E','AVE',7),
    @(3,'MAIN 68000',201,299,'O','ST',9),
    @(4,'1ST 68000',1,99,'B','ST',11),
    @(5,'Q 68000',1,999,'B','',0),
    @(6,'PO BOX 68000',1,999,'B','',20)
)
$n = $rows.Count
$dr=@();$key=@();$lo=@();$hi=@();$e=@();$h=@();$code=@()
foreach ($r in $rows) { $dr+=$r[0];$key+=$r[1];$lo+=$r[2];$hi+=$r[3];$e+=$r[4];$h+=$r[5];$code+=$r[6] }
$index = New-NeIndex $dr $key $lo $hi $e $h $code $null

function M($addr,$city,$zip) { Find-NeCode (Get-NeParse $addr $city $zip) $index }

Write-Host "`n== house range + suffix + odd/even ==" -ForegroundColor Cyan
$r = M '150 Main St' 'Aurora' '68000'; Eq $r.Flag 'OK' '150 Main St flag'; Eq $r.Code 5 '150 Main St code'
$r = M '250 Main Ave' 'Aurora' '68000'; Eq $r.Code 7 '250 Main Ave -> even AVE row'
$r = M '251 Main St' 'Aurora' '68000'; Eq $r.Code 9 '251 Main St -> odd ST row'
$r = M '250 Main St' 'Aurora' '68000'; Eq $r.Flag 'OK' '250 Main St matches via suffix-fallback'; Eq $r.Code 7 '  (even row, suffix dropped)'

Write-Host "`n== fallback / no match / city gate ==" -ForegroundColor Cyan
$r = M '50 Main St' 'Aurora' '68000';  Eq $r.Flag 'FALLBACK' 'house outside every range -> FALLBACK'; Eq $r.Code 0 'FALLBACK code 0'
$r = M '150 Oak St' 'Aurora' '68000';  Eq $r.Flag 'NO MATCH' 'unknown street -> NO MATCH'
$r = M '150 Main St' 'Nowheresville' '68000'; Eq $r.Flag 'NO CODE-CITY' 'non-taxing city -> NO CODE-CITY'

Write-Host "`n== numbered street + PO box ==" -ForegroundColor Cyan
$r = M '50 1st St' 'Aurora' '68000'; Eq $r.Code 11 'numbered street 1st -> 1ST key'
$r = M 'PO BOX 5' 'Aurora' '68000';  Eq $r.Code 20 'PO Box -> PO BOX key, box# as house'
$r = M 'P O Box 5' 'Aurora' '68000'; Eq $r.Code 20 '"P O Box" normalises to PO BOX'

Write-Host "`n== parsing (W / suffix / city gate) ==" -ForegroundColor Cyan
$p = Get-NeParse '4965 E South Street' 'Hastings' '68901'; Eq $p.W 'SOUTH' 'drop leading dir + suffix'
$p = Get-NeParse '604 12th St.' 'Aurora' '68818';          Eq $p.W '12TH' '12th -> 12TH'
$p = Get-NeParse '76195 N. Highway 61' 'Grant' '69140';    Eq $p.W 'HIGHWAY 61' 'interior number kept'
Eq (Get-NeCityKey 'Saint Paul') 'STPAUL' 'Saint Paul -> STPAUL'
Eq (Get-NeCityKey 'St. Paul')   'STPAUL' 'St. Paul -> STPAUL'
Eq (Get-NeCityKey "O'Neill")    'ONEILL' "O'Neill -> ONEILL"
Ok ((Get-NeParse '1 Main St' 'Saint Paul' '68873').CityOk) 'Saint Paul is a taxing city'
Ok ((Get-NeParse '1 Main St' 'St Paul' '68873').CityOk)    'St Paul is a taxing city'

Write-Host "`n== ZIP9 (9-digit zip; concat = zip5 plus4 _ [code]) ==" -ForegroundColor Cyan
# concat has a code appended on the taxing row, none on the base row
$zdr=@(101,102,103); $zkey=@('POTTER 68007','POTTER 68007','MAIN 68791')
$zlo=@(17100,17200,200); $zhi=@(17199,17299,299); $ze=@('B','B','B'); $zh=@('ST','ST','RD')
$zcode=@(0,0,530); $zconcat=@('680071678_','680071680_','687913029_530')
$zidx = New-NeIndex $zdr $zkey $zlo $zhi $ze $zh $zcode $zconcat
$r = Find-NeCode (Get-NeParse '200 Main Rd' 'Wisner' '687913029') $zidx
Eq $r.Flag 'ZIP9' 'coded concat (687913029_530) -> ZIP9'; Eq $r.Code 530 '  code 530 (the _* wildcard finds it)'
$r = Find-NeCode (Get-NeParse '17150 Potter St' 'Bennington' '680071678') $zidx
Eq $r.Flag 'ZIP9' 'uncoded concat (680071678_) -> ZIP9'
$r = Find-NeCode (Get-NeParse '17150 Potter St' 'Bennington' '68007') $zidx
Eq $r.Flag 'OK' '5-digit zip (no plus4) -> street match, not ZIP9'

Write-Host "`n== fuzzy repair ==" -ForegroundColor Cyan
Eq (Repair-NeAddress '150 MAIN STREE').Text '150 MAIN ST' 'misspelled suffix repaired'
Eq (Repair-NeAddress '100 MAIN ST N').Text '100 N MAIN ST' 'direction moved to front'
Eq (Repair-NeAddress '100 FORTY SECOND ST').Text '100 42ND ST' 'spelled-out number'
Eq (Repair-NeAddress '123MAIN ST').Text '123 MAIN ST' 'missing space after house #'
Eq (Repair-NeAddress 'Saint Marys Rd').Text 'ST MARYS RD' 'Saint -> St in address'
# repaired address then matches
$rep = Repair-NeAddress '150 MAIN STREE'
$r = M $rep.Text 'Aurora' '68000'; Eq $r.Code 5 'repaired address matches code 5'

Write-Host ('-' * 50) -ForegroundColor DarkGray
Write-Host ("  PASS: {0}   FAIL: {1}" -f $pass, $fail) -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
if ($fail) { exit 1 }
exit 0
