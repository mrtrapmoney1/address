<#
================================================================================
 Check-Workpaper.ps1  -  scan a finished workpaper for the "scramble" defect,
 using Excel COM (Windows). Same idea as Check-Workpaper.py, no Python needed.

   powershell -File .\tests\Check-Workpaper.ps1 -Path ".\Working Paper.xlsx" `
              -HeaderRow 4 -CodeCol "NE City Code" -FlagCol "Match Flag"

 Exit 0 = clean, 1 = scrambled/invalid rows found, 2 = could not check.
================================================================================
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Path,
    [string]$Sheet,
    [int]$HeaderRow = 4,
    [string]$CodeCol = 'NE City Code',
    [string]$FlagCol = 'Match Flag'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $Path)) { Write-Host "Not found: $Path" -ForegroundColor Red; exit 2 }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open((Resolve-Path $Path).Path, 0, $true)
    $ws = if ($Sheet) { $wb.Worksheets.Item($Sheet) } else { $wb.Worksheets.Item(1) }

    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1
    $codeC = 0; $flagC = 0
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = [string]$ws.Cells($HeaderRow, $c).Value2
        if ($h.Trim() -eq $CodeCol) { $codeC = $c }
        if ($h.Trim() -eq $FlagCol) { $flagC = $c }
    }
    if ($codeC -eq 0) { Write-Host "No '$CodeCol' header on row $HeaderRow." -ForegroundColor Red; exit 2 }

    $jam = [regex]'.+\s-?\d+$'
    $scrambled = 0; $badCode = 0; $exScr = @(); $exBad = @()
    for ($r = $HeaderRow + 1; $r -le $lastRow; $r++) {
        $code = $ws.Cells($r, $codeC).Value2
        $flag = if ($flagC) { [string]$ws.Cells($r, $flagC).Value2 } else { '' }
        if ($flag -and $jam.IsMatch($flag.Trim()) -and $flag -notmatch 'REPAIRED') { $scrambled++; if ($exScr.Count -lt 15) { $exScr += "row $r : $flag" } }
        if ($null -ne $code -and "$code".Trim() -ne '') {
            $n = 0.0
            if (-not [double]::TryParse("$code", [ref]$n) -or $n -lt 0 -or $n -gt 999 -or $n -ne [Math]::Floor($n)) { $badCode++; if ($exBad.Count -lt 15) { $exBad += "row $r : $code" } }
        }
    }

    Write-Host "Workbook : $Path"
    Write-Host ("Sheet    : {0}   header row {1}" -f $ws.Name, $HeaderRow)
    Write-Host ("Code col : {0} (col {1}){2}" -f $CodeCol, $codeC, $(if ($flagC) { "   Flag col: $FlagCol (col $flagC)" } else { "   (no flag column - code only)" }))
    Write-Host ("Data rows: {0}" -f ($lastRow - $HeaderRow))
    Write-Host ('-' * 60)
    if ($scrambled -eq 0 -and $badCode -eq 0) {
        Write-Host "CLEAN - no scrambled cells, every code is a valid 0..999." -ForegroundColor Green
        $code = 0
    } else {
        if ($scrambled) { Write-Host "SCRAMBLED cells (code/row jammed into the flag): $scrambled" -ForegroundColor Red; $exScr | ForEach-Object { Write-Host "   $_" } }
        if ($badCode)   { Write-Host "INVALID codes (not an integer 0..999): $badCode" -ForegroundColor Red; $exBad | ForEach-Object { Write-Host "   $_" } }
        $code = 1
    }
    $wb.Close($false); $excel.Quit()
    exit $code
}
catch {
    Write-Host "Could not check: $($_.Exception.Message)" -ForegroundColor Red
    try { if ($wb) { $wb.Close($false) } } catch {}
    try { if ($excel) { $excel.Quit() } } catch {}
    exit 2
}
