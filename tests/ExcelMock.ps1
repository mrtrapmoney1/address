<#
================================================================================
 ExcelMock.ps1
 A small, faithful stand-in for the Excel COM surface that NeTaxPaste.ps1 uses,
 so the paste can be executed and checked WITHOUT Excel (e.g. on a build box).

 It models the one thing that matters for the scramble bug: assigning a 2-D
 (N x 1) object array to an N-row x 1-column range writes cell (r1+i, col) =
 arr[i,0], and NOTHING else. A wrong-shaped write (a 1-D array, or a 2-D array
 whose dimensions don't match the range) THROWS - which is exactly what you
 want a test to do, because that is how the old code jammed several values into
 one cell.

 When real Excel IS available, Test-Paste.ps1 uses it instead and the same
 assertions run unchanged.
================================================================================
#>

class MockWorksheets : System.Collections.IEnumerable {
    [System.Collections.Generic.List[object]] $Items
    [int] $Count
    MockWorksheets() { $this.Items = [System.Collections.Generic.List[object]]::new(); $this.Count = 0 }
    [object] Add() {
        $s = New-MockSheet ("Sheet" + ($this.Items.Count + 1))
        $s | Add-Member -NotePropertyName _owner -NotePropertyValue $this -Force
        $this.Items.Add($s); $this.Count = $this.Items.Count
        return $s
    }
    [void] Remove([object]$s) { [void]$this.Items.Remove($s); $this.Count = $this.Items.Count }
    [System.Collections.IEnumerator] GetEnumerator() { return $this.Items.GetEnumerator() }
}

function New-MockSheet([string]$name) {
    $sheet = [pscustomobject]@{ Name = $name; _cells = @{}; _fmt = @{} }
    $sheet | Add-Member ScriptMethod _get { param($r, $c) $k = "$r,$c"; if ($this._cells.ContainsKey($k)) { $this._cells[$k] } else { $null } }
    $sheet | Add-Member ScriptMethod _set { param($r, $c, $v) if ($null -eq $v) { [void]$this._cells.Remove("$r,$c") } else { $this._cells["$r,$c"] = $v } }
    $sheet | Add-Member ScriptMethod _fmtset { param($r, $c, $f) $this._fmt["$r,$c"] = $f }
    $sheet | Add-Member ScriptMethod _fmtget { param($r, $c) $k = "$r,$c"; if ($this._fmt.ContainsKey($k)) { $this._fmt[$k] } else { 'General' } }
    $sheet | Add-Member ScriptMethod Cells { param($r, $c) New-MockCell $this $r $c }
    $sheet | Add-Member ScriptMethod Range { param($a, $b) New-MockRange $this $a.Row $a.Col $b.Row $b.Col }
    $sheet | Add-Member ScriptMethod Columns { param($n) New-MockColumn $this $n }
    $sheet | Add-Member ScriptMethod Delete { if ($this._owner) { $this._owner.Remove($this) } }
    return $sheet
}

function New-MockCell($sheet, [int]$r, [int]$c) {
    $cell = [pscustomobject]@{ Row = $r; Col = $c; _sheet = $sheet }
    $cell | Add-Member ScriptProperty Value2 { $this._sheet._get($this.Row, $this.Col) } { param($v) $this._sheet._set($this.Row, $this.Col, $v) }
    return $cell
}

function New-MockRange($sheet, [int]$r1, [int]$c1, [int]$r2, [int]$c2) {
    $rng = [pscustomobject]@{ Sheet = $sheet; R1 = $r1; C1 = $c1; R2 = $r2; C2 = $c2 }
    $rng | Add-Member ScriptProperty Value2 { $null } {
        param($v)
        $rows = $this.R2 - $this.R1 + 1; $cols = $this.C2 - $this.C1 + 1
        if ($v -is [Array] -and $v.Rank -eq 2) {
            $lb0 = $v.GetLowerBound(0); $lb1 = $v.GetLowerBound(1)
            $a = $v.GetUpperBound(0) - $lb0 + 1; $b = $v.GetUpperBound(1) - $lb1 + 1
            if ($a -ne $rows -or $b -ne $cols) { throw "SHAPE MISMATCH: ${a}x${b} array into ${rows}x${cols} range - this is how cells scramble." }
            for ($i = 0; $i -lt $rows; $i++) { for ($j = 0; $j -lt $cols; $j++) { $this.Sheet._set($this.R1 + $i, $this.C1 + $j, $v.GetValue($lb0 + $i, $lb1 + $j)) } }
        } elseif ($v -is [Array]) {
            throw "1-D ARRAY WRITE into a ${rows}x${cols} range - Excel would spread it row-major and scramble. NeTaxPaste must never do this."
        } else {
            for ($i = 0; $i -lt $rows; $i++) { for ($j = 0; $j -lt $cols; $j++) { $this.Sheet._set($this.R1 + $i, $this.C1 + $j, $v) } }
        }
    }
    $rng | Add-Member ScriptProperty NumberFormat { $this.Sheet._fmtget($this.R1, $this.C1) } {
        param($f)
        for ($i = $this.R1; $i -le $this.R2; $i++) { for ($j = $this.C1; $j -le $this.C2; $j++) { $this.Sheet._fmtset($i, $j, $f) } }
    }
    return $rng
}

function New-MockColumn($sheet, [int]$n) {
    $col = [pscustomobject]@{ Sheet = $sheet; N = $n }
    $col | Add-Member ScriptMethod Insert {
        param($shift)
        $newCells = @{}; $newFmt = @{}
        foreach ($k in $this.Sheet._cells.Keys) { $rc = $k -split ','; $r = [int]$rc[0]; $c = [int]$rc[1]; if ($c -ge $this.N) { $c++ }; $newCells["$r,$c"] = $this.Sheet._cells[$k] }
        foreach ($k in $this.Sheet._fmt.Keys)   { $rc = $k -split ','; $r = [int]$rc[0]; $c = [int]$rc[1]; if ($c -ge $this.N) { $c++ }; $newFmt["$r,$c"]   = $this.Sheet._fmt[$k] }
        $this.Sheet._cells = $newCells; $this.Sheet._fmt = $newFmt
    }
    return $col
}

function New-MockWorkbook {
    $wsx = [MockWorksheets]::new()
    $first = $wsx.Add(); $first.Name = 'Sheet1'   # a fresh workbook opens with one sheet
    $wb = [pscustomobject]@{ Worksheets = $wsx; SavedPath = $null; SavedFormat = $null; Closed = $false }
    $wb | Add-Member ScriptMethod SaveAs { param($p, $fmt) $this.SavedPath = $p; $this.SavedFormat = $fmt }
    $wb | Add-Member ScriptMethod Close { param($save) $this.Closed = $true }
    return $wb
}

function New-MockExcel {
    $wf = [pscustomobject]@{}
    $wf | Add-Member ScriptMethod CountA {
        param($range)
        $n = 0
        foreach ($k in $range.Sheet._cells.Keys) {
            $rc = $k -split ','; $r = [int]$rc[0]; $c = [int]$rc[1]
            if ($r -ge $range.R1 -and $r -le $range.R2 -and $c -ge $range.C1 -and $c -le $range.C2) {
                $v = $range.Sheet._cells[$k]
                if ($null -ne $v -and "$v" -ne '') { $n++ }
            }
        }
        return $n
    }
    $wbks = [pscustomobject]@{ _books = @() }
    $wbks | Add-Member ScriptMethod Add { $b = New-MockWorkbook; $this._books += $b; $b }
    $excel = [pscustomobject]@{ WorksheetFunction = $wf; Workbooks = $wbks }
    return $excel
}
