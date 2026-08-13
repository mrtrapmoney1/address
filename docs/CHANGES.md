# Nebraska City-Code Lookup — analysis and v4 rewrite

This is the write-up you asked for: what the current work does, where the
PowerShell was falling down, and exactly what changed. The engine workbook is
good and is **not** touched — every change here is in how the script *moves*
answers in and out.

---

## 1. What you already had (and what's actually solid)

| Piece | Verdict |
|---|---|
| `Address_Tax_Engine.xlsx` (Addresses `F:H` formulas, `TaxData`, parsing lists `AN:BB`) | **Keep as-is.** This is the brain. It parses the address, gates on the city list, does the sorted street+zip lookup, checks the house-number range and odd/even side, and returns `TAX CODE / Match Row# / Flag`. Re-implementing this in code is what burned the earlier versions. |
| The idea "Excel does the thinking, the script does the moving" | **Correct call.** Kept. |
| Pass 2 fuzzy repair (`Repair-Address`) | **Good — kept verbatim.** Fixes misspelled suffixes, wrong-end directions, spelled-out street numbers, stray unit numbers. |
| The one-column-at-a-time writer | **Good pattern — kept.** It's the one thing that reliably avoids the scramble. |

## 2. What was wrong

**A. The engine run was destroyed by the repair pass.**
In the old script, pass 2 wrote its "better" answer straight back over
`$outCodeF / $outRowG / $outFlagH` — the very arrays that held the pass-1
engine answer. Once a row was repaired, the raw template answer was gone. That
is the *"main bulk of the work"* you said the fuzzy match deletes. Confirmed in
the code (old `Get-NeTaxCode.ps1`, pass-2 loop).

**B. One run wrote a scrambled block into the working paper.**
In `Tax_Rate_Check.xlsx`, columns `T:W` (the first run) are mangled — **63 of
116 rows** have the code, row number and flag jammed into a single cell:
`"236652 0"`, `"OK 0"`, `"FALLBACK - row 97027 (unvalidated) 0"`, with `T`/`U`
left empty. Columns `X:AA` (the second run) are clean and match
`_Results_check.csv`. So the machinery *can* produce clean output; one run
flattened a 2-D array the wrong way on the way out.

**C. No control over where results land, and no true insert.**
The old script always appended at the far right and never *inserted*. You
wanted: press Enter to append at the end, **or** name a spot and have new
columns inserted there without overwriting anything.

**D. The save wasn't guaranteed.**
If anything threw between "write" and `$wb.Save()`, the results lived only in
memory. You noticed this — "always save at the end … you currently are not
doing this."

**E. Strictly sequential.** One quarter at a time. Fine for a sample, a
non-starter for millions of rows.

## 3. What v4 changes

### Concurrency (Section 7)
Rows are grouped by quarter (unchanged), then each quarter is handed to its
**own worker in its own Excel process** with its own copy of the engine, via an
STA runspace pool. Quarters run at the same time up to `$MaxParallel`
(auto = half your cores, capped at 6). `-MaxParallel 1` gives you the exact old
sequential behaviour as a safety valve. Because each worker owns a separate
`EXCEL.EXE`, they never contend over one instance.

### The engine (pass-1) answer is preserved and backed up (Sections 8, 10)
Each worker now keeps **two** sets of arrays:

- `Engine*` — the raw pass-1 template answer. **Never overwritten.**
- `Final*` — starts as a copy of pass 1; only *this* is updated when a repair
  does better.

After the run, the pass-1 answer is written to a **separate backup workbook**,
`_Backups\<paper>_ENGINE_<timestamp>.xlsx`, with **one value-pasted sheet per
quarter** (`2024Q4 (engine)`, `2025Q1 (engine)`, …). Values only — no formulas,
nothing that a later repair or re-sort could disturb. That's your durable record
of the engine's own work, exactly as it came out.

### Clean output, your choice of location (Sections 3b, 9)
A new prompt: **press Enter to append at the end**, or type a column
(`T`, or a header name) to **insert** the result column(s) there. Insert uses
`Columns.Insert(xlShiftToRight)`, so existing data slides right and is never
written over. Every result column is written on its own — the `T:W` scramble is
structurally impossible now.

### Minimal output: code only, formatted `000` (Section 9)
The working paper gets **one column by default — `NE City Code`, formatted
`000`** (0 → `000`, 94 → `094`), every cell the same. The **Match Flag** is an
opt-in second column (`-IncludeFlag`, or answer `Y` at the prompt). `Match Row #`
and `Source File` no longer clutter the working paper — they live in the engine
backup and `_Results_check.csv` for auditing. The written answer is the **final
(repaired)** result — your good run.

### The paste is its own tested unit (`NeTaxPaste.ps1` + `tests/`)
All workbook-writing moved into `NeTaxPaste.ps1` (`Add-ResultColumns`,
`Write-ColumnValues`, `Write-ResultBlock`, `New-EngineBackup`), dot-sourced by
the main script. `tests/Test-Paste.ps1` runs those *same* functions against a
throwaway workbook — real Excel if present, otherwise a faithful COM mock
(`tests/ExcelMock.ps1`) — and reads every cell back to check append vs insert,
per-row alignment across chunk boundaries, no-overwrite, the `000` format, and
the value-pasted per-quarter engine backup. `tests/Check-Workpaper.py` (and a
`.ps1` twin) scan a *finished* paper for the scramble signature and invalid
codes; run on the sample they flag the old file's 63 scrambled + 24 invalid
rows and pass the fixed one.

**A real bug the paste test caught before shipping:** in PowerShell `@($array)`
*flattens*, so the code-only path (`$columns = @($finalCode)`) silently arrived
as N scalars and wrote only row 1. Fixed by passing a single column as
`(,$array)` and adding a loud count-mismatch guard in `Write-ResultBlock`.

### It always saves (Section 11)
`Save()` retries up to 4× with backoff; if the file is locked it falls back to
`SaveAs` a `*_RESULTS_*` copy; and if even that fails it tells you in red not to
close Excel. The whole-workbook timestamped backup is still taken *before* any
cell is written (Section 4).

### Better terminal experience
Clear numbered sections, per-quarter queue/finish lines, a live sample of
finished rows, a flag tally, and plain-English flag meanings. Params
(`-WorkingPaper -SheetName -HeaderRow -PasteAt -MaxParallel -NonInteractive`)
let you run it unattended from Task Scheduler for the big files.

## 4. Why still PowerShell + Excel (not pure code)

Because you need millions of rows **accurate**, and the template is the only
thing that's been proven right. Pure-code re-implementations drifted from the
template three times. v4 keeps the template as the single source of truth and
only makes the *moving* faster (parallel) and safer (dual arrays, insert-not-
overwrite, guaranteed save). Correct-and-parallel beats fast-and-wrong.

## 5. Files in this repo

```
Get-NeTaxCode.ps1                     the v4 script
README.txt                            operator instructions (updated)
docs/CHANGES.md                       this document
docs/README_original.txt              your original README, for reference
samples/Tax_Rate_Check_ORIGINAL.xlsx  your sample, untouched (note T:W scramble)
samples/Tax_Rate_Check_FIXED.xlsx     what v4 produces: one clean block +
                                        a value-pasted "2021Q3 (engine)" backup sheet
samples/_Results_check.csv            your CSV, for reference
_Address/                             put the quarterly state files here
_Backups/                             fills itself (workbook + engine backups)
```

> Note: `Address_Tax_Engine.xlsx` (~25 MB) is **not** committed — it's your
> template and it's large. Drop your clean copy next to the script to run.
