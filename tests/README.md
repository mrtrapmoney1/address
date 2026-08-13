# Tests / verification resources

These verify the part that actually broke before: **pasting results into the
workbook**. The matching logic lives in the Excel engine and is not re-tested
here; these target the script's read/shape/write code.

| File | What it checks | Needs |
|---|---|---|
| `Test-Paste.ps1` | **The paste.** Runs the real `NeTaxPaste.ps1` functions against a throwaway workbook and reads every written cell back: append vs insert, per-row alignment across chunk boundaries, no overwrite of existing data, the `000` code format, and the per-quarter value-pasted engine backup. Uses real Excel if present, otherwise a built-in mock. | PowerShell (Excel optional) |
| `Test-Logic.ps1` | The pure functions: column letters, quarter parsing, address repair, error/code normalisation, array-shape reads. | PowerShell |
| `Check-Workpaper.py` | **Post-run validator.** Scans a *finished* workpaper for the scramble signature (code/row jammed into the flag cell) and for any code that isn't an integer 0..999. Run it on real output before trusting a run. | Python + openpyxl |
| `Check-Workpaper.ps1` | Same validator via Excel COM, for Windows shops without Python. | PowerShell + Excel |
| `ExcelMock.ps1` | The Excel-COM stand-in used by `Test-Paste.ps1` when Excel isn't installed. Not run directly. | - |

## Run everything

```powershell
# Windows (uses real Excel where it can):
pwsh -File tests\Test-Logic.ps1
pwsh -File tests\Test-Paste.ps1
python  tests\Check-Workpaper.py ".\Your Working Paper.xlsx" --header-row 4
```

```bash
# Any machine (paste test runs against the mock):
pwsh -File tests/Test-Logic.ps1
pwsh -File tests/Test-Paste.ps1
python3 tests/Check-Workpaper.py samples/Tax_Rate_Check_FIXED.xlsx --header-row 4
```

## What "good" looks like

- `Test-Logic.ps1`  -> `PASS: 58   FAIL: 0`
- `Test-Paste.ps1`  -> `PASS: 30   FAIL: 0`
- `Check-Workpaper.py` on `samples/Tax_Rate_Check_FIXED.xlsx` -> `CLEAN`, exit 0
- `Check-Workpaper.py` on `samples/Tax_Rate_Check_ORIGINAL.xlsx` -> **63 scrambled + 24 invalid**, exit 1
  (that's the old broken run - the detector is supposed to flag it)

## A real bug these caught

`Test-Paste.ps1` found that `@($oneArray)` **flattens** in PowerShell, so the
code-only output silently wrote only the first row. Fixed by passing a single
column as `(,$array)` and adding a loud guard in `Write-ResultBlock`. Without
the paste test this would have shipped and written one cell per run.
