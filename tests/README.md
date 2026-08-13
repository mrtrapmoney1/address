# Tests / verification resources  (PowerShell only)

These verify the part that actually broke before: **pasting results into the
workbook**. The matching logic lives in the Excel engine and is not re-tested
here; these target the script's read/shape/write code. Everything is PowerShell.

| File | What it checks | Needs |
|---|---|---|
| `Test-Paste.ps1` | **The paste.** Runs the real `NeTaxPaste.ps1` functions against a throwaway workbook and reads every written cell back: append vs insert, per-row alignment across chunk boundaries, no overwrite of existing data, the `000` code format, and the per-quarter value-pasted engine backup. Uses real Excel if present, otherwise a built-in mock. | PowerShell (Excel optional) |
| `Test-Logic.ps1` | The pure functions: column letters, quarter parsing, address repair, error/code normalisation, array-shape reads. | PowerShell |
| `Check-Workpaper.ps1` | **Post-run validator.** Opens a *finished* workpaper in Excel and scans for the scramble signature (code/row jammed into the flag cell) and for any code that isn't an integer 0..999. Run it on real output before trusting a run. | PowerShell + Excel |
| `Run-AllTests.ps1` | Runs `Test-Logic` + `Test-Paste` and reports one result. | PowerShell |
| `ExcelMock.ps1` | The Excel-COM stand-in used by `Test-Paste.ps1` when Excel isn't installed. Not run directly. | - |

## Run it

Works in Windows PowerShell 5.1 - you do **not** need `pwsh` (that's PowerShell 7).
From this folder:

```powershell
.\tests\Run-AllTests.ps1                              # logic + paste, one summary
.\tests\Check-Workpaper.ps1 -Path ".\Your Working Paper.xlsx" -HeaderRow 4
```

If a script is blocked, allow it for this session first:
`Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`, or run
`powershell -ExecutionPolicy Bypass -File tests\Run-AllTests.ps1`.
You can also run the two suites individually: `.\tests\Test-Logic.ps1` and
`.\tests\Test-Paste.ps1` (the paste test uses your real Excel automatically).

## What "good" looks like

- `Test-Logic.ps1`  -> `PASS: 58   FAIL: 0`
- `Test-Paste.ps1`  -> `PASS: 30   FAIL: 0`
- `Check-Workpaper.ps1` on a clean run -> `CLEAN`, exit 0
- `Check-Workpaper.ps1` on the old broken output -> scrambled/invalid rows, exit 1

`Test-Paste.ps1` uses a mock when Excel is absent, so it runs on a build box.
On a Windows machine with Excel it automatically switches to real Excel COM and
the same assertions run against the actual product - do this once before a real
run to close the last gap.

## A real bug these caught

`Test-Paste.ps1` found that `@($oneArray)` **flattens** in PowerShell, so the
code-only output silently wrote only the first row. Fixed by passing a single
column as `(,$array)` and adding a loud guard in `Write-ResultBlock`. Without
the paste test this would have shipped and written one cell per run.
