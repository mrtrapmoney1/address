# Nebraska City Tax Code

Address → Nebraska city tax code, written straight onto your working paper.
**Pure PowerShell. No engine workbook. No Python.**

---

## Quick start — get the codes

**1. Put the deploy set + your data in one folder:**

```
<working folder>\
    Get-CityCode.ps1        <- run this
    NeMatch.ps1             <- matching logic
    NeDicts.ps1             <- city / direction / suffix lists
    NeTaxPaste.ps1          <- workbook writer + Report builder
    Your Working Paper.xlsx
    _Address\               <- the quarterly tax files
        2021 Q3 Address Data.xlsx
        2021 Q4 Address Data.xlsx
        ...
    _Backups\               <- created for you
```

**2. Run it:**

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-CityCode.ps1
```

Answer the prompts (it guesses the columns — press Enter to accept): sheet,
header row, ADDRESS / CITY / ZIP / DATE / STATE columns, where to put the result
(Enter = end of sheet, or a column like `T`), and whether to add a Match Flag.

**3. What you get:**

| Output | |
|---|---|
| **NE City Code** | one new column, formatted `000` (0 → `000`, 94 → `094`) |
| **Match Flag** | only if you ask for it |
| **Report** (new tab) | flag summary + every unique address with its code, flag, and count |

The workpaper saves automatically; a full backup is taken first.

**Unattended:**

```powershell
powershell -File .\Get-CityCode.ps1 -WorkingPaper ".\Big.xlsx" `
    -SheetName Data -HeaderRow 1 -PasteAt END -IncludeFlag -NonInteractive
```

---

## How it works (and why the codes are right)

`Get-CityCode.ps1` does **not** use the Excel engine template. It is a faithful
port of that template's Addresses-sheet formulas (columns J..AB and F/G/H) into
PowerShell — `NeMatch.ps1`. The lookup lists (282 taxing cities, directions,
street suffixes) are the exact ones from your engine, in `NeDicts.ps1`.

Per address it does what the sheet does:

1. clean + parse into house number, street name, suffix, zip5;
2. gate on the city (must be a taxing city; `Saint Paul` / `St Paul` both count);
3. find the street+zip block in the quarterly file (sorted key column X);
4. pick the row whose house number is in range, odd/even side matches, and
   suffix agrees (dropping the suffix condition as a fallback);
5. return that row's **City Code (final)**. In range but no house match → `FALLBACK` (code 0).

**Fuzzy match (pass 2):** `NO MATCH` / `FALLBACK` rows are cleaned up (misspelled
suffix, wrong-end direction, spelled-out street number, missing space after the
house number, stray unit number, Saint→St) and retried; a win is flagged
`OK - REPAIRED: <what was fixed>`.

**Many quarters:** rows are grouped by DATE and each quarter is matched against
its own file in `_Address\`.

### Flags

`OK` · `ZIP9` · `… REPAIRED` (spot-check) · `FALLBACK` (code 0) · `NO MATCH` ·
`NO CODE-CITY` (blank) · `OOS` (blank) · `NO DATE` / `NO TAX FILE`

---

## Verify it (all PowerShell, Windows PowerShell 5.1)

```powershell
.\tests\Run-AllTests.ps1
```

| Suite | Covers |
|---|---|
| **Test-Logic** (58) | pure helpers (columns, quarters, repair, normalisation) |
| **Test-Match** (27) | the matcher: ranges, odd/even, suffix fallback, numbered streets, PO boxes, Saint/St, fuzzy repair |
| **Test-Paste** (31–38) | the workbook write (append/insert/no-overwrite/`000`) and the Report sheet — real Excel if present, else a mock |

```powershell
.\tests\Check-Workpaper.ps1 -Path ".\Working Paper.xlsx" -HeaderRow 4
```
scans a finished paper for scrambled cells or non-`000` codes.

> The matcher was cross-checked against an independent second implementation on
> 116 real sample addresses using real tax data — identical on every row.

---

## Repository layout

```
Get-CityCode.ps1   NeMatch.ps1   NeDicts.ps1   NeTaxPaste.ps1   <- the deploy set
_Address\  _Backups\                                            <- runtime folders
tests\             the PowerShell test suite + Check-Workpaper
docs\              CHANGES.md, ARCHITECTURE.md, original README
samples\           example workpapers (original vs fixed) + a results CSV
legacy\            the old engine-COM script, superseded — reference only
```

See `docs/ARCHITECTURE.md` for how the pieces fit together.

---

## Notes

- Quarterly files must be the prepped layout (col G = `Street Name`, col Y =
  `City Code (final)`), sorted by column X. The script checks this and stops
  rather than writing garbage.
- Nothing on the working paper is overwritten: result columns are brand new, and
  the Report sheet is regenerated each run (it's our own output).
- `pwsh` (PowerShell 7) is **not** required — run the `.ps1` files with the
  `powershell` you already have.
