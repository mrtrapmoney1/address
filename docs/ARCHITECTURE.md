# Architecture

How the pieces fit together and where each responsibility lives.

## Data flow

```
                Get-CityCode.ps1   (the only script you run)
                        │
     ┌──────────────────┼───────────────────────────┐
     │ read             │ match                      │ write
     ▼                  ▼                            ▼
  workpaper        NeMatch.ps1                  NeTaxPaste.ps1
  (Excel COM)      + NeDicts.ps1                (Excel COM)
                        │                            │
      group rows by    │  per quarter:              │  - insert/append columns
      quarter (date)   │  read _Address file,       │  - write code (000) [+flag]
                       │  build in-memory index,    │  - build "Report" sheet
                       │  match (incl. fuzzy pass)  │  - save
```

## Files

| File | Responsibility | Tested by |
|---|---|---|
| **Get-CityCode.ps1** | Orchestration: prompts, read the workpaper columns, group by quarter, read each quarterly file, call the matcher, call the writer, save. Excel COM I/O only. | (run on real data) |
| **NeMatch.ps1** | The matching brain — a faithful port of the engine's Addresses-sheet formulas (J..AB, F/G/H): `Get-NeParse`, `Find-NeCode`, `New-NeIndex`, `Repair-NeAddress`, `Get-NeCityKey`. Pure logic, no COM. | `tests/Test-Match.ps1` + cross-check vs an independent port |
| **NeDicts.ps1** | The three lookup lists extracted from the engine: 282 taxing cities (gate), 17 directions, 114 street suffixes. Pure data. | (used by Test-Match) |
| **NeTaxPaste.ps1** | The workbook writer — `Write-ResultBlock` (insert/append, one column at a time, `000` format), `New-NeReport` (the Report sheet), plus `New-EngineBackup` (used by legacy). Excel COM. | `tests/Test-Paste.ps1` (real Excel or mock) |
| **tests/ExcelMock.ps1** | A stand-in for the Excel COM surface so the paste can run without Excel. Not shipped to users. | — |

## Why a formula port (not the engine workbook)

The Excel engine template is correct, but driving it over COM in parallel added
a large surface for failures and produced wrong/hard-to-run results. Porting the
formulas into PowerShell removes the engine workbook, the second-workbook shuffle
and the runspaces. The port is held to the engine's behaviour by:

1. **`tests/Test-Match.ps1`** — one assertion per matching rule (house range,
   odd/even, suffix + suffix-drop fallback, numbered-street ordinals, PO boxes,
   the city gate, Saint/St, and each fuzzy-repair rule).
2. **A cross-check** — an independent second implementation of the same formulas
   was run against real tax data on the 116-row sample workpaper; the two agreed
   on every row (code and flag).

## The matching rule (what `Find-NeCode` does)

Given a parsed address and a quarter's index:

1. If the customer zip is 9-digit and matches a `Concat zip` exactly → `ZIP9`.
2. If the city is not a taxing city → `NO CODE-CITY` (blank).
3. Look up the block of tax rows with the same `street-key + zip5` (the engine's
   sorted column X). None → `NO MATCH`.
4. Within the block, keep the **last** row where the house number is in
   `[low, high]`, the odd/even side matches, and the suffix agrees (or either
   side is blank). If none, retry ignoring the suffix.
5. A match → that row's `City Code (final)`, flag `OK`. A block but no in-range
   row → `FALLBACK`, code `0`.

Rows that come back `NO MATCH`/`FALLBACK` get one fuzzy-repair retry before the
result is finalised.

## Known non-COM footguns (already guarded)

- **`@($array)` flattens** in PowerShell — a single result column must be passed
  as `(,$array)`; `Write-ResultBlock` throws if the column/header counts differ.
- **Variables are case-insensitive** — never pair `$R`/`$r` or `$HDR`/`$hdr`.
  Both bugs were caught by the mock-backed paste test before shipping.
