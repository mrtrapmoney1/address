================================================================================
 NEBRASKA CITY TAX CODE  -  README
 Address -> Nebraska city tax code, straight onto your working paper.
 Pure PowerShell. No engine workbook. No Python.
================================================================================


################################################################################
#  QUICK START  -  GET THE CODES
################################################################################

STEP 1 - PUT THESE IN ONE FOLDER
--------------------------------
    <working folder>\
        Get-CityCode.ps1            <- run this
        NeMatch.ps1                 <- the matching logic (ships beside it)
        NeDicts.ps1                 <- city / direction / suffix lists
        NeTaxPaste.ps1              <- the workbook writer + Report builder
        Your Working Paper.xlsx     <- the sheet you want codes on
        _Address\                   <- the quarterly tax files
            2021 Q3 Address Data.xlsx
            2021 Q4 Address Data.xlsx
            ...
        _Backups\                   <- created for you

STEP 2 - RUN IT
---------------
    powershell -ExecutionPolicy Bypass -File .\Get-CityCode.ps1

Answer the prompts (it guesses the columns - press Enter to accept):
    which sheet / header row
    which columns are ADDRESS / CITY / ZIP / DATE / STATE
    where to put the result  ->  ENTER = end of sheet, or a column (e.g. T)
    add a Match Flag column too?  ->  ENTER = no (code only)

STEP 3 - WHAT YOU GET
---------------------
    NE City Code   one new column, formatted 000 (0 -> 000, 94 -> 094).
    Match Flag     only if you asked for it.
    Report (sheet) a new tab: flag summary + every UNIQUE address with its
                   code, flag and how many times it appears - your analysis
                   view now that there's no engine sheet to eyeball.

The workpaper is saved automatically; a full backup is taken first.

UNATTENDED:
    powershell -File .\Get-CityCode.ps1 -WorkingPaper ".\Big.xlsx" `
        -SheetName Data -HeaderRow 1 -PasteAt END -IncludeFlag -NonInteractive


################################################################################
#  HOW IT WORKS  (and why the codes are right)
################################################################################

Get-CityCode.ps1 does NOT use the Excel engine template. It is a faithful,
line-for-line PORT of that template's Addresses-sheet formulas (columns J..AB
and F/G/H) into PowerShell - NeMatch.ps1. The three lookup lists (282 taxing
cities, directions, street suffixes) are the exact ones lifted out of your
engine, in NeDicts.ps1.

For each address it does what the sheet does:
  1. clean + parse the address into house number, street name, suffix, zip5;
  2. gate on the city (must be a taxing city; Saint Paul / St Paul both count);
  3. find the street+zip block in the quarterly tax file (sorted key column X);
  4. pick the row whose house number is in range, odd/even side matches, and
     suffix agrees (falling back to ignore the suffix if nothing else fits);
  5. return that row's City Code (final). No match in range -> FALLBACK (code 0).

FUZZY MATCH (pass 2): rows that come back NO MATCH / FALLBACK are cleaned up
(misspelled suffix, direction on the wrong end, spelled-out street number,
missing space after the house number, stray unit number, Saint->St) and tried
again. A win is flagged "OK - REPAIRED: <what was fixed>".

MANY QUARTERS: rows are grouped by their DATE into quarters, and each quarter is
matched against its own file in _Address (named "2021 Q3 Address Data.xlsx" etc).

FLAGS
    OK              street + house number + odd/even + suffix all matched
    ZIP9            matched on a full 9-digit zip
    ... REPAIRED    matched only after the fuzzy cleanup (spot-check these)
    FALLBACK        street+zip exist but house number out of range (code 0)
    NO MATCH        nothing lined up
    NO CODE-CITY    city has no local tax (blank code)
    OOS             not Nebraska (blank)
    NO DATE / NO TAX FILE   date unreadable / no quarter file for that period


################################################################################
#  VERIFY IT  (all PowerShell - works in Windows PowerShell 5.1)
################################################################################

    .\tests\Run-AllTests.ps1
        Test-Logic  (58)  pure helpers
        Test-Match  (27)  the ported matcher: ranges, odd/even, suffix fallback,
                          numbered streets, PO boxes, Saint/St, fuzzy repair
        Test-Paste  (38)  the workbook write (append/insert/no-overwrite/000)
                          and the Report sheet - real Excel if present, else a mock

    .\tests\Check-Workpaper.ps1 -Path ".\Working Paper.xlsx" -HeaderRow 4
        scan a finished paper for scrambled cells or non-000 codes

The matcher was cross-checked against an independent second implementation on
116 real sample addresses using real tax data - identical on every row.


################################################################################
#  NOTES
################################################################################

- The quarterly files must be the prepped layout (column G = "Street Name",
  column Y = "City Code (final)"), sorted by column X. The script checks this
  and stops rather than writing garbage.
- Nothing on the working paper is overwritten: result columns are brand new,
  and the Report sheet is regenerated each run (it's our own output).
- The older engine-COM script (Get-NeTaxCode.ps1) is kept for reference, but
  Get-CityCode.ps1 is the recommended path - simpler and no engine workbook.
