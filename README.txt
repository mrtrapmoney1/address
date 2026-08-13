================================================================================
 NE TAX CODE LOOKUP  -  README  (v4)
================================================================================

WHAT THIS IS
------------
A PowerShell script that puts the Nebraska city tax code next to every address
on a working sheet.

It does NOT re-invent the matching logic. It feeds your addresses into your own
Excel template (Address_Tax_Engine.xlsx), lets the template's formulas work them
out, and carries the answers back. Excel does the thinking; the script moves.

WHAT'S NEW IN v4  (full write-up in docs/CHANGES.md)
----------------------------------------------------
  * QUARTERS RUN CONCURRENTLY. Each quarter gets its own Excel process and its
    own copy of the engine. Set -MaxParallel 1 for the old sequential behaviour.
  * THE ENGINE RUN IS BACKED UP. Pass 1 (the raw template answer) is preserved
    and written, value-pasted, to one sheet per quarter in a separate backup
    workbook - BEFORE the fuzzy repair (pass 2) can change anything. The old
    script let the repair overwrite the engine answer; that no longer happens.
  * YOU CHOOSE WHERE RESULTS GO. Press Enter to append the result columns at the
    end of the sheet, or name a column and the script INSERTS new columns there.
    Existing data is shifted right, never overwritten. Columns are written one at
    a time, so they can never swap places.
  * IT ALWAYS SAVES. The working paper and the engine backup are both saved
    before the script hands control back, with the path printed.


================================================================================
 SETUP  (once)
================================================================================

The folder should look like this:

    NE Tax Code Lookup\
        Get-NeTaxCode.ps1           <- the script
        Address_Tax_Engine.xlsx     <- YOU ADD THIS (clean copy of your template)
        README.txt
        Your Working Paper.xlsx     <- whatever you're working on
        _Address\                   <- the quarterly state files
            2020 Q1 Address Data.xlsx
            ...
            2026 Q3 Address Data.xlsx
        _Backups\                   <- fills itself

MAKING Address_Tax_Engine.xlsx
------------------------------
It is just a clean copy of your template:
  1. Copy your Address_Tax_Template into this folder.
  2. Rename the copy to    Address_Tax_Engine.xlsx
  3. On "Addresses", clear C4:E<last> and F4:AL<last>. LEAVE ROW 3 (the master
     formula row) and LEAVE columns AN:BB (the parsing lists) exactly as they are.
     Do NOT delete the rows themselves - the lists live on those same rows.
  4. On "TaxData", clear A3:Y downward (keep the row 1 note and row 2 headers).
  5. Save and close.

The script never writes to this file - it copies it per worker on each run and
throws the copies away. The engine stays clean forever. It also checks the three
parsing lists at startup and refuses to run if any is empty.


================================================================================
 RUNNING IT
================================================================================

Right-click Get-NeTaxCode.ps1 > Run with PowerShell, or from a prompt:

    powershell -ExecutionPolicy Bypass -File .\Get-NeTaxCode.ps1

It asks you:
  - which workbook is the working paper (if there's more than one)
  - which sheet, which header row
  - which columns are ADDRESS / CITY / ZIP / DATE(YYYYMM, optional) / STATE(optional)
  - WHERE THE RESULTS GO:
        press Enter          -> append four columns at the end of the sheet
        type a column, e.g. T -> INSERT four new columns at T (data shifts right)
        type a header name    -> insert in front of that column

Then it backs the whole workbook up, runs every quarter (in parallel), writes:

    NE City Code      the final answer, as a number   (template column F)
    Match Row #       the state-file row it came from  (column G)
    Match Flag        how it got there                 (column H)
    Source File       which quarterly file was used

...and saves. Separately, it drops the ENGINE (pass-1) answer into
_Backups\<paper>_ENGINE_<timestamp>.xlsx, one value-pasted sheet per quarter.

UNATTENDED / SCRIPTED
---------------------
    powershell -File .\Get-NeTaxCode.ps1 `
        -WorkingPaper ".\Big File.xlsx" -SheetName "Data" -HeaderRow 1 `
        -PasteAt END -MaxParallel 4 -NonInteractive

  -PasteAt END        append at the end   |   -PasteAt T   insert at column T
  -MaxParallel 1      strictly sequential (old behaviour)
  -MaxParallel 0      auto (half your cores, capped at 6)
  -NonInteractive     never prompt; auto-detect columns, fail loudly if it can't


================================================================================
 WHAT THE FLAGS MEAN
================================================================================

TRUST THESE
  OK                 Street, house number, odd/even and suffix all confirmed.
  ZIP9               Matched on the full 9-digit zip.

CHECK THESE
  ... REPAIRED: x    Didn't match as typed, so it was cleaned up and retried, and
                     the clean version matched. The note says what was changed.
                     The RAW engine answer for these rows is in the engine backup.

FIX BY HAND
  ENGINE ERROR       Template returned an error; code/row left blank.
  FALLBACK           Street + zip exist, house number outside every range. Code 0.
  NO MATCH           Nothing lined up.
  NO ADDRESS         The address cell was blank.
  BAD CODE           Engine returned something that is not a 0-999 code. Left blank.

NORMAL, NO ACTION
  NO CODE-CITY       That city levies no local sales tax.
  OOS                Outside Nebraska. Code left blank.
  NO DATE            Date cell couldn't be read as a period.
  NO TAX FILE - x    No quarterly file in _Address covers that period.


================================================================================
 FIRST RUN  -  PROVE IT BEFORE YOU TRUST IT
================================================================================

  1. Copy your working paper down to ~1,000 rows.
  2. Run the script on it.
  3. Paste the same 1,000 addresses into your real template by hand.
  4. Compare the code columns. They should match on every row.

Note how long 1,000 rows took, multiply up, and that tells you the big file's
cost. Then raise -MaxParallel and watch it drop.


================================================================================
 SETTINGS  (top of the script)
================================================================================

$BatchSize        rows per pass through the template. 5000 (32-bit) / 20000 (64-bit).
$MaxParallel      quarters at once. Also the -MaxParallel param. 1 = sequential.
$TryRepairs       $true = failed rows cleaned up and retried (pass 2).
$BackupEngineRun  $true = write the pass-1 engine answer to per-quarter backup sheets.
$KeepEngine       $true = save one engine copy per quarter (_Engine_last_run_*.xlsx).
$NoCodeCityZero   $false = NO CODE-CITY rows get a blank code; $true = 0.
$DumpCsv          how many finished rows to also dump to _Results_check.csv.


================================================================================
 THINGS WORTH KNOWING
================================================================================

- PO Boxes work; the template handles them and the repair pass leaves them alone.
- The template's lookups reach TaxData row 400,001 (399,999 data rows). A bigger
  quarter stops the run rather than dropping rows.
- Before loading a quarter, the worker checks column G still says "Street Name"
  and column Y still says "City Code (final)". Layout change -> it stops.
- Quarterly files must stay sorted by column X. They already are; don't re-sort.
- Concurrency needs one Excel process per worker. If a machine is tight on RAM,
  lower -MaxParallel. Correct and a bit slower beats fast and wrong.
