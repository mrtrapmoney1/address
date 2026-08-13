================================================================================
 NE TAX CODE LOOKUP  -  README  (v4)
 Address  ->  Nebraska city tax code, straight onto your working paper.
================================================================================


################################################################################
#  QUICK START  -  GET THE CODES
################################################################################

STEP 1 - PUT THESE IN ONE FOLDER (the working folder)
-----------------------------------------------------
    <working folder>\
        Get-NeTaxCode.ps1           <- the script
        NeTaxPaste.ps1              <- ships next to it (the script needs it)
        Address_Tax_Engine.xlsx     <- YOU ADD THIS (clean copy of your template)
        Your Working Paper.xlsx     <- the sheet you want codes on
        _Address\                   <- the quarterly state files
            2021 Q3 Address Data.xlsx
            2021 Q4 Address Data.xlsx
            ...
        _Backups\                   <- created for you; leave it

(That's the whole setup. _Backups fills itself. Making the engine file is a
one-time step - see "MAKING THE ENGINE" further down.)


STEP 2 - RUN IT
---------------
    powershell -ExecutionPolicy Bypass -File .\Get-NeTaxCode.ps1

Answer the prompts (it guesses the columns - press Enter to accept):
    which sheet / header row
    which columns are ADDRESS / CITY / ZIP / DATE / STATE
    where to put the result  ->  ENTER = end of sheet, or type a column (e.g. T)
    add a Match Flag column too?  ->  ENTER = no (code only)


STEP 3 - WHAT YOU GET
---------------------
    NE City Code   ONE new column, formatted 000 (0 -> 000, 94 -> 094).
                   Blank where there's no local code (out of state, no match).
    Match Flag     ONLY if you asked for it - how each code was found (OK,
                   FALLBACK, NO MATCH, ...).

The working paper is saved automatically. A full backup of it is taken first,
and the raw engine answer is saved to _Backups\...ENGINE... (one sheet per
quarter) in case you want to see it.

UNATTENDED (big scripted runs):
    powershell -File .\Get-NeTaxCode.ps1 -WorkingPaper ".\Big.xlsx" `
        -SheetName Data -HeaderRow 1 -PasteAt END -MaxParallel 4 -NonInteractive
    add -IncludeFlag to also write the Match Flag column.


################################################################################
#  EXPLANATION
################################################################################

WHAT IT IS
----------
The script does NOT re-invent the matching logic. It feeds your addresses into
your own Excel template (Address_Tax_Engine.xlsx), lets the template's formulas
work them out, and carries the answers back. Excel does the thinking; the
script moves data. The workbook-writing half lives in NeTaxPaste.ps1 so it can
be tested on its own (see tests\).

WHAT'S NEW IN v4  (full write-up in docs\CHANGES.md)
----------------------------------------------------
  * CLEAN, MINIMAL OUTPUT. One column by default - the code, formatted 000.
    Match Flag is opt-in. Row # / Source File stay out of the working paper.
  * NO MORE SCRAMBLE. Each column is written on its own as an (N x 1) block, so
    two columns can never swap places or jam into one cell.
  * YOU CHOOSE WHERE IT LANDS. Append at the end, or INSERT at a column you name
    (existing data shifts right, never overwritten).
  * THE ENGINE RUN IS BACKED UP. The raw pass-1 answer is value-pasted to one
    sheet per quarter before the fuzzy repair can change it.
  * QUARTERS RUN CONCURRENTLY. Each quarter in its own Excel process. -MaxParallel 1
    for the old sequential behaviour.
  * IT ALWAYS SAVES, with a whole-workbook backup taken first.

MAKING THE ENGINE  (Address_Tax_Engine.xlsx, one time)
------------------------------------------------------
It's a clean copy of your template:
  1. Copy your Address_Tax_Template into the folder; rename it Address_Tax_Engine.xlsx
  2. On "Addresses": clear C4:E<last> and F4:AL<last>. LEAVE row 3 (the master
     formula row) and LEAVE columns AN:BB (the parsing lists) exactly as they are.
     Do NOT delete the rows themselves - the lists live on those rows.
  3. On "TaxData": clear A3:Y downward (keep row 1 note, row 2 headers).
  4. Save and close.
The script never writes to this file (it copies it per run) and checks the three
parsing lists at startup, refusing to run if any is empty.

WHERE RESULTS GO
----------------
    ENTER / -PasteAt END   append after the last used column
    a column, e.g. T       INSERT the result column(s) at T; data shifts right
    a header name          insert in front of that column
Nothing existing is ever overwritten.

WHAT THE FLAGS MEAN  (only shown if you add the flag column)
------------------------------------------------------------
  TRUST:     OK (full match), ZIP9 (matched on 9-digit zip)
  CHECK:     ... REPAIRED: x  (matched only after cleanup; raw answer is in the
                               engine backup)
  FIX:       FALLBACK (house # outside every range; code 0), NO MATCH, NO ADDRESS,
             BAD CODE (engine returned something that isn't a 0-999 code; blanked)
  NO ACTION: NO CODE-CITY (city has no local tax), OOS (out of state),
             NO DATE, NO TAX FILE - x (no quarterly file for that period)

SETTINGS  (top of Get-NeTaxCode.ps1)
------------------------------------
  $BatchSize        rows per pass through the template. 5000 (32-bit) / 20000 (64-bit)
  $MaxParallel      quarters at once (also -MaxParallel). 1 = sequential
  $TryRepairs       $true = failed rows cleaned up and retried (pass 2)
  $BackupEngineRun  $true = write pass-1 engine answers to per-quarter backup sheets
  $CodeNumberFormat '000' (three digits). 'General' shows 94 and 0
  $NoCodeCityZero   $false = NO CODE-CITY code left blank; $true = 0

VERIFY IT  (tests\ - all PowerShell; works in Windows PowerShell 5.1)
--------------------------------------------------------------------
  .\tests\Run-AllTests.ps1                 logic + paste tests (green = good)
  .\tests\Check-Workpaper.ps1 -Path ".\Working Paper.xlsx" -HeaderRow 4
                                           scan a finished paper for scramble/bad codes
  (or:  powershell -ExecutionPolicy Bypass -File tests\Run-AllTests.ps1 )
  You do NOT need pwsh / PowerShell 7 - 'pwsh' just means PS7, which you may not
  have. Run the .ps1 files directly with the powershell you already have.
See tests\README.md for what each one checks.

THINGS WORTH KNOWING
--------------------
  * PO Boxes work; the repair pass leaves them alone.
  * TaxData lookups reach 399,999 rows; a bigger quarter stops the run, not drops rows.
  * Before loading a quarter it checks column G = "Street Name", Y = "City Code (final)".
  * Quarterly files must stay sorted by column X. They already are; don't re-sort.
  * Concurrency uses one Excel process per worker. Tight on RAM? lower -MaxParallel.
  * Prove it first: run ~1,000 rows, compare to a hand run, then scale up.
