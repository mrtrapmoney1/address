================================================================================
 NE TAX CODE LOOKUP  -  README
================================================================================

WHAT THIS IS
------------
A PowerShell script that puts the Nebraska city tax code next to every address
on a working sheet.

The important part: it does NOT re-invent the matching logic. It feeds your
addresses into your own Excel template and lets the template's formulas work
them out, then carries the answers back. Excel does the thinking, the script
does the moving.

That is the whole reason the earlier version kept coming out wrong - it had
rebuilt the logic in code, and the copy disagreed with the template in five
different places. This one can't drift, because it IS the template.


================================================================================
 SETUP  (once)
================================================================================

The folder should end up looking exactly like this:

    NE Tax Code Lookup\
        Get-NeTaxCode.ps1           <- the script
        Address_Tax_Engine.xlsx     <- YOU ADD THIS (see below)
        README.txt
        Your Working Paper.xlsx     <- whatever you're working on
        _Address\                   <- the 27 quarterly state files
            2020 Q1 Address Data.xlsx
            ...
            2026 Q3 Address Data.xlsx
        _Backups\                   <- fills itself


MAKING Address_Tax_Engine.xlsx
------------------------------
It is just a clean copy of Address_Tax_Template_v3.xlsx:

  1. Copy Address_Tax_Template_v3.xlsx into this folder.
  2. Rename the copy to    Address_Tax_Engine.xlsx
  3. Open it.
  4. On the "Addresses" sheet, strip it back to ONE formula row. Row 3 is the
     master row the script copies down, so row 3 stays exactly as it is.
     Select each of these and press Delete (clear contents):

            C4:E29812        the pasted-in addresses
            F4:AL29812       the answers and every helper column

     (Use whatever your real last row is instead of 29812.)

     *** STOP AT AL. DO NOT TOUCH AN THROUGH BB. ***
     Columns AN:BB are the parsing lists - the ~282 taxing cities (AN),
     the direction dictionary (AX/AY) and the street-suffix dictionary
     (BA/BB). They are lookup tables, not per-row formulas, so they only
     ever occupy rows 3-284 and must be left exactly as they are.

     *** DO NOT DELETE THE ROWS THEMSELVES. ***
     Those lists sit out to the right on the same rows as your data. Delete
     rows 4:29812 and the city list goes with them - after which every row
     comes back NO CODE-CITY. Clear the two ranges above and nothing else.

     The script checks all three lists at startup and refuses to run if any
     of them is empty, so a slip here can't quietly poison a whole run.

  5. On the "TaxData" sheet, clear A3:Y downward. (Row 1 note and row 2
     headers stay.)
  6. Delete the "Match Detail" sheet if it's still there. It isn't used.
  7. Save and close. The file should now be small and open instantly - if
     it's still tens of megabytes, something didn't clear.

You don't have to do this - the script trims the engine itself at startup and
works fine either way. But the engine gets copied to a temp file on every run,
and 30,000 rows of that big AB array formula makes it slow to open and copy.
One row is worth the five minutes.

The script never writes to this file. It copies it to a temp file each run,
works in the copy, and throws the copy away when it's done. So the engine
stays clean forever.


================================================================================
 RUNNING IT
================================================================================

Right-click Get-NeTaxCode.ps1 > Run with PowerShell.

It asks you:
  - which workbook is the working paper (if there's more than one)
  - which sheet
  - which row has the headers
  - which column is the STREET ADDRESS
  - which column is the CITY
  - which column is the ZIP
  - which column is the DATE (YYYYMM)
  - which column is the STATE (optional)

It prints the headers it found first, so you pick by number or by name.
Press Enter to accept its guess when it guesses right.

Then it backs the workbook up, runs, and adds four new columns on the right:

    NE City Code      the answer, as a number      (template column F)
    Match Row #       the row of the state file it came from  (column G)
    Match Flag        how it got there             (column H)
    Source File       which quarterly file was used

Those first three are the template's F, G and H, in that order and untouched.
Nothing is re-ordered or re-derived on the way out.

Nothing existing is ever overwritten. A timestamped copy of the workbook goes
into _Backups before a single cell is written.


================================================================================
 WHAT THE FLAGS MEAN
================================================================================

TRUST THESE
  OK                 Street, house number, odd/even and suffix all confirmed.
  ZIP9               Matched on the full 9-digit zip.

CHECK THESE
  ... REPAIRED: x    The address didn't match as typed, so it was cleaned up
                     and tried again - and the clean version matched. The note
                     says exactly what was changed ("misspelled street suffix",
                     "direction on the wrong end", and so on). Usually right,
                     but it IS a guess. Spot-check these.

FIX BY HAND
  ENGINE ERROR       The template returned an error (#N/A and the like) for
                     that row. Code and row number left blank rather than
                     writing the error's internal number as a tax code.
  FALLBACK           Street and zip exist, but the house number falls outside
                     every range on file. Code is 0 on purpose.
  NO MATCH           Nothing lined up.
  NO ADDRESS         The address cell was blank.

NORMAL, NO ACTION
  NO CODE-CITY       That city levies no local sales tax.
  OOS                Outside Nebraska. Code left blank.
  NO DATE            Date cell couldn't be read as a period.
  NO TAX FILE - x    No quarterly file in _Address covers that period.


================================================================================
 FIRST RUN  -  PROVE IT BEFORE YOU TRUST IT
================================================================================

Do this once. It takes ten minutes and it's the whole reason to bother.

  1. Make a copy of your working paper with about 1,000 rows in it.
  2. Run the script on it.
  3. Take the same 1,000 addresses and paste them into your real template by
     hand, the way you always have.
  4. Compare the two code columns.

They should match on every single row. If they don't, stop and send me the
rows that disagree - that's a bug worth fixing before it touches real work.

While you're there, note how long the 1,000 rows took. Multiply up. That tells
you what the big file is going to cost before you commit to it.


================================================================================
 SETTINGS  (top of the script, first 20 lines)
================================================================================

$BatchSize       How many rows go through the template at a time.
                 5000 is safe on 32-bit Excel. On 64-bit try 20000 - it will
                 run noticeably faster. If Excel runs out of memory, halve it.

$TryRepairs      $true  = failed rows get cleaned up and retried (flagged).
                 $false = pure template output, nothing guessed at all.
                 Set this to $false if you ever need a run that reconciles
                 to the template exactly, row for row.

$NoCodeCityZero  $false = NO CODE-CITY rows get a blank code (what the template
                          does today).
                 $true  = they get 0 instead.

$KeepEngine      $true = after the run, the engine is saved beside your working
                 paper as _Engine_last_run.xlsx, holding the LAST batch that
                 went through it. Open it any time the numbers look off:
                 C:E shows exactly what was fed in, F:H exactly what came
                 back, and TaxData shows which quarter was loaded. That file
                 is the whole audit trail.

$ShowSample      How many finished rows get printed to the screen during the
                 first batch. Lets you spot a wiring problem in ten seconds
                 instead of an hour. 0 turns it off.

$ReadChunk       How many rows are read from / written to the working paper at
                 a time. 50000 is fine. No reason to touch it.


================================================================================
 THINGS WORTH KNOWING
================================================================================

- PO Boxes work, because the template handles them. The script leaves box
  addresses alone rather than "repairing" them.

- The template's lookups reach row 400,001 of TaxData, which is 399,999 data
  rows. If a quarterly file ever grows past that, the script stops and says so
  rather than quietly dropping rows off the end.

- Before loading a quarterly file, the script checks that column G still says
  "Street Name" and column Y still says "City Code (final)". If the state ever
  changes the layout, it stops instead of writing garbage.

- The quarterly files must stay sorted by column X (the _key column). The
  template's fast lookup depends on it. They already are - just don't re-sort
  them by something else.

- It's slower than pure code, because Excel is doing the work. That's the
  trade. Correct and slower beats fast and wrong, and the fast version was
  wrong three times.
