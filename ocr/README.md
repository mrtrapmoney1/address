# OCR — scanned invoices to positioned words

> **This is a standalone tool.** It is a reader for scanned PDFs and invoices,
> and it has nothing to do with the Nebraska city-tax-code matcher in the rest
> of this repository — no shared code, no shared data, no shared runtime.
> It lives in its own folder so the whole thing can be copied somewhere else
> and run on its own. Delete everything outside `ocr\` and it still works.


Point it at a folder of scanned invoices. It reads every page and writes out
**every word with its position on the page**, grouped into lines and columns.

**PowerShell + Node.js + JavaScript only.** Nothing is downloaded when it runs.

---

## The problem this solves

Run any OCR engine over an invoice and ask it for text, and you get this:

```
01Aug2018 Incoming Interac e-Transfer      1454 101,583.92
```

Is `1454` a cheque number or a debit? On the page it is obvious — it sits under
a column heading. In the text it is **unknowable**, and no parsing downstream
can recover it. The information was destroyed the moment the words were
flattened into a string.

The engine knew where every word sat. This keeps that:

```
| Date      | Description               | Number | Debits | Credits | Balance    |
| 01Aug2018 | Incoming Interac e-Transf |        |        | 1454    | 101,583.92 |
```

That is the whole point of the tool.

---

## Setting it up (once)

**1.** Put the `ocr` folder wherever you like — a network share, a USB stick,
your desktop. Nothing is registered or installed into Windows.

**2.** Install [Node.js](https://nodejs.org) (the LTS build) if it is not there.

**3.** Run this once, on a machine with a network:

```powershell
.\Invoke-Ocr.ps1 -Install
```

That puts the OCR engine and the English language data into `ocr\js\node_modules`.
**After this the tool never touches the network again** — which matters, because
the machines this runs on are usually the ones that cannot reach a CDN.

---

## Using it

Just run it:

```powershell
.\Invoke-Ocr.ps1
```

It asks for the folder. Paste the path, or drag the folder onto the window:

```
  Paste the folder holding the PDFs to read
  (drag the folder onto this window, or paste the path, then press Enter)

  folder> C:\Invoices\August

  [ok]   Found 14 file(s), 22.8 MB:
         12 PDF(s)
         2 page image(s)
           ACME-8841.pdf
           ...

  Read these? [Y/n]
```

Quotes around the path are fine — Explorer's "Copy as path" adds them, and so
does dragging a folder whose name has a space in it.

Or skip the prompt:

```powershell
.\Invoke-Ocr.ps1 -Path C:\Invoices\August -Csv -Layout
```

### What you get

A new folder per run, so a second run can never overwrite the first:

```
_OcrRuns\2026-08-24_141530_August\
    ocr\
        ACME-8841.ocr.xml        every word, its box, its column, its confidence
        ACME-8841.words.csv      the same as a spreadsheet        (-Csv)
        ACME-8841.layout.txt     the page drawn back as text      (-Layout)
    pages\                       the page images that were read
    run.log                      everything the run printed
```

---

## Reading the results in PowerShell

```powershell
. .\OcrDocument.ps1
$doc = Import-OcrDocument .\_OcrRuns\...\ocr\ACME-8841.ocr.xml

Show-OcrPage $doc                    # the page, drawn back as text
Get-OcrWords $doc                    # every word with Page/Text/X/Y/W/H/Size
Get-OcrTable $doc                    # the line items, as rows of cells
Get-OcrInvoiceSummary $doc           # one flat row: invoice no, dates, totals
```

`Show-OcrPage` is the first thing to run when something is missing. **If a
number is not visible there, no rule downstream can find it** — the problem is
in the scan, not in the parsing.

---

## What it does to a page, and why

| Stage | Why |
|---|---|
| **Lift the image out of the PDF** | A scanned PDF wraps one photograph per page. Lifting it out gives the scanner's own pixels; *rendering* the page resamples them and softens every letter edge. |
| **Grey, then stretch the contrast** | So the threshold step sees a full black-to-white range instead of a washed-out one. |
| **Scale to 300 dpi** | The single biggest win. Tesseract's own guide: *"the same image will get much better results if you upscale it"*. A 150 dpi fax has eight-pixel-tall letters. |
| **Threshold locally (Sauvola)** | One threshold for the whole page loses a side that is in shadow. A local one does not. |
| **Straighten** | Measured to a tenth of a degree from the ink itself. |
| **De-speckle** | Scanner dust otherwise becomes full-confidence punctuation. |
| **Recognise** | tesseract.js, workers reused across the batch. |
| **Re-read weak lines** | Lines under 70% are read again alone, where the engine can pick settings that suit one line. **A retry is only kept if it beats what it replaces.** |
| **Rebuild the layout** | Lines by box overlap, columns from the white channels between them. |

Measured on the sample bad scan (`samples\page_scan.png` — 45% size, lit
unevenly, noisy, tilted 1.8°):

| | words read correctly | confidence |
|---|---|---|
| handed straight to the engine | 79.2% | 86 |
| after preparation | **100%** | **94** |

That comparison is a test (`js\test\ocr.test.js`), not a claim — if preparation
ever stops helping, the suite fails.

---

## Invoices specifically

The invoice pass runs by default (`-NoInvoice` turns it off).

**Misread digits are repaired.** Tesseract confuses `O`/`0`, `l`/`1`, `S`/`5`,
`B`/`8`. In prose that is a typo you read past; in `1,S00.OO` it is a number
that will not parse. Inside a token that is unmistakably an amount, those are
undone — and **the repair is only kept if the result actually reads as a
number**, so nothing can be mangled by a wrong guess. Every repair is written
into the XML with the original text, so it can be audited.

Words are never touched: `GOODS`, `Invoice`, `ISBN` come through untouched,
because nothing without a digit in it is even considered.

**Columns get meaning** — `date`, `number`, `amount`, `text` — so line items
can be read as figures.

**The headline fields are found by position:** invoice number, invoice date,
due date, PO number, subtotal, freight, discount, tax, tax rate, total. A label
and its value are two separate words; what joins them is that one sits to the
right of, or below, the other.

**And the arithmetic is checked:**

```
subtotal + freight - discount + tax  =  total
```

When it holds, every one of those numbers is confirmed by the others — far
stronger evidence than any per-character confidence score. When it does not,
**the invoice is flagged and nothing is changed**:

> the amounts do not add up: subtotal 853.00 + freight 35.00 - discount 0.00 +
> tax 62.16 = 950.16, but the total reads 999.99 (out by -49.83)

---

## The output format

`docs/OCR-FORMAT.md` has the full description. The short version: coordinates
are **points from the top-left of the page**, the same frame the PDF text
extractor uses — so a scanned invoice and a text one arrive downstream in the
same shape, and part two of this job (getting it into Excel) has one input
format rather than two.

```xml
<Page number="1" width="612" height="792" unit="pt" dpi="300" skew="1.4" confidence="93.3">
  <Region index="0" kind="table" columns="6">
    <Columns count="6"><Boundary index="0" x="105.1" /> ... </Columns>
    <Line index="4" y="27.4" confidence="95">
      <Word x="7.8" y="27.4" w="17.9" h="3.1" size="3.1" conf="96" col="0">01Aug2018</Word>
      <Word x="135.4" y="27.4" w="9.8" h="3.1" size="3.1" conf="94" col="3">36.07</Word>
    </Line>
  </Region>
  <Invoice fields="9" repairs="1" reconciles="yes">
    <Field name="total" value="925.04" label="Total Due" from="right" conf="92">925.04</Field>
    <Arithmetic ok="1" computed="925.04" difference="0" />
    <Repair from="45O.OO" to="450.00" column="3" />
  </Invoice>
</Page>
```

---

## Checking it works

```powershell
.\tests\Run-AllTests.ps1
```

| Suite | Covers |
|---|---|
| **Test-Syntax** (46) | every file parses, is pure ASCII, and uses no PowerShell 7-only syntax |
| **Test-Common** (46) | run folders, logging, input listing, formatting, Node detection |
| **Test-Document** (71) | reading the XML back, missing attributes, culture, arrays that stay arrays |
| **Test-EndToEnd** (25) | a real run over the samples, start to finish |
| **the JavaScript suite** (93) | PNG codec, XML, preprocessing, segmentation, invoices, PDFs, live OCR |

**281 checks.** They pass from a clean copy with nothing installed but Node.

---

## Options

```
-Path PATH            the folder to read (asks if not given)
-Csv                  also write every word as a CSV
-Layout               also draw each page back as text
-Engine tesseract     tesseract (default) | windows | both
-Language eng         eng for Tesseract, en-GB for Windows OCR
-PageSegmentation 3   3 auto (default), 4 single column, 6 one block, 11 sparse
-SourceDpi 0          what the images ARE (read from the file when it says)
-TargetDpi 300        what to scale them to
-Threshold sauvola    sauvola (default) | otsu
-Workers 0            worker threads (default: min(4, cpus))
-RefineBelow 70       re-read lines under this confidence
-NoInvoice            skip the invoice pass
-NoPreprocess         hand images to the engine untouched
-NoDeskew  -NoRefine  -Crop  -KeepImages  -Recurse  -NonInteractive
-Install              the one-time setup
```

---

## Limits, stated plainly

- **PDFs**: scans compressed with **CCITT fax** or **JBIG2** are not decoded
  in JavaScript. On Windows the tool falls back to the built-in renderer
  automatically. On anything else it says so and names the file.
- **TIFF** needs Windows (`System.Drawing`); Tesseract.js cannot read TIFF.
- **Handwriting** is not supported. Tesseract is built for printed text.
- **The Windows OCR engine** reports no per-word confidence, so the review
  queue is less useful with `-Engine windows`.
- **A page is read, not understood.** Fields are found by label and position;
  an invoice with no label next to its total will not yield a total.

---

## Where things live

```
Invoke-Ocr.ps1        the one script you run
OcrCommon.ps1         run folders, logging, finding Node
OcrDocument.ps1       reading the XML back into PowerShell objects
OcrNode.ps1           running the Node half
OcrRaster.ps1         PDF and TIFF rendering (Windows.Data.Pdf, System.Drawing)
OcrWindows.ps1        the Windows OCR engine, with word boxes
js\ocr-cli.js         images in, OcrDocument XML out
js\ocr-segment.js     words in, OcrDocument XML out (shared by both engines)
js\lib\png.js         PNG decode/encode, zlib only
js\lib\preprocess.js  greyscale, upscale, Sauvola, deskew, de-speckle
js\lib\segment.js     lines, columns, reading order   <- the important one
js\lib\invoice.js     digit repair, field finding, the arithmetic check
js\lib\pdfimg.js      lifting page images out of a scanned PDF
js\lib\engine.js      tesseract.js: offline data, worker pool, second pass
js\lib\emit.js        the OcrDocument XML
samples\              two page images the tests use
tests\                the PowerShell suites
```

---

## Credits

- **[tesseract.js](https://github.com/naptha/tesseract.js)** (Apache-2.0) — the
  OCR engine.
- **[Invoke-Transfer](https://github.com/JoelGMSec/Invoke-Transfer)** by
  @JoelGMSec — the clearest published example of driving the WinRT OCR and
  imaging APIs from Windows PowerShell 5.1. `OcrRaster.ps1` and
  `OcrWindows.ps1` follow its approach to loading the projected types and
  blocking on `IAsyncOperation`. Where it takes `OcrResult.Text` — right for
  its purpose — this walks `Lines`/`Words` and keeps every bounding box.
