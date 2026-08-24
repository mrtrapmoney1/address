# Architecture

## The split, and why it is where it is

```
        Invoke-Ocr.ps1                 asks for the folder, runs the batch,
              │                        reports, and owns the run folder
    ┌─────────┼──────────────┐
    ▼         ▼              ▼
 OcrRaster  OcrNode      OcrWindows    PowerShell: I/O, Windows APIs, orchestration
 (PDF/TIFF) (runs node)  (Windows OCR)
              │              │
              ▼              ▼
        js/ocr-cli.js   js/ocr-segment.js
              │              │
              └──────┬───────┘
                     ▼
        png · preprocess · segment · invoice · emit   JavaScript: all the algorithms
```

**PowerShell decides what to read and where results go. JavaScript does the
work.** That line is drawn deliberately: everything algorithmic lives where it
can be unit-tested without Windows, Excel, or a scanner, and the PowerShell
layer stays thin enough to be checked by reading it.

## Files

| File | Responsibility | Tested by |
|---|---|---|
| **Invoke-Ocr.ps1** | The prompt, the run folder, staging, the fallback to rendering, the report | `Test-EndToEnd` |
| **OcrCommon.ps1** | Run folders, logging, finding Node, listing input | `Test-Common` |
| **OcrDocument.ps1** | XML → PowerShell objects; word stream, tables, invoice summary, page view | `Test-Document` |
| **OcrNode.ps1** | Running the Node half; the summary comes back through a **file**, never stdout | `Test-EndToEnd` |
| **OcrRaster.ps1** | `Windows.Data.Pdf` rendering, `System.Drawing` TIFF splitting | (Windows only) |
| **OcrWindows.ps1** | The Windows OCR engine, keeping every word box | (Windows only) |
| **js/lib/png.js** | PNG decode/encode on zlib alone | `png.test.js` |
| **js/lib/preprocess.js** | Greyscale, upscale, Sauvola, deskew, de-speckle, and the inverse transform | `preprocess.test.js` |
| **js/lib/segment.js** | XY-cut, line grouping, column detection, reading order | `segment.test.js` |
| **js/lib/invoice.js** | Digit repair, column meaning, field finding, arithmetic | `invoice.test.js` |
| **js/lib/pdfimg.js** | Lifting page images out of a scanned PDF | `pdf.test.js` |
| **js/lib/engine.js** | tesseract.js: offline language data, worker pool, second pass | `ocr.test.js` |
| **js/lib/emit.js** | The OcrDocument XML, and reading it back | `xml.test.js` |

## Decisions worth knowing about

**The language data is local.** tesseract.js downloads ~2–4 MB from a CDN on
first use, which fails outright behind a corporate proxy — exactly the machines
this runs on. Installed from npm and pointed at with `langPath`, it never
touches the network.

**Workers are reused.** The tesseract.js performance guide is blunt about
creating one per image: *"This is never the correct option."* A scheduler with a
small pool does a hundred pages in the time a naive loop does a dozen.

**The summary comes back through a file.** Capturing a child process's stdout
from PowerShell while also letting its progress reach the console means merging
streams, and one warning from a dependency then lands in the middle of the JSON.
A file has none of those failure modes.

**Both engines share one segmenter.** The Windows engine writes a word list;
`ocr-segment.js` turns it into the same XML. A second implementation of the
layout rules in PowerShell would drift, and then a page would segment
differently depending on which engine read it.

**PDFs go to JavaScript first, Windows second.** Lifting the embedded image out
of a scanned PDF gives the scanner's own pixels with no resampling; rendering
the page decodes and repaints them. The Windows renderer is the fallback for
CCITT and JBIG2 scans, which JavaScript does not decode.

## The segmentation rules, and the bugs behind them

Every one of these was written because the obvious version was wrong on real
pages.

**Lines are grouped by box overlap, never by rounding y.** A line with a
capital and a comma has words whose tops differ by several pixels; any rounding
boundary eventually falls between two words of one line.

**A vertical cut is refused where the rows carry on across it.** The wide white
channel between an invoice's Description and its Amount column looks identical
in a projection profile to the gap between two independent blocks — but cutting
it severs every row from its own numbers.

**That refusal is measured as a run, not a fraction.** A real invoice has its
line-item table above a totals block that sits entirely on the right. Those
totals rows do not span the channel, and they dragged the fraction below the
threshold — so the cut went through and sliced the item table into three
pieces. A band of consecutive rows all reaching across the channel is a table
whatever the rest of the page is doing.

**Column edges come from the white channels, not from the midpoints between
alignment anchors.** A column that produced an anchor for its left edge and one
for its right got a boundary through its own middle: `800.04` and `823.34`,
printed one above the other, landed in different columns.

**And a boundary that cuts through words is rejected outright.** That single
test removes the whole class of error.

**Gutters are measured in rows, not ink.** Weighting by ink made the threshold
impossible to reason about. Counting rows says exactly what is meant — *a
gutter is a strip hardly any row has text in* — and a heading set across the
top of a table then costs one row instead of erasing every column edge below it.

## The PowerShell footguns, all of them caught by running the code

- **Missing XML attributes throw.** Under `Set-StrictMode`,
  `$node.someAttribute` raises `PropertyNotFoundException` when the attribute is
  absent — and absent is the normal case. `Get-OcrAttr` uses `GetAttribute`,
  which returns `''`.
- **`return @($x)` unrolls.** An empty result becomes `$null` and a
  single-element result becomes a bare object, so the caller's `.Count` throws.
  Every array-returning function returns `,@($x)`.
- **`[int]$x + 'm '` is integer addition.** PowerShell decides what `+` means
  from its left operand, and it threw trying to convert `'m '` to a number.
- **`Get-Command npm` returns every match.** On a machine with two Node installs
  `$npm.Source` is an array, and using it as a command produces
  `The term '/opt/node22/bin/npm /usr/local/bin/npm' is not recognized`.
- **`Join-Path $env:ProgramFiles ...` throws when the variable is unset.**
- **A mandatory `[string]` parameter rejects `$null`,** because the cast turns it
  into `''` before `AllowNull` is consulted.
- **A hard-coded list of test files goes stale.** Two suites were added and
  thirty-three tests quietly did not run while the runner still reported
  success. The runner discovers them now.

`tests\Test-Syntax.ps1` additionally enforces that every shipped file parses, is
**pure ASCII** (Windows PowerShell 5.1 reads a BOM-less file as the machine's
ANSI code page, so one smart quote in a comment can stop the whole script
parsing), and uses **no PowerShell 7-only syntax** — no ternary, `??`, `?.`,
`?[`, or `&&`/`||` chains, all of which parse fine in development and fail on
the Windows PowerShell that ships with Windows.

## Part two

This half ends with words that have positions, columns, meaning and an audit
trail. Getting that into Excel is the other half, and it starts from
`Get-OcrInvoiceSummary` — one flat row per page, the same columns every time.
