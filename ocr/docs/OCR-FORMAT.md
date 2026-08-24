# The OcrDocument format

The file the OCR run produces, and the contract the rest of the work is built on.

## Why XML and not text

Plain text throws away the one thing OCR knows that a human reader takes for
granted: **where on the page each word sat**. Once a line of an invoice is a
string, nothing downstream can tell which column a number came from. Every word
here keeps its box.

## The coordinate system

This is the most important decision in the format, and it is not aesthetic.

> **Coordinates are POINTS, measured from the TOP-LEFT of the page.**
> 72 points = 1 inch.

That is exactly the frame and the units the PDF text extractor uses for a text
PDF. So a scanned invoice and a text invoice arrive downstream in the same
shape, and the second half of the job — getting the data into Excel — has one
input format to write against rather than two that will drift apart.

Conversion from the pixels the engine works in:

```
points = pixels * 72 / dpi
```

When a page was enlarged, straightened or cropped before recognition, every box
is mapped **back through those transforms** to the page as it was supplied.
Without that, deskewing would silently move every reported word by millimetres —
worse than not deskewing at all, because the error would be invisible.

## Structure

```
OcrDocument
├── Source        where it came from
├── Engine        what read it, in what language, at what settings
└── Page          one per page
    ├── Region    a block of related text, in reading order
    │   ├── Columns → Boundary    where the column edges sit
    │   └── Line
    │       └── Word              text, box, confidence, column
    ├── Invoice   the invoice reading, when one was done
    │   ├── Field         invoice number, dates, totals
    │   ├── Arithmetic    whether the amounts reconcile
    │   ├── LineItems     which region holds them, and what each column is
    │   ├── Repair        every misread digit that was corrected
    │   └── Note
    └── Text      the engine's own flat text, for reference
```

## Elements

### `<Page>`

| Attribute | Meaning |
|---|---|
| `number` | 1-based page number |
| `width` `height` | page size in points |
| `unit` | always `pt` |
| `pixelWidth` `pixelHeight` | the image as recognised, after preparation |
| `dpi` | the resolution recognition ran at |
| `skew` | tilt found, in degrees; positive is clockwise |
| `confidence` | mean word confidence, 0–100 |
| `words` `regions` | counts |
| `preprocess` | exactly what was done, in order |
| `note` | anything the run wants the reader to know |

### `<Region>`

A block of related text, found by recursive XY-cut. `kind` is `table` or `text`.
`columns` is how many columns its words were assigned to.

A `table` means the rows put words into several distinct columns — not merely
that the block contains numbers.

### `<Columns>` / `<Boundary>`

Where the column edges sit, in points. These are written out because without
them a consumer can see which column a word is in but not where the column edge
was — and so cannot tell **an empty cell from a missing one**. For a
spreadsheet that difference is the whole point.

### `<Word>`

| Attribute | Meaning |
|---|---|
| `x` `y` | top-left of the word, in points |
| `w` `h` | its size, in points |
| `size` | font size in points (the box height, near enough) |
| `conf` | 0–100 from the engine; absent when the engine reports none |
| `col` | which column of its region it fell in |
| `refined` | `1` if it came from the low-confidence second pass |

The text is the element's content, so it needs no attribute escaping and can
contain anything a scan produces.

### `<Invoice>`

Present when the invoice pass ran. `reconciles` is `yes`, `no` or `unchecked`.

`<Field>` carries `name`, `value` (a number for money, ISO `yyyy-mm-dd` for a
date), the `label` it was found by, whether it was found to the `right` of that
label or `below` it, and its confidence. `ambiguous="1"` marks a date whose
day/month order could not be determined.

`<Repair>` records every misread digit corrected — `from` is exactly what the
engine read. **A number that was quietly changed on its way to a spreadsheet
and cannot be traced back is worse than no number at all**; this is what makes
each one auditable.

## Reading it

**PowerShell:**

```powershell
. .\OcrDocument.ps1
$doc = Import-OcrDocument .\page.ocr.xml

Get-OcrWords $doc            # Page/Text/X/Y/W/H/Size/Conf/Column/Region
Get-OcrTable $doc            # rows of cells
Get-OcrInvoiceSummary $doc   # one flat row per page
Show-OcrPage $doc            # the page drawn back as text
```

Note `Get-OcrAttr` rather than dot access throughout the reader: under
`Set-StrictMode`, `$node.someAttribute` **throws** when the attribute is absent,
and absent attributes are the normal case here.

**JavaScript:**

```javascript
const emit = require('./lib/emit.js');
const doc = emit.readDocument(fs.readFileSync('page.ocr.xml', 'utf8'));
const words = emit.toWordStream(doc);   // Page/Text/X/Y/W/H/Size
```

## The words-in format

`ocr-segment.js` takes a word list and produces this XML, so that the Windows
OCR engine and Tesseract both go through **one** implementation of the layout
rules. A second implementation in PowerShell would drift within a month, and
then a page would segment differently depending on which engine read it.

```json
{
  "source": { "path": "C:\\scans\\inv.png", "kind": "image" },
  "engine": { "name": "Windows.Media.Ocr", "lang": "en-GB", "dpi": 300 },
  "pages": [{
    "number": 1, "pixelWidth": 2550, "pixelHeight": 3300, "dpi": 300,
    "widthPt": 612, "heightPt": 792, "skew": 0.0, "text": "...",
    "words": [{ "text": "TOTAL", "x": 900, "y": 600, "w": 150, "h": 30, "conf": 92 }]
  }]
}
```

`x`/`y`/`w`/`h` are **pixels** of the supplied image; the conversion to points
happens here, once.

## Compatibility

`version` is `1`. New attributes and elements may be added within version 1 —
a reader must ignore what it does not recognise. Anything that changes the
meaning of an existing attribute gets a new version number.
