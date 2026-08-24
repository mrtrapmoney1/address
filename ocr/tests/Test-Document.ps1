<#
    Test-Document.ps1 - reading an OcrDocument back into PowerShell.

    These run against a document written here in the test rather than one
    produced by a real OCR run, so they test the READER and never fail
    because an engine read a word differently. There is a separate suite for
    the end-to-end path.
#>

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here

. (Join-Path $here 'OcrTest.ps1')
. (Join-Path $root 'OcrCommon.ps1')
. (Join-Path $root 'OcrDocument.ps1')

# A small statement: a header row and two data rows, three columns.
# Note that most Word elements carry NO "refined" attribute and the first
# line's words carry no "conf" - absent attributes are the normal case and
# the reader must not fall over on them.
$sampleXml = @'
<?xml version="1.0" encoding="utf-8"?>
<OcrDocument version="1" generated="2026-08-24T00:00:00Z" tool="test">
  <Source path="C:\scans\statement.png" kind="image" pages="1" />
  <Engine name="tesseract.js" lang="eng" oem="1" psm="3" dpi="300" />
  <Page number="1" width="612" height="792" unit="pt" pixelWidth="2550" pixelHeight="3300"
        dpi="300" skew="1.4" confidence="91.5" words="9" regions="1"
        preprocess="grayscale &gt; sauvola">
    <Region index="0" kind="table" x="24" y="24" w="460.8" h="48" columns="3">
      <Columns count="3">
        <Boundary index="0" x="76.2" />
        <Boundary index="1" x="290.4" />
      </Columns>
      <Line index="0" x="24" y="24" w="460.8" h="9.6" confidence="90">
        <Word x="24" y="24" w="28.8" h="9.6" size="9.6" col="0">Date</Word>
        <Word x="96" y="24" w="72" h="9.6" size="9.6" col="1">Description</Word>
        <Word x="432" y="24" w="52.8" h="9.6" size="9.6" col="2">Amount</Word>
      </Line>
      <Line index="1" x="24" y="43.2" w="460.8" h="9.6" confidence="93">
        <Word x="24" y="43.2" w="31.2" h="9.6" size="9.6" conf="95" col="0">01Aug</Word>
        <Word x="96" y="43.2" w="52.8" h="9.6" size="9.6" conf="92" col="1">Widgets</Word>
        <Word x="439.2" y="43.2" w="45.6" h="9.6" size="9.6" conf="88" col="2" refined="1">125.00</Word>
      </Line>
      <Line index="2" x="24" y="62.4" w="460.8" h="9.6" confidence="89">
        <Word x="24" y="62.4" w="31.2" h="9.6" size="9.6" conf="90" col="0">02Aug</Word>
        <Word x="96" y="62.4" w="33.6" h="9.6" size="9.6" conf="91" col="1">Bolts</Word>
        <Word x="424.8" y="62.4" w="60" h="9.6" size="9.6" conf="86" col="2">1,240.50</Word>
      </Line>
    </Region>
    <Text>Date Description Amount</Text>
  </Page>
</OcrDocument>
'@

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ocrtest-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $temp -Force
$samplePath = Join-Path $temp 'statement.ocr.xml'
Set-Content -LiteralPath $samplePath -Value $sampleXml -Encoding UTF8

try {
    Start-OcrSection 'ConvertTo-OcrNumber'

    Assert-OcrEqual (ConvertTo-OcrNumber '12.5') 12.5 'a plain decimal'
    Assert-OcrEqual (ConvertTo-OcrNumber '0') 0 'zero'
    Assert-OcrEqual (ConvertTo-OcrNumber '-3.25') -3.25 'a negative'
    Assert-OcrEqual (ConvertTo-OcrNumber $null) 0 'null becomes 0, not an error'
    Assert-OcrEqual (ConvertTo-OcrNumber '') 0 'empty becomes 0'
    Assert-OcrEqual (ConvertTo-OcrNumber '   ') 0 'whitespace becomes 0'
    Assert-OcrEqual (ConvertTo-OcrNumber 'not a number') 0 'nonsense becomes 0 rather than throwing'

    <#
        The one that bites in the field: on a machine set to a locale that
        uses a comma for the decimal point, a naive [double] cast reads the
        "12.5" in the file as 125 - every coordinate on the page silently
        multiplied by ten. The reader must parse as invariant.
    #>
    $previousCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture =
            New-Object System.Globalization.CultureInfo('de-DE')
        Assert-OcrEqual (ConvertTo-OcrNumber '12.5') 12.5 'a decimal point is read the same under a comma locale'
    } finally {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
    }

    Start-OcrSection 'Import-OcrDocument'

    $doc = Import-OcrDocument $samplePath

    Assert-OcrEqual $doc.Version '1' 'format version'
    Assert-OcrEqual $doc.Engine 'tesseract.js' 'engine name'
    Assert-OcrEqual $doc.Language 'eng' 'language'
    Assert-OcrEqual $doc.SourceKind 'image' 'source kind'
    Assert-OcrEqual $doc.Pages.Count 1 'one page'

    $page = $doc.Pages[0]
    Assert-OcrEqual $page.Number 1 'page number'
    Assert-OcrEqual $page.Width 612 'page width in points'
    Assert-OcrEqual $page.Height 792 'page height in points'
    Assert-OcrEqual $page.Dpi 300 'page dpi'
    Assert-OcrEqual $page.Skew 1.4 'recorded skew'
    Assert-OcrEqual $page.Confidence 91.5 'page confidence'
    Assert-OcrEqual $page.Preprocess 'grayscale > sauvola' 'preprocessing history survives escaping'
    Assert-OcrEqual $page.Regions.Count 1 'one region'
    Assert-OcrEqual $page.Text 'Date Description Amount' 'page text'

    $region = $page.Regions[0]
    Assert-OcrEqual $region.Kind 'table' 'the region is a table'
    Assert-OcrEqual $region.ColumnCount 3 'three columns'
    Assert-OcrEqual $region.Boundaries.Count 2 'two boundaries for three columns'
    Assert-OcrEqual $region.Boundaries[0] 76.2 'first boundary position'
    Assert-OcrEqual $region.Lines.Count 3 'three lines'

    Start-OcrSection 'Missing attributes are normal, not errors'

    # Under Set-StrictMode, $node.missingAttribute throws. Most words have no
    # "refined" and the header row has no "conf" - if the reader used dot
    # access this suite would not get this far.
    $header = $region.Lines[0].Words[0]
    Assert-OcrEqual $header.Text 'Date' 'a word with no conf attribute still reads'
    Assert-OcrEqual $header.Conf 0 'a missing conf becomes 0'
    Assert-OcrEqual $header.Refined $false 'a missing refined becomes false'

    $refined = $region.Lines[1].Words[2]
    Assert-OcrEqual $refined.Text '125.00' 'the refined word'
    Assert-OcrEqual $refined.Refined $true 'refined="1" is read as true'

    Start-OcrSection 'Numbers come back as numbers'

    $first = $region.Lines[0].Words[0]
    Assert-OcrTrue ($first.X -is [double]) 'X is a double, not a string'
    Assert-OcrTrue ($first.Column -is [int]) 'Column is an int'
    # The trap this guards: "9.6" + "9.6" is "9.69.6" if they are strings.
    Assert-OcrEqual ($first.X + $first.W) 52.8 'coordinates add up arithmetically'

    Start-OcrSection 'Get-OcrWords'

    $words = Get-OcrWords -Document $doc
    Assert-OcrEqual $words.Count 9 'nine words in the stream'

    # The shape contract: these are exactly the fields the PDF text extractor
    # produces, so downstream code works against either without changes.
    $sample = $words[0]
    foreach ($field in @('Page', 'Text', 'X', 'Y', 'W', 'H', 'Size')) {
        Assert-OcrTrue ($null -ne $sample.PSObject.Properties[$field]) ('the stream carries ' + $field)
    }
    Assert-OcrEqual $sample.Page 1 'Page is the page number'

    $confident = Get-OcrWords -Document $doc -MinConfidence 90
    Assert-OcrEqual $confident.Count 4 'filtering by confidence keeps only the four at 90 or better'

    $wrongPage = Get-OcrWords -Document $doc -Page 7
    Assert-OcrEqual $wrongPage.Count 0 'asking for a page that is not there gives nothing, not an error'

    Start-OcrSection 'Get-OcrTable'

    $rows = Get-OcrTable -Document $doc -Page 1
    Assert-OcrEqual $rows.Count 3 'three rows'
    Assert-OcrEqual $rows[0].Count 3 'three cells per row'
    Assert-OcrEqual $rows[0][0] 'Date' 'header cell 0'
    Assert-OcrEqual $rows[0][2] 'Amount' 'header cell 2'
    Assert-OcrEqual $rows[1][0] '01Aug' 'row 1 date'
    Assert-OcrEqual $rows[1][2] '125.00' 'row 1 amount lands in the amount column'
    Assert-OcrEqual $rows[2][2] '1,240.50' 'row 2 amount lands in the amount column'

    <#
        The whole point of the exercise, stated as an assertion: an amount is
        identifiable BY ITS COLUMN. Read as flat text, "01Aug Widgets 125.00"
        gives no way to know that 125.00 is an amount rather than a quantity
        or a reference number.
    #>
    Assert-OcrEqual $rows[1][1] 'Widgets' 'the description stays in its own column'

    Start-OcrSection 'Export-OcrWords'

    $csvPath = Join-Path $temp 'words.csv'
    $null = Export-OcrWords -Document $doc -Path $csvPath
    Assert-OcrTrue (Test-Path -LiteralPath $csvPath) 'the CSV is written'
    $csv = Import-Csv -LiteralPath $csvPath
    Assert-OcrEqual $csv.Count 9 'the CSV has one row per word'
    Assert-OcrEqual $csv[0].Text 'Date' 'the CSV carries the text'

    Start-OcrSection 'Bad input is refused clearly'

    Assert-OcrThrows { Import-OcrDocument (Join-Path $temp 'does-not-exist.xml') } 'a missing file throws'

    $notXml = Join-Path $temp 'bad.xml'
    Set-Content -LiteralPath $notXml -Value 'this is not XML at all <<<' -Encoding ASCII
    Assert-OcrThrows { Import-OcrDocument $notXml } 'unparseable XML throws'

    $wrongRoot = Join-Path $temp 'wrong.xml'
    Set-Content -LiteralPath $wrongRoot -Value '<?xml version="1.0"?><SomethingElse />' -Encoding ASCII
    Assert-OcrThrows { Import-OcrDocument $wrongRoot } 'XML that is not an OcrDocument throws'

    Start-OcrSection 'A page with no words'

    $emptyXml = '<?xml version="1.0" encoding="utf-8"?>' +
        '<OcrDocument version="1"><Source path="x" kind="image" /><Engine name="t" lang="eng" />' +
        '<Page number="1" width="612" height="792" words="0" regions="0" note="blank page" /></OcrDocument>'
    $emptyPath = Join-Path $temp 'empty.ocr.xml'
    Set-Content -LiteralPath $emptyPath -Value $emptyXml -Encoding UTF8

    $emptyDoc = Import-OcrDocument $emptyPath
    Assert-OcrEqual $emptyDoc.Pages.Count 1 'the blank page is still a page'
    Assert-OcrEqual $emptyDoc.Pages[0].Regions.Count 0 'with no regions'
    Assert-OcrEqual $emptyDoc.Pages[0].Note 'blank page' 'and it says why'
    Assert-OcrEqual (Get-OcrWords -Document $emptyDoc).Count 0 'the word stream is empty, not an error'

    Start-OcrSection 'One-element results stay arrays'

    <#
        The other half of the unrolling bug, and the sneakier one. An empty
        result becoming $null is at least loud. A ONE-element result becoming
        a bare object is silent: .Count returns nothing, foreach still works,
        and the failure surfaces much later somewhere else entirely.
    #>
    $oneXml = '<?xml version="1.0" encoding="utf-8"?>' +
        '<OcrDocument version="1"><Source path="x" kind="image" /><Engine name="t" lang="eng" />' +
        '<Page number="1" width="612" height="792" dpi="300" words="1" regions="1">' +
        '<Region index="0" kind="text" x="1" y="1" w="10" h="5" columns="1">' +
        '<Line index="0" x="1" y="1" w="10" h="5"><Word x="1" y="1" w="10" h="5" size="5" col="0">Only</Word>' +
        '</Line></Region></Page></OcrDocument>'
    $onePath = Join-Path $temp 'one.ocr.xml'
    Set-Content -LiteralPath $onePath -Value $oneXml -Encoding UTF8

    $oneDoc = Import-OcrDocument $onePath
    $oneWords = Get-OcrWords -Document $oneDoc
    Assert-OcrTrue ($oneWords -is [array]) 'a single word still comes back as an array'
    Assert-OcrEqual $oneWords.Count 1 'and its Count is 1'
    Assert-OcrEqual $oneWords[0].Text 'Only' 'and it is indexable'

    $oneRows = Get-OcrTable -Document $oneDoc -Page 1 -Region 0
    Assert-OcrTrue ($oneRows -is [array]) 'a single table row still comes back as an array'
    Assert-OcrEqual $oneRows.Count 1 'with Count 1'

    $noWords = Get-OcrWords -Document $oneDoc -MinConfidence 99
    Assert-OcrTrue ($noWords -is [array]) 'an empty filtered result is still an array'
    Assert-OcrEqual $noWords.Count 0 'with Count 0'

} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

Complete-OcrSuite 'Test-Document'
