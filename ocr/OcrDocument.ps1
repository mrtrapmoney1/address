<#
    OcrDocument.ps1 - reading the OCR result back into PowerShell.

    This is the seam between the two halves of the job. Node does the
    recognition and writes OcrDocument XML; this file turns that XML into
    objects, and the shape it produces is not arbitrary:

        Page  Text  X  Y  W  H  Size

    That is exactly what the PDF text extractor already hands its layout
    engine, in the same units - points from the top-left of the page. A
    scanned invoice therefore arrives downstream indistinguishable from a text
    one, and the second half of the project (getting it into Excel) has one
    input format to deal with instead of two.
#>

Set-StrictMode -Version 2.0

<#
    Read one attribute off an XML element.

    This exists because of a bug that only appears when the code is actually
    run: under Set-StrictMode, $node.someAttribute throws
    PropertyNotFoundException when the attribute is ABSENT rather than
    returning nothing. Attributes here are absent all the time by design - a
    word only carries "refined" when it was re-read, and only carries "col"
    when the page had columns - so dot access blows up on the common case.

    GetAttribute returns an empty string for a missing attribute, which is
    what the callers want.
#>
function Get-OcrAttr {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Node,

        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Name
    )

    if ($null -eq $Node) { return '' }
    # A comment or text node has no attributes at all.
    if (-not ($Node -is [System.Xml.XmlElement])) { return '' }
    return [string] $Node.GetAttribute($Name)
}

<#
    Read one .ocr.xml file.

    [xml] is used rather than a hand-rolled parser because the file is our own
    output and is known to be well-formed; what matters is that the numbers
    come back as numbers. XML attributes are always strings, and a page full
    of string "12.5" values that silently sort as text is a bug that surfaces
    a long way from here.
#>
function Import-OcrDocument {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [string] $Path
    )

    process {
        $full = Resolve-OcrPath $Path
        if (-not (Test-Path -LiteralPath $full)) {
            throw ('OCR document not found: ' + $Path)
        }

        $xml = $null
        try {
            $xml = [xml](Get-Content -LiteralPath $full -Raw -Encoding UTF8)
        } catch {
            throw ('Not readable as XML: ' + $Path + ' (' + $_.Exception.Message + ')')
        }

        $root = $xml.SelectSingleNode('/OcrDocument')
        if ($null -eq $root) {
            throw ('Not an OcrDocument: ' + $Path)
        }
        $pages = New-Object System.Collections.ArrayList

        foreach ($pageNode in @($root.SelectNodes('Page'))) {
            $regions = New-Object System.Collections.ArrayList

            foreach ($regionNode in @($pageNode.SelectNodes('Region'))) {
                $lines = New-Object System.Collections.ArrayList

                foreach ($lineNode in @($regionNode.SelectNodes('Line'))) {
                    $words = New-Object System.Collections.ArrayList

                    foreach ($wordNode in @($lineNode.SelectNodes('Word'))) {
                        $null = $words.Add([pscustomobject]@{
                            Text    = [string] $wordNode.InnerText
                            X       = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'x')
                            Y       = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'y')
                            W       = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'w')
                            H       = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'h')
                            Size    = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'size')
                            Conf    = ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'conf')
                            Column  = [int] (ConvertTo-OcrNumber (Get-OcrAttr $wordNode 'col'))
                            Refined = ((Get-OcrAttr $wordNode 'refined') -eq '1')
                        })
                    }

                    $null = $lines.Add([pscustomobject]@{
                        Index      = [int] (ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'index'))
                        X          = ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'x')
                        Y          = ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'y')
                        W          = ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'w')
                        H          = ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'h')
                        Confidence = ConvertTo-OcrNumber (Get-OcrAttr $lineNode 'confidence')
                        Words      = @($words)
                    })
                }

                $boundaries = New-Object System.Collections.ArrayList
                foreach ($boundaryNode in @($regionNode.SelectNodes('Columns/Boundary'))) {
                    $null = $boundaries.Add((ConvertTo-OcrNumber (Get-OcrAttr $boundaryNode 'x')))
                }

                $null = $regions.Add([pscustomobject]@{
                    Index       = [int] (ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'index'))
                    Kind        = Get-OcrAttr $regionNode 'kind'
                    X           = ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'x')
                    Y           = ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'y')
                    W           = ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'w')
                    H           = ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'h')
                    ColumnCount = [int] (ConvertTo-OcrNumber (Get-OcrAttr $regionNode 'columns'))
                    Boundaries  = @($boundaries)
                    Lines       = @($lines)
                })
            }

            $textNode = $pageNode.SelectSingleNode('Text')
            $text = ''
            if ($textNode) { $text = [string] $textNode.InnerText }

            # The invoice reading, when the run did one.
            $invoice = $null
            $invoiceNode = $pageNode.SelectSingleNode('Invoice')
            if ($invoiceNode) {
                $fields = @{}
                foreach ($fieldNode in @($invoiceNode.SelectNodes('Field'))) {
                    $name = Get-OcrAttr $fieldNode 'name'
                    if (-not $name) { continue }
                    $raw = Get-OcrAttr $fieldNode 'value'
                    $iso = Get-OcrAttr $fieldNode 'iso'

                    # A money field is a number; a date stays the ISO string,
                    # which sorts correctly and cannot be misread by locale.
                    $value = $raw
                    if ($iso) { $value = $iso }
                    elseif ($raw -match '^-?[0-9]+(\.[0-9]+)?$') { $value = ConvertTo-OcrNumber $raw }

                    $fields[$name] = [pscustomobject]@{
                        Name       = $name
                        Value      = $value
                        Text       = [string] $fieldNode.InnerText
                        Label      = Get-OcrAttr $fieldNode 'label'
                        FoundWhere = Get-OcrAttr $fieldNode 'from'
                        Confidence = ConvertTo-OcrNumber (Get-OcrAttr $fieldNode 'conf')
                        Ambiguous  = ((Get-OcrAttr $fieldNode 'ambiguous') -eq '1')
                    }
                }

                $repairs = New-Object System.Collections.ArrayList
                foreach ($repairNode in @($invoiceNode.SelectNodes('Repair'))) {
                    $null = $repairs.Add([pscustomobject]@{
                        From = Get-OcrAttr $repairNode 'from'
                        To   = Get-OcrAttr $repairNode 'to'
                    })
                }

                $notes = New-Object System.Collections.ArrayList
                foreach ($noteNode in @($invoiceNode.SelectNodes('Note'))) {
                    $null = $notes.Add([string] $noteNode.InnerText)
                }

                $arithmetic = $null
                $arithmeticNode = $invoiceNode.SelectSingleNode('Arithmetic')
                if ($arithmeticNode) {
                    $arithmetic = [pscustomobject]@{
                        Ok         = ((Get-OcrAttr $arithmeticNode 'ok') -eq '1')
                        Computed   = ConvertTo-OcrNumber (Get-OcrAttr $arithmeticNode 'computed')
                        Difference = ConvertTo-OcrNumber (Get-OcrAttr $arithmeticNode 'difference')
                        Note       = Get-OcrAttr $arithmeticNode 'note'
                    }
                }

                $invoice = [pscustomobject]@{
                    Fields     = $fields
                    Repairs    = @($repairs)
                    Notes      = @($notes)
                    Arithmetic = $arithmetic
                    Reconciles = Get-OcrAttr $invoiceNode 'reconciles'
                }
            }

            $null = $pages.Add([pscustomobject]@{
                Number     = [int] (ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'number'))
                Width      = ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'width')
                Height     = ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'height')
                Unit       = Get-OcrAttr $pageNode 'unit'
                Dpi        = ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'dpi')
                Skew       = ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'skew')
                Confidence = ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'confidence')
                WordCount  = [int] (ConvertTo-OcrNumber (Get-OcrAttr $pageNode 'words'))
                Preprocess = Get-OcrAttr $pageNode 'preprocess'
                Note       = Get-OcrAttr $pageNode 'note'
                Regions    = @($regions)
                Invoice    = $invoice
                Text       = $text
            })
        }

        $sourceNode = $root.SelectSingleNode('Source')
        $engineNode = $root.SelectSingleNode('Engine')

        return [pscustomobject]@{
            Path       = $full
            Version    = Get-OcrAttr $root 'version'
            Generated  = Get-OcrAttr $root 'generated'
            SourcePath = Get-OcrAttr $sourceNode 'path'
            SourceKind = Get-OcrAttr $sourceNode 'kind'
            Engine     = Get-OcrAttr $engineNode 'name'
            Language   = Get-OcrAttr $engineNode 'lang'
            Pages      = @($pages)
        }
    }
}

<#
    An XML attribute is always a string, and a missing one is $null. Turning
    that into a number in one place keeps the guard out of a dozen call sites.
    A value that is genuinely absent comes back as 0 rather than $null so that
    arithmetic on it never explodes under Set-StrictMode.
#>
function ConvertTo-OcrNumber {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) { return 0 }
    $text = [string] $Value
    if ($text.Trim() -eq '') { return 0 }

    $parsed = 0.0
    # InvariantCulture on purpose: the XML is written with a dot as the
    # decimal separator, and a machine set to a comma locale would otherwise
    # read 12.5 as 125.
    $ok = [double]::TryParse(
        $text,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref] $parsed)
    if ($ok) { return $parsed }
    return 0
}

<#
    The flat word stream.

    Deliberately the same property names the PDF extractor produces, so code
    written against one works unchanged against the other.
#>
function Get-OcrWords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        $Document,

        [Parameter(Mandatory = $false)]
        [int] $Page = 0,

        [Parameter(Mandatory = $false)]
        [double] $MinConfidence = 0
    )

    process {
        $out = New-Object System.Collections.ArrayList

        foreach ($p in $Document.Pages) {
            if ($Page -gt 0 -and $p.Number -ne $Page) { continue }
            foreach ($region in $p.Regions) {
                foreach ($line in $region.Lines) {
                    foreach ($word in $line.Words) {
                        if ($MinConfidence -gt 0 -and $word.Conf -lt $MinConfidence) { continue }
                        $null = $out.Add([pscustomobject]@{
                            Page       = $p.Number
                            Text       = $word.Text
                            X          = $word.X
                            Y          = $word.Y
                            W          = $word.W
                            H          = $word.H
                            Size       = $word.Size
                            Conf       = $word.Conf
                            Column     = $word.Column
                            Region     = $region.Index
                            RegionKind = $region.Kind
                            Line       = $line.Index
                        })
                    }
                }
            }
        }
        # The comma is load-bearing. "return @($out)" unrolls the array on the
        # way out: an empty result becomes $null and a single-word result
        # becomes a bare object, so the caller's .Count throws under
        # Set-StrictMode. Wrapping it keeps an array an array at every size.
        return ,@($out)
    }
}

<#
    The rows of a table region, as arrays of cells.

    This is the shape that goes straight into a worksheet, and it exists only
    because the words kept their positions - the same table read as flat text
    cannot say which column a number belonged to.
#>
function Get-OcrTable {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Document,

        [Parameter(Mandatory = $false)]
        [int] $Page = 1,

        [Parameter(Mandatory = $false)]
        [int] $Region = -1
    )

    $rows = New-Object System.Collections.ArrayList

    foreach ($p in $Document.Pages) {
        if ($p.Number -ne $Page) { continue }
        foreach ($r in $p.Regions) {
            if ($Region -ge 0 -and $r.Index -ne $Region) { continue }
            if ($Region -lt 0 -and $r.Kind -ne 'table') { continue }

            foreach ($line in $r.Lines) {
                $count = $r.ColumnCount
                if ($count -lt 1) { $count = 1 }

                # A fixed-width row of empty strings, then each word dropped
                # into its own column. A cell with nothing in it stays empty
                # rather than shifting everything after it left.
                $cells = New-Object 'string[]' $count
                for ($i = 0; $i -lt $count; $i++) { $cells[$i] = '' }

                foreach ($word in $line.Words) {
                    $column = $word.Column
                    if ($column -lt 0) { $column = 0 }
                    if ($column -ge $count) { $column = $count - 1 }
                    if ($cells[$column] -eq '') { $cells[$column] = $word.Text }
                    else { $cells[$column] = $cells[$column] + ' ' + $word.Text }
                }
                $null = $rows.Add($cells)
            }
        }
    }
    # See Get-OcrWords: the comma stops the array being unrolled.
    return ,@($rows)
}

<#
    Print the page the way it physically sat.

    This is the first thing to run when a value is missing: if the number is
    not visible here, no rule downstream can find it, and the problem is in
    the scan or the preparation rather than in the parsing.
#>
function Show-OcrPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Document,

        [Parameter(Mandatory = $false)]
        [int] $Page = 1,

        [Parameter(Mandatory = $false)]
        [int] $Width = 110
    )

    <#
        Draw the page the way it physically sat.

        The lines are NOT re-derived from the y coordinates here. The document
        already carries the grouping the segmenter worked out, word by word,
        by box overlap - and rounding y back into rows undoes that work badly:
        a line whose height does not divide evenly into the row spacing gets
        split across two rows, so "01Aug2018" ends up on one line of the
        picture and "Clearing Cheque" on the next. Reading the structure that
        is already in the file avoids the whole problem.

        This is the first thing to run when a value is missing. If the number
        is not visible here, no rule downstream can find it, and the fault is
        in the scan or the preparation rather than in the parsing.
    #>

    $target = $null
    foreach ($p in $Document.Pages) {
        if ($p.Number -eq $Page) { $target = $p; break }
    }
    if ($null -eq $target) {
        Write-Host ('  (no page ' + $Page + ' in this document)') -ForegroundColor DarkGray
        return
    }

    $lines = New-Object System.Collections.ArrayList
    foreach ($region in $target.Regions) {
        foreach ($line in $region.Lines) {
            if ($line.Words.Count -gt 0) { $null = $lines.Add($line) }
        }
    }

    if ($lines.Count -eq 0) {
        Write-Host '  (no words on this page)' -ForegroundColor DarkGray
        return
    }

    # Sorted top to bottom: regions come back in reading order, which for a
    # two-column page means down one column then down the next, and that is
    # not the order they sit on the paper.
    $ordered = @($lines | Sort-Object Y, X)

    $minX = ($ordered | Measure-Object -Property X -Minimum).Minimum
    $maxRight = 0
    foreach ($line in $ordered) {
        $right = $line.X + $line.W
        if ($right -gt $maxRight) { $maxRight = $right }
    }
    $spanX = $maxRight - $minX
    if ($spanX -le 0) { $spanX = 1 }
    $scale = ($Width - 1) / $spanX

    # A blank output row wherever the page left more than a line of space, so
    # the shape of the page survives.
    $gapThreshold = ($ordered | Measure-Object -Property H -Average).Average * 1.6
    if ($gapThreshold -le 0) { $gapThreshold = 12 }

    $previousBottom = $null
    foreach ($line in $ordered) {
        if ($null -ne $previousBottom -and ($line.Y - $previousBottom) -gt $gapThreshold) {
            Write-Host ''
        }

        $builder = New-Object System.Text.StringBuilder
        foreach ($word in ($line.Words | Sort-Object X)) {
            $column = [int][Math]::Round((($word.X - $minX) * $scale))

            # -le, not -lt: when the wanted column is exactly where the last
            # word ended, the two would be written with nothing between them
            # and read as one word.
            if ($column -le $builder.Length) { $column = $builder.Length + 1 }
            while ($builder.Length -lt $column) { $null = $builder.Append(' ') }
            $null = $builder.Append($word.Text)
        }
        Write-Host $builder.ToString()
        $previousBottom = $line.Y + $line.H
    }
}

<#
    Everything the run found, as one CSV. Useful on its own, and the thing to
    hand to a person who wants to eyeball the result in Excel before any
    parsing rules are written.
#>
function Export-OcrWords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Document,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $words = Get-OcrWords -Document $Document
    $full = Resolve-OcrPath $Path
    $words | Export-Csv -LiteralPath $full -NoTypeInformation -Encoding ASCII
    return $full
}

<#
    The invoice fields as one flat row - one object per page, one property per
    field. This is the shape that goes into a worksheet, and it is what the
    second half of the job (getting it into Excel) will consume.
#>
function Get-OcrInvoiceSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        $Document
    )

    process {
        $rows = New-Object System.Collections.ArrayList

        foreach ($page in $Document.Pages) {
            $row = [ordered]@{
                Source     = $Document.SourcePath
                Page       = $page.Number
                Confidence = $page.Confidence
                Words      = $page.WordCount
            }

            # Always the same columns, present or not, so a folder of invoices
            # produces a rectangular table rather than a ragged one.
            foreach ($name in @('invoiceNumber', 'invoiceDate', 'dueDate', 'poNumber',
                                'subtotal', 'freight', 'discount', 'tax', 'taxRate',
                                'total', 'terms', 'accountNumber')) {
                $value = $null
                if ($page.Invoice -and $page.Invoice.Fields.ContainsKey($name)) {
                    $value = $page.Invoice.Fields[$name].Value
                }
                $row[$name] = $value
            }

            $row['Reconciles'] = $(if ($page.Invoice) { $page.Invoice.Reconciles } else { 'unchecked' })
            $row['Repairs'] = $(if ($page.Invoice) { $page.Invoice.Repairs.Count } else { 0 })
            $row['Notes'] = $(if ($page.Invoice) { ($page.Invoice.Notes -join '; ') } else { '' })

            $null = $rows.Add([pscustomobject]$row)
        }

        return ,@($rows)
    }
}
