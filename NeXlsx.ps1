<#
================================================================================
 NeXlsx.ps1
 Fast .xlsx reader - no Excel, no external module.

 WHY
   Pulling ~300,000 x 25 cells out of a quarterly file through Excel COM costs
   minutes, because every cell crosses the COM boundary. An .xlsx IS a zip of
   XML (SpreadsheetML), so we read it directly: shared-string table + a single
   streaming pass over the sheet XML. The parser is compiled C# (Add-Type) so
   the per-cell work happens in .NET, not in a PowerShell loop.

   Typical result on a 300k-row quarterly file: seconds instead of minutes.

 USE
   . .\NeXlsx.ps1
   $data = Read-NeXlsxColumns -Path 'C:\...\2024 Q4 Address Data.xlsx' -Columns @(3,4,5,8,23,24,25)
   $data.Rows      # object[] of string[] - one entry per sheet row, in the order asked
   $data.Count     # number of rows returned (includes the header row)

 NOTES
   * Only the requested columns are materialised, so memory stays small.
   * Values come back as strings exactly as stored (numbers unformatted), which
     is what the matcher wants.
================================================================================
#>

if (-not ('NeXlsxReader' -as [type])) {
    $src = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

public static class NeXlsxReader
{
    // Excel column letters ("BA") -> 1-based index
    private static int ColFromRef(string cellRef)
    {
        int n = 0;
        for (int i = 0; i < cellRef.Length; i++)
        {
            char c = cellRef[i];
            if (c >= 'A' && c <= 'Z') n = n * 26 + (c - 'A' + 1);
            else if (c >= 'a' && c <= 'z') n = n * 26 + (c - 'a' + 1);
            else break;
        }
        return n;
    }

    private static List<string> ReadSharedStrings(ZipArchive zip)
    {
        var list = new List<string>();
        var entry = zip.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return list;
        using (var s = entry.Open())
        using (var r = XmlReader.Create(s, new XmlReaderSettings { IgnoreWhitespace = false }))
        {
            var sb = new StringBuilder();
            bool inSi = false;
            while (r.Read())
            {
                if (r.NodeType == XmlNodeType.Element)
                {
                    if (r.LocalName == "si") { inSi = true; sb.Length = 0; }
                    else if (r.LocalName == "t" && inSi)
                    {
                        // <t> can appear directly or inside <r> runs; append them all
                        if (!r.IsEmptyElement) sb.Append(r.ReadElementContentAsString());
                    }
                }
                else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "si")
                {
                    list.Add(sb.ToString()); inSi = false;
                }
            }
        }
        return list;
    }

    // Resolve the first sheet's XML part via workbook.xml + rels; fall back to sheet1.xml
    private static string FirstSheetPath(ZipArchive zip)
    {
        string relId = null;
        var wb = zip.GetEntry("xl/workbook.xml");
        if (wb != null)
        {
            using (var s = wb.Open())
            using (var r = XmlReader.Create(s))
            {
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.LocalName == "sheet")
                    {
                        relId = r.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        if (relId == null) relId = r.GetAttribute("r:id");
                        break;
                    }
                }
            }
        }
        if (relId != null)
        {
            var rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (rels != null)
            {
                using (var s = rels.Open())
                using (var r = XmlReader.Create(s))
                {
                    while (r.Read())
                    {
                        if (r.NodeType == XmlNodeType.Element && r.LocalName == "Relationship")
                        {
                            if (r.GetAttribute("Id") == relId)
                            {
                                string t = r.GetAttribute("Target");
                                if (!string.IsNullOrEmpty(t))
                                {
                                    if (t.StartsWith("/")) return t.TrimStart('/');
                                    if (t.StartsWith("xl/")) return t;
                                    return "xl/" + t;
                                }
                            }
                        }
                    }
                }
            }
        }
        return "xl/worksheets/sheet1.xml";
    }

    /// Returns one string[] per sheet row, holding ONLY the requested 1-based columns.
    public static List<string[]> ReadColumns(string path, int[] columns)
    {
        int maxCol = 0;
        var want = new Dictionary<int, int>();       // sheet column -> slot in the output
        for (int i = 0; i < columns.Length; i++)
        {
            want[columns[i]] = i;
            if (columns[i] > maxCol) maxCol = columns[i];
        }

        var rows = new List<string[]>();
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            var shared = ReadSharedStrings(zip);
            var sheetEntry = zip.GetEntry(FirstSheetPath(zip));
            if (sheetEntry == null) throw new Exception("could not find the first worksheet inside the .xlsx");

            using (var s = sheetEntry.Open())
            using (var r = XmlReader.Create(s))
            {
                string[] cur = null;
                int curCol = 0;
                string cellType = null;
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element)
                    {
                        if (r.LocalName == "row")
                        {
                            cur = new string[columns.Length];
                            if (r.IsEmptyElement) { rows.Add(cur); cur = null; }
                        }
                        else if (r.LocalName == "c" && cur != null)
                        {
                            string cr = r.GetAttribute("r");
                            curCol = (cr != null) ? ColFromRef(cr) : curCol + 1;
                            cellType = r.GetAttribute("t");
                            if (r.IsEmptyElement) { curCol = 0; }
                        }
                        else if ((r.LocalName == "v" || r.LocalName == "t") && cur != null && curCol > 0)
                        {
                            int slot;
                            if (want.TryGetValue(curCol, out slot))
                            {
                                string val = r.ReadElementContentAsString();
                                if (cellType == "s")
                                {
                                    int si;
                                    if (int.TryParse(val, out si) && si >= 0 && si < shared.Count) val = shared[si];
                                }
                                cur[slot] = val;
                            }
                        }
                    }
                    else if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "row" && cur != null)
                    {
                        rows.Add(cur); cur = null;
                    }
                }
            }
        }
        return rows;
    }
}
'@
    # Assembly references differ by edition: Windows PowerShell 5.1 (.NET
    # Framework) resolves simple names from the GAC; PowerShell 7 (.NET Core)
    # already supplies the BCL reference set, and naming it explicitly breaks on
    # type-forwarded types such as List<T>.
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Add-Type -TypeDefinition $src -ErrorAction Stop
    } else {
        foreach ($a in 'System.IO.Compression', 'System.IO.Compression.FileSystem', 'System.Xml') {
            try { Add-Type -AssemblyName $a -ErrorAction Stop } catch { }
        }
        Add-Type -TypeDefinition $src -ReferencedAssemblies @('System.IO.Compression', 'System.IO.Compression.FileSystem', 'System.Xml') -ErrorAction Stop
    }
}

# Read the requested 1-based columns from the first worksheet of an .xlsx.
# Returns @{ Rows = <List[string[]]>; Count = <n> }
function Read-NeXlsxColumns {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [int[]]$Columns
    )
    $rows = [NeXlsxReader]::ReadColumns((Resolve-Path -LiteralPath $Path).Path, $Columns)
    return @{ Rows = $rows; Count = $rows.Count }
}
