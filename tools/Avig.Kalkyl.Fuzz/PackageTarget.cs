using System.IO.Compression;
using System.Text;

namespace Avig.Kalkyl.Fuzz;

/// <summary>
/// <c>xlsx-part</c> — a damaged or unusual file must give a Result failure or an
/// <see cref="InvalidDataException"/>, never anything else; and what the document model saves
/// must open again.
///
/// The input is not a whole .xlsx. Its first byte picks one part of a small, valid template
/// workbook and the rest of the input replaces that part. Mutating a zip file directly gets
/// nowhere: nearly every change breaks the deflate stream or the CRC, so the fuzzer would only
/// ever test the zip reader. This way the mutations land in the XML the package parses.
/// </summary>
internal static class PackageTarget
{
    private const string ContentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        </Types>
        """;

    private const string PackageRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private const string WorkbookXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <workbookPr date1904="0"/>
        <sheets>
        <sheet name="Data" sheetId="1" r:id="rId1"/>
        <sheet name="Calc 2" sheetId="2" state="hidden" r:id="rId2"/>
        </sheets>
        <definedNames>
        <definedName name="Rate">Data!$B$2</definedName>
        <definedName name="Local" localSheetId="1">'Calc 2'!$A$1</definedName>
        </definedNames>
        <calcPr calcId="191029"/>
        </workbook>
        """;

    private const string WorkbookRelationships = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>
        <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
        </Relationships>
        """;

    private const string Styles = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\-mm\-dd\ hh:mm"/></numFmts>
        <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><u/><sz val="14"/><color rgb="FFC00000"/><name val="Arial"/></font></fonts>
        <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>
        <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
        </cellXfs>
        </styleSheet>
        """;

    private const string SharedStrings = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
        <si><t>Name</t></si>
        <si><t xml:space="preserve"> Öst _x0001_ </t></si>
        <si><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> text</t></r></si>
        <si><t/></si>
        </sst>
        """;

    private const string Sheet1 = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetPr><tabColor rgb="FF00B050"/></sheetPr>
        <dimension ref="A1:F6"/>
        <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="15"/>
        <cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="4" width="12.5" customWidth="1"/></cols>
        <sheetData>
        <row r="1" spans="1:6"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>
        <row r="2"><c r="A2" t="inlineStr"><is><t>Inline</t></is></c><c r="B2"><v>0.25</v></c><c r="C2" s="2"><v>46023.5</v></c><c r="D2" s="3"><v>60</v></c><c r="E2" t="b"><v>1</v></c><c r="F2" t="e"><v>#N/A</v></c></row>
        <row r="3" ht="30" customHeight="1"><c r="A3" t="str"><f>A2&amp;"!"</f><v>Inline!</v></c><c r="B3"><f>B2*Rate+SUM(B2:B2)</f><v>0.3125</v></c><c r="C3"><f t="shared" ref="C3:C5" si="0">B3+1</f><v>1.3125</v></c></row>
        <row r="4"><c r="B4"><v>-1.5E-3</v></c><c r="C4"><f t="shared" si="0"/><v>0.9985</v></c></row>
        <row r="5"><c r="B5"><v>1E+300</v></c><c r="C5"><f t="shared" si="0"/><v>1E+300</v></c></row>
        <row r="6"><c r="A6"><f>IFERROR(VLOOKUP("x",A1:B5,2,FALSE),"none")</f></c><c r="B6"><f>'Calc 2'!A1+INDIRECT("B2")</f></c></row>
        </sheetData>
        <autoFilter ref="A1:F6"/>
        <mergeCells count="1"><mergeCell ref="D5:E6"/></mergeCells>
        <conditionalFormatting sqref="B2:B5"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        </worksheet>
        """;

    private const string Sheet2 = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
        <row r="1"><c r="A1"><f>SUM(Data!B:B)</f><v>1E+300</v></c><c r="B1"><f>Local*2</f></c><c r="C1"><f>B1</f></c></row>
        <row r="3"><c r="C3"><f>C3+1</f><v>0</v></c><c r="XFD3"><v>1</v></c></row>
        <row r="1048576"><c r="A1048576" t="inlineStr"><is><t>last</t></is></c></row>
        </sheetData>
        </worksheet>
        """;

    /// <summary>The template's parts. The first input byte picks the one to replace.</summary>
    internal static readonly (string Path, string Content)[] Parts =
    [
        ("xl/worksheets/sheet1.xml", Sheet1),
        ("xl/worksheets/sheet2.xml", Sheet2),
        ("xl/sharedStrings.xml", SharedStrings),
        ("xl/styles.xml", Styles),
        ("xl/workbook.xml", WorkbookXml),
        ("xl/_rels/workbook.xml.rels", WorkbookRelationships),
        ("[Content_Types].xml", ContentTypes),
        ("_rels/.rels", PackageRelationships),
    ];

    /// <summary>Things that are easy to get wrong in a reader of spreadsheet XML.</summary>
    private static readonly string[] Tokens =
    [
        "<", ">", "\"", "/>", "</c>", "</row>", "&amp;", "&#0;", "&#x1F;", "&#xD800;", "&lt;", "&bogus;",
        "<!DOCTYPE x [<!ENTITY e \"eeeeeeeeee\">]>", "&e;", "<![CDATA[x]]>", "<!-- c -->", "<?pi x?>",
        " r=\"A0\"", " r=\"XFD1048577\"", " r=\"XFE1\"", " r=\"ZZZZ1\"", " r=\"a1\"", " r=\"$A$1\"", " r=\"\"", " r=\"1\"",
        " s=\"9999\"", " s=\"-1\"", " s=\"\"", " t=\"s\"", " t=\"e\"", " t=\"d\"", " t=\"b\"", " t=\"str\"", " t=\"n\"", " t=\"x\"",
        "<v>1e309</v>", "<v>-0</v>", "<v>NaN</v>", "<v>99999999999</v>", "<v></v>", "<v>2</v>", "<v>-1</v>", "<v>#BOGUS!</v>",
        "<v>2958466</v>", "<v>2026-01-02T03:04:05</v>", "<v> 1 </v>",
        "<f t=\"shared\" si=\"0\"/>", "<f t=\"shared\" si=\"9\"/>", "<f t=\"shared\" ref=\"A1:XFD1048576\" si=\"0\">A1+1</f>",
        "<f t=\"shared\" ref=\"Z9:A1\" si=\"1\">ROW()</f>", "<f t=\"array\" ref=\"A1:B2\">{1,2;3,4}</f>", "<f>SUM(</f>", "<f/>",
        "<f>A1:XFD1048576</f>", "<f>INDIRECT(\"A\"&amp;ROW())</f>", "<f>OFFSET(A1,1048575,16383)</f>",
        "<c r=\"B2\"/>", "<c/>", "<c r=\"A2\" t=\"inlineStr\"><is><r><t>x</t></r><r/></is></c>",
        "<row r=\"0\">", "<row r=\"1048577\">", "<row r=\"2\">", "<row>", "<row r=\"x\">", "<row r=\"3\" hidden=\"1\">",
        "<si><t>", "</si>", "<si/>", "<r><t>x</t></r>", " xml:space=\"preserve\"", "_x0000_", "_xD800_", "_x005F_x0041_",
        " numFmtId=\"14\"", " numFmtId=\"999\"", " numFmtId=\"-1\"", " fontId=\"99\"", " fillId=\"99\"", " count=\"-1\"",
        " uniqueCount=\"999999999\"", "<numFmt numFmtId=\"165\" formatCode=\"[h]:mm;@\"/>", " formatCode=\"\"",
        " rgb=\"zz\"", " rgb=\"\"", " theme=\"1\"", " indexed=\"64\"", " val=\"-5\"", " val=\"1e999\"",
        "<sheet name=\"Data\" sheetId=\"3\" r:id=\"rId1\"/>", "<sheet name=\"\" sheetId=\"4\" r:id=\"rId9\"/>", " r:id=\"rId3\"",
        "<definedName name=\"Rate\">Rate</definedName>", "<definedName name=\"A1\">1</definedName>", " localSheetId=\"99\"",
        " date1904=\"1\"", " date1904=\"true\"", " Target=\"../../x.xml\"", " Target=\"\"", " Target=\"styles.xml\"",
        "<pane xSplit=\"16384\" ySplit=\"-1\" state=\"frozen\"/>", "<col min=\"0\" max=\"99999\" width=\"-1\"/>",
        "<mergeCell ref=\"A1:XFD1048576\"/>", "<mergeCell ref=\"bogus\"/>", " ref=\"\"",
        "￿", "\uD800", "\0", "Öst", " xmlns=\"\"", " xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\"",

        // Found by the security test plan rather than by this list, which is why they are here now:
        // text longer than a cell holds, and nesting past the reader's depth limit.
        "<is><t>" + new string('a', 32_768) + "</t></is>", "<v>" + new string('7', 40_000) + "</v>",
        string.Concat(Enumerable.Repeat("<r>", 150)) + "<t>x</t>" + string.Concat(Enumerable.Repeat("</r>", 150)),
        string.Concat(Enumerable.Repeat("<x>", 5_000)) + string.Concat(Enumerable.Repeat("</x>", 5_000)),
    ];

    /// <summary>Set by <c>--repro FILE --dump DIR</c>: the built and the saved package are written there to open and compare.</summary>
    public static string DumpDirectory { get; set; } = string.Empty;

    private static async Task DumpAsync(string name, MemoryStream file)
    {
        if (DumpDirectory.Length == 0)
        {
            return;
        }

        Directory.CreateDirectory(DumpDirectory);
        await File.WriteAllBytesAsync(Path.Combine(DumpDirectory, name), file.ToArray());
        file.Position = 0;
    }

    public static FuzzTarget Create()
    {
        return new FuzzTarget("xlsx-part", "one part of a valid workbook replaced; read, recalculate, save, reopen", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        int part = random.Next(Parts.Length);

        // Sheets are where nearly all the parsing happens, so they get half the inputs.
        if (random.Next(2) == 0)
        {
            part = random.Next(2);
        }

        string content = Parts[part].Content;

        // One in ten unchanged: the save invariant matters most on a file that is not damaged,
        // and these are the best starting points for libFuzzer's corpus.
        if (random.Next(10) > 0)
        {
            content = Inputs.Mutate(random, content, Tokens);
        }

        // Text right at the 32 767-character limit of a cell, in a cell where it will be read:
        // a token for it lands inside a cell too rarely to find a missing length check.
        if (random.Next(15) == 0)
        {
            string longText = new('a', random.Next(32_760, 32_800));
            content = content.Replace("<t>Inline</t>", "<t>" + longText + "</t>", StringComparison.Ordinal).Replace("<t>Name</t>", "<t>" + longText + "</t>", StringComparison.Ordinal);
        }

        byte[] text = Encoding.UTF8.GetBytes(content);
        byte[] input = new byte[text.Length + 1];
        input[0] = (byte)part;
        text.CopyTo(input, 1);

        return input;
    }

    internal static async Task RunAsync(byte[] input)
    {
        if (input.Length == 0)
        {
            return;
        }

        await using MemoryStream file = Build(input[0] % Parts.Length, input.AsMemory(1));
        await DumpAsync("input.xlsx", file);
        await ReadEverythingAsync(file);
        file.Position = 0;
        await EditAndSaveAsync(file);
    }

    /// <summary>The template workbook, unchanged, as zip bytes; the starting point of the zip-raw target.</summary>
    internal static byte[] Template()
    {
        using MemoryStream file = Build(-1, ReadOnlyMemory<byte>.Empty);

        return file.ToArray();
    }

    private static MemoryStream Build(int replaced, ReadOnlyMemory<byte> content)
    {
        MemoryStream stream = new();

        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, true))
        {
            for (int index = 0; index < Parts.Length; index++)
            {
                ZipArchiveEntry entry = archive.CreateEntry(Parts[index].Path, CompressionLevel.Fastest);
                using Stream entryStream = entry.Open();

                if (index == replaced)
                {
                    entryStream.Write(content.Span);
                }
                else
                {
                    entryStream.Write(Encoding.UTF8.GetBytes(Parts[index].Content));
                }
            }
        }

        stream.Position = 0;

        return stream;
    }

    /// <summary>Reads every cell of every sheet with the streaming reader.</summary>
    internal static async Task ReadEverythingAsync(Stream file)
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(file, true);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            return;
        }

        await using (reader)
        {
            for (int index = 0; index < reader.Sheets.Count; index++)
            {
                // A sheet that points to a part that is not there is listed and then refused with
                // Invalid; that is a Result failure, which is the contract.
                if (reader.Sheet(index).IsSuccess(out SheetReader? sheet) == false)
                {
                    continue;
                }

                await foreach (Row row in sheet.ReadRowsAsync())
                {
                    Touch(row);
                }
            }
        }
    }

    /// <summary>
    /// Opens the file as a document, recalculates, changes a cell and saves. The saved file must
    /// open again and read through without being called damaged: the package wrote it.
    /// </summary>
    internal static async Task EditAndSaveAsync(Stream file)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(file);

        if (opened.IsSuccess(out Workbook? book) == false || book.Sheets.Count == 0)
        {
            return;
        }

        await book.RecalculateAsync(RecalculationScope.All);
        book.Sheets[0][0, 0] = Cell.Of("changed").Bold();
        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        await DumpAsync("saved.xlsx", saved);
        saved.Position = 0;

        try
        {
            Result<Workbook> reopened = await Workbook.OpenAsync(saved);

            if (reopened.IsSuccess(out Workbook? again) == false)
            {
                throw new InvariantViolationException($"A saved workbook cannot be opened again: {reopened.Error}");
            }

            if (again.Sheets.Count != book.Sheets.Count)
            {
                throw new InvariantViolationException($"The workbook had {book.Sheets.Count} sheets and {again.Sheets.Count} after saving.");
            }

            if (again.Sheets[0][0, 0].Value != (CellValue)"changed")
            {
                throw new InvariantViolationException($"The changed cell reads back as {again.Sheets[0][0, 0].Value}.");
            }

            saved.Position = 0;
            await ReadEverythingAsync(saved);
        }
        catch (InvalidDataException exception)
        {
            throw new InvariantViolationException("A workbook the package saved is called damaged by the package.", exception);
        }
    }

    private static void Touch(Row row)
    {
        foreach (Cell cell in row.Cells)
        {
            _ = cell.Value.ToString();
            _ = cell.Style.GetHashCode();
            _ = cell.Formula?.Length;
        }
    }
}
