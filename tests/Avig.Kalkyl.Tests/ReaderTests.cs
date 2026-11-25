namespace Avig.Kalkyl.Tests;

/// <summary>
/// Files built by hand the way Excel and other programs write them: shared
/// strings, cells without addresses, shared formulas, custom date formats and Strict OOXML.
/// </summary>
public class ReaderTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static MemoryStream ExcelLikeFile()
    {
        return TestFiles.Package(
            ("_rels/.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationships}">
                  <Relationship Id="rId1" Type="{Relationships}/officeDocument" Target="/xl/workbook.xml"/>
                </Relationships>
                """),
            ("xl/workbook.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="{Main}" xmlns:r="{Relationships}">
                  <workbookPr defaultThemeVersion="166925"/>
                  <sheets>
                    <sheet name="Data" sheetId="1" r:id="rId1"/>
                    <sheet name="Diagram" sheetId="3" r:id="rId4"/>
                    <sheet name="Gömd" sheetId="2" state="hidden" r:id="rId2"/>
                  </sheets>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationships}">
                  <Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="{Relationships}/worksheet" Target="/xl/worksheets/sheet2.xml"/>
                  <Relationship Id="rId3" Type="{Relationships}/sharedStrings" Target="sharedStrings.xml"/>
                  <Relationship Id="rId4" Type="{Relationships}/chartsheet" Target="chartsheets/sheet1.xml"/>
                  <Relationship Id="rId5" Type="{Relationships}/styles" Target="styles.xml"/>
                </Relationships>
                """),
            ("xl/sharedStrings.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="{Main}" count="3" uniqueCount="3">
                  <si><t>Namn</t></si>
                  <si><r><rPr><b/></rPr><t xml:space="preserve">Hej </t></r><r><t>världen</t></r><rPh sb="0" eb="1"><t>XX</t></rPh></si>
                  <si><t>_x0001_</t></si>
                </sst>
                """),
            ("xl/styles.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="{Main}">
                  <numFmts count="2">
                    <numFmt numFmtId="164" formatCode="yyyy\-mm\-dd\ hh:mm"/>
                    <numFmt numFmtId="165" formatCode="0.00\ &quot;kr&quot;"/>
                  </numFmts>
                  <cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs>
                  <cellXfs count="5">
                    <xf numFmtId="0"/>
                    <xf numFmtId="14" applyNumberFormat="1"/>
                    <xf numFmtId="164" applyNumberFormat="1"/>
                    <xf numFmtId="165" applyNumberFormat="1"/>
                    <xf numFmtId="46" applyNumberFormat="1"/>
                  </cellXfs>
                </styleSheet>
                """),
            ("xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="{Main}" xmlns:r="{Relationships}">
                  <dimension ref="A1:F6"/>
                  <sheetData>
                    <row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>
                    <row r="3"><c r="A3" s="1"><v>45000</v></c><c s="2"><v>45000.5</v></c><c s="3"><v>12.5</v></c><c s="4"><v>1.5</v></c></row>
                    <row><c t="b"><v>1</v></c><c t="e"><v>#DIV/0!</v></c><c t="str"><f>A1&amp;"x"</f><v>Namnx</v></c><c><f t="shared" ref="D4:D5" si="0">$A$1&amp;A3</f><v>2</v></c></row>
                    <row r="5"><c r="D5"><f t="shared" si="0"/><v>2</v></c><c r="E5" t="inlineStr"><is><t>inline</t></is></c><c r="F5" t="s"><v>2</v></c></row>
                    <row r="6" ht="30" customHeight="1"/>
                  </sheetData>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                </worksheet>
                """),
            ("xl/worksheets/sheet2.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="{Main}"><sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData></worksheet>
                """));
    }

    [Fact]
    public async Task Worksheets_are_listed_in_tab_order_without_chart_sheets()
    {
        await using MemoryStream file = ExcelLikeFile();
        await using XlsxReader reader = await TestFiles.OpenAsync(file);

        Assert.Equal(2, reader.Sheets.Count);
        Assert.Equal(new SheetInfo("Data", 0, SheetVisibility.Visible), reader.Sheets[0]);
        Assert.Equal(new SheetInfo("Gömd", 1, SheetVisibility.Hidden), reader.Sheets[1]);
        Assert.False(reader.Uses1904Dates);
    }

    [Fact]
    public async Task Cells_are_read_the_way_Excel_wrote_them()
    {
        await using MemoryStream file = ExcelLikeFile();
        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal([0, 2, 3, 4, 5], rows.Select(row => row.Index));

        Assert.Equal((CellValue)"Namn", rows[0][0].Value);
        Assert.Equal((CellValue)"Hej världen", rows[0][2].Value);

        Assert.Equal((CellValue)new DateTime(2023, 3, 15), rows[1][0].Value);
        Assert.Equal((CellValue)new DateTime(2023, 3, 15, 12, 0, 0), rows[1][1].Value);
        Assert.Equal((CellValue)12.5, rows[1][2].Value);
        Assert.Equal("0.00\\ \"kr\"", rows[1][2].Style.NumberFormat);
        Assert.Equal((CellValue)1.5, rows[1][3].Value);

        Assert.True(rows[2].TryGetBool(0, out bool flag));
        Assert.True(flag);
        Assert.True(rows[2][1].TryGetError(out string? error));
        Assert.Equal("#DIV/0!", error);
        Assert.Equal((CellValue)"Namnx", rows[2][2].Value);
        Assert.Equal("A1&\"x\"", rows[2][2].Formula);
        Assert.Equal("$A$1&A3", rows[2][3].Formula);

        Assert.Equal("$A$1&A4", rows[3][3].Formula);
        Assert.Equal((CellValue)2, rows[3][3].Value);
        Assert.Equal((CellValue)"inline", rows[3][4].Value);
        Assert.Equal((CellValue)"\u0001", rows[3][5].Value);

        Assert.True(rows[4].IsEmpty);
    }

    [Fact]
    public async Task Sheets_are_found_by_name_regardless_of_case()
    {
        await using MemoryStream file = ExcelLikeFile();
        await using XlsxReader reader = await TestFiles.OpenAsync(file);

        Assert.True(reader.Sheet("gömd").IsSuccess(out SheetReader? hidden));
        Assert.Equal("Gömd", hidden.Info.Name);

        Result<SheetReader> missing = reader.Sheet("Saknas");
        Assert.True(missing.IsFailure(out string? message, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
        Assert.Contains("'Data', 'Gömd'", message, StringComparison.Ordinal);

        Assert.True(reader.Sheet(2).IsFailure(out _, out ResultStatus beyond));
        Assert.Equal(ResultStatus.NotFound, beyond);
    }

    [Fact]
    public async Task A_sheet_can_be_read_more_than_once()
    {
        await using MemoryStream file = ExcelLikeFile();
        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Sheet(0).IsSuccess(out SheetReader? sheet));
        int first = 0;
        int second = 0;

        await foreach (Row row in sheet.ReadRowsAsync())
        {
            first++;
        }

        await foreach (Row row in sheet.ReadRowsAsync())
        {
            second++;
        }

        Assert.Equal(5, first);
        Assert.Equal(first, second);
    }

    [Fact]
    public async Task Strict_OOXML_and_the_1904_date_system_are_read()
    {
        const string strictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";

        await using MemoryStream file = TestFiles.Package(
            ("_rels/.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", $"""<workbook xmlns="{strictMain}" xmlns:r="{strictRelationships}"><workbookPr date1904="1"/><sheets><sheet name="Strikt" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{strictRelationships}/styles" Target="styles.xml"/></Relationships>"""),
            ("xl/styles.xml", $"""<styleSheet xmlns="{strictMain}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{strictMain}"><sheetData><row r="1"><c r="A1" s="1"><v>0</v></c><c r="B1" t="d"><v>2026-11-23T10:15:00</v></c></row></sheetData></worksheet>"""));

        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Uses1904Dates);
        Assert.Equal("Strikt", reader.Sheets[0].Name);

        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        Assert.Equal((CellValue)new DateTime(1904, 1, 1), rows[0][0].Value);
        Assert.Equal((CellValue)new DateTime(2026, 11, 23, 10, 15, 0), rows[0][1].Value);
    }

    [Fact]
    public async Task A_missing_file_is_NotFound()
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx"));

        Assert.True(opened.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
    }

    [Theory]
    [InlineData(new byte[] { 0x48, 0x65, 0x6A, 0x21, 0x0A }, "not an .xlsx")]
    [InlineData(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0 }, ".xls")]
    [InlineData(new byte[] { 0x50, 0x4B, 0x03, 0x04, 1, 2, 3 }, "too short to be a zip file")]
    public async Task Content_that_is_not_an_xlsx_is_Invalid_with_a_reason(byte[] content, string reason)
    {
        await using MemoryStream file = new(content);
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(file);

        Assert.True(opened.IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Contains(reason, error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_zip_without_a_workbook_is_Invalid()
    {
        await using MemoryStream file = TestFiles.Package(("word/document.xml", "<document/>"));
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(file);

        Assert.True(opened.IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Contains("no workbook part", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_stream_that_cannot_seek_is_read_too()
    {
        await using MemoryStream file = ExcelLikeFile();
        await using ForwardOnlyStream forwardOnly = new(file.ToArray());
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(forwardOnly);

        Assert.True(opened.IsSuccess(out XlsxReader? reader), opened.Error);
        await reader.DisposeAsync();
    }

    [Fact]
    public async Task Damage_inside_a_sheet_names_the_cell()
    {
        await using MemoryStream file = TestFiles.Package(
            ("xl/workbook.xml", $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{Main}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2" t="s"><v>9</v></c></row></sheetData></worksheet>"""));

        InvalidDataException exception = await Assert.ThrowsAsync<InvalidDataException>(() => TestFiles.ReadRowsAsync(file));

        Assert.Contains("cell B2 (row 1, column 1)", exception.Message, StringComparison.Ordinal);
    }

    private sealed class ForwardOnlyStream(byte[] content) : MemoryStream(content)
    {
        public override bool CanSeek => false;
    }
}
