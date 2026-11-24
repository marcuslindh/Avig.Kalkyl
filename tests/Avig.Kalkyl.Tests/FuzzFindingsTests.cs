using Avig.Kalkyl.Formulas;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Every bug the fuzz harness (tools/Avig.Kalkyl.Fuzz) has found, as a test, so that it stays
/// fixed without the fuzzer having to find it again.
/// </summary>
public class FuzzFindingsTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string ContentTypes = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private const string Workbook1 = $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>
        """;

    private const string WorkbookRelationships = $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="{PackageRelationships}">
        <Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="{Relationships}/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private const string Styles = $"""
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="{Main}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></styleSheet>
        """;

    private static MemoryStream File(string sheetData, string? workbook = null, string? relationships = null, string? contentTypes = null, string? styles = null, string sheetPrefix = "")
    {
        return TestFiles.Package(
            ("[Content_Types].xml", contentTypes ?? ContentTypes),
            ("_rels/.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", workbook ?? Workbook1),
            ("xl/_rels/workbook.xml.rels", relationships ?? WorkbookRelationships),
            ("xl/styles.xml", styles ?? Styles),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{Main}">{sheetPrefix}<sheetData>{sheetData}</sheetData></worksheet>"""));
    }

    private static async Task<Workbook> OpenAsync(Stream file)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        return book;
    }

    private static async Task<Workbook> SaveAndReopenAsync(Workbook book)
    {
        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        return await OpenAsync(saved);
    }

    [Fact]
    public async Task A_broken_content_types_part_is_reported_when_the_file_is_opened_not_thrown_when_it_is_saved()
    {
        await using MemoryStream file = File("""<row r="1"><c r="A1"><v>1</v></c></row>""", contentTypes: "<Types><Default Extension=\"xml\" <");

        Result<Workbook> opened = await Workbook.OpenAsync(file);

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("[Content_Types].xml", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<row r=\"1\"><c r=\"A1\"><v>1<</v></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>x</is></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"A1\"><f>A2</c></row>")]
    public async Task Broken_xml_inside_a_cell_is_an_InvalidDataException_like_broken_xml_between_cells(string sheetData)
    {
        await using MemoryStream file = File(sheetData);

        await Assert.ThrowsAsync<InvalidDataException>(() => TestFiles.ReadRowsAsync(file));
    }

    [Theory]
    [InlineData("1e309")]
    [InlineData("-1E+400")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    public async Task A_number_Excel_cannot_hold_makes_the_cell_damaged(string stored)
    {
        await using MemoryStream file = File($"""<row r="1"><c r="A1"><v>{stored}</v></c></row>""");

        await Assert.ThrowsAsync<InvalidDataException>(() => TestFiles.ReadRowsAsync(file));
        file.Position = 0;
        Assert.True((await Workbook.OpenAsync(file)).IsFailure(out _, out _));
    }

    [Fact]
    public async Task A_row_after_the_last_row_of_a_sheet_makes_the_file_damaged()
    {
        await using MemoryStream file = File("""<row r="1048577"><c><f>1</f></c></row>""");

        await Assert.ThrowsAsync<InvalidDataException>(() => TestFiles.ReadRowsAsync(file));
        file.Position = 0;
        Assert.True((await Workbook.OpenAsync(file)).IsFailure(out _, out _));
    }

    [Theory]
    [InlineData("2147483647")]
    [InlineData("0")]
    [InlineData("-5")]
    public async Task A_font_size_Excel_cannot_show_is_ignored(string size)
    {
        string styles = $"""<styleSheet xmlns="{Main}"><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="{size}"/></font></fonts><cellXfs count="2"><xf fontId="0"/><xf fontId="1"/></cellXfs></styleSheet>""";
        await using MemoryStream file = File("""<row r="1"><c r="A1" s="1"><v>1</v></c></row>""", styles: styles);

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.True(rows[0][0].Style.Bold);
        Assert.Null(rows[0][0].Style.FontSize);
    }

    [Fact]
    public async Task An_unknown_element_before_sheetData_does_not_lose_the_edit()
    {
        await using MemoryStream file = File("""<row r="1"><c r="A1"><v>1</v></c></row>""", sheetPrefix: "<x14:unknown xmlns:x14=\"urn:other\"/><si/>");
        Workbook book = await OpenAsync(file);
        book.Sheets[0]["A1"] = "changed";

        Workbook again = await SaveAndReopenAsync(book);

        Assert.Equal((CellValue)"changed", again.Sheets[0]["A1"].Value);
    }

    [Fact]
    public async Task An_unknown_element_before_sheets_does_not_list_every_sheet_twice()
    {
        string workbook = $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><workbookPr_x/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""";
        await using MemoryStream file = File("""<row r="1"><c r="A1"><v>1</v></c></row>""", workbook);
        Workbook book = await OpenAsync(file);
        book.Sheets[0]["A1"] = 2;

        Workbook again = await SaveAndReopenAsync(book);

        Assert.Single(again.Sheets);
    }

    [Fact]
    public async Task A_relationship_without_a_target_is_saved_without_one()
    {
        string relationships = $"""
            <Relationships xmlns="{PackageRelationships}">
            <Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/>
            <Relationship Id="rId1" Type="{Relationships}/other"/>
            </Relationships>
            """;
        await using MemoryStream file = File("""<row r="1"><c r="A1"><v>1</v></c></row>""", relationships: relationships);
        Workbook book = await OpenAsync(file);
        book.Sheets[0]["A1"] = 2;

        Workbook again = await SaveAndReopenAsync(book);

        Assert.Single(again.Sheets);
        Assert.Equal((CellValue)2, again.Sheets[0]["A1"].Value);
    }

    [Theory]
    [InlineData("a\rb")]
    [InlineData("\r")]
    [InlineData("line\r\nbreak")]
    [InlineData("_xc1e5\uD800")]
    [InlineData("_x0041\u0001")]
    public async Task Text_reads_back_as_it_was_written(string text)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([text]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal((CellValue)text, rows[0][0].Value);
        Assert.Equal(text, CellText.Decode(CellText.Encode(text)));
    }

    [Fact]
    public async Task The_last_millisecond_Excel_can_hold_reads_back_as_a_date()
    {
        DateTime last = new(9999, 12, 31, 23, 59, 59, 999, DateTimeKind.Unspecified);
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([last]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.True(rows[0].TryGetDate(0, out DateTime read));
        Assert.Equal(last, read);
    }

    [Theory]
    [InlineData(2_958_465.9999999999, false)]
    [InlineData(2_958_466, false)]
    [InlineData(2_957_003.9999999999, true)]
    [InlineData(2_957_004, true)]
    [InlineData(2_958_000, true)]
    public void A_serial_past_the_end_of_9999_is_not_a_date_and_does_not_throw(double serial, bool uses1904)
    {
        Assert.False(ExcelDate.TryFromSerial(serial, uses1904, out _));
    }

    [Theory]
    [InlineData("=REPT(\"\",2147483647)", "")]
    [InlineData("=REPT(\"\",2147483648)", "#VALUE!")]
    [InlineData("=REPT(\"\",1E+15)", "#VALUE!")]
    public void REPT_of_empty_text_is_empty_up_to_Excels_largest_count(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    [Theory]
    [InlineData("=REPT(\"ab\",1E+15)")]
    [InlineData("=REPT(\"a\",2147483648)")]
    [InlineData("=REPT(\"ab\",16384)")]
    public void REPT_past_the_text_limit_is_VALUE(string formula)
    {
        Assert.Equal(CellValue.Error("#VALUE!"), Formula.Evaluate(formula));
    }

    [Theory]
    [InlineData("Other!#REF!")]
    [InlineData("'Other sheet'!#REF!+1")]
    [InlineData("SUM(Data!#ref!)")]
    public void A_sheet_reference_to_deleted_cells_parses_and_is_REF(string formula)
    {
        Assert.True(FormulaParser.Parse(formula).IsSuccess(out _));
        Assert.Equal(CellValue.Error("#REF!"), Formula.Evaluate(formula));
    }

    /// <summary>
    /// These used to overflow the stack, which ends the process: no catch runs, so one formula
    /// typed by a user or read from a file brought down the whole program.
    /// </summary>
    [Theory]
    [InlineData("parentheses", 256)]
    [InlineData("parentheses", 4000)]
    [InlineData("signs", 256)]
    [InlineData("signs", 8000)]
    [InlineData("functions", 65)]
    [InlineData("functions", 1500)]
    public void Nesting_past_the_limit_is_Invalid_and_does_not_end_the_process(string kind, int depth)
    {
        string formula = Nested(kind, depth, "1");

        Assert.True(FormulaParser.Parse(formula).IsFailure(out string? error, out _));
        Assert.Contains("levels", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("parentheses", 255)]
    [InlineData("signs", 255)]
    [InlineData("functions", 64)]
    public void Nesting_up_to_the_limit_is_calculated(string kind, int depth)
    {
        Assert.Equal(1, Math.Abs(Formula.Number(Nested(kind, depth, "1"))));
    }

    private static string Nested(string kind, int depth, string inner)
    {
        return kind switch
        {
            "parentheses" => new string('(', depth) + inner + new string(')', depth),
            "signs" => new string('-', depth) + inner,
            _ => string.Concat(Enumerable.Repeat("ABS(", depth)) + inner + new string(')', depth),
        };
    }

    [Theory]
    [InlineData("+", 4000, 4000)]
    [InlineData("-", 4000, -3998)]
    [InlineData("*", 4000, 1)]
    [InlineData("^", 4000, 1)]
    public void A_chain_of_thousands_of_operators_is_calculated(string op, int terms, double expected)
    {
        Assert.Equal(expected, Formula.Number(string.Join(op, Enumerable.Repeat("1", terms))));
    }

    [Fact]
    public async Task A_chain_of_thousands_of_operators_is_recalculated_in_a_cell()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 2;
        sheet["B1"] = Cell.FromFormula(string.Join("+", Enumerable.Repeat("A1", 2000)));

        await book.RecalculateAsync();

        Assert.Equal((CellValue)4000, sheet["B1"].Value);
    }

    [Fact]
    public void A_formula_moved_off_the_sheet_still_parses()
    {
        Assert.True(FormulaShifter.Shift("Other!A1+'Data'!A1:A3", -5, 0).IsSuccess(out string? moved));

        Assert.Equal("Other!#REF!+Data!#REF!", moved);
        Assert.True(FormulaParser.Parse(moved).IsSuccess(out _));
    }
}
