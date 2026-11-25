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

    [Fact]
    public async Task Rows_outside_sheetData_are_not_rows_of_the_sheet_in_either_reader()
    {
        // Found by the night run (xlsx-part): the document model keeps an unknown element as it is,
        // and the streaming reader read the rows inside it, so a workbook the package saved was damaged.
        const string unknown = """<x:extra xmlns:x="urn:example"><row r="1048576"><c r="A1048576"><v>9</v></c></row><row r="2"/></x:extra>""";
        await using MemoryStream file = File("""<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>""", sheetPrefix: unknown);

        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        Assert.Equal([0, 2], rows.Select(row => row.Index));

        file.Position = 0;
        Workbook book = await OpenAsync(file);
        Workbook again = await SaveAndReopenAsync(book);
        Assert.Equal((CellValue)3, again.Sheets[0][2, 0].Value);

        await using MemoryStream saved = new();
        await again.SaveAsync(saved);
        saved.Position = 0;
        Assert.Equal([0, 2], (await TestFiles.ReadRowsAsync(saved)).Select(row => row.Index));
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
    [InlineData(59.99999999999999, "1900-02-28T00:00:00")]
    [InlineData(59.9999999999, "1900-02-28T00:00:00")]
    [InlineData(59.999, "1900-02-28T23:58:33.6")]
    [InlineData(60.99999999999999, "1900-03-01T00:00:00")]
    public void A_serial_that_rounds_onto_a_side_of_the_made_up_day_is_read_on_that_side(double serial, string expected)
    {
        // Found by the night run (excel-date): 59.99999999999999 was read as before the made-up
        // 1900-02-29 and then rounded to 1900-03-01, which writes back as 61.
        Assert.True(ExcelDate.TryFromSerial(serial, false, out DateTime date));

        Assert.Equal(DateTime.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), date);
        Assert.True(ExcelDate.ToSerial(date) <= Math.Round(serial, 6), "A date never writes back as a later serial than it was read from.");
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
    public async Task A_formula_of_arrays_just_under_the_limit_stops_soon_after_it_is_cancelled()
    {
        // Found by the night run (formula, timeouts): each comparison makes a new array of 9 million
        // values, under KalkylLimits.MaxArrayCells, so a chain of them takes seconds. Time is the
        // caller's to limit, and the token has to stop it within one row of the array.
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        for (int row = 0; row < 8; row++)
        {
            sheet[row, 3] = row;
        }

        string formula = "rDF:P&6=2^64:811" + string.Concat(Enumerable.Repeat("<D1", 200));
        using CancellationTokenSource cancel = new(TimeSpan.FromMilliseconds(300));
        System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Task.Run(() => book.Evaluate(formula, sheet, cancel.Token)));

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(3), $"Cancelled after {watch.Elapsed.TotalSeconds:N1} s.");

        sheet["F1"] = Cell.FromFormula(formula);
        using CancellationTokenSource cancelRecalculation = new(TimeSpan.FromMilliseconds(300));
        watch.Restart();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => book.RecalculateAsync(RecalculationScope.All, cancelRecalculation.Token));

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(3), $"Recalculation cancelled after {watch.Elapsed.TotalSeconds:N1} s.");
    }

    [Fact]
    public void A_formula_moved_off_the_sheet_still_parses()
    {
        Assert.True(FormulaShifter.Shift("Other!A1+'Data'!A1:A3", -5, 0).IsSuccess(out string? moved));

        Assert.Equal("Other!#REF!+Data!#REF!", moved);
        Assert.True(FormulaParser.Parse(moved).IsSuccess(out _));
    }

    /// <summary>
    /// Found by xlsx-part and zip-raw: the reader finds xl/workbook.xml also when the package root
    /// has no relationship to it, but RemoveActiveContent only kept what the root led to, so the
    /// shared strings were left out and the saved file was damaged.
    /// </summary>
    [Theory]
    [InlineData("missing")]
    [InlineData("empty")]
    public async Task Without_root_relationships_RemoveActiveContent_keeps_what_the_workbook_uses(string root)
    {
        List<(string Path, string Content)> parts =
        [
            ("[Content_Types].xml", ContentTypes),
            ("xl/workbook.xml", Workbook1),
            ("xl/_rels/workbook.xml.rels", WorkbookRelationships.Replace("</Relationships>", $"<Relationship Id=\"rId3\" Type=\"{Relationships}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>", StringComparison.Ordinal)),
            ("xl/styles.xml", Styles),
            ("xl/sharedStrings.xml", $"""<sst xmlns="{Main}" count="1" uniqueCount="1"><si><t>kept</t></si></sst>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{Main}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>"""),
        ];

        if (root == "empty")
        {
            parts.Add(("_rels/.rels", $"""<Relationships xmlns="{PackageRelationships}"/>"""));
        }

        await using MemoryStream file = TestFiles.Package([.. parts]);
        Workbook book = await OpenAsync(file);
        Assert.True(book.RemoveActiveContent().IsSuccess(out _, out _));

        Workbook again = await SaveAndReopenAsync(book);

        Assert.Equal((CellValue)"kept", again.Sheets[0]["A1"].Value);
    }

    [Theory]
    [InlineData("A1:'Data'!B2", "A1:'Data'!B2")]
    [InlineData("=FALSE:'Data'!A1:A3", "FALSE:'Data'!A1:A3")]
    [InlineData("XFD1:'Other'!B:B", "XFD1:'Other'!B:B")]
    [InlineData("1+'Data'!A1:A3", "1+Data!A1:A3")]
    public void A_sheet_name_after_a_colon_stays_quoted_so_it_does_not_become_a_3d_reference(string formula, string expected)
    {
        // Found by the smoke run after the 3D fix (formula-shift): A1:'Data'!B2, the area from A1
        // to Data!B2, was written A1:Data!B2, which reads as the 3D reference 'A1:Data'!B2.
        Assert.True(FormulaShifter.Shift(formula, 0, 0).IsSuccess(out string? once));
        Assert.Equal(expected, once);

        Assert.True(FormulaShifter.Shift(once, 0, 0).IsSuccess(out string? twice));
        Assert.Equal(once, twice);
    }

    [Theory]
    [InlineData("Jan:Mar!A1", 2, 1, "Jan:Mar!B3")]
    [InlineData("A2:A47485!A1", 38, 49, "'A2:A47485'!AX39")]
    [InlineData("T&A:A2147:A3!AA1", 9, -10, "T&A:'A2147:A3'!Q10")]
    [InlineData("SUM('Jan 2026:Mar 2026'!B2:B5)", 1, 0, "SUM('Jan 2026:Mar 2026'!B3:B6)")]
    public void A_3d_reference_is_one_reference_and_moves_there_and_back(string formula, int rows, int columns, string expected)
    {
        // Found by the night run (formula-shift): the first sheet of the range was read as a name
        // or a cell, and written back so that it read as the other the next time.
        Assert.True(FormulaShifter.Shift(formula, rows, columns).IsSuccess(out string? moved));
        Assert.Equal(expected, moved);

        Assert.True(FormulaShifter.Shift(moved, -rows, -columns).IsSuccess(out string? back));
        Assert.True(FormulaShifter.Shift(formula, 0, 0).IsSuccess(out string? canonical));
        Assert.Equal(canonical, back);
    }
}
