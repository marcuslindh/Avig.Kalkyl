using System.Diagnostics;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Bugs found by the compatibility tests (tests/Avig.Kalkyl.Compatibility), which read real
/// tutorial workbooks with Avig.Kalkyl and four other libraries. Kept here so that they stay
/// fixed without the downloaded files.
/// </summary>
public class CompatibilityFindingsTests
{
    [Fact]
    public async Task A_formulas_text_result_keeps_its_spaces_when_saved()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = "age ";
        sheet["B1"] = Cell.FromFormula("PROPER(A1)");
        await book.RecalculateAsync();
        await using MemoryStream file = new();

        await book.SaveAsync(file);

        Assert.Contains("<v xml:space=\"preserve\">Age </v>", TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_line_break_is_written_as_it_is_on_every_operating_system()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["first\nsecond"]);
        });

        string xml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");

        Assert.Contains("first\nsecond", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("first\r\nsecond", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_time_without_a_date_gets_a_time_format()
    {
        DateTime time = new(1899, 12, 31, 8, 30, 0, DateTimeKind.Unspecified);
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([time]);
        });

        Assert.Contains("formatCode=\"hh:mm:ss\"", TestFiles.ReadPart(file, "xl/styles.xml"), StringComparison.Ordinal);
        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        Assert.True(rows[0].TryGetDate(0, out DateTime read));
        Assert.Equal(time, read);
    }

    /// <summary>
    /// The used area of a sheet was found by walking every row on every range read, so 200 000
    /// SUM(G2:I2) on 30 000 rows took minutes. Generous limit: it is a few hundred milliseconds now.
    /// </summary>
    [Fact]
    public async Task Recalculating_a_long_sheet_does_not_walk_the_sheet_for_every_range()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        for (int row = 0; row < 20_000; row++)
        {
            sheet[row, 0] = row;
            sheet[row, 1] = row * 2;
            sheet[row, 2] = Cell.FromFormula($"SUM(A{row + 1}:B{row + 1})");
        }

        Stopwatch watch = Stopwatch.StartNew();
        await book.RecalculateAsync(RecalculationScope.All);

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(10), $"Recalculation took {watch.Elapsed}.");
        Assert.Equal((CellValue)(19_999 * 3), sheet[19_999, 2].Value);
    }

    [Fact]
    public void Clearing_the_last_column_shrinks_the_column_count()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 1;
        sheet["D3"] = 2;
        Assert.Equal(4, sheet.ColumnCount);
        Assert.Equal(3, sheet.RowCount);

        sheet.Clear(2, 3);

        Assert.Equal(1, sheet.ColumnCount);
    }

    [Theory]
    [InlineData("=SUMPRODUCT(1/COUNTIF(A1:A6,A1:A6))", 3)]
    [InlineData("=SUM(COUNTIF(A1:A6,{\"a\",\"b\"}))", 5)]
    [InlineData("=INDEX(COUNTIF(A1:A6,A1:A6),3)", 1)]
    public void COUNTIF_with_a_range_of_criteria_counts_each_one(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, ("A1", "a"), ("A2", "b"), ("A3", "c"), ("A4", "a"), ("A5", "b"), ("A6", "a")), 10);
    }

    [Fact]
    public void CHOOSE_with_an_array_of_indexes_lets_VLOOKUP_look_to_the_left()
    {
        CellValue value = Formula.Evaluate(
            "=VLOOKUP(\"Bo\",CHOOSE({1,2},B1:B3,A1:A3),2,FALSE)",
            ("A1", 101), ("B1", "Ada"), ("A2", 102), ("B2", "Bo"), ("A3", 103), ("B3", "Cy"));

        Assert.Equal((CellValue)102, value);
    }

    [Fact]
    public void CHOOSE_with_an_index_out_of_range_in_the_array_is_VALUE()
    {
        Assert.Equal(CellValue.Error("#VALUE!"), Formula.Evaluate("=INDEX(CHOOSE({1,3},A1:A2,B1:B2),1,1)", ("A1", 1)));
    }

    private static readonly (string Address, CellValue Value)[] Database =
    [
        ("A1", "Name"), ("B1", "Region"), ("C1", "Sales"),
        ("A2", "Ada"), ("B2", "East"), ("C2", 100),
        ("A3", "Bo"), ("B3", "West"), ("C3", 250),
        ("A4", "Cy"), ("B4", "East"), ("C4", 300),
        ("A5", "Di"), ("B5", "North"), ("C5", "n/a"),
        ("E1", "Region"), ("F1", "Sales"),
        ("E2", "East"), ("F2", ">150"),
        ("E3", "West"),
        ("H1", "region"),
        ("H2", "East"),
        ("J1", "Region"),
    ];

    [Theory]
    [InlineData("=DSUM(A1:C5,\"Sales\",H1:H2)", 400)]
    [InlineData("=DSUM(A1:C5,3,E1:F2)", 300)]
    [InlineData("=DSUM(A1:C5,\"Sales\",E1:F3)", 550)]
    [InlineData("=DAVERAGE(A1:C5,\"Sales\",H1:H2)", 200)]
    [InlineData("=DCOUNT(A1:C5,\"Sales\",J1:J2)", 3)]
    [InlineData("=DCOUNTA(A1:C5,\"Sales\",J1:J2)", 4)]
    [InlineData("=DCOUNT(A1:C5,,H1:H2)", 2)]
    [InlineData("=DMAX(A1:C5,\"Sales\",H1:H2)", 300)]
    [InlineData("=DMIN(A1:C5,\"Sales\",H1:H2)", 100)]
    [InlineData("=DPRODUCT(A1:C5,\"Sales\",H1:H2)", 30000)]
    [InlineData("=DVARP(A1:C5,\"Sales\",H1:H2)", 10000)]
    [InlineData("=DVAR(A1:C5,\"Sales\",H1:H2)", 20000)]
    [InlineData("=DSTDEVP(A1:C5,\"Sales\",H1:H2)", 100)]
    public void Database_functions_select_records_by_the_criteria_range(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Database), 9);
    }

    private static CellValue InCulture(string culture, string formula, params (string Address, CellValue Value)[] cells)
    {
        Workbook book = Workbook.Create();
        book.Culture = System.Globalization.CultureInfo.GetCultureInfo(culture);
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        foreach ((string address, CellValue value) in cells)
        {
            sheet[address] = value;
        }

        Result<CellValue> result = book.Evaluate(formula, sheet);
        Assert.True(result.IsSuccess(out CellValue evaluated), result.Error);

        return evaluated;
    }

    [Theory]
    [InlineData("en-GB", "=YEAR(\"31/12/2018\")", 2018)]
    [InlineData("en-GB", "=MONTH(\"3/2/2013\")", 2)]
    [InlineData("en-US", "=MONTH(\"3/2/2013\")", 3)]
    [InlineData("sv-SE", "=\"1,5\"*2", 3)]
    [InlineData("en-US", "=\"1,5\"*2", 30)]
    [InlineData("de-DE", "=\"1.000,5\"+0", 1000.5)]
    [InlineData("sv-SE", "=YEAR(\"2026-02-01\")", 2026)]
    [InlineData("en-GB", "=YEAR(\"2026-02-01\")", 2026)]
    [InlineData("en-GB", "=HOUR(\"8:30\")", 8)]
    public void Text_is_read_as_numbers_and_dates_in_the_workbooks_culture(string culture, string formula, double expected)
    {
        Assert.True(InCulture(culture, formula).TryGetNumber(out double number), formula);
        Assert.Equal(expected, number, 9);
    }

    [Theory]
    [InlineData("en-GB", "=YEAR(\"1/13/1966\")")]
    [InlineData("en-US", "=YEAR(\"1975-02-23T02:58:41.000Z\")")]
    [InlineData("sv-SE", "=HOUR(\"1975-02-23T02:58:41\")")]
    public void Text_Excel_does_not_read_as_a_date_is_VALUE(string culture, string formula)
    {
        Assert.Equal(CellValue.Error("#VALUE!"), InCulture(culture, formula));
    }

    [Fact]
    public void Date_criteria_are_read_in_the_workbooks_culture()
    {
        (string, CellValue)[] cells = [("A1", new DateTime(2013, 2, 1)), ("A2", new DateTime(2013, 2, 3)), ("A3", new DateTime(2013, 3, 2)), ("A4", new DateTime(2013, 6, 2))];

        Assert.Equal((CellValue)3, InCulture("en-GB", "=COUNTIFS(A1:A4,\">=3/2/2013\")", cells));
        Assert.Equal((CellValue)2, InCulture("en-US", "=COUNTIFS(A1:A4,\">=3/2/2013\")", cells));
    }

    [Fact]
    public void DGET_gives_the_one_match_or_an_error()
    {
        Assert.Equal((CellValue)"Bo", Formula.Evaluate("=DGET(A1:C5,\"Name\",{\"Region\";\"West\"})", Database));
        Assert.Equal(CellValue.Error("#NUM!"), Formula.Evaluate("=DGET(A1:C5,\"Name\",H1:H2)", Database));
        Assert.Equal(CellValue.Error("#VALUE!"), Formula.Evaluate("=DGET(A1:C5,\"Name\",{\"Region\";\"South\"})", Database));
        Assert.Equal(CellValue.Error("#VALUE!"), Formula.Evaluate("=DGET(A1:C5,\"Nobody\",H1:H2)", Database));
    }
}
