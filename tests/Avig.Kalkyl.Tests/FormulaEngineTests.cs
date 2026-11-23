using Avig.Kalkyl.Formulas;

namespace Avig.Kalkyl.Tests;

/// <summary>Evaluates formulas the way the function tests need: on a fresh workbook with given cells.</summary>
internal static class Formula
{
    /// <summary>Evaluates a formula on a sheet named Data with the given cells set first.</summary>
    public static CellValue Evaluate(string formula, params (string Address, CellValue Value)[] cells)
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        foreach ((string address, CellValue value) in cells)
        {
            sheet[address] = value;
        }

        Result<CellValue> result = book.Evaluate(formula, sheet);
        Assert.True(result.IsSuccess(out CellValue evaluated), result.Error);

        return evaluated;
    }

    /// <summary>The result as a number, failing the test when it is something else.</summary>
    public static double Number(string formula, params (string Address, CellValue Value)[] cells)
    {
        CellValue value = Evaluate(formula, cells);
        Assert.True(value.TryGetNumber(out double number), $"{formula} gave {value} ({value.Kind}), not a number.");

        return number;
    }
}

public class FormulaEngineTests
{
    private static readonly string[] CircularCells = ["Data!A1", "Data!B1", "Data!C1"];

    private static readonly string[] UnknownNames = ["NOTREAL"];

    [Theory]
    [InlineData("=-2^2", 4)]
    [InlineData("=2+3*4", 14)]
    [InlineData("=(2+3)*4", 20)]
    [InlineData("=10/4", 2.5)]
    [InlineData("=2^3^2", 64)]
    [InlineData("=50%", 0.5)]
    [InlineData("=-50%*2", -1)]
    [InlineData("=TRUE+1", 2)]
    [InlineData("=\"5\"+1", 6)]
    [InlineData("=\"1,000\"*2", 2000)]
    [InlineData("=\"10%\"*10", 1)]
    [InlineData("=1E3+1", 1001)]
    [InlineData("=SUM(1,2,3)", 6)]
    [InlineData("=SUM({1,2;3,4})", 10)]
    [InlineData("=ROUND(2.675,2)", 2.68)]
    [InlineData("=ROUND(-2.5,0)", -3)]
    [InlineData("=ROUND(1234.5,-2)", 1200)]
    [InlineData("=ABS(-3)", 3)]
    public void Arithmetic_follows_Excels_precedence_and_conversions(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 12);
    }

    [Theory]
    [InlineData("=1/0", "#DIV/0!")]
    [InlineData("=\"x\"+1", "#VALUE!")]
    [InlineData("=0^0", "#NUM!")]
    [InlineData("=NOSUCHFUNCTION(1)", "#NAME?")]
    [InlineData("=NoSuchName+1", "#NAME?")]
    [InlineData("=SUM(1,#N/A)", "#N/A")]
    [InlineData("=Saknas!A1", "#REF!")]
    [InlineData("=SUM(1)/SUM(0)", "#DIV/0!")]
    public void Errors_are_Excels_errors(string formula, string error)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(error, code);
    }

    [Theory]
    [InlineData("=\"a\"&1&TRUE", "a1TRUE")]
    [InlineData("=1/3&\"\"", "0.333333333333333")]
    [InlineData("=IF(1>2,\"ja\",\"nej\")", "nej")]
    public void Text_results(string formula, string expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula));
    }

    [Theory]
    [InlineData("=\"A\"=\"a\"", true)]
    [InlineData("=1<\"a\"", true)]
    [InlineData("=\"z\"<TRUE", true)]
    [InlineData("=A1=0", true)]
    [InlineData("=A1=\"\"", true)]
    [InlineData("=AND(TRUE,1,\"TRUE\")", true)]
    [InlineData("=OR(FALSE,0)", false)]
    [InlineData("=NOT(0)", true)]
    [InlineData("=IF(TRUE,1,1/0)=1", true)]
    [InlineData("=IFERROR(1/0,7)=7", true)]
    [InlineData("=IF(FALSE,1)", false)]
    public void Comparisons_and_logic(string formula, bool expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula));
    }

    [Fact]
    public void References_ranges_and_arrays()
    {
        (string, CellValue)[] cells = [("A1", 1), ("A2", 2), ("A3", 3), ("B1", "text"), ("B2", true), ("C1", 10)];

        Assert.Equal(6, Formula.Number("=SUM(A1:A3)", cells));
        Assert.Equal(6, Formula.Number("=SUM(A:A)", cells));
        Assert.Equal(13, Formula.Number("=SUM(A1:C1, A2)", cells));
        Assert.Equal(12, Formula.Number("=SUM(A1:A3*2)", cells));
        Assert.Equal(14, Formula.Number("=SUM(A1:A3*{1;2;3})", cells));
        Assert.Equal(6, Formula.Number("=SUM(A1:A2:A3)", cells));
        Assert.Equal(3, Formula.Number("=ROW(A3)", cells));
        Assert.Equal(1, Formula.Number("=ROW()", cells));
        Assert.Equal(3, Formula.Number("=COLUMN(C9)", cells));
        Assert.Equal(10, Formula.Number("=INDIRECT(\"C1\")", cells));
        Assert.Equal(5, Formula.Number("=SUM(OFFSET(A1,1,0,2,1))", cells));
    }

    [Fact]
    public async Task Dates_are_serial_numbers_and_a_date_format_turns_results_into_dates()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = new DateTime(2026, 11, 23);
        sheet["B1"] = Cell.FromFormula("A1+7").NumberFormat("yyyy-mm-dd");
        sheet["C1"] = Cell.FromFormula("A1+7");

        RecalculationReport report = await RecalculateAsync(book);

        Assert.Equal(2, report.Calculated);
        Assert.Equal((CellValue)new DateTime(2026, 11, 30), sheet["B1"].Value);
        Assert.Equal((CellValue)46356, sheet["C1"].Value);
    }

    [Fact]
    public async Task Recalculation_follows_dependencies_and_only_touches_what_changed()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        Assert.True(book.AddSheet("Andra blad").IsSuccess(out Worksheet? other));
        sheet["C1"] = Cell.FromFormula("B1+1");
        sheet["B1"] = Cell.FromFormula("A1*2");
        sheet["A1"] = 1;
        sheet["D1"] = Cell.FromFormula("10");
        other["A1"] = Cell.FromFormula("Data!C1*10");

        RecalculationReport first = await RecalculateAsync(book);
        Assert.Equal(4, first.Calculated);
        Assert.Equal((CellValue)3, sheet["C1"].Value);
        Assert.Equal((CellValue)30, other["A1"].Value);

        sheet["A1"] = 2;
        RecalculationReport second = await RecalculateAsync(book);
        Assert.Equal(3, second.Calculated);
        Assert.Equal((CellValue)5, sheet["C1"].Value);
        Assert.Equal((CellValue)50, other["A1"].Value);

        Assert.Equal(0, (await RecalculateAsync(book)).Calculated);
        Assert.Equal(4, (await RecalculateAsync(book, RecalculationScope.All)).Calculated);
    }

    [Fact]
    public async Task Circular_references_get_zero_and_are_reported()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = Cell.FromFormula("B1+1");
        sheet["B1"] = Cell.FromFormula("A1+1");
        sheet["C1"] = Cell.FromFormula("C1+1");
        sheet["D1"] = Cell.FromFormula("1+1");

        RecalculationReport report = await RecalculateAsync(book);

        Assert.Equal(CircularCells, report.CircularReferences.Order(StringComparer.Ordinal));
        Assert.Equal((CellValue)2, sheet["D1"].Value);
    }

    [Fact]
    public async Task Unknown_functions_and_broken_formulas_are_reported()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = Cell.FromFormula("_xlfn.NOTREAL(1)");
        sheet["A2"] = Cell.FromFormula("SUM(1,");

        RecalculationReport report = await RecalculateAsync(book);

        Assert.Equal(UnknownNames, report.UnknownFunctions);
        Assert.Single(report.InvalidFormulas);
        Assert.Equal((CellValue)CellValue.Error("#NAME?"), sheet["A1"].Value);
        Assert.Equal((CellValue)CellValue.Error("#NAME?"), sheet["A2"].Value);
    }

    [Fact]
    public async Task Volatile_functions_are_recalculated_every_time()
    {
        Workbook book = Workbook.Create();
        FakeClock clock = new(new DateTimeOffset(2026, 11, 23, 9, 30, 0, TimeSpan.Zero));
        book.Clock = clock;
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = Cell.FromFormula("TODAY()").NumberFormat("yyyy-mm-dd");
        sheet["B1"] = Cell.FromFormula("A1+1");

        await RecalculateAsync(book);
        Assert.Equal((CellValue)new DateTime(2026, 11, 23), sheet["A1"].Value);

        clock.Now = clock.Now.AddDays(1);
        RecalculationReport report = await RecalculateAsync(book);
        Assert.Equal(2, report.Calculated);
        Assert.Equal((CellValue)new DateTime(2026, 11, 24), sheet["A1"].Value);
    }

    [Fact]
    public async Task A_cell_reached_through_INDIRECT_is_calculated_before_it_is_read()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = Cell.FromFormula("INDIRECT(\"B1\")*2");
        sheet["B1"] = Cell.FromFormula("C1+1");
        sheet["C1"] = 4;

        await RecalculateAsync(book, RecalculationScope.All);

        Assert.Equal((CellValue)10, sheet["A1"].Value);
    }

    [Fact]
    public async Task Defined_names_are_resolved()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 100;
        book.DefinedNames.Add(new DefinedName("Moms", -1, "0.25"));
        book.DefinedNames.Add(new DefinedName("Pris", -1, "Data!$A$1"));
        sheet["B1"] = Cell.FromFormula("Pris*(1+Moms)");

        await RecalculateAsync(book);

        Assert.Equal((CellValue)125, sheet["B1"].Value);
    }

    [Fact]
    public async Task Calculated_values_are_saved_with_the_file()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 20;
        sheet["A2"] = Cell.FromFormula("A1*2");
        await book.RecalculateAsync();

        await using MemoryStream file = new();
        await book.SaveAsync(file);
        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal((CellValue)40, rows[1][0].Value);
        Assert.Equal("A1*2", rows[1][0].Formula);
    }

    private static async Task<RecalculationReport> RecalculateAsync(Workbook book, RecalculationScope scope = RecalculationScope.Changed)
    {
        return await book.RecalculateAsync(scope);
    }

    private sealed class FakeClock(DateTimeOffset now) : TimeProvider
    {
        public DateTimeOffset Now { get; set; } = now;

        public override DateTimeOffset GetUtcNow()
        {
            return Now;
        }

        public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    }
}
