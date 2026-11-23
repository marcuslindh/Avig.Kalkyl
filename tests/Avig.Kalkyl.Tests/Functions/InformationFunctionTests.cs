namespace Avig.Kalkyl.Tests;

public class InformationFunctionTests
{
    private static readonly (string, CellValue)[] Cells =
    [
        ("A1", 5),
        ("A2", "text"),
        ("A3", ""),
        ("A4", true),
        ("A5", new DateTime(2026, 1, 2)),
        ("A6", CellValue.Error("#N/A")),
    ];

    [Theory]
    [InlineData("=ISBLANK(B1)", true)]
    [InlineData("=ISBLANK(A1)", false)]
    [InlineData("=ISBLANK(A3)", false)]
    [InlineData("=ISBLANK(\"\")", false)]
    [InlineData("=ISERROR(1/0)", true)]
    [InlineData("=ISERROR(A6)", true)]
    [InlineData("=ISERROR(A1)", false)]
    [InlineData("=ISERR(NA())", false)]
    [InlineData("=ISERR(1/0)", true)]
    [InlineData("=ISERR(5)", false)]
    [InlineData("=ISNA(NA())", true)]
    [InlineData("=ISNA(A6)", true)]
    [InlineData("=ISNA(1/0)", false)]
    [InlineData("=ISNUMBER(A1)", true)]
    [InlineData("=ISNUMBER(A5)", true)]
    [InlineData("=ISNUMBER(\"5\")", false)]
    [InlineData("=ISNUMBER(B1)", false)]
    [InlineData("=ISTEXT(A2)", true)]
    [InlineData("=ISTEXT(A3)", true)]
    [InlineData("=ISTEXT(A1)", false)]
    [InlineData("=ISNONTEXT(B1)", true)]
    [InlineData("=ISNONTEXT(A1)", true)]
    [InlineData("=ISNONTEXT(A2)", false)]
    [InlineData("=ISLOGICAL(A4)", true)]
    [InlineData("=ISLOGICAL(1)", false)]
    [InlineData("=ISEVEN(-1)", false)]
    [InlineData("=ISEVEN(2.5)", true)]
    [InlineData("=ISEVEN(0)", true)]
    [InlineData("=ISEVEN(B1)", true)]
    [InlineData("=ISEVEN(\"4\")", true)]
    [InlineData("=ISODD(5)", true)]
    [InlineData("=ISODD(-3.9)", true)]
    [InlineData("=ISODD(2)", false)]
    [InlineData("=ISREF(A1)", true)]
    [InlineData("=ISREF(A1:B2)", true)]
    [InlineData("=ISREF(\"A1\")", false)]
    [InlineData("=ISREF(1/0)", false)]
    public void Tests(string formula, bool expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula, Cells));
    }

    [Theory]
    [InlineData("=ERROR.TYPE(#NULL!)", 1)]
    [InlineData("=ERROR.TYPE(1/0)", 2)]
    [InlineData("=ERROR.TYPE(\"a\"+1)", 3)]
    [InlineData("=ERROR.TYPE(#REF!)", 4)]
    [InlineData("=ERROR.TYPE(#NAME?)", 5)]
    [InlineData("=ERROR.TYPE(#NUM!)", 6)]
    [InlineData("=ERROR.TYPE(NA())", 7)]
    [InlineData("=ERROR.TYPE(A6)", 7)]
    [InlineData("=TYPE(A1)", 1)]
    [InlineData("=TYPE(B1)", 1)]
    [InlineData("=TYPE(A2)", 2)]
    [InlineData("=TYPE(\"Smith\")", 2)]
    [InlineData("=TYPE(A4)", 4)]
    [InlineData("=TYPE(1/0)", 16)]
    [InlineData("=TYPE({1,2;3,4})", 64)]
    [InlineData("=TYPE(A1:A3)", 64)]
    [InlineData("=N(7)", 7)]
    [InlineData("=N(A1)", 5)]
    [InlineData("=N(\"7\")", 0)]
    [InlineData("=N(A2)", 0)]
    [InlineData("=N(A4)", 1)]
    [InlineData("=N(FALSE)", 0)]
    [InlineData("=N(A5)", 46024)]
    [InlineData("=N(B1)", 0)]
    public void Numbers(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Cells), 9);
    }

    [Theory]
    [InlineData("=NA()", "#N/A")]
    [InlineData("=ERROR.TYPE(1)", "#N/A")]
    [InlineData("=ERROR.TYPE(\"x\")", "#N/A")]
    [InlineData("=N(1/0)", "#DIV/0!")]
    [InlineData("=ISEVEN(\"x\")", "#VALUE!")]
    [InlineData("=ISEVEN(TRUE)", "#VALUE!")]
    [InlineData("=ISODD(1/0)", "#DIV/0!")]
    [InlineData("=ISFORMULA(1)", "#VALUE!")]
    public void Errors(string formula, string expected)
    {
        CellValue value = Formula.Evaluate(formula, Cells);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(expected, code);
    }

    [Fact]
    public void Isformula_reads_the_referenced_cell()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 5;
        sheet.SetFormula(1, 0, "A1*2");

        Assert.Equal((CellValue)true, Evaluate(book, sheet, "=ISFORMULA(A2)"));
        Assert.Equal((CellValue)false, Evaluate(book, sheet, "=ISFORMULA(A1)"));
        Assert.Equal((CellValue)false, Evaluate(book, sheet, "=ISFORMULA(C9)"));
    }

    private static CellValue Evaluate(Workbook book, Worksheet sheet, string formula)
    {
        Result<CellValue> result = book.Evaluate(formula, sheet);
        Assert.True(result.IsSuccess(out CellValue value), result.Error);

        return value;
    }
}
