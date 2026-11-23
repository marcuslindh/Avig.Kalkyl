using Avig.Kalkyl.Formulas;

namespace Avig.Kalkyl.Tests;

/// <summary>Functions from different categories working together in one recalculated, saved workbook.</summary>
public class IntegrationTests
{
    [Fact]
    public void The_library_has_more_than_three_hundred_functions()
    {
        Assert.True(FunctionLibrary.Names.Count() > 300, $"Only {FunctionLibrary.Names.Count()} functions.");
    }

    [Theory]
    [InlineData("=INDEX(MODE.MULT({1,2,2,3,3,4}),2)", 3)]
    [InlineData("=SUMPRODUCT((A2:A5=\"Öst\")*C2:C5)", 300)]
    [InlineData("=SUMIFS(C2:C5,A2:A5,\"Väst\",B2:B5,\">=2026-02-01\")", 400)]
    [InlineData("=VLOOKUP(\"Norr\",A2:C5,3,FALSE)", 50)]
    [InlineData("=XLOOKUP(MAX(C2:C5),C2:C5,A2:A5)=\"Väst\"", 1)]
    [InlineData("=ROUND(STDEV.S(C2:C5),4)", 154.7848)]
    [InlineData("=LEN(TEXTJOIN(\",\",TRUE,A2:A5))", 17)]
    [InlineData("=YEAR(EDATE(B2,11))+MONTH(B5)", 2030)]
    [InlineData("=ROUND(PMT(5%/12,60,-SUM(C2:C5)),2)", 14.15)]
    [InlineData("=HEX2DEC(DEC2HEX(COUNTIF(A2:A5,\"*st\")*100))", 300)]
    public void Functions_combine_across_categories(string formula, double expected)
    {
        (string, CellValue)[] data =
        [
            ("A2", "Öst"), ("B2", new DateTime(2026, 1, 15)), ("C2", 100),
            ("A3", "Väst"), ("B3", new DateTime(2026, 2, 3)), ("C3", 400),
            ("A4", "Öst"), ("B4", new DateTime(2026, 3, 9)), ("C4", 200),
            ("A5", "Norr"), ("B5", new DateTime(2026, 4, 1)), ("C5", 50),
        ];

        CellValue value = Formula.Evaluate(formula, data);

        if (value.TryGetBool(out bool flag) == true)
        {
            Assert.Equal(expected == 1, flag);

            return;
        }

        Assert.True(value.TryGetNumber(out double number), $"{formula} gave {value}.");
        Assert.Equal(expected, number, 4);
    }

    [Fact]
    public async Task A_budget_is_recalculated_saved_and_valid()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Budget").IsSuccess(out Worksheet? budget));
        Assert.True(book.AddSheet("Sammanfattning").IsSuccess(out Worksheet? summary));
        string[] months = ["Jan", "Feb", "Mar"];

        for (int index = 0; index < months.Length; index++)
        {
            budget[index + 1, 0] = months[index];
            budget[index + 1, 1] = (index + 1) * 1000;
            budget[index + 1, 2] = Cell.FromFormula($"B{index + 2}*1.25").NumberFormat("0.00");
        }

        budget["A5"] = Cell.Of("Summa").Bold();
        budget["B5"] = Cell.FromFormula("SUBTOTAL(9,B2:B4)");
        budget["C5"] = Cell.FromFormula("SUBTOTAL(9,C2:C4)");
        summary["A1"] = Cell.FromFormula("TEXT(Budget!C5,\"# ##0.00\")&\" kr\"");
        summary["A2"] = Cell.FromFormula("IFERROR(Budget!C5/Budget!B9,\"saknas\")");
        summary["A3"] = Cell.FromFormula("AVERAGE(Budget!B2:B5)");

        RecalculationReport report = await book.RecalculateAsync();

        Assert.Empty(report.UnknownFunctions);
        Assert.Empty(report.CircularReferences);
        Assert.Equal((CellValue)6000, budget["B5"].Value);
        Assert.Equal((CellValue)7500, budget["C5"].Value);
        Assert.Equal((CellValue)"7 500.00 kr", summary["A1"].Value);
        Assert.Equal((CellValue)"saknas", summary["A2"].Value);
        Assert.Equal((CellValue)3000, summary["A3"].Value);

        await using MemoryStream file = new();
        await book.SaveAsync(file);
        Assert.Empty(TestFiles.Validate(file));
    }
}
