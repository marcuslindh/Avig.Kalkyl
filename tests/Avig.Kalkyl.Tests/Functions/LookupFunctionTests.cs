namespace Avig.Kalkyl.Tests;

public class LookupFunctionTests
{
    /// <summary>A1:D5 is a table sorted on A; F1:I2 is a sideways table sorted on row 1.</summary>
    private static readonly (string, CellValue)[] Table =
    [
        ("A1", 101), ("B1", "Kovacs"), ("C1", 10), ("D1", 1),
        ("A2", 102), ("B2", "Brown"), ("C2", 20), ("D2", 2),
        ("A3", 103), ("B3", "Dalton"), ("C3", 30), ("D3", 3),
        ("A4", 104), ("B4", "Ricci"), ("C4", 40), ("D4", 4),
        ("A5", 105), ("B5", "Brown"), ("C5", 50), ("D5", 5),
        ("F1", "Axles"), ("G1", "Bearings"), ("H1", "Bolts"), ("I1", "Gears"),
        ("F2", 4), ("G2", 4), ("H2", 9), ("I2", 7),
    ];

    [Theory]
    [InlineData("=VLOOKUP(102,A1:D5,3,FALSE)", 20)]
    [InlineData("=VLOOKUP(102.5,A1:D5,3)", 20)]
    [InlineData("=VLOOKUP(999,A1:D5,4,TRUE)", 5)]
    [InlineData("=VLOOKUP(\"brown\",B1:C5,2,FALSE)", 20)]
    [InlineData("=VLOOKUP(\"R*\",B1:C5,2,FALSE)", 40)]
    [InlineData("=VLOOKUP(\"?alton\",B1:D5,3,0)", 3)]
    [InlineData("=VLOOKUP(103,A:D,4,FALSE)", 3)]
    [InlineData("=HLOOKUP(\"Bolts\",F1:I2,2,FALSE)", 9)]
    [InlineData("=HLOOKUP(\"C\",F1:I2,2)", 9)]
    [InlineData("=HLOOKUP(\"b*\",F1:I2,2,FALSE)", 4)]
    [InlineData("=MATCH(25,{10,20,30},1)", 2)]
    [InlineData("=MATCH(25,{10,20,30})", 2)]
    [InlineData("=MATCH(30,{10,20,30},0)", 3)]
    [InlineData("=MATCH(25,{30,20,10},-1)", 1)]
    [InlineData("=MATCH(20,{30,20,10},-1)", 2)]
    [InlineData("=MATCH(2,{1,2,2,3},1)", 3)]
    [InlineData("=MATCH(\"b?\",{\"a\",\"bc\",\"c\"},0)", 2)]
    [InlineData("=MATCH(\"BC\",{\"a\",\"bc\",\"c\"},0)", 2)]
    [InlineData("=MATCH(104,A1:A5,0)", 4)]
    [InlineData("=MATCH(\"Gears\",F1:I1,0)", 4)]
    [InlineData("=MATCH(104,{\"x\",101,\"y\",103,105},1)", 4)]
    [InlineData("=LOOKUP(4.19,{4.14,4.19,5.17,5.77,6.39},{1,2,3,4,5})", 2)]
    [InlineData("=LOOKUP(5,{4.14,4.19,5.17,5.77,6.39},{1,2,3,4,5})", 2)]
    [InlineData("=LOOKUP(7.66,{4.14,4.19,5.17,5.77,6.39},{1,2,3,4,5})", 5)]
    [InlineData("=LOOKUP(\"C\",{\"a\",\"b\",\"c\",\"d\";1,2,3,4})", 3)]
    [InlineData("=LOOKUP(\"bump\",{\"a\",1;\"b\",2;\"c\",3})", 2)]
    [InlineData("=LOOKUP(103.5,A1:A5,C1:C5)", 30)]
    [InlineData("=XLOOKUP(104,A1:A5,C1:C5)", 40)]
    [InlineData("=XLOOKUP(\"brown\",B1:B5,C1:C5)", 20)]
    [InlineData("=XLOOKUP(\"brown\",B1:B5,C1:C5,,0,-1)", 50)]
    [InlineData("=XLOOKUP(\"br*\",B1:B5,C1:C5,,2)", 20)]
    [InlineData("=XLOOKUP(102.5,A1:A5,C1:C5,,-1)", 20)]
    [InlineData("=XLOOKUP(102.5,A1:A5,C1:C5,,1)", 30)]
    [InlineData("=XLOOKUP(103,A1:A5,C1:C5,,0,2)", 30)]
    [InlineData("=XLOOKUP(999,A1:A5,C1:C5,-1)", -1)]
    [InlineData("=SUM(XLOOKUP(103,A1:A5,C1:D5))", 33)]
    [InlineData("=ROW(XLOOKUP(103,A1:A5,C1:D5))", 3)]
    [InlineData("=XLOOKUP(\"Bolts\",F1:I1,F2:I2)", 9)]
    [InlineData("=XMATCH(4,{5,4,3,2,1})", 2)]
    [InlineData("=XMATCH(4.5,{5,4,3,2,1},1)", 1)]
    [InlineData("=XMATCH(4.5,{5,4,3,2,1},-1)", 2)]
    [InlineData("=XMATCH(3,{1,3,3,5},0,-1)", 3)]
    [InlineData("=XMATCH(3,{1,3,3,5},0,1)", 2)]
    [InlineData("=XMATCH(4,{1,2,3,4,5},0,2)", 4)]
    [InlineData("=XMATCH(4,{5,4,3,2,1},0,-2)", 2)]
    [InlineData("=XMATCH(3.5,{1,2,3,4,5},-1,2)", 3)]
    [InlineData("=XMATCH(3.5,{1,2,3,4,5},1,2)", 4)]
    [InlineData("=XMATCH(3.5,{5,4,3,2,1},-1,-2)", 3)]
    [InlineData("=XMATCH(3.5,{5,4,3,2,1},1,-2)", 2)]
    [InlineData("=XMATCH(\"b*\",{\"apple\",\"banana\"},2)", 2)]
    [InlineData("=INDEX({1,2;3,4},2,2)", 4)]
    [InlineData("=INDEX({1,2,3},2)", 2)]
    [InlineData("=INDEX({1;2;3},3)", 3)]
    [InlineData("=SUM(INDEX({1,2;3,4},0,2))", 6)]
    [InlineData("=SUM(INDEX({1,2;3,4},1))", 3)]
    [InlineData("=INDEX(A1:D5,2,3)", 20)]
    [InlineData("=INDEX(C1:C5,4)", 40)]
    [InlineData("=INDEX(F2:I2,3)", 9)]
    [InlineData("=SUM(INDEX(C1:D5,0,1))", 150)]
    [InlineData("=SUM(INDEX(C1:D5,2,0))", 22)]
    [InlineData("=SUM(INDEX(C1:D5,0,0))", 165)]
    [InlineData("=ROWS(INDEX(A1:D5,0,2))", 5)]
    [InlineData("=ROW(INDEX(A1:D5,3,2))", 3)]
    [InlineData("=COLUMN(INDEX(A1:D5,3,2))", 2)]
    [InlineData("=SUM(A1:INDEX(A1:A5,2))", 203)]
    [InlineData("=INDEX(A1:D5,MATCH(104,A1:A5,0),3)", 40)]
    [InlineData("=CHOOSE(2,10,20,30)", 20)]
    [InlineData("=CHOOSE(2.9,10,20,30)", 20)]
    [InlineData("=SUM(CHOOSE(2,A1:A2,C1:C5))", 150)]
    [InlineData("=ROWS(CHOOSE(1,A1:A3))", 3)]
    [InlineData("=ROWS(A1:D5)", 5)]
    [InlineData("=COLUMNS(A1:D5)", 4)]
    [InlineData("=ROWS({1,2;3,4;5,6})", 3)]
    [InlineData("=COLUMNS({1,2,3})", 3)]
    [InlineData("=ROWS(A:A)", 1048576)]
    [InlineData("=COLUMNS(1:1)", 16384)]
    [InlineData("=ROWS(5)", 1)]
    [InlineData("=INDEX(TRANSPOSE({1,2,3}),3,1)", 3)]
    [InlineData("=ROWS(TRANSPOSE({1,2,3}))", 3)]
    [InlineData("=SUM(TRANSPOSE(C1:D5))", 165)]
    [InlineData("=INDEX(TRANSPOSE(A1:D5),3,2)", 20)]
    [InlineData("=AREAS(A1:B2)", 1)]
    public void Lookup_functions_give_Excels_results(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Table), 9);
    }

    [Theory]
    [InlineData("=VLOOKUP(102,A1:D5,2,FALSE)", "Brown")]
    [InlineData("=VLOOKUP(101.9,A1:D5,2)", "Kovacs")]
    [InlineData("=XLOOKUP(999,A1:A5,B1:B5,\"none\")", "none")]
    [InlineData("=CHOOSE(3,\"a\",\"b\",\"c\")", "c")]
    [InlineData("=INDEX(B1:B5,3)", "Dalton")]
    [InlineData("=LOOKUP(102,A1:B5)", "Brown")]
    [InlineData("=ADDRESS(2,3)", "$C$2")]
    [InlineData("=ADDRESS(2,3,2)", "C$2")]
    [InlineData("=ADDRESS(2,3,3)", "$C2")]
    [InlineData("=ADDRESS(1,27,4)", "AA1")]
    [InlineData("=ADDRESS(2,3,2,FALSE)", "R2C[3]")]
    [InlineData("=ADDRESS(2,3,4,FALSE)", "R[2]C[3]")]
    [InlineData("=ADDRESS(2,3,1,FALSE)", "R2C3")]
    [InlineData("=ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")", "'EXCEL SHEET'!R2C3")]
    [InlineData("=ADDRESS(1,1,1,TRUE,\"Data\")", "Data!$A$1")]
    [InlineData("=ADDRESS(1048576,16384)", "$XFD$1048576")]
    [InlineData("=HYPERLINK(\"https://example.com\",\"Click\")", "Click")]
    [InlineData("=HYPERLINK(\"https://example.com\")", "https://example.com")]
    public void Lookup_functions_give_Excels_text_results(string formula, string expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula, Table));
    }

    [Theory]
    [InlineData("=VLOOKUP(99,A1:D5,2)", "#N/A")]
    [InlineData("=VLOOKUP(106,A1:D5,2,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(102,A1:D5,5,FALSE)", "#REF!")]
    [InlineData("=VLOOKUP(102,A1:D5,0,FALSE)", "#VALUE!")]
    [InlineData("=VLOOKUP(\"102\",A1:D5,2,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(Z1,A1:D5,2,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(1/0,A1:D5,2,FALSE)", "#DIV/0!")]
    [InlineData("=HLOOKUP(\"Bolts\",F1:I2,3,FALSE)", "#REF!")]
    [InlineData("=HLOOKUP(\"Nuts\",F1:I2,2,FALSE)", "#N/A")]
    [InlineData("=MATCH(5,{10,20,30},1)", "#N/A")]
    [InlineData("=MATCH(40,{30,20,10},-1)", "#N/A")]
    [InlineData("=MATCH(1,{1,2;3,4},0)", "#N/A")]
    [InlineData("=MATCH(\"b*\",{\"a\",\"c\"},0)", "#N/A")]
    [InlineData("=LOOKUP(0,{4.14,4.19},{1,2})", "#N/A")]
    [InlineData("=XLOOKUP(999,A1:A5,C1:C5)", "#N/A")]
    [InlineData("=XLOOKUP(101,A1:A5,C1:C4)", "#VALUE!")]
    [InlineData("=XLOOKUP(101,A1:A5,C1:C5,,4)", "#VALUE!")]
    [InlineData("=XLOOKUP(101,A1:A5,C1:C5,,0,0)", "#VALUE!")]
    [InlineData("=XMATCH(\"b*\",{\"apple\",\"banana\"})", "#N/A")]
    [InlineData("=XMATCH(1,{1,2;3,4})", "#VALUE!")]
    [InlineData("=XMATCH(6,{1,2,3,4,5},1)", "#N/A")]
    [InlineData("=INDEX(A1:D5,6,1)", "#REF!")]
    [InlineData("=INDEX(A1:D5,1,5)", "#REF!")]
    [InlineData("=INDEX(A1:A5,-1)", "#VALUE!")]
    [InlineData("=INDEX(A1:D5,1,1,2)", "#REF!")]
    [InlineData("=INDEX({1,2;3,4},3,1)", "#REF!")]
    [InlineData("=CHOOSE(4,1,2,3)", "#VALUE!")]
    [InlineData("=CHOOSE(0,1,2,3)", "#VALUE!")]
    [InlineData("=ADDRESS(0,1)", "#VALUE!")]
    [InlineData("=ADDRESS(1,16385)", "#VALUE!")]
    [InlineData("=ADDRESS(1,1,5)", "#VALUE!")]
    [InlineData("=AREAS(1)", "#VALUE!")]
    [InlineData("=FORMULATEXT(A1)", "#N/A")]
    [InlineData("=FORMULATEXT(1)", "#VALUE!")]
    public void Lookup_errors_are_Excels_errors(string formula, string error)
    {
        CellValue value = Formula.Evaluate(formula, Table);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(error, code);
    }

    [Fact]
    public void Formulatext_gives_the_formula_of_the_referenced_cell()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 5;
        sheet["B1"] = Cell.FromFormula("SUM(A1,2)*3");

        Result<CellValue> result = book.Evaluate("FORMULATEXT(B1)", sheet);

        Assert.True(result.IsSuccess(out CellValue value), result.Error);
        Assert.Equal((CellValue)"=SUM(A1,2)*3", value);
    }

    [Fact]
    public void Xlookup_finds_blank_cells_and_lookups_ignore_case()
    {
        (string, CellValue)[] cells = [("A1", "x"), ("A3", "y"), ("B1", 1), ("B2", 2), ("B3", 3)];

        Assert.Equal(2, Formula.Number("=XLOOKUP(A2,A1:A3,B1:B3)", cells));
        Assert.Equal(3, Formula.Number("=XLOOKUP(\"Y\",A1:A3,B1:B3)", cells));
        Assert.Equal(3, Formula.Number("=MATCH(\"Y\",A1:A3,0)", cells));
    }
}
