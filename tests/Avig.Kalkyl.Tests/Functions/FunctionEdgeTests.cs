namespace Avig.Kalkyl.Tests;

/// <summary>
/// Zero, empty, too large and the other edges of the math, lookup, date and financial functions: values checked in
/// Excel 365, and relations between functions that must hold.
/// </summary>
public class FunctionEdgeTests
{
    [Theory]
    [InlineData("=FLOOR(0,5)", 0)]
    [InlineData("=CEILING.MATH(0,5)", 0)]
    [InlineData("=FLOOR.MATH(7,0)", 0)]
    [InlineData("=CEILING.PRECISE(0,2)", 0)]
    [InlineData("=FLOOR.PRECISE(5,0)", 0)]
    [InlineData("=ISO.CEILING(5,0)", 0)]
    [InlineData("=ROUND(0,2)", 0)]
    [InlineData("=ROUNDUP(0,1)", 0)]
    [InlineData("=MROUND(0,3)", 0)]
    [InlineData("=MROUND(5,0)", 0)]
    [InlineData("=COMBINA(5,0)", 1)]
    [InlineData("=DDB(1000,100,1,1)", 900)]
    [InlineData("=YEARFRAC(DATE(2011,1,31),DATE(2011,3,31),0)", 60.0 / 360)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 211.0 / 366)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),0)", 209.0 / 360)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)", 211.0 / 365)]
    public void The_edges_have_their_exact_values(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 12);
    }

    [Theory]
    [InlineData("=ROMAN(499,0)", "CDXCIX")]
    [InlineData("=ROMAN(499,1)", "LDVLIV")]
    [InlineData("=ROMAN(499,2)", "XDIX")]
    [InlineData("=ROMAN(499,3)", "VDIV")]
    [InlineData("=ROMAN(499,4)", "ID")]
    [InlineData("=ADDRESS(1,1,1,TRUE,\"\")", "!$A$1")]
    public void Text_results_are_Excels(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    [Theory]
    [InlineData("=MOD(1E+16,3)", "#NUM!")]
    [InlineData("=COMBIN(2000,1000)", "#NUM!")]
    [InlineData("=SUMPRODUCT({1,2},)", "#VALUE!")]
    [InlineData("=SUMX2MY2({1,#N/A},{1,2})", "#N/A")]
    [InlineData("=AGGREGATE(4,6,{#N/A})", "#VALUE!")]
    [InlineData("=AGGREGATE(12,6,{#N/A})", "#VALUE!")]
    [InlineData("=ARABIC(REPT(\"I\",256))", "#VALUE!")]
    [InlineData("=DECIMAL(REPT(\"1\",256),2)", "#VALUE!")]
    [InlineData("=DECIMAL(\"1*\",10)", "#NUM!")]
    [InlineData("=XLOOKUP(2,{1,2,3},{4,5})", "#VALUE!")]
    [InlineData("=XLOOKUP(2,{1,2,3},{4;5;6})", "#VALUE!")]
    [InlineData("=XLOOKUP(2,{1;2;3},{4;5})", "#VALUE!")]
    [InlineData("=INDIRECT(\"1+2\")", "#REF!")]
    [InlineData("=DATE(10000,1,1)", "#NUM!")]
    [InlineData("=EDATE(DATE(9999,12,1),1)", "#NUM!")]
    [InlineData("=WORKDAY(1,1E+10)", "#NUM!")]
    [InlineData("=DATEVALUE(\"10%\")", "#VALUE!")]
    [InlineData("=DATEVALUE(\"$5\")", "#VALUE!")]
    [InlineData("=NETWORKDAYS(1,10,{\"x\"})", "#VALUE!")]
    [InlineData("=NETWORKDAYS(1,10,{-1})", "#NUM!")]
    [InlineData("=NPER(0.1,-50,1000)", "#NUM!")]
    [InlineData("=RATE(12,100,1000)", "#NUM!")]
    [InlineData("=IRR({100,200,300})", "#NUM!")]
    [InlineData("=XIRR({100,50},{40000,40100})", "#NUM!")]
    [InlineData("=XNPV(-1,{-100,50},{40000,40100})", "#NUM!")]
    [InlineData("=XNPV(0.1,{-100,TRUE},{40000,40100})", "#NUM!")]
    [InlineData("=FVSCHEDULE(1,{0.1,TRUE})", "#VALUE!")]
    public void Excels_errors_come_from_the_edges(string formula, string error)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code) == true && code == error, $"{formula} gave {value}.");
    }

    [Fact]
    public void Column_without_an_argument_is_the_column_of_the_formula()
    {
        Assert.Equal(Formula.Number("=ROW()"), Formula.Number("=COLUMN()"));
    }

    [Fact]
    public void Offset_moves_and_resizes_a_reference_and_refuses_to_leave_the_sheet()
    {
        (string, CellValue)[] cells = [("B2", 1), ("C2", 2), ("B3", 3), ("C3", 4)];

        Assert.Equal(4, Formula.Number("=OFFSET(A1,2,2)", cells));
        Assert.Equal(10, Formula.Number("=SUM(OFFSET(A1,1,1,2,2))", cells));
        Assert.Equal(10, Formula.Number("=SUM(OFFSET(C3,0,0,-2,-2))", cells));
        Assert.Equal("#REF!", Error("=OFFSET(A1,-1,0)", cells));
        Assert.Equal("#REF!", Error("=OFFSET(A1,0,0,0,1)", cells));
        Assert.Equal("#DIV/0!", Error("=OFFSET(A1,1/0,0)", cells));
        Assert.Equal("#VALUE!", Error("=OFFSET(A1,\"x\",0)", cells));
    }

    [Fact]
    public void Lookup_of_an_empty_cell_is_NA()
    {
        Assert.Equal("#N/A", Error("=LOOKUP(A5,{1,2,3})"));
    }

    [Theory]
    [InlineData(0.1, 3, 10, 1000)]
    [InlineData(0.05, 2, 24, -5000)]
    [InlineData(0.01, 12, 12, 250)]
    public void Interest_and_principal_add_up_to_the_payment_in_advance(double rate, int period, int periods, double present)
    {
        string arguments = $"{Text(rate)},{period},{periods},{Text(present)},0,1";
        double interest = Formula.Number($"=IPMT({arguments})");
        double principal = Formula.Number($"=PPMT({arguments})");
        double payment = Formula.Number($"=PMT({Text(rate)},{periods},{Text(present)},0,1)");

        Assert.Equal(payment, interest + principal, 9);
    }

    [Theory]
    [InlineData(0.5, 1.5)]
    [InlineData(1.25, 3.75)]
    public void Vdb_over_a_fractional_period_is_the_sum_of_its_parts(double middle, double end)
    {
        double whole = Formula.Number($"=VDB(2400,300,10,0,{Text(end)})");
        double first = Formula.Number($"=VDB(2400,300,10,0,{Text(middle)})");
        double second = Formula.Number($"=VDB(2400,300,10,{Text(middle)},{Text(end)})");

        Assert.Equal(whole, first + second, 9);
    }

    private static string Error(string formula, params (string Address, CellValue Value)[] cells)
    {
        CellValue value = Formula.Evaluate(formula, cells);
        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");

        return code;
    }

    private static string Text(double number)
    {
        return number.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
    }
}
