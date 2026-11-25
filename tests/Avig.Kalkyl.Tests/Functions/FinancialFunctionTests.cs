namespace Avig.Kalkyl.Tests;

public class FinancialFunctionTests
{
    [Theory]
    [InlineData("=PMT(0.08/12,10,10000)", -1037.0320893)]
    [InlineData("=PMT(0.08/12,10,10000,0,1)", -1030.1643271)]
    [InlineData("=PMT(0.06/12,18*12,0,50000)", -129.0811609)]
    [InlineData("=PMT(0,10,1000)", -100)]
    [InlineData("=IPMT(0.1/12,1,3*12,8000)", -66.6666667)]
    [InlineData("=IPMT(0.1,3,3,8000)", -292.4471299)]
    [InlineData("=IPMT(0.1,1,3,8000,0,1)", 0)]
    [InlineData("=PPMT(0.1/12,1,2*12,2000)", -75.6231860)]
    [InlineData("=PPMT(0.08,10,10,200000)", -27598.0534624)]
    [InlineData("=PV(0.08/12,12*20,500)", -59777.1458511)]
    [InlineData("=PV(0,10,100)", -1000)]
    [InlineData("=FV(0.06/12,10,-200,-500,1)", 2581.4033741)]
    [InlineData("=FV(0.12/12,12,-1000)", 12682.5030132)]
    [InlineData("=FV(0,12,-100,-1000)", 2200)]
    [InlineData("=NPER(0.12/12,-100,-1000,10000,1)", 59.6738657)]
    [InlineData("=NPER(0.12/12,-100,-1000)", -9.5785940)]
    [InlineData("=NPER(0,-100,1000)", 10)]
    [InlineData("=RATE(4*12,-200,8000)", 0.0077014)]
    [InlineData("=RATE(10,-100,1000)", -0.0000000)]
    [InlineData("=RATE(10,-120,1000)", 0.0346015)]
    [InlineData("=NPV(0.1,-10000,3000,4200,6800)", 1188.4434123)]
    [InlineData("=SLN(30000,7500,10)", 2250)]
    [InlineData("=SYD(30000,7500,10,1)", 4090.9090909)]
    [InlineData("=SYD(30000,7500,10,10)", 409.0909091)]
    [InlineData("=DB(1000000,100000,6,1,7)", 186083.3333333)]
    [InlineData("=DB(1000000,100000,6,2,7)", 259639.4166667)]
    [InlineData("=DB(1000000,100000,6,3,7)", 176814.44275)]
    [InlineData("=DB(1000000,100000,6,7,7)", 15845.0984738)]
    [InlineData("=DDB(2400,300,10*365,1)", 1.3150685)]
    [InlineData("=DDB(2400,300,10*12,1,2)", 40)]
    [InlineData("=DDB(2400,300,10,1,2)", 480)]
    [InlineData("=DDB(2400,300,10,2,1.5)", 306)]
    [InlineData("=DDB(2400,300,10,10)", 22.1225472)]
    [InlineData("=VDB(2400,300,10*365,0,1)", 1.3150685)]
    [InlineData("=VDB(2400,300,10*12,0,1)", 40)]
    [InlineData("=VDB(2400,300,10,0,1)", 480)]
    [InlineData("=VDB(2400,300,10*12,6,18)", 396.3060533)]
    [InlineData("=VDB(2400,300,10*12,6,18,1.5)", 311.8089367)]
    [InlineData("=VDB(2400,300,10,0,0.875,1.5)", 315)]
    [InlineData("=VDB(2400,0,10,0,10)", 2400)]
    [InlineData("=VDB(2400,0,10,0,10,2,TRUE)", 2142.3019622)]
    [InlineData("=EFFECT(0.0525,4)", 0.0535427)]
    [InlineData("=NOMINAL(0.053543,4)", 0.0525003)]
    [InlineData("=CUMIPMT(0.09/12,30*12,125000,13,24,0)", -11135.2321307)]
    [InlineData("=CUMIPMT(0.09/12,30*12,125000,1,1,0)", -937.5)]
    [InlineData("=CUMPRINC(0.09/12,30*12,125000,13,24,0)", -934.1071234)]
    [InlineData("=CUMPRINC(0.09/12,30*12,125000,1,1,0)", -68.2782712)]
    [InlineData("=FVSCHEDULE(1,{0.09,0.11,0.1})", 1.33089)]
    [InlineData("=PDURATION(2.5%,2000,2200)", 3.8598661)]
    [InlineData("=RRI(96,10000,11000)", 0.0009933)]
    [InlineData("=ISPMT(0.1/12,0,3*12,8000000)", -66666.6666667)]
    [InlineData("=ISPMT(0.1,1,3,8000000)", -533333.3333333)]
    [InlineData("=DOLLARDE(1.02,16)", 1.125)]
    [InlineData("=DOLLARDE(1.1,32)", 1.3125)]
    [InlineData("=DOLLARDE(-1.02,16)", -1.125)]
    [InlineData("=DOLLARFR(1.125,16)", 1.02)]
    [InlineData("=DOLLARFR(1.125,32)", 1.04)]
    public void Financial_functions_give_Excels_values(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 6);
    }

    [Fact]
    public void Npv_and_irr_read_ranges_and_skip_text()
    {
        (string, CellValue)[] cells =
        [
            ("A1", -70000), ("A2", 12000), ("A3", 15000), ("A4", 18000), ("A5", 21000), ("A6", 26000), ("A7", "note"),
        ];

        Assert.Equal(0.0866309, Formula.Number("=IRR(A1:A6)", cells), 6);
        Assert.Equal(0.0866309, Formula.Number("=IRR(A1:A7)", cells), 6);
        Assert.Equal(-0.0212448, Formula.Number("=IRR(A1:A5)", cells), 6);
        Assert.Equal(-0.4435069, Formula.Number("=IRR(A1:A3,-10%)", cells), 6);
        Assert.Equal(1922.0615549, Formula.Number("=NPV(0.08,A2:A6)+A1", ("A1", -40000), ("A2", 8000), ("A3", 9200), ("A4", 10000), ("A5", 12000), ("A6", 14500)), 6);
    }

    [Fact]
    public void Mirr_uses_finance_and_reinvestment_rates()
    {
        (string, CellValue)[] cells =
        [
            ("A1", -120000), ("A2", 39000), ("A3", 30000), ("A4", 21000), ("A5", 37000), ("A6", 46000),
        ];

        Assert.Equal(0.1260941, Formula.Number("=MIRR(A1:A6,10%,12%)", cells), 6);
        Assert.Equal(-0.0480446, Formula.Number("=MIRR(A1:A4,10%,12%)", cells), 6);
        Assert.Equal(0.1347591, Formula.Number("=MIRR(A1:A6,10%,14%)", cells), 6);
    }

    [Fact]
    public void Xnpv_and_xirr_take_dates_as_serials()
    {
        (string, CellValue)[] cells =
        [
            ("A1", -10000), ("A2", 2750), ("A3", 4250), ("A4", 3250), ("A5", 2750),
            ("B1", new DateTime(2008, 1, 1)), ("B2", new DateTime(2008, 3, 1)), ("B3", new DateTime(2008, 10, 30)), ("B4", new DateTime(2009, 2, 15)), ("B5", new DateTime(2009, 4, 1)),
        ];

        Assert.Equal(2086.6476020, Formula.Number("=XNPV(0.09,A1:A5,B1:B5)", cells), 6);
        Assert.Equal(0.3733625335, Formula.Number("=XIRR(A1:A5,B1:B5)", cells), 6);
        Assert.Equal(0.3733625335, Formula.Number("=XIRR(A1:A5,B1:B5,0.5)", cells), 6);
    }

    [Theory]
    [InlineData("=IPMT(0.1,0,3,8000)", "#NUM!")]
    [InlineData("=PPMT(0.1,4,3,8000)", "#NUM!")]
    [InlineData("=PMT(\"x\",10,1000)", "#VALUE!")]
    [InlineData("=NPER(0,0,1000)", "#DIV/0!")]
    [InlineData("=RATE(10,100,1000)", "#NUM!")]
    [InlineData("=RATE(0,-100,1000)", "#NUM!")]
    [InlineData("=NPV(-1,100)", "#DIV/0!")]
    [InlineData("=IRR({100,200})", "#NUM!")]
    [InlineData("=IRR({-100,#N/A})", "#VALUE!")]
    [InlineData("=XNPV(0.1,{-100,50,60},{40000,40100})", "#NUM!")]
    [InlineData("=XNPV(0.1,{-100,50},{40000,39000})", "#NUM!")]
    [InlineData("=XNPV(0.1,{-100,\"x\"},{40000,40100})", "#NUM!")]
    [InlineData("=XIRR({100,50},{40000,40100})", "#NUM!")]
    [InlineData("=MIRR({100,200},0.1,0.1)", "#DIV/0!")]
    [InlineData("=SLN(100,10,0)", "#DIV/0!")]
    [InlineData("=SYD(100,10,5,6)", "#NUM!")]
    [InlineData("=DB(100,10,5,7)", "#NUM!")]
    [InlineData("=DB(100,10,5,1,13)", "#NUM!")]
    [InlineData("=DB(-100,10,5,1)", "#NUM!")]
    [InlineData("=DDB(100,10,5,6)", "#NUM!")]
    [InlineData("=DDB(100,10,5,1,0)", "#NUM!")]
    [InlineData("=VDB(100,10,5,3,2)", "#NUM!")]
    [InlineData("=VDB(100,10,5,0,6)", "#NUM!")]
    [InlineData("=EFFECT(0,4)", "#NUM!")]
    [InlineData("=EFFECT(0.05,0.5)", "#NUM!")]
    [InlineData("=NOMINAL(-0.05,4)", "#NUM!")]
    [InlineData("=CUMIPMT(0.01,36,1000,1,12,2)", "#NUM!")]
    [InlineData("=CUMIPMT(0.01,36,1000,13,12,0)", "#NUM!")]
    [InlineData("=CUMPRINC(0,36,1000,1,12,0)", "#NUM!")]
    [InlineData("=FVSCHEDULE(1,{0.1,\"x\"})", "#VALUE!")]
    [InlineData("=PDURATION(0,100,200)", "#NUM!")]
    [InlineData("=RRI(0,100,200)", "#NUM!")]
    [InlineData("=ISPMT(0.1,1,0,1000)", "#DIV/0!")]
    [InlineData("=DOLLARDE(1.02,0)", "#DIV/0!")]
    [InlineData("=DOLLARFR(1.02,-1)", "#NUM!")]
    public void Financial_errors_are_Excels_errors(string formula, string expected)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(expected, code);
    }
}
