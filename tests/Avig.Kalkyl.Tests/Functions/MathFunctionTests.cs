namespace Avig.Kalkyl.Tests;

public class MathFunctionTests
{
    private static readonly (string Address, CellValue Value)[] Incomes =
    [
        ("A1", 100000), ("A2", 200000), ("A3", 300000), ("A4", 400000),
        ("B1", 7000), ("B2", 14000), ("B3", 21000), ("B4", 28000),
        ("C1", "apples"), ("C2", "avocado"), ("C3", "banana"), ("C4", "apricot"),
    ];

    private static readonly (string Address, CellValue Value)[] Subtotals =
    [
        ("A1", 120), ("A2", 10), ("A3", 150), ("A4", "text"),
    ];

    private static readonly (string Address, CellValue Value)[] WithErrors =
    [
        ("A1", CellValue.Error("#DIV/0!")), ("A2", 72), ("A3", 30), ("A4", CellValue.Error("#NUM!")), ("A5", 31), ("A6", 96),
        ("A7", 32), ("A8", 81), ("A9", 33), ("A10", 53), ("A11", 34),
    ];

    private static readonly (string Address, CellValue Value)[] Coefficients =
    [
        ("A1", 1), ("A2", -0.5), ("A3", 1.0 / 24), ("A4", -1.0 / 720),
    ];

    [Theory]
    [InlineData("=ACOS(-0.5)", 2.094395102393195)]
    [InlineData("=ACOSH(10)", 2.993222846126381)]
    [InlineData("=ACOT(2)", 0.4636476090008061)]
    [InlineData("=ACOTH(6)", 0.1682361183106064)]
    [InlineData("=ASIN(-0.5)", -0.5235987755982989)]
    [InlineData("=ASINH(10)", 2.99822295029797)]
    [InlineData("=ATAN(1)", 0.7853981633974483)]
    [InlineData("=ATAN2(1,1)", 0.7853981633974483)]
    [InlineData("=ATAN2(-1,-1)", -2.356194490192345)]
    [InlineData("=ATAN2(1,0)", 0)]
    [InlineData("=ATANH(0.5)", 0.5493061443340549)]
    [InlineData("=COS(PI()/3)", 0.5)]
    [InlineData("=COSH(4)", 27.30823283601649)]
    [InlineData("=COT(30)", -0.1561199522)]
    [InlineData("=COTH(2)", 1.037314720727548)]
    [InlineData("=CSC(15)", 1.537780562)]
    [InlineData("=CSCH(1.5)", 0.4696424406)]
    [InlineData("=SEC(45)", 1.903594407)]
    [InlineData("=SECH(1)", 0.6480542736638855)]
    [InlineData("=SIN(PI()/2)", 1)]
    [InlineData("=SINH(1)", 1.175201193643801)]
    [InlineData("=TAN(PI()/4)", 1)]
    [InlineData("=TANH(-2)", -0.9640275800758169)]
    [InlineData("=DEGREES(PI())", 180)]
    [InlineData("=RADIANS(270)", 4.71238898038469)]
    [InlineData("=PI()", Math.PI)]
    public void Trigonometry_matches_Excel(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=ACOS(2)", "#NUM!")]
    [InlineData("=ACOSH(0.5)", "#NUM!")]
    [InlineData("=ACOTH(1)", "#NUM!")]
    [InlineData("=ATANH(1)", "#NUM!")]
    [InlineData("=ATAN2(0,0)", "#DIV/0!")]
    [InlineData("=COT(0)", "#DIV/0!")]
    [InlineData("=CSC(0)", "#DIV/0!")]
    [InlineData("=COTH(0)", "#DIV/0!")]
    [InlineData("=CSCH(0)", "#DIV/0!")]
    [InlineData("=SIN(2^27)", "#NUM!")]
    [InlineData("=COS(\"x\")", "#VALUE!")]
    public void Trigonometry_errors(string formula, string error)
    {
        AssertError(formula, error);
    }

    [Theory]
    [InlineData("=CEILING(2.5,1)", 3)]
    [InlineData("=CEILING(-2.5,-2)", -4)]
    [InlineData("=CEILING(-2.5,2)", -2)]
    [InlineData("=CEILING(1.5,0.1)", 1.5)]
    [InlineData("=CEILING(0.234,0.01)", 0.24)]
    [InlineData("=CEILING(1.1,0.1)", 1.1)]
    [InlineData("=CEILING(5,0)", 0)]
    [InlineData("=CEILING.MATH(24.3,5)", 25)]
    [InlineData("=CEILING.MATH(6.7)", 7)]
    [InlineData("=CEILING.MATH(-8.1,2)", -8)]
    [InlineData("=CEILING.MATH(-5.5,2,-1)", -6)]
    [InlineData("=CEILING.PRECISE(4.3)", 5)]
    [InlineData("=CEILING.PRECISE(-4.3)", -4)]
    [InlineData("=CEILING.PRECISE(4.3,-2)", 6)]
    [InlineData("=CEILING.PRECISE(-4.3,-2)", -4)]
    [InlineData("=ISO.CEILING(4.3)", 5)]
    [InlineData("=ISO.CEILING(-4.3,2)", -4)]
    [InlineData("=FLOOR(3.7,2)", 2)]
    [InlineData("=FLOOR(-2.5,-2)", -2)]
    [InlineData("=FLOOR(-2.5,2)", -4)]
    [InlineData("=FLOOR(1.58,0.1)", 1.5)]
    [InlineData("=FLOOR(0.234,0.01)", 0.23)]
    [InlineData("=FLOOR.MATH(24.3,5)", 20)]
    [InlineData("=FLOOR.MATH(6.7)", 6)]
    [InlineData("=FLOOR.MATH(-8.1,2)", -10)]
    [InlineData("=FLOOR.MATH(-5.5,2,-1)", -4)]
    [InlineData("=FLOOR.PRECISE(-3.2,-1)", -4)]
    [InlineData("=FLOOR.PRECISE(3.2,1)", 3)]
    [InlineData("=FLOOR.PRECISE(-3.2,1)", -4)]
    [InlineData("=FLOOR.PRECISE(3.2)", 3)]
    [InlineData("=INT(8.9)", 8)]
    [InlineData("=INT(-8.9)", -9)]
    [InlineData("=TRUNC(8.9)", 8)]
    [InlineData("=TRUNC(-8.9)", -8)]
    [InlineData("=TRUNC(0.45)", 0)]
    [InlineData("=TRUNC(3.14159,2)", 3.14)]
    [InlineData("=ROUNDDOWN(3.2,0)", 3)]
    [InlineData("=ROUNDDOWN(76.9,0)", 76)]
    [InlineData("=ROUNDDOWN(3.14159,3)", 3.141)]
    [InlineData("=ROUNDDOWN(-3.14159,1)", -3.1)]
    [InlineData("=ROUNDDOWN(31415.92654,-2)", 31400)]
    [InlineData("=ROUNDUP(3.2,0)", 4)]
    [InlineData("=ROUNDUP(76.9,0)", 77)]
    [InlineData("=ROUNDUP(3.14159,3)", 3.142)]
    [InlineData("=ROUNDUP(-3.14159,1)", -3.2)]
    [InlineData("=ROUNDUP(31415.92654,-2)", 31500)]
    [InlineData("=MROUND(10,3)", 9)]
    [InlineData("=MROUND(-10,-3)", -9)]
    [InlineData("=MROUND(1.3,0.2)", 1.4)]
    [InlineData("=EVEN(1.5)", 2)]
    [InlineData("=EVEN(3)", 4)]
    [InlineData("=EVEN(2)", 2)]
    [InlineData("=EVEN(-1)", -2)]
    [InlineData("=ODD(1.5)", 3)]
    [InlineData("=ODD(3)", 3)]
    [InlineData("=ODD(2)", 3)]
    [InlineData("=ODD(-1)", -1)]
    [InlineData("=ODD(-2)", -3)]
    [InlineData("=ODD(0)", 1)]
    [InlineData("=SIGN(10)", 1)]
    [InlineData("=SIGN(4-4)", 0)]
    [InlineData("=SIGN(-0.00001)", -1)]
    public void Rounding_matches_Excel(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 12);
    }

    [Theory]
    [InlineData("=CEILING(2.5,-2)", "#NUM!")]
    [InlineData("=FLOOR(2.5,-2)", "#NUM!")]
    [InlineData("=FLOOR(5,0)", "#DIV/0!")]
    [InlineData("=MROUND(5,-2)", "#NUM!")]
    [InlineData("=ROUNDUP(\"x\",1)", "#VALUE!")]
    public void Rounding_errors(string formula, string error)
    {
        AssertError(formula, error);
    }

    [Theory]
    [InlineData("=EXP(1)", Math.E)]
    [InlineData("=EXP(2)", 7.38905609893065)]
    [InlineData("=LN(86)", 4.454347296253507)]
    [InlineData("=LOG(10)", 1)]
    [InlineData("=LOG(8,2)", 3)]
    [InlineData("=LOG(1000,10)", 3)]
    [InlineData("=LOG10(86)", 1.934498451243568)]
    [InlineData("=LOG10(1E5)", 5)]
    [InlineData("=POWER(5,2)", 25)]
    [InlineData("=POWER(4,5/4)", 5.656854249492381)]
    [InlineData("=POWER(-8,1/3)", -2)]
    [InlineData("=POWER(-2,3)", -8)]
    [InlineData("=SQRT(16)", 4)]
    [InlineData("=SQRTPI(1)", 1.772453850905516)]
    [InlineData("=SQRTPI(2)", 2.506628274631)]
    [InlineData("=MOD(3,2)", 1)]
    [InlineData("=MOD(-3,2)", 1)]
    [InlineData("=MOD(3,-2)", -1)]
    [InlineData("=MOD(-3,-2)", -1)]
    [InlineData("=MOD(5.5,2)", 1.5)]
    [InlineData("=QUOTIENT(5,2)", 2)]
    [InlineData("=QUOTIENT(4.5,3.1)", 1)]
    [InlineData("=QUOTIENT(-10,3)", -3)]
    [InlineData("=GCD(5,2)", 1)]
    [InlineData("=GCD(24,36)", 12)]
    [InlineData("=GCD(5,0)", 5)]
    [InlineData("=GCD(24.9,36)", 12)]
    [InlineData("=LCM(5,2)", 10)]
    [InlineData("=LCM(24,36)", 72)]
    [InlineData("=LCM(3,0)", 0)]
    [InlineData("=FACT(5)", 120)]
    [InlineData("=FACT(1.9)", 1)]
    [InlineData("=FACT(0)", 1)]
    [InlineData("=FACTDOUBLE(6)", 48)]
    [InlineData("=FACTDOUBLE(7)", 105)]
    [InlineData("=FACTDOUBLE(-1)", 1)]
    [InlineData("=COMBIN(8,2)", 28)]
    [InlineData("=COMBIN(10,0)", 1)]
    [InlineData("=COMBINA(4,3)", 20)]
    [InlineData("=COMBINA(10,3)", 220)]
    [InlineData("=MULTINOMIAL(2,3,4)", 1260)]
    public void Powers_and_integers_match_Excel(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=LN(0)", "#NUM!")]
    [InlineData("=LOG(-1)", "#NUM!")]
    [InlineData("=LOG(5,1)", "#DIV/0!")]
    [InlineData("=LOG10(0)", "#NUM!")]
    [InlineData("=POWER(0,0)", "#NUM!")]
    [InlineData("=POWER(0,-1)", "#DIV/0!")]
    [InlineData("=POWER(-4,0.5)", "#NUM!")]
    [InlineData("=SQRT(-16)", "#NUM!")]
    [InlineData("=SQRTPI(-1)", "#NUM!")]
    [InlineData("=MOD(1,0)", "#DIV/0!")]
    [InlineData("=QUOTIENT(1,0)", "#DIV/0!")]
    [InlineData("=GCD(-1,2)", "#NUM!")]
    [InlineData("=LCM(-1,2)", "#NUM!")]
    [InlineData("=FACT(-1)", "#NUM!")]
    [InlineData("=FACT(171)", "#NUM!")]
    [InlineData("=FACTDOUBLE(-2)", "#NUM!")]
    [InlineData("=COMBIN(2,3)", "#NUM!")]
    [InlineData("=COMBIN(-1,1)", "#NUM!")]
    [InlineData("=COMBINA(-1,2)", "#NUM!")]
    [InlineData("=MULTINOMIAL(1,-1)", "#NUM!")]
    [InlineData("=EXP(1000)", "#NUM!")]
    public void Power_and_integer_errors(string formula, string error)
    {
        AssertError(formula, error);
    }

    [Theory]
    [InlineData("=PRODUCT(5,15,30)", 2250)]
    [InlineData("=PRODUCT({5,15,30},2)", 4500)]
    [InlineData("=SUMSQ(3,4)", 25)]
    [InlineData("=SUMPRODUCT({3,4;8,6;1,9},{2,7;6,7;5,3})", 156)]
    [InlineData("=SUMPRODUCT({1,\"a\"},{2,3})", 2)]
    [InlineData("=SUMPRODUCT({1,2,3})", 6)]
    [InlineData("=SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", -55)]
    [InlineData("=SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 521)]
    [InlineData("=SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})", 79)]
    public void Sums_and_products_match_Excel(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=SUMPRODUCT({1,2},{1,2,3})", "#VALUE!")]
    [InlineData("=SUMPRODUCT({1,#N/A},{1,2})", "#N/A")]
    [InlineData("=SUMX2MY2({1,2},{1,2,3})", "#N/A")]
    [InlineData("=SERIESSUM(1,0,1,{1,\"a\"})", "#VALUE!")]
    public void Sum_errors(string formula, string error)
    {
        AssertError(formula, error);
    }

    [Fact]
    public void Product_skips_text_in_ranges_and_is_zero_without_numbers()
    {
        Assert.Equal(2250, Formula.Number("=PRODUCT(A1:A4)", ("A1", 5), ("A2", 15), ("A3", 30), ("A4", "x")));
        Assert.Equal(0, Formula.Number("=PRODUCT(A1:A2)"));
    }

    [Fact]
    public void Seriessum_approximates_the_cosine()
    {
        Assert.Equal(0.707103215, Formula.Number("=SERIESSUM(PI()/4,0,2,A1:A4)", Coefficients), 6);
    }

    [Theory]
    [InlineData("=SUMIF(A1:A4,\">160000\",B1:B4)", 63000)]
    [InlineData("=SUMIF(A1:A4,\">160000\")", 900000)]
    [InlineData("=SUMIF(A1:A4,300000,B1:B4)", 21000)]
    [InlineData("=SUMIF(A1:A4,\">160000\",B1)", 63000)]
    [InlineData("=SUMIF(C1:C4,\"a*\",B1:B4)", 49000)]
    [InlineData("=SUMIF(C1:C4,\"banana\",B1:B2)", 21000)]
    [InlineData("=SUMIF(A1:A4,\"<>200000\",B1:B4)", 56000)]
    [InlineData("=SUMIFS(B1:B4,A1:A4,\">150000\",A1:A4,\"<350000\")", 35000)]
    [InlineData("=SUMIFS(B1:B4,C1:C4,\"a*\",A1:A4,\">100000\")", 42000)]
    [InlineData("=SUMIFS(B1:B4,C1:C4,\"cherry\")", 0)]
    public void Conditional_sums_follow_the_criteria(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Incomes), 9);
    }

    [Fact]
    public void Sumifs_with_ranges_of_different_shapes_is_a_value_error()
    {
        CellValue value = Formula.Evaluate("=SUMIFS(B1:B3,A1:A4,\">0\")", Incomes);

        Assert.True(value.TryGetError(out string? code));
        Assert.Equal("#VALUE!", code);
    }

    [Fact]
    public void Rand_and_randbetween_stay_in_their_ranges()
    {
        for (int attempt = 0; attempt < 20; attempt++)
        {
            double random = Formula.Number("=RAND()");
            double whole = Formula.Number("=RANDBETWEEN(1,10)");

            Assert.InRange(random, 0, 0.9999999999999999);
            Assert.InRange(whole, 1, 10);
            Assert.Equal(Math.Floor(whole), whole);
        }

        Assert.Equal(-3, Formula.Number("=RANDBETWEEN(-3,-3)"));
        AssertError("=RANDBETWEEN(5,1)", "#NUM!");
    }

    [Theory]
    [InlineData("=SUBTOTAL(9,A1:A4)", 280)]
    [InlineData("=SUBTOTAL(109,A1:A4)", 280)]
    [InlineData("=SUBTOTAL(1,A1:A4)", 93.333333333333)]
    [InlineData("=SUBTOTAL(2,A1:A4)", 3)]
    [InlineData("=SUBTOTAL(3,A1:A4)", 4)]
    [InlineData("=SUBTOTAL(4,A1:A4)", 150)]
    [InlineData("=SUBTOTAL(5,A1:A4)", 10)]
    [InlineData("=SUBTOTAL(6,A1:A4)", 180000)]
    [InlineData("=SUBTOTAL(7,A1:A4)", 73.711147958319)]
    [InlineData("=SUBTOTAL(8,A1:A4)", 60.184900284226)]
    [InlineData("=SUBTOTAL(10,A1:A4)", 5433.333333333)]
    [InlineData("=SUBTOTAL(11,A1:A4)", 3622.222222222)]
    [InlineData("=SUBTOTAL(9,A1:A2,A3)", 280)]
    public void Subtotal_computes_each_function(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Subtotals), 6);
    }

    [Theory]
    [InlineData("=SUBTOTAL(12,A1:A4)", "#VALUE!")]
    [InlineData("=SUBTOTAL(0,A1:A4)", "#VALUE!")]
    [InlineData("=SUBTOTAL(1,B1:B2)", "#DIV/0!")]
    [InlineData("=SUBTOTAL(7,A1)", "#DIV/0!")]
    public void Subtotal_errors(string formula, string error)
    {
        AssertError(formula, error, Subtotals);
    }

    [Fact]
    public void Subtotal_skips_cells_that_hold_a_subtotal()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 1;
        sheet["A2"] = 2;
        sheet.SetFormula(2, 0, "SUBTOTAL(9,A1:A2)");

        Assert.True(book.Evaluate("=SUBTOTAL(9,A1:A3)", sheet).IsSuccess(out CellValue value));
        Assert.True(value.TryGetNumber(out double number));
        Assert.Equal(3, number);
    }

    [Theory]
    [InlineData("=AGGREGATE(4,6,A1:A11)", 96)]
    [InlineData("=AGGREGATE(14,6,A1:A11,3)", 72)]
    [InlineData("=AGGREGATE(9,6,A1:A11)", 462)]
    [InlineData("=AGGREGATE(9,7,A1:A11)", 462)]
    [InlineData("=AGGREGATE(12,6,A1:A11)", 34)]
    [InlineData("=AGGREGATE(5,6,A1:A11)", 30)]
    [InlineData("=AGGREGATE(1,6,A1:A11)", 51.333333333333)]
    [InlineData("=AGGREGATE(2,0,A1:A11)", 9)]
    [InlineData("=AGGREGATE(3,6,A1:A11)", 9)]
    [InlineData("=AGGREGATE(3,0,A1:A11)", 11)]
    [InlineData("=AGGREGATE(9,6,A2:A3,A5)", 133)]
    public void Aggregate_ignores_errors_when_asked(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, WithErrors), 9);
    }

    [Theory]
    [InlineData("=AGGREGATE(4,0,A1:A11)", "#DIV/0!")]
    [InlineData("=AGGREGATE(9,4,A1:A11)", "#DIV/0!")]
    [InlineData("=AGGREGATE(13,6,A1:A11)", "#N/A")]
    [InlineData("=AGGREGATE(14,6,A1:A11)", "#VALUE!")]
    [InlineData("=AGGREGATE(14,6,A1:A11,20)", "#NUM!")]
    [InlineData("=AGGREGATE(0,6,A1:A11)", "#VALUE!")]
    [InlineData("=AGGREGATE(9,8,A1:A11)", "#VALUE!")]
    public void Aggregate_errors(string formula, string error)
    {
        AssertError(formula, error, WithErrors);
    }

    [Fact]
    public void Aggregate_mode_finds_the_most_frequent_number()
    {
        Assert.Equal(2, Formula.Number("=AGGREGATE(13,0,{1,2,2,3,3})"));
    }

    [Theory]
    [InlineData("=ROMAN(499)", "CDXCIX")]
    [InlineData("=ROMAN(499,0)", "CDXCIX")]
    [InlineData("=ROMAN(499,1)", "LDVLIV")]
    [InlineData("=ROMAN(499,2)", "XDIX")]
    [InlineData("=ROMAN(499,3)", "VDIV")]
    [InlineData("=ROMAN(499,4)", "ID")]
    [InlineData("=ROMAN(499,TRUE)", "CDXCIX")]
    [InlineData("=ROMAN(499,FALSE)", "ID")]
    [InlineData("=ROMAN(1999)", "MCMXCIX")]
    [InlineData("=ROMAN(3999)", "MMMCMXCIX")]
    [InlineData("=ROMAN(2026.9)", "MMXXVI")]
    [InlineData("=ROMAN(0)", "")]
    [InlineData("=BASE(7,2)", "111")]
    [InlineData("=BASE(100,16)", "64")]
    [InlineData("=BASE(15,2,10)", "0000001111")]
    [InlineData("=BASE(0,2)", "0")]
    [InlineData("=BASE(35,36)", "Z")]
    public void Number_systems_give_text(string formula, string expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula));
    }

    [Theory]
    [InlineData("=ARABIC(\"LVII\")", 57)]
    [InlineData("=ARABIC(\"mcmxii\")", 1912)]
    [InlineData("=ARABIC(\"\")", 0)]
    [InlineData("=ARABIC(\"-MMXI\")", -2011)]
    [InlineData("=ARABIC(\" CDXCIX \")", 499)]
    [InlineData("=ARABIC(ROMAN(499,4))", 499)]
    [InlineData("=DECIMAL(\"FF\",16)", 255)]
    [InlineData("=DECIMAL(111,2)", 7)]
    [InlineData("=DECIMAL(\"zap\",36)", 45745)]
    public void Number_systems_give_numbers(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula));
    }

    [Theory]
    [InlineData("=ROMAN(4000)", "#VALUE!")]
    [InlineData("=ROMAN(-1)", "#VALUE!")]
    [InlineData("=ROMAN(5,5)", "#VALUE!")]
    [InlineData("=ARABIC(\"ABC\")", "#VALUE!")]
    [InlineData("=BASE(-1,2)", "#NUM!")]
    [InlineData("=BASE(5,1)", "#NUM!")]
    [InlineData("=BASE(5,37)", "#NUM!")]
    [InlineData("=BASE(5,2,256)", "#NUM!")]
    [InlineData("=DECIMAL(\"2\",2)", "#NUM!")]
    [InlineData("=DECIMAL(\"10\",37)", "#NUM!")]
    public void Number_system_errors(string formula, string error)
    {
        AssertError(formula, error);
    }

    private static void AssertError(string formula, string error, params (string Address, CellValue Value)[] cells)
    {
        CellValue value = Formula.Evaluate(formula, cells);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(error, code);
    }
}
