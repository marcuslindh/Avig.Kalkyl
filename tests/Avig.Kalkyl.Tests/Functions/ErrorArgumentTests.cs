namespace Avig.Kalkyl.Tests;

/// <summary>
/// Excel's rule for arguments: an error passed to a function comes back out of it, and text that is not
/// a number where a number belongs is #VALUE!. Every function is called once per argument with that
/// argument replaced, starting from a call that works. The exceptions below are the ones Excel 365 showed when every
/// one of these calls was evaluated in it.
/// </summary>
public class ErrorArgumentTests
{
    /// <summary>One call per function whose arguments are all constants, most taken from the other function tests.</summary>
    private static readonly string[] Calls =
    [
        "=ABS(-3)",
        "=ACOS(-0.5)",
        "=ACOSH(10)",
        "=ACOT(2)",
        "=ACOTH(6)",
        "=ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")",
        "=AND(TRUE,1,\"TRUE\")",
        "=ARABIC(\"LVII\")",
        "=ASIN(-0.5)",
        "=ASINH(10)",
        "=ATAN(1)",
        "=ATAN2(1,1)",
        "=ATANH(0.5)",
        "=AVEDEV(4,5,6,7,5,4,3)",
        "=AVERAGE(10,7,9,27,2)",
        "=BASE(15,2,10)",
        "=BETA.DIST(2,8,10,TRUE,1,3)",
        "=BIN2DEC(1100100)",
        "=BIN2HEX(11111011,4)",
        "=BIN2OCT(1001,3)",
        "=BINOM.DIST(6,10,0.5,FALSE)",
        "=BINOMDIST(6,10,0.5,TRUE)",
        "=BITAND(1,5)",
        "=BITLSHIFT(4,2)",
        "=BITOR(23,10)",
        "=BITRSHIFT(13,2)",
        "=BITXOR(5,3)",
        "=CEILING(2.5,1)",
        "=CEILING.MATH(-5.5,2,-1)",
        "=CEILING.PRECISE(4.3,-2)",
        "=CHAR(65)",
        "=CHISQ.DIST(0.5,1,TRUE)",
        "=CHISQ.DIST.RT(18.307,10)",
        "=CHISQ.INV(0.93,1)",
        "=CHISQ.INV.RT(0.050001,10)",
        "=CHOOSE(2,10,20,30)",
        "=CODE(\"A\")",
        "=COLUMNS({1,2,3})",
        "=COMBIN(8,2)",
        "=COMBINA(4,3)",
        "=CONCAT(\"x\",{1,2})",
        "=CONCATENATE(\"a\",1,TRUE)",
        "=CONFIDENCE(0.05,2.5,50)",
        "=CONFIDENCE.NORM(0.05,2.5,50)",
        "=CORREL({3,2,4,5,6},{9,7,12,15,17})",
        "=COS(0)",
        "=COSH(4)",
        "=COT(30)",
        "=COTH(2)",
        "=COVAR({3,2,4,5,6},{9,7,12,15,17})",
        "=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})",
        "=COVARIANCE.S({2,4,8},{5,11,12})",
        "=CSC(15)",
        "=CSCH(1.5)",
        "=CUMIPMT(0.0075,360,125000,13,24,0)",
        "=CUMPRINC(0.0075,360,125000,13,24,0)",
        "=DATE(2008,1,1)",
        "=DATEDIF(\"2001-01-01\",\"2003-01-01\",\"Y\")",
        "=DATEVALUE(\"2008-08-22\")",
        "=DAY(60)",
        "=DAYS(\"2021-03-15\",\"2021-02-01\")",
        "=DAYS360(\"2011-01-30\",\"2011-12-31\")",
        "=DB(1000000,100000,6,1,7)",
        "=DDB(2400,300,10,1,2)",
        "=DEC2BIN(9,4)",
        "=DEC2HEX(100,4)",
        "=DEC2OCT(58,3)",
        "=DECIMAL(\"FF\",16)",
        "=DELTA(5,4)",
        "=DEVSQ(4,5,8,7,11,4,3)",
        "=DOLLAR(1234.567,2)",
        "=DOLLARDE(1.02,16)",
        "=DOLLARFR(1.125,16)",
        "=EDATE(\"2011-01-15\",-1)",
        "=EFFECT(0.0525,4)",
        "=EOMONTH(40544,1)",
        "=ERF(1,2)",
        "=ERF.PRECISE(0.745)",
        "=ERFC(1)",
        "=ERFC.PRECISE(1)",
        "=ERROR.TYPE(#N/A)",
        "=EVEN(1.5)",
        "=EXACT(\"word\",\"word\")",
        "=EXP(1)",
        "=EXPON.DIST(0.2,10,TRUE)",
        "=EXPONDIST(0.2,10,TRUE)",
        "=F.DIST(15.2069,6,4,TRUE)",
        "=F.DIST.RT(15.2068649,6,4)",
        "=FACT(5)",
        "=FACTDOUBLE(6)",
        "=FIND(\"M\",\"Miriam McGovern\",3)",
        "=FISHER(0.75)",
        "=FISHERINV(0.972955)",
        "=FIXED(-1234.567,-1,TRUE)",
        "=FLOOR(3.7,2)",
        "=FLOOR.MATH(-5.5,2,-1)",
        "=FLOOR.PRECISE(-3.2,-1)",
        "=FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})",
        "=FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})",
        "=FV(0,12,-100,-1000)",
        "=FVSCHEDULE(1,{0.09,0.11,0.1})",
        "=GAMMA(2.5)",
        "=GAMMA.DIST(10.00001131,9,2,FALSE)",
        "=GAMMALN(4)",
        "=GAMMALN.PRECISE(4)",
        "=GAUSS(2)",
        "=GCD(5,2)",
        "=GEOMEAN(4,5,8,7,11,4,3)",
        "=GESTEP(5,4)",
        "=HARMEAN(4,5,8,7,11,4,3)",
        "=HEX2BIN(\"F\",8)",
        "=HEX2DEC(\"A5\")",
        "=HEX2OCT(\"F\",3)",
        "=HOUR(\"8:30\")",
        "=HYPERLINK(\"https://example.com\",\"Click\")",
        "=IF(FALSE,1)",
        "=IFNA(\"ok\",\"missing\")",
        "=IFS(FALSE,1,TRUE,2)",
        "=INDEX({1,2;3,4},2,2)",
        "=INDIRECT(\"C1\")",
        "=INT(8.9)",
        "=INTERCEPT({2,3,9,1,8},{6,5,11,7,5})",
        "=IPMT(0.1,1,3,8000,0,1)",
        "=IRR({-100,30,40,50})",
        "=ISBLANK(\"\")",
        "=ISERR(5)",
        "=ISEVEN(-1)",
        "=ISLOGICAL(1)",
        "=ISNUMBER(\"5\")",
        "=ISO.CEILING(-4.3,2)",
        "=ISODD(5)",
        "=ISOWEEKNUM(\"2012-03-09\")",
        "=ISPMT(0.1,1,3,8000000)",
        "=ISREF(\"A1\")",
        "=KURT(3,4,5,2,3,4,5,6,4,7)",
        "=LARGE({3,5,3,5,4,4,2,4,6,7},3)",
        "=LCM(5,2)",
        "=LEFT(\"Sale Price\",4)",
        "=LEN(\"Phoenix, AZ\")",
        "=LN(86)",
        "=LOG(8,2)",
        "=LOG10(86)",
        "=LOGNORM.DIST(4,3.5,1.2,TRUE)",
        "=LOGNORM.INV(0.039084,3.5,1.2)",
        "=LOOKUP(4.19,{4.14,4.19,5.17,5.77,6.39},{1,2,3,4,5})",
        "=LOWER(\"E. E. Cummings\")",
        "=MATCH(25,{10,20,30},1)",
        "=MEDIAN(1,2,3,4,5,6)",
        "=MID(\"Fluid Flow\",1,5)",
        "=MINUTE(\"12:45\")",
        "=MIRR({-120000,39000,30000,21000,37000,46000},0.1,0.12)",
        "=MOD(3,2)",
        "=MODE(1,2,2,3)",
        "=MODE.MULT(1,2,2,3)",
        "=MODE.SNGL({5.6,4,4,3,2,4})",
        "=MONTH(\"3/2/2013\")",
        "=MROUND(10,3)",
        "=MULTINOMIAL(2,3,4)",
        "=N(7)",
        "=NETWORKDAYS(1,10,{3})",
        "=NETWORKDAYS.INTL(1,10,1,{3})",
        "=NOMINAL(0.053543,4)",
        "=NORM.DIST(42,40,1.5,TRUE)",
        "=NORM.INV(0.908789,40,1.5)",
        "=NORM.S.DIST(1.333333,TRUE)",
        "=NORM.S.INV(0.908789)",
        "=NORMDIST(42,40,1.5,TRUE)",
        "=NORMINV(0.908789,40,1.5)",
        "=NORMSDIST(1.333333)",
        "=NORMSINV(0.908789)",
        "=NOT(0)",
        "=NPER(0,-100,1000)",
        "=NPV(0.1,-10000,3000,4200,6800)",
        "=NUMBERVALUE(\"2.500,27\",\",\",\".\")",
        "=OCT2BIN(3,3)",
        "=OCT2DEC(54)",
        "=OCT2HEX(100,4)",
        "=ODD(1.5)",
        "=OR(FALSE,0)",
        "=PDURATION(2.5%,2000,2200)",
        "=PEARSON({9,7,5,3,1},{10,6,1,5,3})",
        "=PERCENTILE({1,3,2,4},1)",
        "=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)",
        "=PERCENTILE.INC({1,3,2,4},0.3)",
        "=PERCENTRANK({13,12,11,8,4,3,2,1,1,1},4,1)",
        "=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)",
        "=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},2)",
        "=PERMUT(100,3)",
        "=PERMUTATIONA(3,2)",
        "=PHI(0.75)",
        "=PI()",
        "=PMT(0,10,1000)",
        "=POISSON(2,5,TRUE)",
        "=POISSON.DIST(2,5,TRUE)",
        "=POWER(5,2)",
        "=PPMT(0.08,10,10,200000)",
        "=PRODUCT(5,15,30)",
        "=PROPER(\"this is a TITLE\")",
        "=PV(0,10,100)",
        "=QUARTILE({1,2,4,7,8,9,10,12},1)",
        "=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)",
        "=QUARTILE.INC({1,2,4,7,8,9,10,12},4)",
        "=QUOTIENT(5,2)",
        "=RADIANS(270)",
        "=RAND()",
        "=RANDBETWEEN(1,10)",
        "=RATE(10,-100,1000)",
        "=REPLACE(\"abcdefghijk\",6,5,\"*\")",
        "=REPT(\"ab\",3)",
        "=RIGHT(\"Stock Number\",6)",
        "=ROMAN(499,0)",
        "=ROUND(2.675,2)",
        "=ROUNDDOWN(3.2,0)",
        "=ROUNDUP(3.2,0)",
        "=ROW()",
        "=ROWS({1,2;3,4;5,6})",
        "=RRI(96,10000,11000)",
        "=SEARCH(\"e\",\"Statements\",6)",
        "=SEC(45)",
        "=SECH(1)",
        "=SECOND(\"4:48:18 PM\")",
        "=SERIESSUM(1,0,1,{1,2,3})",
        "=SIGN(10)",
        "=SINH(1)",
        "=SKEW(3,4,5,2,3,4,5,6,4,7)",
        "=SKEW.P(3,4,5,2,3,4,5,6,4,7)",
        "=SLN(30000,7500,10)",
        "=SLOPE({2,\"x\",4,6},{1,9,2,3})",
        "=SMALL({3,4,5,2,3,4,6,4,7},4)",
        "=SQRT(16)",
        "=SQRTPI(1)",
        "=STANDARDIZE(42,40,1.5)",
        "=STDEV({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=STDEV.P({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=STDEV.S({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=STDEVP({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=STEYX({2,3,9,1,8},{6,5,11,7,5})",
        "=SUBSTITUTE(\"Quarter 1, 2008\",\"1\",\"2\",1)",
        "=SUM(1,2,3)",
        "=SUMPRODUCT({3,4;8,6;1,9},{2,7;6,7;5,3})",
        "=SUMSQ(3,4)",
        "=SUMX2MY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
        "=SUMX2PY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
        "=SUMXMY2({2,3,9,1,8,7,5},{6,5,11,7,5,4,4})",
        "=SWITCH(2,1,\"one\",2,\"two\",\"none\")",
        "=SYD(30000,7500,10,1)",
        "=T(\"Rainfall\")",
        "=T.DIST(60,1,TRUE)",
        "=T.DIST.2T(1.959999998,60)",
        "=T.DIST.RT(1.959999998,60)",
        "=T.INV(0.75,2)",
        "=T.INV.2T(0.546449,60)",
        "=TANH(-2)",
        "=TEXT(1234.5,\"0.00\")",
        "=TEXTAFTER(\"abc\",\"x\",1,0,1)",
        "=TEXTBEFORE(\"abc\",\"x\",1,0,1)",
        "=TEXTJOIN({\"-\",\"+\"},TRUE,\"a\",\"b\",\"c\",\"d\")",
        "=TIME(12,0,0)",
        "=TIMEVALUE(\"6:35 AM\")",
        "=TRIM(\"  First   Quarter  Earnings \")",
        "=TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)",
        "=TRUNC(3.14159,2)",
        "=TYPE(\"Smith\")",
        "=UNICHAR(66)",
        "=UNICODE(\"B\")",
        "=UPPER(\"total\")",
        "=VALUE(\"$1,000\")",
        "=VALUETOTEXT(\"a\",1)",
        "=VAR({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=VAR.P({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=VAR.S({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=VARP({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})",
        "=VDB(2400,0,10,0,10,2,TRUE)",
        "=WEEKDAY(\"2008-02-14\",2)",
        "=WEEKNUM(\"2012-03-09\",2)",
        "=WEIBULL.DIST(105,20,100,TRUE)",
        "=XIRR({-100,50,60},{40000,40100,40200})",
        "=XMATCH(3,{1,3,3,5},0,-1)",
        "=XNPV(0.1,{-100,50,60},{40000,40100,40200})",
        "=XOR(TRUE,TRUE,TRUE)",
        "=YEAR(43465)",
        "=YEARFRAC(1,400,1)",
        "=AVERAGEA(1,2,3)",
        "=CLEAN(\"text\")",
        "=DEGREES(1)",
        "=MAX(1,2,3)",
        "=MAXA(1,2,3)",
        "=MIN(1,2,3)",
        "=MINA(1,2,3)",
        "=RANK(2,{1,2,3})",
        "=RANK.AVG(2,{1,2,3},1)",
        "=RANK.EQ(2,{1,2,3},1)",
        "=RSQ({1,2,3},{2,4,7})",
        "=SIN(1)",
        "=STDEVA(1,2,3)",
        "=STDEVPA(1,2,3)",
        "=TAN(1)",
        "=TRANSPOSE({1,2})",
        "=VARA(1,2,3)",
        "=VARPA(1,2,3)",
        "=VLOOKUP(2,{1,\"a\";2,\"b\"},2,FALSE)",
        "=HLOOKUP(2,{1,2;\"a\",\"b\"},2,FALSE)",
        "=WORKDAY(1,5,{3})",
        "=WORKDAY.INTL(1,5,1,{3})",
        "=XLOOKUP(2,{1,2,3},{4,5,6},0,0,1)",
        "=AGGREGATE(14,6,{1,2,3},2)",
        "=TRUE()",
        "=FALSE()",
        "=NOW()",
        "=TODAY()",
        "=IFERROR(1,2)",
        "=ISERROR(1)",
        "=ISNA(1)",
        "=ISTEXT(1)",
        "=ISNONTEXT(1)",
        "=COUNT(1,2)",
        "=COUNTA(1,2)",
    ];

    /// <summary>Functions a call of constants cannot test: they need ranges, formulas in cells or a pivot table, or are always an error (NA). The other tests cover them.</summary>
    private static readonly HashSet<string> Uncalled =
    [
        "AVERAGEIF", "AVERAGEIFS", "COLUMN", "COUNTBLANK", "COUNTIF", "COUNTIFS", "MAXIFS", "MINIFS", "OFFSET", "SUBTOTAL", "SUMIF", "SUMIFS",
        "AREAS", "FORMULATEXT", "ISFORMULA", "NA", "GETPIVOTDATA",
        "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP",
    ];

    /// <summary>
    /// Where Excel does not return the error: the IS functions, IFERROR, IFNA, ERROR.TYPE, TYPE and the counts take errors as
    /// values; IF, IFS, CHOOSE, SWITCH and XLOOKUP's if_not_found only evaluate the argument that is used; VALUETOTEXT shows it.
    /// The weekend of NETWORKDAYS.INTL and WORKDAY.INTL turns an error into #VALUE! and #NUM! (FunctionEdgeTests).
    /// "NAME" is every argument, "NAME:n" only argument n.
    /// </summary>
    private static readonly HashSet<string> TakesErrors =
    [
        "CHOOSE:1", "CHOOSE:2", "CHOOSE:3", "COUNT", "COUNTA", "ERROR.TYPE", "IF", "IFERROR", "IFNA", "IFS", "ISBLANK", "ISERR", "ISERROR", "ISLOGICAL",
        "ISNA", "ISNONTEXT", "ISNUMBER", "ISREF", "ISTEXT", "SWITCH", "TYPE", "VALUETOTEXT", "XLOOKUP:3", "NETWORKDAYS.INTL:2", "WORKDAY.INTL:2",
    ];

    /// <summary>
    /// Where text in a number's place is not #VALUE!: the functions above, values that are looked up or compared (#N/A when
    /// not found), digits of another base (#NUM!), N (0), TEXT of text (the text) and the values of CHOOSE and CONCATENATE.
    /// AND, OR, XOR and MODE skip text written as an argument, as they skip it in a range.
    /// </summary>
    private static readonly HashSet<string> TakesText =
    [
        "BIN2DEC", "BIN2HEX", "BIN2OCT", "OCT2BIN", "OCT2DEC", "OCT2HEX", "CHOOSE:1", "CHOOSE:2", "CHOOSE:3", "CONCATENATE", "COUNT", "COUNTA", "ERROR.TYPE",
        "IF", "IFERROR", "IFS", "ISERR", "ISERROR", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISTEXT", "N", "SWITCH",
        "AND", "OR", "XOR", "MODE", "MODE.MULT", "MODE.SNGL",
        "HLOOKUP:0", "LOOKUP:0", "MATCH:0", "TEXT:0", "VLOOKUP:0", "XLOOKUP:0", "XLOOKUP:3", "XMATCH:0",
    ];

    public static TheoryData<string, int> ErrorCases()
    {
        return Cases(TakesErrors, argument => true);
    }

    public static TheoryData<string, int> TextCases()
    {
        return Cases(TakesText, Numeric);
    }

    [Theory]
    [MemberData(nameof(ErrorCases))]
    public void An_error_argument_comes_back_as_the_result(string call, int position)
    {
        string formula = Replace(call, position, "#DIV/0!");
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code) == true && code == "#DIV/0!", $"{formula} gave {value}.");
    }

    [Theory]
    [MemberData(nameof(TextCases))]
    public void Text_where_a_number_belongs_is_VALUE(string call, int position)
    {
        string formula = Replace(call, position, "\"abc\"");
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code) == true && code == "#VALUE!", $"{formula} gave {value}.");
    }

    [Fact]
    public void Every_call_works_before_an_argument_is_replaced()
    {
        List<string> failing = [.. Calls.Where(call => Formula.Evaluate(call).TryGetError(out _) == true)];

        Assert.True(failing.Count == 0, string.Join(" ", failing));
    }

    [Fact]
    public void Every_function_has_a_call()
    {
        HashSet<string> called = [.. Calls.Select(Name)];
        List<string> missing = [];

        foreach (string name in FunctionNames())
        {
            if (called.Contains(name) == false && Uncalled.Contains(name) == false)
            {
                missing.Add(name);
            }
        }

        Assert.Empty(missing);
    }

    private static TheoryData<string, int> Cases(HashSet<string> exceptions, Func<string, bool> include)
    {
        TheoryData<string, int> cases = [];

        foreach (string call in Calls)
        {
            List<string> arguments = Arguments(call);

            for (int position = 0; position < arguments.Count; position++)
            {
                string name = Name(call);

                if (include(arguments[position].Trim()) == true && exceptions.Contains(name) == false && exceptions.Contains($"{name}:{position}") == false)
                {
                    cases.Add(call, position);
                }
            }
        }

        return cases;
    }

    private static IEnumerable<string> FunctionNames()
    {
        return Avig.Kalkyl.Formulas.FunctionLibrary.Names;
    }

    private static bool Numeric(string argument)
    {
        return double.TryParse(argument.TrimEnd('%'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out _);
    }

    private static string Name(string call)
    {
        return call[1..call.IndexOf('(', StringComparison.Ordinal)];
    }

    private static string Replace(string call, int position, string replacement)
    {
        List<string> arguments = Arguments(call);
        arguments[position] = replacement;

        return $"={Name(call)}({string.Join(",", arguments)})";
    }

    /// <summary>The arguments of a call, split at the commas outside strings and arrays.</summary>
    private static List<string> Arguments(string call)
    {
        int open = call.IndexOf('(', StringComparison.Ordinal);
        string text = call[(open + 1)..^1];
        List<string> parts = [];

        if (text.Length == 0)
        {
            return parts;
        }

        int depth = 0;
        bool quoted = false;
        int start = 0;

        for (int index = 0; index < text.Length; index++)
        {
            char character = text[index];

            if (character == '"')
            {
                quoted = quoted == false;
            }
            else if (quoted == false && character == '{')
            {
                depth++;
            }
            else if (quoted == false && character == '}')
            {
                depth--;
            }
            else if (quoted == false && depth == 0 && character == ',')
            {
                parts.Add(text[start..index]);
                start = index + 1;
            }
        }

        parts.Add(text[start..]);

        return parts;
    }
}
