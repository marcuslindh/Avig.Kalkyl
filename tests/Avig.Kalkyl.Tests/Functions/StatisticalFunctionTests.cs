namespace Avig.Kalkyl.Tests;

public class StatisticalFunctionTests
{
    private static readonly (string, CellValue)[] Mixed =
    [
        ("A1", 10), ("A2", 7), ("A3", 9), ("A4", 27), ("A5", 2), ("A6", "text"), ("A7", true), ("A8", ""), ("A10", new DateTime(2026, 1, 2)),
    ];

    private static readonly (string, CellValue)[] Property =
    [
        ("A2", 100000), ("A3", 200000), ("A4", 300000), ("A5", 400000),
        ("B2", 7000), ("B3", 14000), ("B4", 21000), ("B5", 28000),
        ("C2", "east"), ("C3", "west"), ("C4", "east"), ("C5", "west"),
    ];

    private static readonly (string, CellValue)[] Grades =
    [
        ("A2", 89), ("A3", 93), ("A4", 96), ("A5", 85), ("A6", 91), ("A7", 88),
        ("B2", 1), ("B3", 2), ("B4", 2), ("B5", 3), ("B6", 1), ("B7", 1),
    ];

    private static readonly (string, CellValue)[] Ranks = [("A2", 7), ("A3", 3.5), ("A4", 3.5), ("A5", 1), ("A6", 2)];

    private static readonly (string, CellValue)[] Temperatures = [("A1", 89), ("A2", 88), ("A3", 92), ("A4", 101), ("A5", 94), ("A6", 97), ("A7", 95)];

    private static readonly (string, CellValue)[] Pairs =
    [
        ("A1", 2), ("A2", 3), ("A3", 9), ("A4", 1), ("A5", 8), ("A6", 7), ("A7", 5),
        ("B1", 6), ("B2", 5), ("B3", 11), ("B4", 7), ("B5", 5), ("B6", 4), ("B7", 4),
        ("C1", "skip"), ("D1", 3),
    ];


    [Theory]
    [InlineData("=AVERAGE(10,7,9,27,2)", 11)]
    [InlineData("=AVERAGE(A1:A10)", 7679.833333333333)]
    [InlineData("=AVERAGE(A1:A5,5)", 10)]
    [InlineData("=AVERAGE(\"4\",TRUE)", 2.5)]
    [InlineData("=AVERAGEA(A1:A7)", 8)]
    [InlineData("=COUNT(A1:A10)", 6)]
    [InlineData("=COUNT(A1:A10,\"1\",TRUE,\"x\",#N/A)", 8)]
    [InlineData("=COUNT(A:A)", 6)]
    [InlineData("=COUNTA(A1:A10)", 9)]
    [InlineData("=COUNTA(1,\"\",#N/A)", 3)]
    [InlineData("=COUNTBLANK(A1:A10)", 2)]
    [InlineData("=COUNTBLANK(B1:C3)", 6)]
    [InlineData("=COUNTBLANK(A:A)", 1048568)]
    [InlineData("=MAX(A1:A5,30)", 30)]
    [InlineData("=MAX(A6:A8)", 0)]
    [InlineData("=MAXA(A5:A7)", 2)]
    [InlineData("=MIN(A1:A5)", 2)]
    [InlineData("=MINA(A5:A7)", 0)]
    [InlineData("=MIN(\"-3\",A1:A5)", -3)]
    public void Averages_counts_and_extremes(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Mixed), 9);
    }

    [Theory]
    [InlineData("=AVERAGEIF(B2:B5,\"<23000\")", 14000)]
    [InlineData("=AVERAGEIF(A2:A5,\"<250000\")", 150000)]
    [InlineData("=AVERAGEIF(A2:A5,\">250000\",B2:B5)", 24500)]
    [InlineData("=AVERAGEIF(A2:A5,\">250000\",B2)", 24500)]
    [InlineData("=AVERAGEIFS(B2:B5,C2:C5,\"east\",A2:A5,\">100000\")", 21000)]
    [InlineData("=COUNTIF(C2:C5,\"east\")", 2)]
    [InlineData("=COUNTIF(C2:C5,\"*st\")", 4)]
    [InlineData("=COUNTIF(A2:A5,\">=\"&A3)", 3)]
    [InlineData("=COUNTIF(A1:A6,\"\")", 2)]
    [InlineData("=COUNTIF(A:A,\"<>100000\")", 1048575)]
    [InlineData("=COUNTIFS(C2:C5,\"west\",B2:B5,\">20000\")", 1)]
    [InlineData("=COUNTIFS(A:A,\"\",B:B,\"\")", 1048572)]
    [InlineData("=MAXIFS(B2:B5,C2:C5,\"east\")", 21000)]
    [InlineData("=MINIFS(B2:B5,C2:C5,\"west\")", 14000)]
    [InlineData("=MAXIFS(B2:B5,C2:C5,\"north\")", 0)]
    public void Conditional_aggregates(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Property), 9);
    }

    [Theory]
    [InlineData("=MAXIFS(A2:A7,B2:B7,1)", 91)]
    [InlineData("=MINIFS(A2:A7,B2:B7,1)", 88)]
    [InlineData("=MAXIFS(A2:A7,B2:B7,\">1\",A2:A7,\"<95\")", 93)]
    public void Maxifs_and_minifs_follow_the_documentation(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Grades), 9);
    }

    [Theory]
    [InlineData("=AVERAGEIF(A2:A5,\"<95000\")", "#DIV/0!")]
    [InlineData("=AVERAGEIFS(B2:B5,C2:C4,\"east\")", "#VALUE!")]
    [InlineData("=COUNTIFS(C2:C5,\"east\",A2:A4,\">1\")", "#VALUE!")]
    [InlineData("=MAXIFS(B2:B4,C2:C5,\"east\")", "#VALUE!")]
    [InlineData("=COUNTIFS(C2:C5,\"east\",A2:A5)", "#VALUE!")]
    [InlineData("=AVERAGE(A1:A1)", "#DIV/0!")]
    [InlineData("=AVERAGE(\"x\")", "#VALUE!")]
    [InlineData("=MAX(1,#N/A)", "#N/A")]
    public void Aggregate_errors(string formula, string expected)
    {
        AssertError(expected, Formula.Evaluate(formula, Property));
    }

    [Theory]
    [InlineData("=MEDIAN(1,2,3,4,5)", 3)]
    [InlineData("=MEDIAN(1,2,3,4,5,6)", 3.5)]
    [InlineData("=MODE.SNGL({5.6,4,4,3,2,4})", 4)]
    [InlineData("=MODE({1,2,2,3,3})", 2)]
    [InlineData("=LARGE({3,5,3,5,4,4,2,4,6,7},3)", 5)]
    [InlineData("=LARGE({3,5,3,5,4,4,2,4,6,7},7)", 4)]
    [InlineData("=SMALL({3,4,5,2,3,4,6,4,7},4)", 4)]
    [InlineData("=SMALL({1,4,8,3,7,12,54,8,23},2)", 3)]
    [InlineData("=SMALL({1,2,3},1.2)", 1)]
    [InlineData("=PERCENTILE.INC({1,3,2,4},0.3)", 1.9)]
    [InlineData("=PERCENTILE({1,3,2,4},1)", 4)]
    [InlineData("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)", 2.5)]
    [InlineData("=QUARTILE({1,2,4,7,8,9,10,12},1)", 3.5)]
    [InlineData("=QUARTILE.INC({1,2,4,7,8,9,10,12},4)", 12)]
    [InlineData("=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)", 15)]
    [InlineData("=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)", 43)]
    [InlineData("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},2)", 0.333)]
    [InlineData("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},4)", 0.555)]
    [InlineData("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},8)", 0.666)]
    [InlineData("=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},5)", 0.583)]
    [InlineData("=PERCENTRANK({13,12,11,8,4,3,2,1,1,1},4,1)", 0.5)]
    [InlineData("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)", 0.7)]
    [InlineData("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)", 0.381)]
    [InlineData("=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)", 0.3)]
    [InlineData("=TRIMMEAN({4,5,6,7,2,3,4,5,1,2,3},0.2)", 3.777777778)]
    public void Order_statistics(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Fact]
    public void Ranks_follow_the_documentation()
    {
        Assert.Equal(5, Formula.Number("=RANK.EQ(A2,A2:A6,1)", Ranks));
        Assert.Equal(4, Formula.Number("=RANK.EQ(A6,A2:A6)", Ranks));
        Assert.Equal(3, Formula.Number("=RANK.EQ(A3,A2:A6,1)", Ranks));
        Assert.Equal(2, Formula.Number("=RANK(3.5,A2:A6)", Ranks));
        Assert.Equal(2.5, Formula.Number("=RANK.AVG(3.5,A2:A6)", Ranks));
        Assert.Equal(4, Formula.Number("=RANK.AVG(94,A1:A7)", Temperatures));
        AssertError("#N/A", Formula.Evaluate("=RANK(4,A2:A6)", Ranks));
    }

    [Fact]
    public void Mode_mult_returns_every_mode_as_an_array()
    {
        Assert.Equal(1, Formula.Number("=MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1})"));
        Assert.Equal(3, Formula.Number("=COUNT(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}))"));
        Assert.Equal(6, Formula.Number("=SUM(MODE.MULT({1,2,3,4,3,2,1,2,3,5,6,1}))"));
    }

    [Theory]
    [InlineData("=MEDIAN(A1)", "#NUM!")]
    [InlineData("=MODE(1,2,3)", "#N/A")]
    [InlineData("=MODE.MULT(1,2,3)", "#N/A")]
    [InlineData("=LARGE({1,2},3)", "#NUM!")]
    [InlineData("=SMALL({1,2},0)", "#NUM!")]
    [InlineData("=PERCENTILE.INC({1,2},1.5)", "#NUM!")]
    [InlineData("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0)", "#NUM!")]
    [InlineData("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.01)", "#NUM!")]
    [InlineData("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},2)", "#NUM!")]
    [InlineData("=QUARTILE({1,2},5)", "#NUM!")]
    [InlineData("=QUARTILE.EXC({1,2,3},0)", "#NUM!")]
    [InlineData("=PERCENTRANK({1,2,3},4)", "#N/A")]
    [InlineData("=PERCENTRANK({1,2,3},2,0)", "#NUM!")]
    [InlineData("=TRIMMEAN({1,2,3},1)", "#NUM!")]
    public void Order_statistic_errors(string formula, string expected)
    {
        AssertError(expected, Formula.Evaluate(formula));
    }

    [Theory]
    [InlineData("=STDEV.S({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 27.46391572)]
    [InlineData("=STDEV({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 27.46391572)]
    [InlineData("=STDEV.P({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 26.05455814)]
    [InlineData("=STDEVP({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 26.05455814)]
    [InlineData("=VAR.S({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 754.2666667)]
    [InlineData("=VAR({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 754.2666667)]
    [InlineData("=VAR.P({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 678.84)]
    [InlineData("=VARP({1345,1301,1368,1322,1310,1370,1318,1350,1303,1299})", 678.84)]
    [InlineData("=STDEVA(A1:A3)", 1)]
    [InlineData("=STDEVPA(A1:A3)", 0.816496581)]
    [InlineData("=VARA(A1:A3)", 1)]
    [InlineData("=VARPA(A1:A3)", 0.666666667)]
    [InlineData("=STDEV(A1:A3)", 0.707106781)]
    [InlineData("=AVEDEV(4,5,6,7,5,4,3)", 1.020408163)]
    [InlineData("=DEVSQ(4,5,8,7,11,4,3)", 48)]
    [InlineData("=GEOMEAN(4,5,8,7,11,4,3)", 5.47698697)]
    [InlineData("=HARMEAN(4,5,8,7,11,4,3)", 5.028375962)]
    [InlineData("=KURT(3,4,5,2,3,4,5,6,4,7)", -0.151799637)]
    [InlineData("=SKEW(3,4,5,2,3,4,5,6,4,7)", 0.359543071)]
    [InlineData("=SKEW.P(3,4,5,2,3,4,5,6,4,7)", 0.303193339)]
    [InlineData("=STANDARDIZE(42,40,1.5)", 1.333333333)]
    public void Spread_and_shape(string formula, double expected)
    {
        (string, CellValue)[] cells = [("A1", 1), ("A2", 2), ("A3", "x")];

        AssertClose(expected, Formula.Number(formula, cells), formula);
    }

    [Theory]
    [InlineData("=CORREL({3,2,4,5,6},{9,7,12,15,17})", 0.997054486)]
    [InlineData("=PEARSON({9,7,5,3,1},{10,6,1,5,3})", 0.699379)]
    [InlineData("=COVAR({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    [InlineData("=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})", 5.2)]
    [InlineData("=COVARIANCE.S({2,4,8},{5,11,12})", 9.666666667)]
    [InlineData("=SLOPE(A1:A7,B1:B7)", 0.305556)]
    [InlineData("=INTERCEPT({2,3,9,1,8},{6,5,11,7,5})", 0.0483871)]
    [InlineData("=RSQ(A1:A7,B1:B7)", 0.05795)]
    [InlineData("=STEYX(A1:A7,B1:B7)", 3.305719)]
    [InlineData("=FORECAST(30,{6,7,9,15,21},{20,28,31,38,40})", 10.607253)]
    [InlineData("=FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40})", 10.607253)]
    public void Regression(string formula, double expected)
    {
        AssertClose(expected, Formula.Number(formula, Pairs), formula);
    }

    [Fact]
    public void Regression_skips_pairs_with_text_and_needs_equal_sizes()
    {
        Assert.Equal(2, Formula.Number("=SLOPE({2,\"x\",4,6},{1,9,2,3})"), 12);
        AssertError("#N/A", Formula.Evaluate("=SLOPE(A1:A2,C1:E1)", Pairs));
    }

    [Theory]
    [InlineData("=STDEV(1)", "#DIV/0!")]
    [InlineData("=VAR.P(A9)", "#DIV/0!")]
    [InlineData("=AVEDEV(A9)", "#NUM!")]
    [InlineData("=GEOMEAN(1,0)", "#NUM!")]
    [InlineData("=HARMEAN(-1,2)", "#NUM!")]
    [InlineData("=KURT(1,2,3)", "#DIV/0!")]
    [InlineData("=SKEW(1,1,1)", "#DIV/0!")]
    [InlineData("=STANDARDIZE(1,1,0)", "#NUM!")]
    [InlineData("=CORREL({1,2},{1,2,3})", "#N/A")]
    [InlineData("=CORREL({1,1},{1,2})", "#DIV/0!")]
    [InlineData("=COVARIANCE.S({1},{2})", "#DIV/0!")]
    [InlineData("=STEYX({1,2},{1,2})", "#DIV/0!")]
    [InlineData("=SLOPE({1,2},{3,3})", "#DIV/0!")]
    public void Spread_errors(string formula, string expected)
    {
        AssertError(expected, Formula.Evaluate(formula));
    }

    [Theory]
    [InlineData("=PERMUT(100,3)", 970200)]
    [InlineData("=PERMUT(3,2)", 6)]
    [InlineData("=PERMUT(4.9,2.1)", 12)]
    [InlineData("=PERMUTATIONA(3,2)", 9)]
    [InlineData("=PERMUTATIONA(2,2)", 4)]
    [InlineData("=FISHER(0.75)", 0.972955)]
    [InlineData("=FISHERINV(0.972955)", 0.75)]
    [InlineData("=GAMMA(2.5)", 1.329340388)]
    [InlineData("=GAMMA(-3.75)", 0.267866129)]
    [InlineData("=GAMMA(5)", 24)]
    [InlineData("=GAMMA(0.5)", 1.772453851)]
    [InlineData("=GAMMALN(4)", 1.791759469)]
    [InlineData("=GAMMALN.PRECISE(4)", 1.791759)]
    [InlineData("=GAMMALN(0.5)", 0.572364943)]
    [InlineData("=PHI(0.75)", 0.301137432)]
    [InlineData("=GAUSS(2)", 0.47724987)]
    [InlineData("=GAUSS(-2)", -0.47724987)]
    [InlineData("=NORM.DIST(42,40,1.5,TRUE)", 0.9087888)]
    [InlineData("=NORM.DIST(42,40,1.5,FALSE)", 0.10934)]
    [InlineData("=NORMDIST(42,40,1.5,TRUE)", 0.9087888)]
    [InlineData("=NORM.S.DIST(1.333333,TRUE)", 0.908788726)]
    [InlineData("=NORM.S.DIST(1.333333,FALSE)", 0.164010148)]
    [InlineData("=NORMSDIST(1.333333)", 0.908788726)]
    [InlineData("=NORM.S.DIST(1,TRUE)", 0.841344746)]
    [InlineData("=NORM.INV(0.908789,40,1.5)", 42.000002)]
    [InlineData("=NORMINV(0.908789,40,1.5)", 42.000002)]
    [InlineData("=NORM.S.INV(0.908789)", 1.3333347)]
    [InlineData("=NORMSINV(0.908789)", 1.3333347)]
    [InlineData("=NORM.S.INV(0.975)", 1.959963985)]
    [InlineData("=NORM.S.INV(0.001)", -3.090232306)]
    [InlineData("=LOGNORM.DIST(4,3.5,1.2,TRUE)", 0.0390836)]
    [InlineData("=LOGNORM.DIST(4,3.5,1.2,FALSE)", 0.0176176)]
    [InlineData("=LOGNORM.INV(0.039084,3.5,1.2)", 4.0000252)]
    [InlineData("=EXPON.DIST(0.2,10,TRUE)", 0.86466472)]
    [InlineData("=EXPON.DIST(0.2,10,FALSE)", 1.35335283)]
    [InlineData("=EXPONDIST(0.2,10,TRUE)", 0.86466472)]
    [InlineData("=POISSON.DIST(2,5,TRUE)", 0.124652019)]
    [InlineData("=POISSON.DIST(2,5,FALSE)", 0.084224337)]
    [InlineData("=POISSON(2,5,TRUE)", 0.124652019)]
    [InlineData("=BINOM.DIST(6,10,0.5,FALSE)", 0.205078125)]
    [InlineData("=BINOMDIST(6,10,0.5,TRUE)", 0.828125)]
    [InlineData("=CONFIDENCE.NORM(0.05,2.5,50)", 0.692951912)]
    [InlineData("=CONFIDENCE(0.05,2.5,50)", 0.692951912)]
    [InlineData("=WEIBULL.DIST(105,20,100,TRUE)", 0.929581)]
    [InlineData("=WEIBULL.DIST(105,20,100,FALSE)", 0.035589)]
    [InlineData("=GAMMA.DIST(10.00001131,9,2,FALSE)", 0.032639)]
    [InlineData("=GAMMA.DIST(10.00001131,9,2,TRUE)", 0.068094)]
    [InlineData("=BETA.DIST(2,8,10,TRUE,1,3)", 0.6854706)]
    [InlineData("=BETA.DIST(2,8,10,FALSE,1,3)", 1.4837646)]
    [InlineData("=T.DIST(60,1,TRUE)", 0.99469533)]
    [InlineData("=T.DIST(8,3,FALSE)", 0.00073691)]
    [InlineData("=T.DIST.2T(1.959999998,60)", 0.054644930)]
    [InlineData("=T.DIST.RT(1.959999998,60)", 0.027322465)]
    [InlineData("=T.INV(0.75,2)", 0.8164966)]
    [InlineData("=T.INV.2T(0.546449,60)", 0.606533)]
    [InlineData("=CHISQ.DIST(0.5,1,TRUE)", 0.52049988)]
    [InlineData("=CHISQ.DIST(2,3,FALSE)", 0.20755375)]
    [InlineData("=CHISQ.DIST.RT(18.307,10)", 0.0500006)]
    [InlineData("=CHISQ.INV(0.93,1)", 3.283020287)]
    [InlineData("=CHISQ.INV(0.6,2)", 1.832581464)]
    [InlineData("=CHISQ.INV.RT(0.050001,10)", 18.30697346)]
    [InlineData("=F.DIST(15.2069,6,4,TRUE)", 0.99)]
    [InlineData("=F.DIST(15.2069,6,4,FALSE)", 0.0012238)]
    [InlineData("=F.DIST.RT(15.2068649,6,4)", 0.01)]
    public void Distributions_match_Excels_documented_values(string formula, double expected)
    {
        AssertClose(expected, Formula.Number(formula), formula);
    }

    [Theory]
    [InlineData("=NORM.S.DIST(1,TRUE)", 0.841344746068543)]
    [InlineData("=NORM.S.DIST(-3,TRUE)", 0.00134989803163009)]
    [InlineData("=NORM.S.INV(0.841344746068543)", 1)]
    [InlineData("=NORM.S.INV(0.00134989803163009)", -3)]
    [InlineData("=NORM.S.INV(1E-10)", -6.36134090240406)]
    [InlineData("=T.DIST(T.INV(0.995,10),10,TRUE)", 0.995)]
    [InlineData("=T.INV(0.975,1)", 12.7062047361747)]
    [InlineData("=T.INV(0.75,2)", 0.816496580927726)]
    [InlineData("=CHISQ.INV(CHISQ.DIST(7.5,4,TRUE),4)", 7.5)]
    [InlineData("=NORM.S.DIST(-5,TRUE)*1E7", 2.86651571879194)]
    [InlineData("=CHISQ.DIST.RT(3.84145882069412,1)", 0.05)]
    [InlineData("=GAMMALN(100)", 359.134205369575)]
    public void Distributions_have_full_precision(string formula, double expected)
    {
        double actual = Formula.Number(formula);

        Assert.True(Math.Abs(expected - actual) <= 1e-12 * Math.Max(1, Math.Abs(expected)), $"{formula} gave {actual:R}, expected {expected}.");
    }

    [Theory]
    [InlineData("=PERMUT(2,3)", "#NUM!")]
    [InlineData("=PERMUT(0,1)", "#NUM!")]
    [InlineData("=PERMUTATIONA(-1,2)", "#NUM!")]
    [InlineData("=FISHER(1)", "#NUM!")]
    [InlineData("=GAMMA(0)", "#NUM!")]
    [InlineData("=GAMMA(-2)", "#NUM!")]
    [InlineData("=GAMMALN(0)", "#NUM!")]
    [InlineData("=NORM.DIST(1,0,0,TRUE)", "#NUM!")]
    [InlineData("=NORM.INV(0,0,1)", "#NUM!")]
    [InlineData("=NORM.S.INV(1)", "#NUM!")]
    [InlineData("=LOGNORM.DIST(0,1,1,TRUE)", "#NUM!")]
    [InlineData("=LOGNORM.INV(0.5,1,-1)", "#NUM!")]
    [InlineData("=EXPON.DIST(-1,1,TRUE)", "#NUM!")]
    [InlineData("=POISSON.DIST(-1,1,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(11,10,0.5,TRUE)", "#NUM!")]
    [InlineData("=BINOM.DIST(1,10,1.5,TRUE)", "#NUM!")]
    [InlineData("=CONFIDENCE(1,1,10)", "#NUM!")]
    [InlineData("=CONFIDENCE(0.05,1,0.5)", "#NUM!")]
    [InlineData("=WEIBULL.DIST(1,0,1,TRUE)", "#NUM!")]
    [InlineData("=GAMMA.DIST(-1,1,1,TRUE)", "#NUM!")]
    [InlineData("=BETA.DIST(4,8,10,TRUE,1,3)", "#NUM!")]
    [InlineData("=T.DIST(1,0,TRUE)", "#NUM!")]
    [InlineData("=T.DIST.2T(-1,5)", "#NUM!")]
    [InlineData("=T.INV(0,5)", "#NUM!")]
    [InlineData("=T.INV.2T(2,5)", "#NUM!")]
    [InlineData("=CHISQ.DIST(-1,2,TRUE)", "#NUM!")]
    [InlineData("=CHISQ.INV(1.5,2)", "#NUM!")]
    [InlineData("=CHISQ.INV.RT(0,2)", "#NUM!")]
    [InlineData("=F.DIST(-1,2,3,TRUE)", "#NUM!")]
    [InlineData("=F.DIST.RT(1,0,3)", "#NUM!")]
    [InlineData("=NORM.DIST(\"x\",0,1,TRUE)", "#VALUE!")]
    public void Distribution_errors(string formula, string expected)
    {
        AssertError(expected, Formula.Evaluate(formula));
    }

    /// <summary>Equal to the digits the expected value is written with (Excel's documentation rounds its examples).</summary>
    private static void AssertClose(double expected, double actual, string formula)
    {
        int digits = Digits(expected);

        Assert.True(Math.Abs(expected - actual) <= (0.5 * Math.Pow(10, -digits)) + 1e-12, $"{formula} gave {actual:R}, expected {expected}.");
    }

    private static int Digits(double expected)
    {
        string text = expected.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
        int point = text.IndexOf('.', StringComparison.Ordinal);

        if (text.Contains('E', StringComparison.Ordinal) == true)
        {
            return 15;
        }

        if (point < 0)
        {
            return 9;
        }

        return text.Length - point - 1;
    }

    private static void AssertError(string expected, CellValue value)
    {
        Assert.True(value.TryGetError(out string? code), $"Expected {expected}, got {value}.");
        Assert.Equal(expected, code);
    }
}
