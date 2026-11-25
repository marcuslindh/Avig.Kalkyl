namespace Avig.Kalkyl.Tests;

/// <summary>
/// The edges of the distributions: probabilities of 0 and 1, the point 0, negative values and overflow. The values are
/// Excel 365's, which at the ends of an interval are not always the mathematics: a density that is finite there can be
/// #NUM! or 0. Elsewhere a relation that must hold is tested.
/// </summary>
public class DistributionEdgeTests
{
    [Theory]
    [InlineData("=POISSON.DIST(0,0,FALSE)", 1)]
    [InlineData("=POISSON.DIST(2,0,FALSE)", 0)]
    [InlineData("=POISSON.DIST(2,0,TRUE)", 1)]
    [InlineData("=BINOM.DIST(0,10,0,FALSE)", 1)]
    [InlineData("=BINOM.DIST(3,10,0,FALSE)", 0)]
    [InlineData("=BINOM.DIST(10,10,1,FALSE)", 1)]
    [InlineData("=BINOM.DIST(3,10,1,FALSE)", 0)]
    [InlineData("=BINOM.DIST(3,10,0,TRUE)", 1)]
    [InlineData("=BINOM.DIST(3,10,1,TRUE)", 0)]
    [InlineData("=BINOM.DIST(10,10,0.5,TRUE)", 1)]
    [InlineData("=GAMMA.DIST(0,2,2,FALSE)", 0)]
    [InlineData("=CHISQ.DIST(0,2,FALSE)", 0.5)]
    [InlineData("=WEIBULL.DIST(0,0.5,2,FALSE)", 0)]
    [InlineData("=WEIBULL.DIST(0,1,2,FALSE)", 0)]
    [InlineData("=BETA.DIST(0,2,3,FALSE)", 0)]
    [InlineData("=BETA.DIST(1,3,2,FALSE)", 0)]
    [InlineData("=BETA.DIST(2,1,3,FALSE,2,4)", 0)]
    [InlineData("=BETA.DIST(4,3,1,FALSE,2,4)", 0)]
    [InlineData("=BETA.DIST(2,0.5,3,FALSE,2,4)", 0)]
    [InlineData("=BETA.DIST(2,1,1,FALSE,0,2)", 0)]
    [InlineData("=BETA.DIST(1,2,2,FALSE,1,2)", 0)]
    [InlineData("=BETA.DIST(0.5,1,1,FALSE)", 1)]
    [InlineData("=F.DIST(0,2,5,FALSE)", 1)]
    [InlineData("=F.DIST(0,3,5,FALSE)", 0)]
    [InlineData("=T.INV(0.5,7)", 0)]
    [InlineData("=CHISQ.INV(0,2)", 0)]
    [InlineData("=CHISQ.INV.RT(1,2)", 0)]
    public void The_edges_have_their_exact_values(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 12);
    }

    [Theory]
    [InlineData("=GAMMA.DIST(0,0.5,1,FALSE)")]
    [InlineData("=GAMMA.DIST(0,1,2,FALSE)")]
    [InlineData("=CHISQ.DIST(0,1,FALSE)")]
    [InlineData("=BETA.DIST(0,1,3,FALSE)")]
    [InlineData("=BETA.DIST(1,3,1,FALSE)")]
    [InlineData("=BETA.DIST(0,1,1,FALSE)")]
    [InlineData("=BETA.DIST(1,1,1,FALSE)")]
    [InlineData("=BETA.DIST(0,1,3,FALSE,0,2)")]
    [InlineData("=BETA.DIST(1,2,1,FALSE,1,2)")]
    [InlineData("=BETA.DIST(1,2,0.5,FALSE,-1,1)")]
    [InlineData("=BETA.DIST(0,0.5,3,FALSE)")]
    [InlineData("=BETA.DIST(1,3,0.5,FALSE)")]
    [InlineData("=F.DIST(0,1,5,FALSE)")]
    [InlineData("=CHISQ.DIST.RT(-1,2)")]
    [InlineData("=CHISQ.INV.RT(0,2)")]
    [InlineData("=PERMUT(1000,500)")]
    [InlineData("=POISSON.DIST(-1,2,FALSE)")]
    [InlineData("=POISSON.DIST(1,-2,FALSE)")]
    public void Where_Excel_finds_the_density_infinite_or_the_result_too_large_it_is_NUM(string formula)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code) == true && code == "#NUM!", $"{formula} gave {value}.");
    }

    [Theory]
    [InlineData(0.3, 1)]
    [InlineData(1.5, 4)]
    [InlineData(2.75, 30)]
    public void The_t_distribution_is_symmetric(double number, double freedom)
    {
        string value = Text(number);
        string degrees = Text(freedom);

        Assert.Equal(1, Formula.Number($"=T.DIST({value},{degrees},TRUE)+T.DIST(-{value},{degrees},TRUE)"), 12);
        Assert.Equal(1, Formula.Number($"=T.DIST.RT({value},{degrees})+T.DIST.RT(-{value},{degrees})"), 12);
        Assert.Equal(Formula.Number($"=T.DIST({value},{degrees},FALSE)"), Formula.Number($"=T.DIST(-{value},{degrees},FALSE)"), 12);
        Assert.Equal(-Formula.Number($"=T.INV(0.8,{degrees})"), Formula.Number($"=T.INV(0.2,{degrees})"), 9);
        Assert.Equal(-number, Formula.Number($"=T.INV(T.DIST(-{value},{degrees},TRUE),{degrees})"), 7);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(0.5)]
    [InlineData(0.9)]
    [InlineData(0.999)]
    public void The_chi_square_inverse_with_two_degrees_of_freedom_is_a_logarithm(double probability)
    {
        double expected = -2 * Math.Log(1 - probability);
        string chance = Text(probability);

        Assert.Equal(expected, Formula.Number($"=CHISQ.INV({chance},2)"), 8);
        Assert.Equal(expected, Formula.Number($"=CHISQ.INV.RT({Text(1 - probability)},2)"), 8);
        Assert.Equal(probability, Formula.Number($"=CHISQ.DIST({Text(expected)},2,TRUE)"), 10);
        Assert.Equal(1 - probability, Formula.Number($"=CHISQ.DIST.RT({Text(expected)},2)"), 10);
    }

    private static string Text(double number)
    {
        return number.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
    }
}
