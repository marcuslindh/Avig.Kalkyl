using Special = Avig.Kalkyl.Formulas.Functions.StatisticalFunctionsSpecial;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Counting and special functions (PERMUT, GAMMA, FISHER …) and the probability distributions:
/// normal, lognormal, exponential, Poisson, binomial, Weibull, gamma, beta, t, chi-squared and F.
/// Out-of-domain arguments give #NUM!, as in Excel; degrees of freedom are truncated to integers.
/// </summary>
internal static class StatisticalFunctionsDistributions
{
    public static void Register(FunctionTable table)
    {
        table.Add("PERMUT", 2, 2, Permut);
        table.Add("PERMUTATIONA", 2, 2, PermutationA);
        table.Add("FISHER", 1, 1, Fisher);
        table.Add("FISHERINV", 1, 1, call => MathFunctions.Unary(call, Math.Tanh));
        table.Add("GAMMA", 1, 1, call => MathFunctions.Unary(call, Special.Gamma));
        table.Add("GAMMALN", 1, 1, call => MathFunctions.Unary(call, Special.LogGamma));
        table.Add("GAMMALN.PRECISE", 1, 1, call => MathFunctions.Unary(call, Special.LogGamma));
        table.Add("PHI", 1, 1, call => MathFunctions.Unary(call, Special.NormalDensity));
        table.Add("GAUSS", 1, 1, call => MathFunctions.Unary(call, Gauss));

        table.Add("NORM.DIST", 4, 4, call => NormalDistribution(call, 0));
        table.Add("NORMDIST", 4, 4, call => NormalDistribution(call, 0));
        table.Add("NORM.S.DIST", 2, 2, call => NormalDistribution(call, 1));
        table.Add("NORMSDIST", 1, 1, call => NormalDistribution(call, 2));
        table.Add("NORM.INV", 3, 3, NormalInverse);
        table.Add("NORMINV", 3, 3, NormalInverse);
        table.Add("NORM.S.INV", 1, 1, StandardNormalInverse);
        table.Add("NORMSINV", 1, 1, StandardNormalInverse);
        table.Add("LOGNORM.DIST", 4, 4, LognormalDistribution);
        table.Add("LOGNORM.INV", 3, 3, LognormalInverse);
        table.Add("EXPON.DIST", 3, 3, ExponentialDistribution);
        table.Add("EXPONDIST", 3, 3, ExponentialDistribution);
        table.Add("POISSON.DIST", 3, 3, PoissonDistribution);
        table.Add("POISSON", 3, 3, PoissonDistribution);
        table.Add("BINOM.DIST", 4, 4, BinomialDistribution);
        table.Add("BINOMDIST", 4, 4, BinomialDistribution);
        table.Add("CONFIDENCE", 3, 3, Confidence);
        table.Add("CONFIDENCE.NORM", 3, 3, Confidence);
        table.Add("WEIBULL.DIST", 4, 4, WeibullDistribution);
        table.Add("GAMMA.DIST", 4, 4, GammaDistribution);
        table.Add("BETA.DIST", 4, 6, BetaDistribution);

        table.Add("T.DIST", 3, 3, TDistribution);
        table.Add("T.DIST.2T", 2, 2, call => TTail(call, true));
        table.Add("T.DIST.RT", 2, 2, call => TTail(call, false));
        table.Add("T.INV", 2, 2, TInverse);
        table.Add("T.INV.2T", 2, 2, TInverseTwoTailed);
        table.Add("CHISQ.DIST", 3, 3, ChiSquareDistribution);
        table.Add("CHISQ.DIST.RT", 2, 2, ChiSquareRightTail);
        table.Add("CHISQ.INV", 2, 2, call => ChiSquareInverse(call, false));
        table.Add("CHISQ.INV.RT", 2, 2, call => ChiSquareInverse(call, true));
        table.Add("F.DIST", 4, 4, FDistribution);
        table.Add("F.DIST.RT", 3, 3, FRightTail);
    }

    private static double Gauss(double value)
    {
        if (value < 0)
        {
            return Special.NormalCdf(value) - 0.5;
        }

        return 0.5 - Special.NormalCdf(-value);
    }

    /// <summary>PERMUT(n, k): n!/(n−k)! with both truncated; n ≤ 0, k &lt; 0 or n &lt; k is #NUM!.</summary>
    private static Value Permut(FunctionCall call)
    {
        if (call.TryNumber(0, out double total, out Value error) == false || call.TryNumber(1, out double chosen, out error) == false)
        {
            return error;
        }

        total = Math.Truncate(total);
        chosen = Math.Truncate(chosen);

        if (total <= 0 || chosen < 0 || total < chosen)
        {
            return Value.Num;
        }

        double result = 1;

        for (double factor = total - chosen + 1; factor <= total; factor++)
        {
            result *= factor;

            if (double.IsFinite(result) == false)
            {
                return Value.Num;
            }
        }

        return Value.FromNumber(result);
    }

    /// <summary>PERMUTATIONA(n, k): n^k with both truncated; negative arguments are #NUM!.</summary>
    private static Value PermutationA(FunctionCall call)
    {
        if (call.TryNumber(0, out double total, out Value error) == false || call.TryNumber(1, out double chosen, out error) == false)
        {
            return error;
        }

        total = Math.Truncate(total);
        chosen = Math.Truncate(chosen);

        if (total < 0 || chosen < 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Pow(total, chosen));
    }

    private static Value Fisher(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        if (number <= -1 || number >= 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(0.5 * Math.Log((1 + number) / (1 - number)));
    }

    /// <summary>The normal distribution. Form 0: (x, mean, sd, cumulative); 1: (z, cumulative); 2: (z), cumulative.</summary>
    private static Value NormalDistribution(FunctionCall call, int form)
    {
        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        double mean = 0;
        double deviation = 1;
        bool cumulative = true;

        if (form == 0)
        {
            if (call.TryNumber(1, out mean, out error) == false || call.TryNumber(2, out deviation, out error) == false || call.TryBool(3, out cumulative, out error) == false)
            {
                return error;
            }
        }
        else if (form == 1 && call.TryBool(1, out cumulative, out error) == false)
        {
            return error;
        }

        if (deviation <= 0)
        {
            return Value.Num;
        }

        double standardized = (number - mean) / deviation;

        if (cumulative == true)
        {
            return Value.FromNumber(Special.NormalCdf(standardized));
        }

        return Value.FromNumber(Special.NormalDensity(standardized) / deviation);
    }

    private static Value NormalInverse(FunctionCall call)
    {
        if (call.TryNumber(0, out double probability, out Value error) == false
            || call.TryNumber(1, out double mean, out error) == false
            || call.TryNumber(2, out double deviation, out error) == false)
        {
            return error;
        }

        if (probability <= 0 || probability >= 1 || deviation <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(mean + (deviation * Special.NormalInverse(probability)));
    }

    private static Value StandardNormalInverse(FunctionCall call)
    {
        if (call.TryNumber(0, out double probability, out Value error) == false)
        {
            return error;
        }

        if (probability <= 0 || probability >= 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(Special.NormalInverse(probability));
    }

    private static Value LognormalDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double mean, out error) == false
            || call.TryNumber(2, out double deviation, out error) == false
            || call.TryBool(3, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (number <= 0 || deviation <= 0)
        {
            return Value.Num;
        }

        double standardized = (Math.Log(number) - mean) / deviation;

        if (cumulative == true)
        {
            return Value.FromNumber(Special.NormalCdf(standardized));
        }

        return Value.FromNumber(Special.NormalDensity(standardized) / (deviation * number));
    }

    private static Value LognormalInverse(FunctionCall call)
    {
        if (call.TryNumber(0, out double probability, out Value error) == false
            || call.TryNumber(1, out double mean, out error) == false
            || call.TryNumber(2, out double deviation, out error) == false)
        {
            return error;
        }

        if (probability <= 0 || probability >= 1 || deviation <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Exp(mean + (deviation * Special.NormalInverse(probability))));
    }

    private static Value ExponentialDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double rate, out error) == false
            || call.TryBool(2, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (number < 0 || rate <= 0)
        {
            return Value.Num;
        }

        if (cumulative == true)
        {
            return Value.FromNumber(-double.ExpM1(-rate * number));
        }

        return Value.FromNumber(rate * Math.Exp(-rate * number));
    }

    /// <summary>POISSON.DIST(x, mean, cumulative): x is truncated; the cumulative form is Q(x+1, mean).</summary>
    private static Value PoissonDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double mean, out error) == false
            || call.TryBool(2, out bool cumulative, out error) == false)
        {
            return error;
        }

        double events = Math.Truncate(number);

        if (events < 0 || mean < 0)
        {
            return Value.Num;
        }

        if (cumulative == true)
        {
            return Value.FromNumber(Special.GammaQ(events + 1, mean));
        }

        if (mean == 0)
        {
            if (events == 0)
            {
                return Value.FromNumber(1);
            }

            return Value.FromNumber(0);
        }

        return Value.FromNumber(Math.Exp((events * Math.Log(mean)) - mean - Special.LogGamma(events + 1)));
    }

    /// <summary>BINOM.DIST(successes, trials, p, cumulative): counts are truncated; successes outside [0, trials] or p outside [0, 1] is #NUM!.</summary>
    private static Value BinomialDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double successValue, out Value error) == false
            || call.TryNumber(1, out double trialValue, out error) == false
            || call.TryNumber(2, out double probability, out error) == false
            || call.TryBool(3, out bool cumulative, out error) == false)
        {
            return error;
        }

        double successes = Math.Truncate(successValue);
        double trials = Math.Truncate(trialValue);

        if (successes < 0 || successes > trials || probability < 0 || probability > 1)
        {
            return Value.Num;
        }

        if (cumulative == true)
        {
            if (successes == trials || probability == 0)
            {
                return Value.FromNumber(1);
            }

            if (probability == 1)
            {
                return Value.FromNumber(0);
            }

            return Value.FromNumber(Special.BetaRegularized(trials - successes, successes + 1, 1 - probability));
        }

        return Value.FromNumber(BinomialProbability(successes, trials, probability));
    }

    private static double BinomialProbability(double successes, double trials, double probability)
    {
        if (probability == 0)
        {
            if (successes == 0)
            {
                return 1;
            }

            return 0;
        }

        if (probability == 1)
        {
            if (successes == trials)
            {
                return 1;
            }

            return 0;
        }

        double logCombinations = Special.LogGamma(trials + 1) - Special.LogGamma(successes + 1) - Special.LogGamma(trials - successes + 1);

        return Math.Exp(logCombinations + (successes * Math.Log(probability)) + ((trials - successes) * double.LogP1(-probability)));
    }

    /// <summary>CONFIDENCE(alpha, sd, size): the half-width of the normal confidence interval; size is truncated.</summary>
    private static Value Confidence(FunctionCall call)
    {
        if (call.TryNumber(0, out double alpha, out Value error) == false
            || call.TryNumber(1, out double deviation, out error) == false
            || call.TryNumber(2, out double sizeValue, out error) == false)
        {
            return error;
        }

        double size = Math.Truncate(sizeValue);

        if (alpha <= 0 || alpha >= 1 || deviation <= 0 || size < 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(-Special.NormalInverse(alpha / 2) * deviation / Math.Sqrt(size));
    }

    private static Value WeibullDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double shape, out error) == false
            || call.TryNumber(2, out double scale, out error) == false
            || call.TryBool(3, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (number < 0 || shape <= 0 || scale <= 0)
        {
            return Value.Num;
        }

        double power = Math.Pow(number / scale, shape);

        if (cumulative == true)
        {
            return Value.FromNumber(-double.ExpM1(-power));
        }

        return Value.FromNumber(shape / scale * Math.Pow(number / scale, shape - 1) * Math.Exp(-power));
    }

    /// <summary>GAMMA.DIST(x, alpha, beta, cumulative), with beta the scale.</summary>
    private static Value GammaDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double shape, out error) == false
            || call.TryNumber(2, out double scale, out error) == false
            || call.TryBool(3, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (number < 0 || shape <= 0 || scale <= 0)
        {
            return Value.Num;
        }

        if (cumulative == true)
        {
            return Value.FromNumber(Special.GammaP(shape, number / scale));
        }

        return Value.FromNumber(GammaDensity(number, shape, scale));
    }

    private static double GammaDensity(double number, double shape, double scale)
    {
        if (number == 0)
        {
            if (shape < 1)
            {
                return double.PositiveInfinity;
            }

            if (shape == 1)
            {
                return 1 / scale;
            }

            return 0;
        }

        return Math.Exp(((shape - 1) * Math.Log(number)) - (number / scale) - Special.LogGamma(shape) - (shape * Math.Log(scale)));
    }

    /// <summary>BETA.DIST(x, alpha, beta, cumulative, [A], [B]) on the interval [A, B], by default [0, 1].</summary>
    private static Value BetaDistribution(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double alpha, out error) == false
            || call.TryNumber(2, out double beta, out error) == false
            || call.TryBool(3, out bool cumulative, out error) == false
            || call.TryNumber(4, 0, out double lower, out error) == false
            || call.TryNumber(5, 1, out double upper, out error) == false)
        {
            return error;
        }

        if (alpha <= 0 || beta <= 0 || lower >= upper || number < lower || number > upper)
        {
            return Value.Num;
        }

        double scaled = (number - lower) / (upper - lower);

        if (cumulative == true)
        {
            return Value.FromNumber(Special.BetaRegularized(alpha, beta, scaled));
        }

        if (scaled == 0 || scaled == 1)
        {
            double edgeExponent = alpha;

            if (scaled == 1)
            {
                edgeExponent = beta;
            }

            if (edgeExponent < 1)
            {
                return Value.Num;
            }

            if (edgeExponent > 1)
            {
                return Value.FromNumber(0);
            }
        }

        double logDensity = ((alpha - 1) * Math.Log(scaled)) + ((beta - 1) * Math.Log(1 - scaled)) - Special.LogBeta(alpha, beta);

        if (scaled == 0 || scaled == 1)
        {
            logDensity = -Special.LogBeta(alpha, beta);
        }

        return Value.FromNumber(Math.Exp(logDensity) / (upper - lower));
    }

    /// <summary>Reads x and the degrees of freedom (truncated; below 1 is #NUM!).</summary>
    private static bool TryValueAndFreedom(FunctionCall call, out double number, out double freedom, out Value error)
    {
        freedom = 0;

        if (call.TryNumber(0, out number, out error) == false || call.TryNumber(1, out double freedomValue, out error) == false)
        {
            return false;
        }

        freedom = Math.Truncate(freedomValue);

        if (freedom < 1 || freedom > 1e10)
        {
            error = Value.Num;

            return false;
        }

        return true;
    }

    /// <summary>P(T &gt; t) for t ≥ 0.</summary>
    private static double TUpperTail(double number, double freedom)
    {
        return 0.5 * Special.BetaRegularized(freedom / 2, 0.5, freedom / (freedom + (number * number)));
    }

    private static Value TDistribution(FunctionCall call)
    {
        if (TryValueAndFreedom(call, out double number, out double freedom, out Value error) == false || call.TryBool(2, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (cumulative == true)
        {
            double tail = TUpperTail(Math.Abs(number), freedom);

            if (number < 0)
            {
                return Value.FromNumber(tail);
            }

            return Value.FromNumber(1 - tail);
        }

        double logDensity = Special.LogGamma((freedom + 1) / 2) - Special.LogGamma(freedom / 2) - (0.5 * Math.Log(freedom * Math.PI)) - ((freedom + 1) / 2 * double.LogP1(number * number / freedom));

        return Value.FromNumber(Math.Exp(logDensity));
    }

    /// <summary>T.DIST.2T (x ≥ 0, both tails) and T.DIST.RT (the right tail).</summary>
    private static Value TTail(FunctionCall call, bool twoTailed)
    {
        if (TryValueAndFreedom(call, out double number, out double freedom, out Value error) == false)
        {
            return error;
        }

        if (twoTailed == true)
        {
            if (number < 0)
            {
                return Value.Num;
            }

            return Value.FromNumber(2 * TUpperTail(number, freedom));
        }

        double tail = TUpperTail(Math.Abs(number), freedom);

        if (number < 0)
        {
            return Value.FromNumber(1 - tail);
        }

        return Value.FromNumber(tail);
    }

    /// <summary>The t ≥ 0 whose upper tail is <paramref name="tail"/> (at most 1/2).</summary>
    private static double TUpperInverse(double tail, double freedom)
    {
        if (tail >= 0.5)
        {
            return 0;
        }

        return Special.Solve(value => TUpperTail(value, freedom), tail, 0, 1, false);
    }

    private static Value TInverse(FunctionCall call)
    {
        if (TryValueAndFreedom(call, out double probability, out double freedom, out Value error) == false)
        {
            return error;
        }

        if (probability <= 0 || probability >= 1)
        {
            return Value.Num;
        }

        if (probability < 0.5)
        {
            return Value.FromNumber(-TUpperInverse(probability, freedom));
        }

        return Value.FromNumber(TUpperInverse(1 - probability, freedom));
    }

    private static Value TInverseTwoTailed(FunctionCall call)
    {
        if (TryValueAndFreedom(call, out double probability, out double freedom, out Value error) == false)
        {
            return error;
        }

        if (probability <= 0 || probability > 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(TUpperInverse(probability / 2, freedom));
    }

    private static Value ChiSquareDistribution(FunctionCall call)
    {
        if (TryValueAndFreedom(call, out double number, out double freedom, out Value error) == false || call.TryBool(2, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (number < 0)
        {
            return Value.Num;
        }

        if (cumulative == true)
        {
            return Value.FromNumber(Special.GammaP(freedom / 2, number / 2));
        }

        return Value.FromNumber(GammaDensity(number, freedom / 2, 2));
    }

    private static Value ChiSquareRightTail(FunctionCall call)
    {
        if (TryValueAndFreedom(call, out double number, out double freedom, out Value error) == false)
        {
            return error;
        }

        if (number < 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(Special.GammaQ(freedom / 2, number / 2));
    }

    /// <summary>CHISQ.INV (left-tailed probability) and CHISQ.INV.RT (right-tailed), solved on whichever tail is smaller.</summary>
    private static Value ChiSquareInverse(FunctionCall call, bool rightTail)
    {
        if (TryValueAndFreedom(call, out double probability, out double freedom, out Value error) == false)
        {
            return error;
        }

        if (probability < 0 || probability > 1)
        {
            return Value.Num;
        }

        double lower = probability;
        double upper = 1 - probability;

        if (rightTail == true)
        {
            lower = 1 - probability;
            upper = probability;
        }

        if (lower == 0)
        {
            return Value.FromNumber(0);
        }

        if (upper == 0)
        {
            return Value.Num;
        }

        double shape = freedom / 2;
        double half;

        if (lower <= 0.5)
        {
            half = Special.Solve(value => Special.GammaP(shape, value), lower, 0, Math.Max(1, shape), true);
        }
        else
        {
            half = Special.Solve(value => Special.GammaQ(shape, value), upper, 0, Math.Max(1, shape), false);
        }

        return Value.FromNumber(2 * half);
    }

    /// <summary>Reads x and the two degrees of freedom of the F distribution.</summary>
    private static bool TryFArguments(FunctionCall call, out double number, out double first, out double second, out Value error)
    {
        first = 0;
        second = 0;

        if (call.TryNumber(0, out number, out error) == false
            || call.TryNumber(1, out double firstValue, out error) == false
            || call.TryNumber(2, out double secondValue, out error) == false)
        {
            return false;
        }

        first = Math.Truncate(firstValue);
        second = Math.Truncate(secondValue);

        if (number < 0 || first < 1 || second < 1 || first > 1e10 || second > 1e10)
        {
            error = Value.Num;

            return false;
        }

        return true;
    }

    private static Value FDistribution(FunctionCall call)
    {
        if (TryFArguments(call, out double number, out double first, out double second, out Value error) == false || call.TryBool(3, out bool cumulative, out error) == false)
        {
            return error;
        }

        if (cumulative == true)
        {
            return Value.FromNumber(Special.BetaRegularized(first / 2, second / 2, first * number / ((first * number) + second)));
        }

        if (number == 0)
        {
            if (first < 2)
            {
                return Value.Num;
            }

            if (first == 2)
            {
                return Value.FromNumber(1);
            }

            return Value.FromNumber(0);
        }

        double logDensity = (0.5 * ((first * Math.Log(first * number)) + (second * Math.Log(second)) - ((first + second) * Math.Log((first * number) + second))))
            - Math.Log(number) - Special.LogBeta(first / 2, second / 2);

        return Value.FromNumber(Math.Exp(logDensity));
    }

    private static Value FRightTail(FunctionCall call)
    {
        if (TryFArguments(call, out double number, out double first, out double second, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(Special.BetaRegularized(second / 2, first / 2, second / ((first * number) + second)));
    }
}
