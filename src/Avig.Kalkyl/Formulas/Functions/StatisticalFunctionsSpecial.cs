namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// The special functions behind the distributions: log-gamma, the regularized incomplete gamma
/// and beta functions, the normal distribution and its inverse, and a bracketing root finder.
/// All aim at full double precision.
/// </summary>
internal static class StatisticalFunctionsSpecial
{
    private const double Epsilon = 1e-16;
    private const double Tiny = 1e-300;
    private const int MaxIterations = 100_000;
    private const double HalfLogTwoPi = 0.91893853320467274178;
    private const double SqrtTwoPi = 2.50662827463100050242;

    private static readonly double[] StirlingCoefficients =
    [
        1.0 / 12, -1.0 / 360, 1.0 / 1260, -1.0 / 1680, 1.0 / 1188, -691.0 / 360360, 1.0 / 156, -3617.0 / 122400,
    ];

    private static readonly double[] AcklamA = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00];
    private static readonly double[] AcklamB = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01];
    private static readonly double[] AcklamC = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00];
    private static readonly double[] AcklamD = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00];

    /// <summary>ln Γ(x) for x &gt; 0: the argument is shifted up to 15 and the Stirling series used there.</summary>
    public static double LogGamma(double value)
    {
        if (value <= 0 || double.IsFinite(value) == false)
        {
            return double.NaN;
        }

        double shifted = value;
        double product = 1;

        while (shifted < 15)
        {
            product *= shifted;
            shifted += 1;
        }

        double inverse = 1 / shifted;
        double inverseSquared = inverse * inverse;
        double series = 0;
        double power = inverse;

        foreach (double coefficient in StirlingCoefficients)
        {
            series += coefficient * power;
            power *= inverseSquared;
        }

        double result = ((shifted - 0.5) * Math.Log(shifted)) - shifted + HalfLogTwoPi + series;

        return result - Math.Log(product);
    }

    /// <summary>Γ(x), with the reflection formula for negative arguments. NaN at 0 and the negative integers.</summary>
    public static double Gamma(double value)
    {
        if (value == Math.Floor(value) && value <= 0)
        {
            return double.NaN;
        }

        if (value < 0.5)
        {
            return Math.PI / (Math.Sin(Math.PI * value) * Gamma(1 - value));
        }

        if (value == Math.Floor(value) && value <= 25)
        {
            double factorial = 1;

            for (int factor = 2; factor < value; factor++)
            {
                factorial *= factor;
            }

            return factorial;
        }

        return Math.Exp(LogGamma(value));
    }

    /// <summary>ln B(a, b).</summary>
    public static double LogBeta(double first, double second)
    {
        return LogGamma(first) + LogGamma(second) - LogGamma(first + second);
    }

    /// <summary>The regularized lower incomplete gamma function P(a, x).</summary>
    public static double GammaP(double shape, double value)
    {
        if (value <= 0)
        {
            return 0;
        }

        if (value < shape + 1)
        {
            return GammaSeries(shape, value);
        }

        return 1 - GammaContinuedFraction(shape, value);
    }

    /// <summary>The regularized upper incomplete gamma function Q(a, x) = 1 − P(a, x), accurate in the tail.</summary>
    public static double GammaQ(double shape, double value)
    {
        if (value <= 0)
        {
            return 1;
        }

        if (value < shape + 1)
        {
            return 1 - GammaSeries(shape, value);
        }

        return GammaContinuedFraction(shape, value);
    }

    /// <summary>The regularized incomplete beta function I_x(a, b).</summary>
    public static double BetaRegularized(double first, double second, double value)
    {
        if (value <= 0)
        {
            return 0;
        }

        if (value >= 1)
        {
            return 1;
        }

        double front = Math.Exp((first * Math.Log(value)) + (second * Math.Log(1 - value)) - LogBeta(first, second));

        if (value < (first + 1) / (first + second + 2))
        {
            return front * BetaContinuedFraction(first, second, value) / first;
        }

        return 1 - (front * BetaContinuedFraction(second, first, 1 - value) / second);
    }

    /// <summary>The standard normal density.</summary>
    public static double NormalDensity(double value)
    {
        return Math.Exp(-0.5 * value * value) / SqrtTwoPi;
    }

    /// <summary>The standard normal distribution Φ(x), from erfc so that both tails keep their precision.</summary>
    public static double NormalCdf(double value)
    {
        return 0.5 * Erfc(-value / Math.Sqrt(2));
    }

    /// <summary>The inverse of Φ: Acklam's approximation followed by two Halley steps on Φ.</summary>
    public static double NormalInverse(double probability)
    {
        if (probability <= 0 || probability >= 1)
        {
            return double.NaN;
        }

        const double low = 0.02425;
        double result;

        if (probability < low)
        {
            double root = Math.Sqrt(-2 * Math.Log(probability));
            result = Rational(AcklamC, AcklamD, root);
        }
        else if (probability > 1 - low)
        {
            double root = Math.Sqrt(-2 * Math.Log(1 - probability));
            result = -Rational(AcklamC, AcklamD, root);
        }
        else
        {
            double centered = probability - 0.5;
            double squared = centered * centered;
            result = centered * Rational(AcklamA, AcklamB, squared);
        }

        for (int step = 0; step < 2; step++)
        {
            double difference;

            if (result < 0)
            {
                difference = NormalCdf(result) - probability;
            }
            else
            {
                difference = (1 - probability) - NormalCdf(-result);
            }

            double correction = difference * SqrtTwoPi * Math.Exp(result * result / 2);
            result -= correction / (1 + (result * correction / 2));
        }

        return result;
    }

    /// <summary>
    /// Solves <c>function(x) = target</c> for x in [low, ∞) where the function is monotonic,
    /// by expanding the bracket and bisecting to full precision. NaN when it cannot be bracketed.
    /// </summary>
    public static double Solve(Func<double, double> function, double target, double low, double high, bool increasing)
    {
        int expansions = 0;

        while (Beyond(function(high), target, increasing) == false)
        {
            low = high;
            high *= 2;
            expansions++;

            if (expansions > 2000 || double.IsFinite(high) == false)
            {
                return double.NaN;
            }
        }

        for (int iteration = 0; iteration < 2000; iteration++)
        {
            double middle = low + ((high - low) / 2);

            if (middle <= low || middle >= high)
            {
                break;
            }

            if (Beyond(function(middle), target, increasing) == true)
            {
                high = middle;
            }
            else
            {
                low = middle;
            }
        }

        return low + ((high - low) / 2);
    }

    private static bool Beyond(double value, double target, bool increasing)
    {
        if (increasing == true)
        {
            return value >= target;
        }

        return value <= target;
    }

    private static double Rational(double[] numerator, double[] denominator, double value)
    {
        double top = 0;

        foreach (double coefficient in numerator)
        {
            top = (top * value) + coefficient;
        }

        double bottom = 0;

        foreach (double coefficient in denominator)
        {
            bottom = (bottom * value) + coefficient;
        }

        bottom = (bottom * value) + 1;

        return top / bottom;
    }

    /// <summary>erfc(x) through the incomplete gamma function: erfc(x) = Q(1/2, x²) for x ≥ 0.</summary>
    private static double Erfc(double value)
    {
        if (value < 0)
        {
            return 2 - Erfc(-value);
        }

        if (value == 0)
        {
            return 1;
        }

        if (value > 27)
        {
            return 0;
        }

        return GammaQ(0.5, value * value);
    }

    private static double GammaFront(double shape, double value)
    {
        return Math.Exp((shape * Math.Log(value)) - value - LogGamma(shape));
    }

    private static double GammaSeries(double shape, double value)
    {
        double term = 1 / shape;
        double sum = term;
        double denominator = shape;

        for (int iteration = 0; iteration < MaxIterations; iteration++)
        {
            denominator += 1;
            term *= value / denominator;
            sum += term;

            if (Math.Abs(term) < Math.Abs(sum) * Epsilon)
            {
                break;
            }
        }

        return sum * GammaFront(shape, value);
    }

    private static double GammaContinuedFraction(double shape, double value)
    {
        double denominatorTerm = value + 1 - shape;
        double current = 1 / Tiny;
        double previous = 1 / denominatorTerm;
        double result = previous;

        for (int iteration = 1; iteration < MaxIterations; iteration++)
        {
            double numeratorTerm = -iteration * (iteration - shape);
            denominatorTerm += 2;
            previous = (numeratorTerm * previous) + denominatorTerm;

            if (Math.Abs(previous) < Tiny)
            {
                previous = Tiny;
            }

            current = denominatorTerm + (numeratorTerm / current);

            if (Math.Abs(current) < Tiny)
            {
                current = Tiny;
            }

            previous = 1 / previous;
            double delta = previous * current;
            result *= delta;

            if (Math.Abs(delta - 1) < Epsilon)
            {
                break;
            }
        }

        return GammaFront(shape, value) * result;
    }

    private static double BetaContinuedFraction(double first, double second, double value)
    {
        double sum = first + second;
        double plusOne = first + 1;
        double minusOne = first - 1;
        double current = 1;
        double previous = 1 - (sum * value / plusOne);

        if (Math.Abs(previous) < Tiny)
        {
            previous = Tiny;
        }

        previous = 1 / previous;
        double result = previous;

        for (int iteration = 1; iteration < MaxIterations; iteration++)
        {
            int doubled = 2 * iteration;
            double coefficient = iteration * (second - iteration) * value / ((minusOne + doubled) * (first + doubled));
            previous = 1 + (coefficient * previous);

            if (Math.Abs(previous) < Tiny)
            {
                previous = Tiny;
            }

            current = 1 + (coefficient / current);

            if (Math.Abs(current) < Tiny)
            {
                current = Tiny;
            }

            previous = 1 / previous;
            result *= previous * current;
            coefficient = -(first + iteration) * (sum + iteration) * value / ((first + doubled) * (plusOne + doubled));
            previous = 1 + (coefficient * previous);

            if (Math.Abs(previous) < Tiny)
            {
                previous = Tiny;
            }

            current = 1 + (coefficient / current);

            if (Math.Abs(current) < Tiny)
            {
                current = Tiny;
            }

            previous = 1 / previous;
            double delta = previous * current;
            result *= delta;

            if (Math.Abs(delta - 1) < Epsilon)
            {
                break;
            }
        }

        return result;
    }
}
