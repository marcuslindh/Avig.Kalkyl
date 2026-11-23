namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Spread and shape (STDEV, VAR, KURT, SKEW …), other means, and the paired regression functions.</summary>
internal static class StatisticalFunctionsSpread
{
    public static void Register(FunctionTable table)
    {
        table.Add("STDEV", 1, FunctionTable.Many, call => Spread(call, false, true, true));
        table.Add("STDEV.S", 1, FunctionTable.Many, call => Spread(call, false, true, true));
        table.Add("STDEV.P", 1, FunctionTable.Many, call => Spread(call, false, false, true));
        table.Add("STDEVP", 1, FunctionTable.Many, call => Spread(call, false, false, true));
        table.Add("STDEVA", 1, FunctionTable.Many, call => Spread(call, true, true, true));
        table.Add("STDEVPA", 1, FunctionTable.Many, call => Spread(call, true, false, true));
        table.Add("VAR", 1, FunctionTable.Many, call => Spread(call, false, true, false));
        table.Add("VAR.S", 1, FunctionTable.Many, call => Spread(call, false, true, false));
        table.Add("VAR.P", 1, FunctionTable.Many, call => Spread(call, false, false, false));
        table.Add("VARP", 1, FunctionTable.Many, call => Spread(call, false, false, false));
        table.Add("VARA", 1, FunctionTable.Many, call => Spread(call, true, true, false));
        table.Add("VARPA", 1, FunctionTable.Many, call => Spread(call, true, false, false));
        table.Add("AVEDEV", 1, FunctionTable.Many, AverageDeviation);
        table.Add("DEVSQ", 1, FunctionTable.Many, DeviationSquares);
        table.Add("GEOMEAN", 1, FunctionTable.Many, GeometricMean);
        table.Add("HARMEAN", 1, FunctionTable.Many, HarmonicMean);
        table.Add("KURT", 1, FunctionTable.Many, Kurtosis);
        table.Add("SKEW", 1, FunctionTable.Many, call => Skew(call, false));
        table.Add("SKEW.P", 1, FunctionTable.Many, call => Skew(call, true));
        table.Add("STANDARDIZE", 3, 3, Standardize);
        table.Add("CORREL", 2, 2, call => Paired(call, 0, 1, Correlation));
        table.Add("PEARSON", 2, 2, call => Paired(call, 0, 1, Correlation));
        table.Add("RSQ", 2, 2, call => Paired(call, 1, 0, RSquared));
        table.Add("COVAR", 2, 2, call => Paired(call, 0, 1, sums => Covariance(sums, false)));
        table.Add("COVARIANCE.P", 2, 2, call => Paired(call, 0, 1, sums => Covariance(sums, false)));
        table.Add("COVARIANCE.S", 2, 2, call => Paired(call, 0, 1, sums => Covariance(sums, true)));
        table.Add("SLOPE", 2, 2, call => Paired(call, 1, 0, Slope));
        table.Add("INTERCEPT", 2, 2, call => Paired(call, 1, 0, Intercept));
        table.Add("STEYX", 2, 2, call => Paired(call, 1, 0, StandardError));
        table.Add("FORECAST", 3, 3, Forecast);
        table.Add("FORECAST.LINEAR", 3, 3, Forecast);
    }

    /// <summary>STDEV and VAR in all their forms: sample (n−1) or population (n); the …A forms count text as 0 and booleans in references.</summary>
    private static Value Spread(FunctionCall call, bool includeTextAndBooleans, bool sample, bool root)
    {
        if (call.TryCollectNumbers(0, includeTextAndBooleans, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        int divisor = numbers.Count;

        if (sample == true)
        {
            divisor--;
        }

        if (divisor < 1)
        {
            return Value.Div0;
        }

        double variance = StatisticalFunctions.SumOfSquares(numbers) / divisor;

        if (root == true)
        {
            return Value.FromNumber(Math.Sqrt(variance));
        }

        return Value.FromNumber(variance);
    }

    private static Value AverageDeviation(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        if (numbers.Count == 0)
        {
            return Value.Num;
        }

        double mean = StatisticalFunctions.Mean(numbers);
        double sum = 0;

        foreach (double number in numbers)
        {
            sum += Math.Abs(number - mean);
        }

        return Value.FromNumber(sum / numbers.Count);
    }

    private static Value DeviationSquares(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        if (numbers.Count == 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(StatisticalFunctions.SumOfSquares(numbers));
    }

    /// <summary>GEOMEAN and HARMEAN need positive numbers (#NUM! otherwise).</summary>
    private static Value GeometricMean(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        if (numbers.Count == 0 || numbers.Exists(number => number <= 0) == true)
        {
            return Value.Num;
        }

        double logSum = 0;

        foreach (double number in numbers)
        {
            logSum += Math.Log(number);
        }

        return Value.FromNumber(Math.Exp(logSum / numbers.Count));
    }

    private static Value HarmonicMean(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        if (numbers.Count == 0 || numbers.Exists(number => number <= 0) == true)
        {
            return Value.Num;
        }

        double inverseSum = 0;

        foreach (double number in numbers)
        {
            inverseSum += 1 / number;
        }

        return Value.FromNumber(numbers.Count / inverseSum);
    }

    /// <summary>KURT: the sample excess kurtosis; fewer than four numbers or no spread is #DIV/0!.</summary>
    private static Value Kurtosis(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        double count = numbers.Count;

        if (count < 4)
        {
            return Value.Div0;
        }

        double deviation = Math.Sqrt(StatisticalFunctions.SumOfSquares(numbers) / (count - 1));

        if (deviation == 0)
        {
            return Value.Div0;
        }

        double sum = StandardizedPowerSum(numbers, deviation, 4);
        double factor = count * (count + 1) / ((count - 1) * (count - 2) * (count - 3));
        double correction = 3 * (count - 1) * (count - 1) / ((count - 2) * (count - 3));

        return Value.FromNumber((factor * sum) - correction);
    }

    /// <summary>SKEW (sample, at least three numbers) and SKEW.P (population); no spread is #DIV/0!.</summary>
    private static Value Skew(FunctionCall call, bool population)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        double count = numbers.Count;

        if ((population == false && count < 3) || count < 1)
        {
            return Value.Div0;
        }

        if (population == true)
        {
            double populationDeviation = Math.Sqrt(StatisticalFunctions.SumOfSquares(numbers) / count);

            if (populationDeviation == 0)
            {
                return Value.Div0;
            }

            return Value.FromNumber(StandardizedPowerSum(numbers, populationDeviation, 3) / count);
        }

        double deviation = Math.Sqrt(StatisticalFunctions.SumOfSquares(numbers) / (count - 1));

        if (deviation == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(count / ((count - 1) * (count - 2)) * StandardizedPowerSum(numbers, deviation, 3));
    }

    private static double StandardizedPowerSum(List<double> numbers, double deviation, int power)
    {
        double mean = StatisticalFunctions.Mean(numbers);
        double sum = 0;

        foreach (double number in numbers)
        {
            sum += Math.Pow((number - mean) / deviation, power);
        }

        return sum;
    }

    private static Value Standardize(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || call.TryNumber(1, out double mean, out error) == false
            || call.TryNumber(2, out double deviation, out error) == false)
        {
            return error;
        }

        if (deviation <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber((number - mean) / deviation);
    }

    /// <summary>Paired functions on (x, y): the argument indexes say which one holds the x values.</summary>
    private static Value Paired(FunctionCall call, int xIndex, int yIndex, Func<PairSums, Value> compute)
    {
        if (StatisticalFunctions.TryPairs(call, xIndex, yIndex, out List<double> xs, out List<double> ys, out Value error) == false)
        {
            return error;
        }

        return compute(PairSums.From(xs, ys));
    }

    private static Value Correlation(PairSums sums)
    {
        if (sums.Count == 0 || sums.Sxx == 0 || sums.Syy == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(sums.Sxy / Math.Sqrt(sums.Sxx * sums.Syy));
    }

    private static Value RSquared(PairSums sums)
    {
        if (sums.Count == 0 || sums.Sxx == 0 || sums.Syy == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(sums.Sxy * sums.Sxy / (sums.Sxx * sums.Syy));
    }

    private static Value Covariance(PairSums sums, bool sample)
    {
        int divisor = sums.Count;

        if (sample == true)
        {
            divisor--;
        }

        if (divisor < 1)
        {
            return Value.Div0;
        }

        return Value.FromNumber(sums.Sxy / divisor);
    }

    private static Value Slope(PairSums sums)
    {
        if (sums.Count == 0 || sums.Sxx == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(sums.Sxy / sums.Sxx);
    }

    private static Value Intercept(PairSums sums)
    {
        if (sums.Count == 0 || sums.Sxx == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(sums.MeanY - (sums.Sxy / sums.Sxx * sums.MeanX));
    }

    /// <summary>STEYX: the standard error of the predicted y; fewer than three pairs is #DIV/0!.</summary>
    private static Value StandardError(PairSums sums)
    {
        if (sums.Count < 3 || sums.Sxx == 0)
        {
            return Value.Div0;
        }

        double residual = sums.Syy - (sums.Sxy * sums.Sxy / sums.Sxx);

        return Value.FromNumber(Math.Sqrt(Math.Max(0, residual) / (sums.Count - 2)));
    }

    /// <summary>FORECAST(x, known_y's, known_x's): the linear regression's value at x.</summary>
    private static Value Forecast(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || StatisticalFunctions.TryPairs(call, 2, 1, out List<double> xs, out List<double> ys, out error) == false)
        {
            return error;
        }

        PairSums sums = PairSums.From(xs, ys);

        if (sums.Count == 0 || sums.Sxx == 0)
        {
            return Value.Div0;
        }

        double slope = sums.Sxy / sums.Sxx;

        return Value.FromNumber(sums.MeanY + (slope * (number - sums.MeanX)));
    }

    /// <summary>Centered sums of a set of pairs, computed in two passes.</summary>
    private readonly record struct PairSums(int Count, double MeanX, double MeanY, double Sxx, double Syy, double Sxy)
    {
        public static PairSums From(List<double> xs, List<double> ys)
        {
            if (xs.Count == 0)
            {
                return new PairSums(0, 0, 0, 0, 0, 0);
            }

            double meanX = StatisticalFunctions.Mean(xs);
            double meanY = StatisticalFunctions.Mean(ys);
            double sxx = 0;
            double syy = 0;
            double sxy = 0;

            for (int index = 0; index < xs.Count; index++)
            {
                double dx = xs[index] - meanX;
                double dy = ys[index] - meanY;
                sxx += dx * dx;
                syy += dy * dy;
                sxy += dx * dy;
            }

            return new PairSums(xs.Count, meanX, meanY, sxx, syy, sxy);
        }
    }
}
