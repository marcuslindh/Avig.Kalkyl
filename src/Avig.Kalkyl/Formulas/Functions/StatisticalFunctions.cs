namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Statistical functions: averages, counts, extremes, spread, distributions and regression.
/// The bodies live in the StatisticalFunctions*.cs files next to this one.
/// </summary>
internal static class StatisticalFunctions
{
    public static void Register(FunctionTable table)
    {
        ArgumentNullException.ThrowIfNull(table);

        StatisticalFunctionsAggregates.Register(table);
        StatisticalFunctionsOrder.Register(table);
        StatisticalFunctionsSpread.Register(table);
        StatisticalFunctionsDistributions.Register(table);
    }

    /// <summary>
    /// The numbers in one argument, with the rule of <see cref="FunctionCall.TryCollectNumbers"/>:
    /// in a reference or array only numbers count (or, with <paramref name="includeTextAndBooleans"/>,
    /// text as 0 and booleans as 1/0); a value written directly is converted. An error stops it.
    /// </summary>
    internal static bool TryNumbersOf(FunctionCall call, int index, bool includeTextAndBooleans, out List<double> numbers, out Value error)
    {
        numbers = [];
        error = Value.Blank;

        if (call.IsMissing(index) == true)
        {
            return true;
        }

        if (call.IsRangeOrArray(index) == true)
        {
            foreach (Value item in call.Values(index))
            {
                switch (item.Kind)
                {
                    case ValueKind.Number:
                        numbers.Add(item.Number);

                        break;

                    case ValueKind.Error:
                        error = item;

                        return false;

                    case ValueKind.Bool when includeTextAndBooleans == true:
                        numbers.Add(item.Number);

                        break;

                    case ValueKind.Text when includeTextAndBooleans == true:
                        numbers.Add(0);

                        break;
                }
            }

            return true;
        }

        Value scalar = call.Scalar(index);

        if (scalar.IsBlank == true)
        {
            return true;
        }

        Value converted = Coerce.ToNumber(scalar, call.Uses1904);

        if (converted.IsError == true)
        {
            error = converted;

            return false;
        }

        numbers.Add(converted.Number);

        return true;
    }

    /// <summary>
    /// The pairs of numbers in two arguments of the same size, as CORREL, SLOPE and the like take
    /// them: a pair where either side is not a number is skipped; different sizes give #N/A.
    /// </summary>
    internal static bool TryPairs(FunctionCall call, int firstIndex, int secondIndex, out List<double> firsts, out List<double> seconds, out Value error)
    {
        firsts = [];
        seconds = [];
        error = Value.Blank;

        Value firstRaw = call.Raw(firstIndex);
        Value secondRaw = call.Raw(secondIndex);

        if (firstRaw.IsError == true)
        {
            error = firstRaw;

            return false;
        }

        if (secondRaw.IsError == true)
        {
            error = secondRaw;

            return false;
        }

        List<Value> firstValues = [.. call.Values(firstIndex)];
        List<Value> secondValues = [.. call.Values(secondIndex)];

        if (firstValues.Count != secondValues.Count)
        {
            error = Value.NA;

            return false;
        }

        for (int index = 0; index < firstValues.Count; index++)
        {
            Value first = firstValues[index];
            Value second = secondValues[index];

            if (first.IsError == true)
            {
                error = first;

                return false;
            }

            if (second.IsError == true)
            {
                error = second;

                return false;
            }

            if (first.Kind == ValueKind.Number && second.Kind == ValueKind.Number)
            {
                firsts.Add(first.Number);
                seconds.Add(second.Number);
            }
        }

        return true;
    }

    internal static double Mean(List<double> numbers)
    {
        double sum = 0;

        foreach (double number in numbers)
        {
            sum += number;
        }

        return sum / numbers.Count;
    }

    /// <summary>The sum of squared deviations from the mean, in two passes for accuracy.</summary>
    internal static double SumOfSquares(List<double> numbers)
    {
        double mean = Mean(numbers);
        double sum = 0;

        foreach (double number in numbers)
        {
            double deviation = number - mean;
            sum += deviation * deviation;
        }

        return sum;
    }
}
