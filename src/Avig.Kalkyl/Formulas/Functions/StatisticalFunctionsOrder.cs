namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Order statistics: MEDIAN, MODE, LARGE, SMALL, RANK, PERCENTILE, PERCENTRANK, QUARTILE and TRIMMEAN.</summary>
internal static class StatisticalFunctionsOrder
{
    public static void Register(FunctionTable table)
    {
        table.Add("MEDIAN", 1, FunctionTable.Many, Median);
        table.Add("MODE", 1, FunctionTable.Many, Mode);
        table.Add("MODE.SNGL", 1, FunctionTable.Many, Mode);
        table.Add("MODE.MULT", 1, FunctionTable.Many, ModeMultiple);
        table.Add("LARGE", 2, 2, call => Nth(call, true));
        table.Add("SMALL", 2, 2, call => Nth(call, false));
        table.Add("RANK", 2, 3, call => Rank(call, false));
        table.Add("RANK.EQ", 2, 3, call => Rank(call, false));
        table.Add("RANK.AVG", 2, 3, call => Rank(call, true));
        table.Add("PERCENTILE", 2, 2, call => Percentile(call, false));
        table.Add("PERCENTILE.INC", 2, 2, call => Percentile(call, false));
        table.Add("PERCENTILE.EXC", 2, 2, call => Percentile(call, true));
        table.Add("PERCENTRANK", 2, 3, call => PercentRank(call, false));
        table.Add("PERCENTRANK.INC", 2, 3, call => PercentRank(call, false));
        table.Add("PERCENTRANK.EXC", 2, 3, call => PercentRank(call, true));
        table.Add("QUARTILE", 2, 2, call => Quartile(call, false));
        table.Add("QUARTILE.INC", 2, 2, call => Quartile(call, false));
        table.Add("QUARTILE.EXC", 2, 2, call => Quartile(call, true));
        table.Add("TRIMMEAN", 2, 2, TrimMean);
    }

    private static Value Median(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        if (numbers.Count == 0)
        {
            return Value.Num;
        }

        numbers.Sort();
        int middle = numbers.Count / 2;

        if (numbers.Count % 2 == 1)
        {
            return Value.FromNumber(numbers[middle]);
        }

        return Value.FromNumber((numbers[middle - 1] + numbers[middle]) / 2);
    }

    /// <summary>The values that occur most often (at least twice), in the order they first appear.</summary>
    private static List<double> Modes(List<double> numbers)
    {
        Dictionary<double, int> counts = [];
        List<double> order = [];

        foreach (double number in numbers)
        {
            if (counts.TryGetValue(number, out int count) == true)
            {
                counts[number] = count + 1;
            }
            else
            {
                counts[number] = 1;
                order.Add(number);
            }
        }

        int most = 0;

        foreach (int count in counts.Values)
        {
            most = Math.Max(most, count);
        }

        if (most < 2)
        {
            return [];
        }

        return order.FindAll(number => counts[number] == most);
    }

    /// <summary>MODE: the most frequent number, the first one when several tie; #N/A when no number repeats.</summary>
    private static Value Mode(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error, skipText: true) == false)
        {
            return error;
        }

        List<double> modes = Modes(numbers);

        if (modes.Count == 0)
        {
            return Value.NA;
        }

        return Value.FromNumber(modes[0]);
    }

    /// <summary>MODE.MULT: every most frequent number, as a vertical array.</summary>
    private static Value ModeMultiple(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error, skipText: true) == false)
        {
            return error;
        }

        List<double> modes = Modes(numbers);

        if (modes.Count == 0)
        {
            return Value.NA;
        }

        ValueArray result = new(modes.Count, 1);

        for (int index = 0; index < modes.Count; index++)
        {
            result[index, 0] = Value.FromNumber(modes[index]);
        }

        return Value.FromArray(result);
    }

    /// <summary>LARGE and SMALL: k is rounded up, and must be between 1 and the count (#NUM! otherwise).</summary>
    private static Value Nth(FunctionCall call, bool largest)
    {
        if (StatisticalFunctions.TryNumbersOf(call, 0, false, out List<double> numbers, out Value error) == false || call.TryNumber(1, out double position, out error) == false)
        {
            return error;
        }

        return Nth(numbers, position, largest);
    }

    /// <summary>
    /// The k-th largest or smallest number. Excel truncates k for SMALL and counts LARGE from the top,
    /// so LARGE rounds k up; below 1 or past the count both are #NUM!.
    /// </summary>
    internal static Value Nth(List<double> numbers, double position, bool largest)
    {
        double rank = Math.Floor(position);

        if (largest == true)
        {
            rank = Math.Ceiling(position);
        }

        if (numbers.Count == 0 || position < 1 || rank > numbers.Count)
        {
            return Value.Num;
        }

        numbers.Sort();

        if (largest == true)
        {
            return Value.FromNumber(numbers[numbers.Count - (int)rank]);
        }

        return Value.FromNumber(numbers[(int)rank - 1]);
    }

    /// <summary>RANK(number, ref, [order]): descending unless order is non-zero; #N/A when the number is not in ref. RANK.AVG averages ties.</summary>
    private static Value Rank(FunctionCall call, bool average)
    {
        if (call.TryNumber(0, out double number, out Value error) == false
            || StatisticalFunctions.TryNumbersOf(call, 1, false, out List<double> numbers, out error) == false
            || call.TryNumber(2, 0, out double order, out error) == false)
        {
            return error;
        }

        int before = 0;
        int equal = 0;

        foreach (double item in numbers)
        {
            if (item == number)
            {
                equal++;
            }
            else if ((order == 0 && item > number) || (order != 0 && item < number))
            {
                before++;
            }
        }

        if (equal == 0)
        {
            return Value.NA;
        }

        if (average == true)
        {
            return Value.FromNumber(before + ((equal + 1) / 2.0));
        }

        return Value.FromNumber(before + 1);
    }

    private static Value Percentile(FunctionCall call, bool exclusive)
    {
        if (StatisticalFunctions.TryNumbersOf(call, 0, false, out List<double> numbers, out Value error) == false || call.TryNumber(1, out double fraction, out error) == false)
        {
            return error;
        }

        return Percentile(numbers, fraction, exclusive);
    }

    /// <summary>
    /// The k-th percentile with linear interpolation: the inclusive form puts it at k·(n−1) in the
    /// sorted values (k in [0, 1]); the exclusive form at k·(n+1) − 1, which must lie inside them.
    /// </summary>
    internal static Value Percentile(List<double> numbers, double fraction, bool exclusive)
    {
        if (numbers.Count == 0 || fraction < 0 || fraction > 1)
        {
            return Value.Num;
        }

        numbers.Sort();
        double position;

        if (exclusive == true)
        {
            double rank = fraction * (numbers.Count + 1);

            if (rank < 1 || rank > numbers.Count)
            {
                return Value.Num;
            }

            position = rank - 1;
        }
        else
        {
            position = fraction * (numbers.Count - 1);
        }

        int lower = (int)Math.Floor(position);
        double weight = position - lower;

        if (lower + 1 >= numbers.Count)
        {
            return Value.FromNumber(numbers[^1]);
        }

        return Value.FromNumber(numbers[lower] + (weight * (numbers[lower + 1] - numbers[lower])));
    }

    /// <summary>QUARTILE: quart is truncated; 0 to 4 for the inclusive form, 1 to 3 for the exclusive one.</summary>
    private static Value Quartile(FunctionCall call, bool exclusive)
    {
        if (StatisticalFunctions.TryNumbersOf(call, 0, false, out List<double> numbers, out Value error) == false || call.TryNumber(1, out double quart, out error) == false)
        {
            return error;
        }

        return Quartile(numbers, quart, exclusive);
    }

    /// <summary>The quartile of the numbers; quart is truncated.</summary>
    internal static Value Quartile(List<double> numbers, double quart, bool exclusive)
    {
        double whole = Math.Truncate(quart);

        if (whole < 0 || whole > 4 || (exclusive == true && (whole < 1 || whole > 3)))
        {
            return Value.Num;
        }

        return Percentile(numbers, whole / 4, exclusive);
    }

    /// <summary>
    /// PERCENTRANK(array, x, [significance]): the rank of x as a fraction, interpolated between the
    /// values around it, truncated to significance digits (3 by default). x outside the values is #N/A.
    /// The inclusive form ranks the i-th smallest (0-based) as i/(n−1), the exclusive one as (i+1)/(n+1).
    /// </summary>
    private static Value PercentRank(FunctionCall call, bool exclusive)
    {
        if (StatisticalFunctions.TryNumbersOf(call, 0, false, out List<double> numbers, out Value error) == false
            || call.TryNumber(1, out double number, out error) == false
            || call.TryInteger(2, 3, out long significance, out error) == false)
        {
            return error;
        }

        if (numbers.Count == 0 || significance < 1)
        {
            return Value.Num;
        }

        numbers.Sort();

        if (number < numbers[0] || number > numbers[^1])
        {
            return Value.NA;
        }

        int count = numbers.Count;
        double position = 0;

        for (int index = 0; index < count; index++)
        {
            if (numbers[index] == number)
            {
                position = index;

                break;
            }

            if (numbers[index] < number && index + 1 < count && numbers[index + 1] > number)
            {
                position = index + ((number - numbers[index]) / (numbers[index + 1] - numbers[index]));

                break;
            }
        }

        double result;

        if (exclusive == true)
        {
            result = (position + 1) / (count + 1);
        }
        else if (count == 1)
        {
            result = 1;
        }
        else
        {
            result = position / (count - 1);
        }

        return Value.FromNumber(TruncateDigits(result, (int)Math.Min(significance, 15)));
    }

    /// <summary>Truncates to a number of decimals, without letting 0.7 stored as 0.6999… fall to 0.6.</summary>
    private static double TruncateDigits(double number, int digits)
    {
        double scale = Math.Pow(10, digits);
        double scaled = number * scale;
        double nearest = Math.Round(scaled);

        if (Math.Abs(scaled - nearest) < 1e-9 * Math.Max(1, Math.Abs(scaled)))
        {
            scaled = nearest;
        }

        return Math.Floor(scaled) / scale;
    }

    /// <summary>TRIMMEAN(array, percent): drops floor(n·percent/2) values at each end; percent must be in [0, 1).</summary>
    private static Value TrimMean(FunctionCall call)
    {
        if (StatisticalFunctions.TryNumbersOf(call, 0, false, out List<double> numbers, out Value error) == false || call.TryNumber(1, out double percent, out error) == false)
        {
            return error;
        }

        if (numbers.Count == 0 || percent < 0 || percent >= 1)
        {
            return Value.Num;
        }

        numbers.Sort();
        int trimmed = (int)Math.Floor(numbers.Count * percent / 2);
        List<double> kept = numbers.GetRange(trimmed, numbers.Count - (2 * trimmed));

        return Value.FromNumber(StatisticalFunctions.Mean(kept));
    }
}
