namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Averages, counts and extremes, with their conditional *IF and *IFS forms.</summary>
internal static class StatisticalFunctionsAggregates
{
    public static void Register(FunctionTable table)
    {
        table.Add("AVERAGE", 1, FunctionTable.Many, call => Average(call, false));
        table.Add("AVERAGEA", 1, FunctionTable.Many, call => Average(call, true));
        table.Add("AVERAGEIF", 2, 3, AverageIf);
        table.Add("AVERAGEIFS", 3, FunctionTable.Many, call => ConditionalTarget(call, Average));
        table.Add("COUNT", 1, FunctionTable.Many, Count);
        table.Add("COUNTA", 1, FunctionTable.Many, CountA);
        table.Add("COUNTBLANK", 1, 1, CountBlank);
        table.Add("COUNTIF", 2, 2, CountIfs);
        table.Add("COUNTIFS", 2, FunctionTable.Many, CountIfs);
        table.Add("MAX", 1, FunctionTable.Many, call => Extreme(call, false, true));
        table.Add("MAXA", 1, FunctionTable.Many, call => Extreme(call, true, true));
        table.Add("MAXIFS", 3, FunctionTable.Many, call => ConditionalTarget(call, numbers => Extreme(numbers, true)));
        table.Add("MIN", 1, FunctionTable.Many, call => Extreme(call, false, false));
        table.Add("MINA", 1, FunctionTable.Many, call => Extreme(call, true, false));
        table.Add("MINIFS", 3, FunctionTable.Many, call => ConditionalTarget(call, numbers => Extreme(numbers, false)));
    }

    private static Value Average(FunctionCall call, bool includeTextAndBooleans)
    {
        if (call.TryCollectNumbers(0, includeTextAndBooleans, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        return Average(numbers);
    }

    private static Value Average(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(StatisticalFunctions.Mean(numbers));
    }

    /// <summary>MAX and MIN give 0 when there are no numbers.</summary>
    private static Value Extreme(FunctionCall call, bool includeTextAndBooleans, bool largest)
    {
        if (call.TryCollectNumbers(0, includeTextAndBooleans, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        return Extreme(numbers, largest);
    }

    private static Value Extreme(List<double> numbers, bool largest)
    {
        if (numbers.Count == 0)
        {
            return Value.FromNumber(0);
        }

        if (largest == true)
        {
            return Value.FromNumber(numbers.Max());
        }

        return Value.FromNumber(numbers.Min());
    }

    /// <summary>
    /// COUNT: numbers (dates included) in references and arrays; a value written directly counts
    /// when it converts to a number, so =COUNT("1",TRUE,"x") is 2. Errors are not counted.
    /// </summary>
    private static Value Count(FunctionCall call)
    {
        long count = 0;

        for (int index = 0; index < call.Count; index++)
        {
            if (call.IsMissing(index) == true)
            {
                continue;
            }

            if (call.IsRangeOrArray(index) == true)
            {
                foreach (Value item in call.Values(index))
                {
                    if (item.Kind == ValueKind.Number)
                    {
                        count++;
                    }
                }

                continue;
            }

            Value scalar = call.Scalar(index);

            if (scalar.IsBlank == false && scalar.IsError == false && Coerce.ToNumber(scalar, call.Uses1904).IsError == false)
            {
                count++;
            }
        }

        return Value.FromNumber(count);
    }

    /// <summary>COUNTA: every value that is not blank, errors and "" included.</summary>
    private static Value CountA(FunctionCall call)
    {
        long count = 0;

        for (int index = 0; index < call.Count; index++)
        {
            if (call.IsMissing(index) == true)
            {
                continue;
            }

            if (call.IsRangeOrArray(index) == true)
            {
                foreach (Value item in call.Values(index))
                {
                    if (item.IsBlank == false)
                    {
                        count++;
                    }
                }

                continue;
            }

            count++;
        }

        return Value.FromNumber(count);
    }

    /// <summary>COUNTBLANK: blank cells and cells with "".</summary>
    private static Value CountBlank(FunctionCall call)
    {
        if (StatisticalFunctionsGrid.TryCreate(call, 0, out StatisticalFunctionsGrid grid, out Value error) == false)
        {
            return error;
        }

        long count = grid.CellCount - ((long)grid.UsedRows * grid.UsedColumns);

        for (int row = 0; row < grid.UsedRows; row++)
        {
            for (int column = 0; column < grid.UsedColumns; column++)
            {
                Value item = grid.Get(row, column);

                if (item.IsBlank == true || (item.Kind == ValueKind.Text && item.Text!.Length == 0))
                {
                    count++;
                }
            }
        }

        return Value.FromNumber(count);
    }

    private static Value CountIfs(FunctionCall call)
    {
        if (TryMatch(call, 0, call.Count, null, out _, out long count, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(count);
    }

    /// <summary>AVERAGEIF(range, criteria, [average_range]): average_range is read from its top-left cell with the size of range.</summary>
    private static Value AverageIf(FunctionCall call)
    {
        if (StatisticalFunctionsGrid.TryCreate(call, 0, out StatisticalFunctionsGrid range, out Value error) == false)
        {
            return error;
        }

        StatisticalFunctionsGrid target = range;

        if (call.IsMissing(2) == false)
        {
            if (StatisticalFunctionsGrid.TryCreate(call, 2, out StatisticalFunctionsGrid averageRange, out error) == false)
            {
                return error;
            }

            target = averageRange.Resized(range);
        }

        if (TryMatch(call, 0, 2, target, out List<Value> matched, out _, out error) == false)
        {
            return error;
        }

        if (TryNumbersOfMatches(matched, out List<double> numbers, out error) == false)
        {
            return error;
        }

        return Average(numbers);
    }

    /// <summary>AVERAGEIFS, MAXIFS, MINIFS: the target range first, then range/criteria pairs of the same shape.</summary>
    private static Value ConditionalTarget(FunctionCall call, Func<List<double>, Value> aggregate)
    {
        if (StatisticalFunctionsGrid.TryCreate(call, 0, out StatisticalFunctionsGrid target, out Value error) == false)
        {
            return error;
        }

        if (TryMatch(call, 1, call.Count, target, out List<Value> matched, out _, out error) == false)
        {
            return error;
        }

        if (TryNumbersOfMatches(matched, out List<double> numbers, out error) == false)
        {
            return error;
        }

        return aggregate(numbers);
    }

    private static bool TryNumbersOfMatches(List<Value> matched, out List<double> numbers, out Value error)
    {
        numbers = [];
        error = Value.Blank;

        foreach (Value item in matched)
        {
            if (item.IsError == true)
            {
                error = item;

                return false;
            }

            if (item.Kind == ValueKind.Number)
            {
                numbers.Add(item.Number);
            }
        }

        return true;
    }

    /// <summary>
    /// Walks the range/criteria pairs in the arguments [first, end) cell by cell. Every range must
    /// have the shape of the first (and of <paramref name="target"/>), otherwise #VALUE!.
    /// Gives the target's values where every criterion holds, and the number of such cells.
    /// </summary>
    private static bool TryMatch(FunctionCall call, int first, int end, StatisticalFunctionsGrid? target, out List<Value> matched, out long count, out Value error)
    {
        matched = [];
        count = 0;
        error = Value.Blank;

        if ((end - first) % 2 != 0)
        {
            error = Value.ValueError;

            return false;
        }

        List<StatisticalFunctionsGrid> ranges = [];
        List<Criteria> criteria = [];

        for (int index = first; index < end; index += 2)
        {
            if (StatisticalFunctionsGrid.TryCreate(call, index, out StatisticalFunctionsGrid range, out error) == false)
            {
                return false;
            }

            if ((ranges.Count > 0 && range.SameShape(ranges[0]) == false) || (target is not null && range.SameShape(target) == false))
            {
                error = Value.ValueError;

                return false;
            }

            ranges.Add(range);
            criteria.Add(Criteria.Parse(call.Scalar(index + 1), call.Uses1904));
        }

        StatisticalFunctionsGrid shape = ranges[0];
        int usedRows = 0;
        int usedColumns = 0;

        foreach (StatisticalFunctionsGrid range in ranges)
        {
            usedRows = Math.Max(usedRows, range.UsedRows);
            usedColumns = Math.Max(usedColumns, range.UsedColumns);
        }

        if (target is not null)
        {
            usedRows = Math.Max(usedRows, target.UsedRows);
            usedColumns = Math.Max(usedColumns, target.UsedColumns);
        }

        usedRows = Math.Min(usedRows, shape.Rows);
        usedColumns = Math.Min(usedColumns, shape.Columns);

        for (int row = 0; row < usedRows; row++)
        {
            for (int column = 0; column < usedColumns; column++)
            {
                if (MatchesAll(ranges, criteria, row, column) == false)
                {
                    continue;
                }

                count++;

                if (target is not null)
                {
                    matched.Add(target.Get(row, column));
                }
            }
        }

        // The cells beyond the used area are blank in every range.
        if (criteria.TrueForAll(item => item.Matches(Value.Blank)) == true)
        {
            count += shape.CellCount - ((long)usedRows * usedColumns);
        }

        return true;
    }

    private static bool MatchesAll(List<StatisticalFunctionsGrid> ranges, List<Criteria> criteria, int row, int column)
    {
        for (int index = 0; index < ranges.Count; index++)
        {
            if (criteria[index].Matches(ranges[index].Get(row, column)) == false)
            {
                return false;
            }
        }

        return true;
    }
}
