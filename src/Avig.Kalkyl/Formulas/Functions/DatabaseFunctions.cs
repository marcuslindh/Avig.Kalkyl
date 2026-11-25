namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// The database functions: DSUM(database, field, criteria) and the rest. The database is a
/// range with labels in its first row and one record per row below. The criteria range has
/// labels in its first row too; the conditions on one row must all hold, and a record counts
/// when any row holds. An empty condition holds for everything.
/// </summary>
internal static class DatabaseFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("DAVERAGE", 3, 3, call => Aggregate(call, numbers => Average(numbers)));
        table.Add("DCOUNT", 2, 3, call => Count(call, false));
        table.Add("DCOUNTA", 2, 3, call => Count(call, true));
        table.Add("DGET", 3, 3, Get);
        table.Add("DMAX", 3, 3, call => Aggregate(call, numbers => Extreme(numbers, true)));
        table.Add("DMIN", 3, 3, call => Aggregate(call, numbers => Extreme(numbers, false)));
        table.Add("DPRODUCT", 3, 3, call => Aggregate(call, Product));
        table.Add("DSTDEV", 3, 3, call => Aggregate(call, numbers => Spread(numbers, true, true)));
        table.Add("DSTDEVP", 3, 3, call => Aggregate(call, numbers => Spread(numbers, false, true)));
        table.Add("DSUM", 3, 3, call => Aggregate(call, numbers => Value.FromNumber(numbers.Sum())));
        table.Add("DVAR", 3, 3, call => Aggregate(call, numbers => Spread(numbers, true, false)));
        table.Add("DVARP", 3, 3, call => Aggregate(call, numbers => Spread(numbers, false, false)));
    }

    /// <summary>The numbers in the field of the matching records, through <paramref name="aggregate"/>. Text and booleans are left out, as in Excel.</summary>
    private static Value Aggregate(FunctionCall call, Func<List<double>, Value> aggregate)
    {
        if (TryField(call, out List<Value> values, out Value error) == false)
        {
            return error;
        }

        List<double> numbers = [];

        foreach (Value value in values)
        {
            if (value.Kind == ValueKind.Number)
            {
                numbers.Add(value.Number);
            }
        }

        return aggregate(numbers);
    }

    /// <summary>DCOUNT counts the numbers in the field, DCOUNTA everything that is not empty. Without a field, the matching records.</summary>
    private static Value Count(FunctionCall call, bool everything)
    {
        if (call.IsMissing(1) == true)
        {
            if (TryRecords(call, out _, out List<int> records, out Value missingError) == false)
            {
                return missingError;
            }

            return Value.FromNumber(records.Count);
        }

        if (TryField(call, out List<Value> values, out Value error) == false)
        {
            return error;
        }

        int count = 0;

        foreach (Value value in values)
        {
            if ((everything == true && value.IsBlank == false) || value.Kind == ValueKind.Number)
            {
                count++;
            }
        }

        return Value.FromNumber(count);
    }

    /// <summary>DGET: the one value that matches. #VALUE! when none does, #NUM! when several do.</summary>
    private static Value Get(FunctionCall call)
    {
        if (TryField(call, out List<Value> values, out Value error) == false)
        {
            return error;
        }

        if (values.Count == 0)
        {
            return Value.ValueError;
        }

        if (values.Count > 1)
        {
            return Value.Num;
        }

        return values[0];
    }

    /// <summary>The field's values in the records that match the criteria.</summary>
    private static bool TryField(FunctionCall call, out List<Value> values, out Value error)
    {
        values = [];

        if (TryRecords(call, out ValueArray database, out List<int> records, out error) == false)
        {
            return false;
        }

        if (TryFieldColumn(call, database, out int column, out error) == false)
        {
            return false;
        }

        foreach (int record in records)
        {
            values.Add(database[record, column]);
        }

        return true;
    }

    /// <summary>The field as a label in the first row (case ignored) or a column number from 1.</summary>
    private static bool TryFieldColumn(FunctionCall call, ValueArray database, out int column, out Value error)
    {
        column = -1;
        error = Value.ValueError;
        Value field = call.Scalar(1);

        if (field.IsError == true)
        {
            error = field;

            return false;
        }

        if (field.Kind == ValueKind.Number)
        {
            int index = (int)Math.Truncate(field.Number);

            if (index < 1 || index > database.Columns)
            {
                return false;
            }

            column = index - 1;

            return true;
        }

        column = FindLabel(database, Coerce.ToText(field).Text ?? string.Empty);

        return column >= 0;
    }

    /// <summary>The rows of the database (1 and down; row 0 holds the labels) that meet the criteria.</summary>
    private static bool TryRecords(FunctionCall call, out ValueArray database, out List<int> records, out Value error)
    {
        records = [];
        database = null!;
        ValueArray criteria;

        if (call.TryArray(0, out database, out error) == false || call.TryArray(2, out criteria, out error) == false)
        {
            return false;
        }

        if (call.IsMissing(1) == false && call.Scalar(1).IsError == true)
        {
            error = call.Scalar(1);

            return false;
        }

        // The criteria must be cells: Excel gives #VALUE! for an array or a single value.
        if (database.Rows < 1 || criteria.Rows < 1 || call.Raw(2).Kind != ValueKind.Reference)
        {
            error = Value.ValueError;

            return false;
        }

        // Each criteria column becomes the database column with the same label. A criteria
        // label the database does not have leaves that column's conditions unmet.
        int[] columns = new int[criteria.Columns];

        for (int column = 0; column < criteria.Columns; column++)
        {
            columns[column] = FindLabel(database, Coerce.ToText(criteria[0, column]).Text ?? string.Empty);
        }

        List<(int Column, Criteria Condition)[]> rows = [];

        for (int row = 1; row < criteria.Rows; row++)
        {
            List<(int Column, Criteria Condition)> conditions = [];

            for (int column = 0; column < criteria.Columns; column++)
            {
                Value condition = criteria[row, column];

                if (condition.IsBlank == true || (condition.Kind == ValueKind.Text && condition.Text!.Length == 0))
                {
                    continue;
                }

                conditions.Add((columns[column], Criteria.Parse(condition, call.Conversion)));
            }

            rows.Add([.. conditions]);
        }

        for (int record = 1; record < database.Rows; record++)
        {
            if (Matches(database, record, rows) == true)
            {
                records.Add(record);
            }
        }

        return true;
    }

    private static bool Matches(ValueArray database, int record, List<(int Column, Criteria Condition)[]> rows)
    {
        // A criteria range with only its label row selects nothing, as in Excel.
        foreach ((int Column, Criteria Condition)[] conditions in rows)
        {
            bool all = true;

            foreach ((int column, Criteria condition) in conditions)
            {
                if (column < 0 || condition.Matches(database[record, column]) == false)
                {
                    all = false;

                    break;
                }
            }

            if (all == true)
            {
                return true;
            }
        }

        return false;
    }

    private static int FindLabel(ValueArray database, string label)
    {
        string wanted = label.Trim();

        for (int column = 0; column < database.Columns; column++)
        {
            string text = Coerce.ToText(database[0, column]).Text ?? string.Empty;

            if (string.Equals(text.Trim(), wanted, StringComparison.OrdinalIgnoreCase) == true)
            {
                return column;
            }
        }

        return -1;
    }

    private static Value Average(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(numbers.Sum() / numbers.Count);
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

    private static Value Product(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return Value.FromNumber(0);
        }

        double product = 1;

        foreach (double number in numbers)
        {
            product *= number;
        }

        return Value.FromNumber(product);
    }

    /// <summary>Variance or standard deviation, of a sample (n - 1) or of the whole population (n).</summary>
    private static Value Spread(List<double> numbers, bool sample, bool root)
    {
        int divisor = numbers.Count;

        if (sample == true)
        {
            divisor--;
        }

        if (divisor <= 0)
        {
            return Value.Div0;
        }

        double mean = numbers.Average();
        double squares = numbers.Sum(number => (number - mean) * (number - mean));
        double variance = squares / divisor;

        if (root == true)
        {
            return Value.FromNumber(Math.Sqrt(variance));
        }

        return Value.FromNumber(variance);
    }
}
