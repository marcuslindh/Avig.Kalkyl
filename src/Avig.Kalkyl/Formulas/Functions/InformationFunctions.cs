namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Information functions: IS… tests, TYPE, N, NA, ERROR.TYPE.</summary>
internal static class InformationFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("ISBLANK", 1, 1, IsBlank);
        table.Add("ISERROR", 1, 1, call => Test(call, value => value.IsError == true));
        table.Add("ISERR", 1, 1, call => Test(call, value => value.IsError == true && value.Text != Value.NA.Text));
        table.Add("ISNA", 1, 1, call => Test(call, value => value.IsError == true && value.Text == Value.NA.Text));
        table.Add("ISNUMBER", 1, 1, call => Test(call, value => value.Kind == ValueKind.Number));
        table.Add("ISTEXT", 1, 1, call => Test(call, value => value.Kind == ValueKind.Text));
        table.Add("ISNONTEXT", 1, 1, call => Test(call, value => value.Kind != ValueKind.Text));
        table.Add("ISLOGICAL", 1, 1, call => Test(call, value => value.Kind == ValueKind.Bool));
        table.Add("ISEVEN", 1, 1, call => Parity(call, true));
        table.Add("ISODD", 1, 1, call => Parity(call, false));
        table.Add("ISREF", 1, 1, call => Value.FromBool(call.Raw(0).Kind == ValueKind.Reference));
        table.Add("ISFORMULA", 1, 1, IsFormula);
        table.Add("NA", 0, 0, _ => Value.NA);
        table.Add("ERROR.TYPE", 1, 1, ErrorType);
        table.Add("TYPE", 1, 1, TypeOf);
        table.Add("N", 1, 1, ToN);
    }

    private static Value Test(FunctionCall call, Func<Value, bool> test)
    {
        return Value.FromBool(test(call.Scalar(0)));
    }

    /// <summary>ISBLANK: TRUE only for a reference to an empty cell; "" and values written in the formula are not blank.</summary>
    private static Value IsBlank(FunctionCall call)
    {
        if (call.Raw(0).Kind != ValueKind.Reference)
        {
            return Value.False;
        }

        return Value.FromBool(call.Scalar(0).IsBlank);
    }

    /// <summary>ISEVEN and ISODD: the number truncated; booleans and text that is not a number give #VALUE!.</summary>
    private static Value Parity(FunctionCall call, bool even)
    {
        if (call.Scalar(0).Kind == ValueKind.Bool)
        {
            return Value.ValueError;
        }

        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        bool isEven = Math.Truncate(Math.Abs(number)) % 2 == 0;

        return Value.FromBool(isEven == even);
    }

    private static Value IsFormula(FunctionCall call)
    {
        if (call.TryReference(0, out RangeReference range, out Value error) == false)
        {
            return error;
        }

        return Value.FromBool(range.Sheet[range.FirstRow, range.FirstColumn].Formula is not null);
    }

    private static Value ErrorType(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (value.IsError == false)
        {
            return Value.NA;
        }

        int number = value.Text switch
        {
            "#NULL!" => 1,
            "#DIV/0!" => 2,
            "#VALUE!" => 3,
            "#REF!" => 4,
            "#NAME?" => 5,
            "#NUM!" => 6,
            "#N/A" => 7,
            "#GETTING_DATA" => 8,
            "#SPILL!" => 9,
            "#CONNECT!" => 10,
            "#BLOCKED!" => 11,
            "#UNKNOWN!" => 12,
            "#FIELD!" => 13,
            "#CALC!" => 14,
            _ => 0,
        };

        if (number == 0)
        {
            return Value.NA;
        }

        return Value.FromNumber(number);
    }

    /// <summary>TYPE: 1 number (and blank), 2 text, 4 logical, 16 error, 64 array or multi-cell range.</summary>
    private static Value TypeOf(FunctionCall call)
    {
        Value raw = call.Raw(0);

        if (raw.Kind == ValueKind.Array || (raw.Kind == ValueKind.Reference && raw.Reference!.IsSingleCell == false))
        {
            return Value.FromNumber(64);
        }

        Value value = call.Scalar(0);

        int type = value.Kind switch
        {
            ValueKind.Text => 2,
            ValueKind.Bool => 4,
            ValueKind.Error => 16,
            _ => 1,
        };

        return Value.FromNumber(type);
    }

    /// <summary>N: numbers as they are, TRUE/FALSE as 1/0, errors pass through, anything else 0.</summary>
    private static Value ToN(FunctionCall call)
    {
        Value value = call.Scalar(0);

        return value.Kind switch
        {
            ValueKind.Number => value,
            ValueKind.Bool => Value.FromNumber(value.Number),
            ValueKind.Error => value,
            _ => Value.FromNumber(0),
        };
    }
}
