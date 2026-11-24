namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Logical functions: IF, IFS, AND, OR, XOR, NOT, IFERROR, IFNA, SWITCH, TRUE, FALSE.</summary>
internal static class LogicalFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("IF", 1, 3, If);
        table.Add("IFERROR", 2, 2, IfError);
        table.Add("AND", 1, FunctionTable.Many, call => Combine(call, (all, item) => all && item, true));
        table.Add("OR", 1, FunctionTable.Many, call => Combine(call, (any, item) => any || item, false));
        table.Add("NOT", 1, 1, Not);
        table.Add("TRUE", 0, 0, _ => Value.True);
        table.Add("FALSE", 0, 0, _ => Value.False);
        table.Add("IFS", 2, FunctionTable.Many, Ifs);
        table.Add("XOR", 1, FunctionTable.Many, call => Combine(call, (odd, item) => odd ^ item, false));
        table.Add("IFNA", 2, 2, IfNa);
        table.Add("SWITCH", 3, FunctionTable.Many, Switch);
    }

    /// <summary>IFS: the value after the first TRUE condition; #N/A when none is TRUE.</summary>
    private static Value Ifs(FunctionCall call)
    {
        for (int index = 0; index < call.Count; index += 2)
        {
            if (call.TryBool(index, out bool condition, out Value error) == false)
            {
                return error;
            }

            if (condition == false)
            {
                continue;
            }

            if (index + 1 >= call.Count)
            {
                return Value.NA;
            }

            if (call.IsMissing(index + 1) == true)
            {
                return Value.FromNumber(0);
            }

            return call.Raw(index + 1);
        }

        return Value.NA;
    }

    private static Value IfNa(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (value.IsError == true && value.Text == Value.NA.Text)
        {
            return call.Raw(1);
        }

        return call.Raw(0);
    }

    /// <summary>
    /// SWITCH(expression, value1, result1, …, [default]): the result for the first value equal to
    /// the expression (text without regard to case, as =); the default or #N/A when none matches.
    /// </summary>
    private static Value Switch(FunctionCall call)
    {
        Value expression = call.Scalar(0);

        if (expression.IsError == true)
        {
            return expression;
        }

        int index = 1;

        while (index + 1 < call.Count)
        {
            Value candidate = call.Scalar(index);

            if (candidate.IsError == true)
            {
                return candidate;
            }

            if (Same(expression, candidate) == true)
            {
                return call.Raw(index + 1);
            }

            index += 2;
        }

        if (index < call.Count)
        {
            return call.Raw(index);
        }

        return Value.NA;
    }

    private static bool Same(Value left, Value right)
    {
        bool bothSet = left.IsBlank == false && right.IsBlank == false;

        if (bothSet == true && left.Kind != right.Kind)
        {
            return false;
        }

        return Coerce.Compare(left, right) == 0;
    }

    private static Value If(FunctionCall call)
    {
        if (call.TryBool(0, out bool condition, out Value error) == false)
        {
            return error;
        }

        int branch = 2;

        if (condition == true)
        {
            branch = 1;
        }

        // IF(x) gives TRUE/FALSE; IF(x,,y) with the branch written but empty gives 0.
        if (branch >= call.Count)
        {
            return Value.FromBool(condition);
        }

        if (call.IsMissing(branch) == true)
        {
            return Value.FromNumber(0);
        }

        return call.Raw(branch);
    }

    private static Value IfError(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (value.IsError == true)
        {
            return call.Raw(1);
        }

        return call.Raw(0);
    }

    private static Value Not(FunctionCall call)
    {
        if (call.TryBool(0, out bool flag, out Value error) == false)
        {
            return error;
        }

        return Value.FromBool(flag == false);
    }

    /// <summary>
    /// AND and OR: booleans and numbers count, text in references is skipped, text written
    /// directly must read TRUE or FALSE. Without any value to combine the result is #VALUE!.
    /// </summary>
    private static Value Combine(FunctionCall call, Func<bool, bool, bool> combine, bool start)
    {
        bool result = start;
        bool any = false;

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
                    if (item.IsError == true)
                    {
                        return item;
                    }

                    if (item.Kind is ValueKind.Bool or ValueKind.Number)
                    {
                        result = combine(result, item.Bool);
                        any = true;
                    }
                }

                continue;
            }

            Value scalar = call.Scalar(index);

            if (scalar.IsError == true)
            {
                return scalar;
            }

            Value converted = Coerce.ToBool(scalar);

            // Text written directly that is not TRUE or FALSE is skipped, as in a range: =AND(TRUE,"x") is TRUE.
            if (converted.IsError == true)
            {
                continue;
            }

            result = combine(result, converted.Bool);
            any = true;
        }

        if (any == false)
        {
            return Value.ValueError;
        }

        return Value.FromBool(result);
    }
}
