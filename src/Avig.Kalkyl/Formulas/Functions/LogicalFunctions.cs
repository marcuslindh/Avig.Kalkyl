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

            if (call.TryBool(index, out bool flag, out Value error) == false)
            {
                return error;
            }

            result = combine(result, flag);
            any = true;
        }

        if (any == false)
        {
            return Value.ValueError;
        }

        return Value.FromBool(result);
    }
}
