namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Math and trigonometry functions.</summary>
internal static class MathFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("SUM", 1, FunctionTable.Many, Sum);
        table.Add("ABS", 1, 1, call => Unary(call, Math.Abs));
        table.Add("ROUND", 2, 2, Round);
    }

    private static Value Sum(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(numbers.Sum());
    }

    private static Value Round(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryInteger(1, out long digits, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(Numeric.Round(number, (int)Math.Clamp(digits, -30, 30)));
    }

    /// <summary>A function of one number. A result that is not finite becomes #NUM!.</summary>
    internal static Value Unary(FunctionCall call, Func<double, double> function)
    {
        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(function(number));
    }
}
