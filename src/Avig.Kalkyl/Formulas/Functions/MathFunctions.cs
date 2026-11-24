using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Math and trigonometry functions.</summary>
internal static class MathFunctions
{
    /// <summary>Excel refuses SIN, COS, TAN and their reciprocals for |x| at or above 2^27.</summary>
    private const double TrigonometryLimit = 134217728;

    /// <summary>The largest whole number a double holds exactly: GCD, LCM, BASE and DECIMAL stop there.</summary>
    private const double ExactLimit = 9007199254740992;

    private static readonly char[] RomanLetters = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];

    private static readonly int[] RomanValues = [1000, 500, 100, 50, 10, 5, 1];

    public static void Register(FunctionTable table)
    {
        table.Add("SUM", 1, FunctionTable.Many, Sum);
        table.Add("ABS", 1, 1, call => Unary(call, Math.Abs));
        table.Add("ROUND", 2, 2, Round);

        RegisterTrigonometry(table);
        RegisterRounding(table);
        RegisterPowersAndIntegers(table);
        RegisterSums(table);

        table.Add("RAND", 0, 0, _ => Value.FromNumber(Random.Shared.NextDouble()), true);
        table.Add("RANDBETWEEN", 2, 2, RandBetween, true);
        table.Add("SUBTOTAL", 2, FunctionTable.Many, Subtotal);
        table.Add("AGGREGATE", 3, FunctionTable.Many, Aggregate);
        table.Add("ROMAN", 1, 2, Roman);
        table.Add("ARABIC", 1, 1, Arabic);
        table.Add("BASE", 2, 3, ToBase);
        table.Add("DECIMAL", 2, 2, FromBase);
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

    private static void RegisterTrigonometry(FunctionTable table)
    {
        table.Add("PI", 0, 0, _ => Value.FromNumber(Math.PI));
        table.Add("DEGREES", 1, 1, call => Unary(call, number => number * 180 / Math.PI));
        table.Add("RADIANS", 1, 1, call => Unary(call, number => number * Math.PI / 180));
        table.Add("SIN", 1, 1, call => Periodic(call, Math.Sin, false));
        table.Add("COS", 1, 1, call => Periodic(call, Math.Cos, false));
        table.Add("TAN", 1, 1, call => Periodic(call, Math.Tan, false));
        table.Add("COT", 1, 1, call => Periodic(call, Math.Tan, true));
        table.Add("CSC", 1, 1, call => Periodic(call, Math.Sin, true));
        table.Add("SEC", 1, 1, call => Periodic(call, Math.Cos, true));
        table.Add("ASIN", 1, 1, call => Unary(call, Math.Asin));
        table.Add("ACOS", 1, 1, call => Unary(call, Math.Acos));
        table.Add("ATAN", 1, 1, call => Unary(call, Math.Atan));
        table.Add("ACOT", 1, 1, call => Unary(call, number => (Math.PI / 2) - Math.Atan(number)));
        table.Add("ATAN2", 2, 2, Atan2);
        table.Add("SINH", 1, 1, call => Unary(call, Math.Sinh));
        table.Add("COSH", 1, 1, call => Unary(call, Math.Cosh));
        table.Add("TANH", 1, 1, call => Unary(call, Math.Tanh));
        table.Add("COTH", 1, 1, call => Reciprocal(call, Math.Tanh));
        table.Add("CSCH", 1, 1, call => Reciprocal(call, Math.Sinh));
        table.Add("SECH", 1, 1, call => Reciprocal(call, Math.Cosh));
        table.Add("ASINH", 1, 1, call => Unary(call, Math.Asinh));
        table.Add("ACOSH", 1, 1, call => Unary(call, Math.Acosh));
        table.Add("ATANH", 1, 1, call => Unary(call, Math.Atanh));
        table.Add("ACOTH", 1, 1, Acoth);
    }

    private static void RegisterRounding(FunctionTable table)
    {
        table.Add("CEILING", 2, 2, Ceiling);
        table.Add("CEILING.MATH", 1, 3, call => MathRounding(call, true));
        table.Add("CEILING.PRECISE", 1, 2, call => PreciseRounding(call, true));
        table.Add("ISO.CEILING", 1, 2, call => PreciseRounding(call, true));
        table.Add("FLOOR", 2, 2, Floor);
        table.Add("FLOOR.MATH", 1, 3, call => MathRounding(call, false));
        table.Add("FLOOR.PRECISE", 1, 2, call => PreciseRounding(call, false));
        table.Add("INT", 1, 1, call => Unary(call, Math.Floor));
        table.Add("TRUNC", 1, 2, call => RoundAtDigits(call, false, true));
        table.Add("ROUNDDOWN", 2, 2, call => RoundAtDigits(call, false, false));
        table.Add("ROUNDUP", 2, 2, call => RoundAtDigits(call, true, false));
        table.Add("MROUND", 2, 2, MRound);
        table.Add("EVEN", 1, 1, call => Unary(call, Even));
        table.Add("ODD", 1, 1, call => Unary(call, Odd));
        table.Add("SIGN", 1, 1, call => Unary(call, number => Math.Sign(number)));
    }

    private static void RegisterPowersAndIntegers(FunctionTable table)
    {
        table.Add("EXP", 1, 1, call => Unary(call, Math.Exp));
        table.Add("LN", 1, 1, call => Logarithm(call, Math.E));
        table.Add("LOG10", 1, 1, call => Logarithm(call, 10));
        table.Add("LOG", 1, 2, Log);
        table.Add("POWER", 2, 2, Power);
        table.Add("SQRT", 1, 1, call => Unary(call, Math.Sqrt));
        table.Add("SQRTPI", 1, 1, call => Unary(call, number => Math.Sqrt(number * Math.PI)));
        table.Add("MOD", 2, 2, Mod);
        table.Add("QUOTIENT", 2, 2, Quotient);
        table.Add("GCD", 1, FunctionTable.Many, call => Divisors(call, true));
        table.Add("LCM", 1, FunctionTable.Many, call => Divisors(call, false));
        table.Add("FACT", 1, 1, Fact);
        table.Add("FACTDOUBLE", 1, 1, FactDouble);
        table.Add("COMBIN", 2, 2, call => Combinations(call, false));
        table.Add("COMBINA", 2, 2, call => Combinations(call, true));
        table.Add("MULTINOMIAL", 1, FunctionTable.Many, Multinomial);
    }

    private static void RegisterSums(FunctionTable table)
    {
        table.Add("PRODUCT", 1, FunctionTable.Many, Product);
        table.Add("SUMSQ", 1, FunctionTable.Many, SumSquares);
        table.Add("SUMPRODUCT", 1, FunctionTable.Many, SumProduct);
        table.Add("SUMIF", 2, 3, StatisticalFunctionsAggregates.SumIf);
        table.Add("SUMIFS", 3, FunctionTable.Many, StatisticalFunctionsAggregates.SumIfs);
        table.Add("SUMX2MY2", 2, 2, call => PairSum(call, (left, right) => (left * left) - (right * right)));
        table.Add("SUMX2PY2", 2, 2, call => PairSum(call, (left, right) => (left * left) + (right * right)));
        table.Add("SUMXMY2", 2, 2, call => PairSum(call, (left, right) => (left - right) * (left - right)));
        table.Add("SERIESSUM", 4, 4, SeriesSum);
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

    /// <summary>A function of one number that returns its result as a Value, for the cases with their own errors.</summary>
    private static Value UnaryValue(FunctionCall call, Func<double, Value> function)
    {
        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        return function(number);
    }

    /// <summary>The number as Excel shows it: rounded to 15 significant digits, which removes binary noise such as 0.30000000000000004.</summary>
    private static double Shown(double number)
    {
        if (double.IsFinite(number) == false)
        {
            return number;
        }

        return double.Parse(number.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    }

    // ---- Trigonometry ----

    /// <summary>SIN, COS, TAN, or (with <paramref name="reciprocal"/>) CSC, SEC, COT: #NUM! for |x| ≥ 2^27, #DIV/0! where the reciprocal is undefined.</summary>
    private static Value Periodic(FunctionCall call, Func<double, double> function, bool reciprocal)
    {
        return UnaryValue(call, number =>
        {
            if (Math.Abs(number) >= TrigonometryLimit)
            {
                return Value.Num;
            }

            double result = function(number);

            if (reciprocal == false)
            {
                return Value.FromNumber(result);
            }

            if (result == 0)
            {
                return Value.Div0;
            }

            return Value.FromNumber(1 / result);
        });
    }

    /// <summary>COTH, CSCH, SECH: one over the hyperbolic function, #DIV/0! where that is zero.</summary>
    private static Value Reciprocal(FunctionCall call, Func<double, double> function)
    {
        return UnaryValue(call, number =>
        {
            double result = function(number);

            if (result == 0)
            {
                return Value.Div0;
            }

            return Value.FromNumber(1 / result);
        });
    }

    /// <summary>ATAN2(x, y): the angle of the point (x, y); #DIV/0! when both are zero.</summary>
    private static Value Atan2(FunctionCall call)
    {
        if (call.TryNumber(0, out double horizontal, out Value error) == false || call.TryNumber(1, out double vertical, out error) == false)
        {
            return error;
        }

        if (horizontal == 0 && vertical == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(Math.Atan2(vertical, horizontal));
    }

    private static Value Acoth(FunctionCall call)
    {
        return UnaryValue(call, number =>
        {
            if (Math.Abs(number) <= 1)
            {
                return Value.Num;
            }

            return Value.FromNumber(0.5 * Math.Log((number + 1) / (number - 1)));
        });
    }

    // ---- Rounding ----

    /// <summary>Rounds to a multiple of <paramref name="significance"/> with <paramref name="rounding"/> applied to the quotient.</summary>
    private static double ToMultiple(double number, double significance, Func<double, double> rounding)
    {
        double ratio = Shown(number / significance);

        return Shown(rounding(ratio) * significance);
    }

    /// <summary>
    /// CEILING(number, significance), Excel 2010 rules: a positive number with a negative
    /// significance is #NUM!; a negative number with a negative significance rounds away from zero.
    /// </summary>
    private static Value Ceiling(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double significance, out error) == false)
        {
            return error;
        }

        if (number == 0 || significance == 0)
        {
            return Value.FromNumber(0);
        }

        if (number > 0 && significance < 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(ToMultiple(number, significance, Math.Ceiling));
    }

    /// <summary>
    /// FLOOR(number, significance): significance 0 is #DIV/0!; a positive number with a negative
    /// significance is #NUM!; a negative number with a negative significance rounds toward zero.
    /// </summary>
    private static Value Floor(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double significance, out error) == false)
        {
            return error;
        }

        if (number == 0)
        {
            return Value.FromNumber(0);
        }

        if (significance == 0)
        {
            return Value.Div0;
        }

        if (number > 0 && significance < 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(ToMultiple(number, significance, Math.Floor));
    }

    /// <summary>
    /// CEILING.MATH and FLOOR.MATH: the significance's sign is ignored (default 1). For negative
    /// numbers a non-zero mode flips the direction: CEILING.MATH away from zero, FLOOR.MATH toward it.
    /// </summary>
    private static Value MathRounding(FunctionCall call, bool up)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, 1, out double significance, out error) == false || call.TryNumber(2, 0, out double mode, out error) == false)
        {
            return error;
        }

        significance = Math.Abs(significance);

        if (significance == 0 || number == 0)
        {
            return Value.FromNumber(0);
        }

        bool towardPositive = up;

        if (number < 0 && mode != 0)
        {
            towardPositive = up == false;
        }

        if (towardPositive == true)
        {
            return Value.FromNumber(ToMultiple(number, significance, Math.Ceiling));
        }

        return Value.FromNumber(ToMultiple(number, significance, Math.Floor));
    }

    /// <summary>CEILING.PRECISE, ISO.CEILING and FLOOR.PRECISE: toward +∞ or −∞ whatever the significance's sign; significance 0 gives 0.</summary>
    private static Value PreciseRounding(FunctionCall call, bool up)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, 1, out double significance, out error) == false)
        {
            return error;
        }

        significance = Math.Abs(significance);

        if (significance == 0 || number == 0)
        {
            return Value.FromNumber(0);
        }

        if (up == true)
        {
            return Value.FromNumber(ToMultiple(number, significance, Math.Ceiling));
        }

        return Value.FromNumber(ToMultiple(number, significance, Math.Floor));
    }

    /// <summary>ROUNDUP, ROUNDDOWN and TRUNC (digits optional, default 0), at 15 significant digits as Excel works.</summary>
    private static Value RoundAtDigits(FunctionCall call, bool awayFromZero, bool digitsOptional)
    {
        long digits;
        Value error;

        if (call.TryNumber(0, out double number, out error) == false)
        {
            return error;
        }

        if (digitsOptional == true)
        {
            if (call.TryInteger(1, 0, out digits, out error) == false)
            {
                return error;
            }
        }
        else if (call.TryInteger(1, out digits, out error) == false)
        {
            return error;
        }

        int places = (int)Math.Clamp(digits, -300, 300);
        double shown = Shown(number);

        if (shown == 0)
        {
            return Value.FromNumber(0);
        }

        // Already whole at this many digits: nothing to cut.
        if (Math.Abs(shown) * Math.Pow(10, places) >= 1e15)
        {
            return Value.FromNumber(shown);
        }

        if (awayFromZero == true)
        {
            return Value.FromNumber(Shown(Numeric.RoundUp(shown, places)));
        }

        return Value.FromNumber(Shown(Numeric.RoundDown(shown, places)));
    }

    /// <summary>MROUND: half away from zero to a multiple; #NUM! when number and multiple have different signs.</summary>
    private static Value MRound(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double multiple, out error) == false)
        {
            return error;
        }

        if (multiple == 0 || number == 0)
        {
            return Value.FromNumber(0);
        }

        if (Math.Sign(number) != Math.Sign(multiple))
        {
            return Value.Num;
        }

        return Value.FromNumber(Shown(Numeric.Round(number / multiple, 0) * multiple));
    }

    /// <summary>EVEN: away from zero to the next even integer.</summary>
    private static double Even(double number)
    {
        double magnitude = Math.Ceiling(Shown(Math.Abs(number)) / 2) * 2;

        return Math.Sign(number) * magnitude;
    }

    /// <summary>ODD: away from zero to the next odd integer; ODD(0) is 1.</summary>
    private static double Odd(double number)
    {
        double magnitude = Math.Ceiling(Shown(Math.Abs(number)));

        if (magnitude % 2 == 0)
        {
            magnitude++;
        }

        if (number < 0)
        {
            return -magnitude;
        }

        return magnitude;
    }

    // ---- Powers, logarithms and integers ----

    /// <summary>LN and LOG10: #NUM! for zero and negative numbers.</summary>
    private static Value Logarithm(FunctionCall call, double logBase)
    {
        return UnaryValue(call, number =>
        {
            if (number <= 0)
            {
                return Value.Num;
            }

            return Value.FromNumber(Math.Log(number, logBase));
        });
    }

    /// <summary>LOG(number, [base=10]): #NUM! for a non-positive number or base, #DIV/0! for base 1.</summary>
    private static Value Log(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, 10, out double logBase, out error) == false)
        {
            return error;
        }

        if (number <= 0 || logBase <= 0)
        {
            return Value.Num;
        }

        if (logBase == 1)
        {
            return Value.Div0;
        }

        return Value.FromNumber(Math.Log(number) / Math.Log(logBase));
    }

    /// <summary>
    /// POWER: 0^0 is #NUM!, 0 to a negative power #DIV/0!. A negative base with a fractional power
    /// is #NUM! unless the power is one over an odd integer, as in POWER(-8,1/3) = -2.
    /// </summary>
    private static Value Power(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double power, out error) == false)
        {
            return error;
        }

        if (number == 0 && power == 0)
        {
            return Value.Num;
        }

        if (number == 0 && power < 0)
        {
            return Value.Div0;
        }

        if (number < 0 && power != Math.Truncate(power))
        {
            double root = Math.Round(1 / power);

            if (Math.Abs((1 / power) - root) < 1e-10 && Math.Abs(root % 2) == 1)
            {
                return Value.FromNumber(-Math.Pow(-number, power));
            }

            return Value.Num;
        }

        return Value.FromNumber(Math.Pow(number, power));
    }

    /// <summary>MOD: the result has the divisor's sign; #DIV/0! for divisor 0.</summary>
    private static Value Mod(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double divisor, out error) == false)
        {
            return error;
        }

        if (divisor == 0)
        {
            return Value.Div0;
        }

        double quotient = Math.Floor(number / divisor);

        if (Math.Abs(quotient) >= 1.125899906842624E+15)
        {
            return Value.Num;
        }

        return Value.FromNumber(number - (divisor * quotient));
    }

    private static Value Quotient(FunctionCall call)
    {
        if (call.TryNumber(0, out double numerator, out Value error) == false || call.TryNumber(1, out double denominator, out error) == false)
        {
            return error;
        }

        if (denominator == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(Math.Truncate(numerator / denominator));
    }

    /// <summary>GCD and LCM of truncated numbers: #NUM! for negatives and values of 2^53 or more; LCM with a 0 is 0.</summary>
    private static Value Divisors(FunctionCall call, bool greatest)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        double result = 0;

        if (greatest == false)
        {
            result = 1;
        }

        foreach (double item in numbers)
        {
            double whole = Math.Truncate(item);

            if (whole < 0 || whole >= ExactLimit)
            {
                return Value.Num;
            }

            if (greatest == true)
            {
                result = GreatestDivisor(result, whole);

                continue;
            }

            if (whole == 0 || result == 0)
            {
                result = 0;

                continue;
            }

            result = result / GreatestDivisor(result, whole) * whole;
        }

        return Value.FromNumber(result);
    }

    private static double GreatestDivisor(double first, double second)
    {
        while (second != 0)
        {
            double remainder = first % second;
            first = second;
            second = remainder;
        }

        return first;
    }

    /// <summary>FACT: the factorial of the truncated number; #NUM! for negatives and above 170.</summary>
    private static Value Fact(FunctionCall call)
    {
        return UnaryValue(call, number =>
        {
            double whole = Math.Truncate(number);

            if (whole < 0 || whole > 170)
            {
                return Value.Num;
            }

            double result = 1;

            for (int factor = 2; factor <= whole; factor++)
            {
                result *= factor;
            }

            return Value.FromNumber(result);
        });
    }

    /// <summary>FACTDOUBLE: n·(n−2)·(n−4)…; #NUM! below −1; FACTDOUBLE(0) and FACTDOUBLE(−1) are 1.</summary>
    private static Value FactDouble(FunctionCall call)
    {
        return UnaryValue(call, number =>
        {
            double whole = Math.Truncate(number);

            if (whole < -1 || whole > 1000)
            {
                return Value.Num;
            }

            double result = 1;

            for (double factor = whole; factor > 1; factor -= 2)
            {
                result *= factor;
            }

            return Value.FromNumber(result);
        });
    }

    /// <summary>The binomial coefficient, rounded to a whole number when it is exactly representable.</summary>
    private static double Binomial(double total, double chosen)
    {
        double smaller = Math.Min(chosen, total - chosen);
        double result = 1;

        for (double step = 1; step <= smaller; step++)
        {
            result = result * (total - smaller + step) / step;

            if (double.IsFinite(result) == false)
            {
                return result;
            }
        }

        if (result < ExactLimit)
        {
            return Math.Round(result);
        }

        return result;
    }

    /// <summary>COMBIN(n, k) and COMBINA(n, k) (with repetitions, C(n+k−1, k)) on truncated arguments.</summary>
    private static Value Combinations(FunctionCall call, bool withRepetition)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double chosen, out error) == false)
        {
            return error;
        }

        number = Math.Truncate(number);
        chosen = Math.Truncate(chosen);

        if (number < 0 || chosen < 0)
        {
            return Value.Num;
        }

        if (withRepetition == true)
        {
            if (chosen == 0)
            {
                return Value.FromNumber(1);
            }

            if (number == 0)
            {
                return Value.FromNumber(0);
            }

            return Value.FromNumber(Binomial(number + chosen - 1, chosen));
        }

        if (chosen > number)
        {
            return Value.Num;
        }

        return Value.FromNumber(Binomial(number, chosen));
    }

    /// <summary>MULTINOMIAL: (a+b+…)! / (a!·b!·…) on truncated numbers; #NUM! for negatives.</summary>
    private static Value Multinomial(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        double total = 0;
        double result = 1;

        foreach (double item in numbers)
        {
            double whole = Math.Truncate(item);

            if (whole < 0)
            {
                return Value.Num;
            }

            total += whole;
            result *= Binomial(total, whole);
        }

        return Value.FromNumber(result);
    }

    // ---- Sums and products ----

    /// <summary>PRODUCT: the product of the numbers; 0 when there are none.</summary>
    private static Value Product(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(ProductOf(numbers));
    }

    private static double ProductOf(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return 0;
        }

        double result = 1;

        foreach (double item in numbers)
        {
            result *= item;
        }

        return result;
    }

    private static Value SumSquares(FunctionCall call)
    {
        if (call.TryCollectNumbers(0, false, out List<double> numbers, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(numbers.Sum(item => item * item));
    }

    /// <summary>SUMPRODUCT: element-wise products of same-sized arrays (#VALUE! otherwise); anything but a number counts as 0, an error is returned.</summary>
    private static Value SumProduct(FunctionCall call)
    {
        List<ValueArray> arrays = [];

        for (int index = 0; index < call.Count; index++)
        {
            if (call.IsMissing(index) == true)
            {
                return Value.ValueError;
            }

            if (call.TryArray(index, out ValueArray array, out Value error) == false)
            {
                return error;
            }

            if (arrays.Count > 0 && (array.Rows != arrays[0].Rows || array.Columns != arrays[0].Columns))
            {
                return Value.ValueError;
            }

            arrays.Add(array);
        }

        double sum = 0;

        for (int row = 0; row < arrays[0].Rows; row++)
        {
            for (int column = 0; column < arrays[0].Columns; column++)
            {
                double product = 1;

                foreach (ValueArray array in arrays)
                {
                    Value item = array[row, column];

                    if (item.IsError == true)
                    {
                        return item;
                    }

                    if (item.Kind == ValueKind.Number)
                    {
                        product *= item.Number;
                    }
                    else
                    {
                        product = 0;
                    }
                }

                sum += product;
            }
        }

        return Value.FromNumber(sum);
    }

    /// <summary>SUMX2MY2, SUMX2PY2, SUMXMY2: pairs where both values are numbers; #N/A when the arrays hold different numbers of values.</summary>
    private static Value PairSum(FunctionCall call, Func<double, double, double> term)
    {
        if (call.TryArray(0, out ValueArray first, out Value error) == false || call.TryArray(1, out ValueArray second, out error) == false)
        {
            return error;
        }

        Value[] left = [.. first.All()];
        Value[] right = [.. second.All()];

        if (left.Length != right.Length)
        {
            return Value.NA;
        }

        double sum = 0;

        for (int index = 0; index < left.Length; index++)
        {
            if (left[index].IsError == true)
            {
                return left[index];
            }

            if (right[index].IsError == true)
            {
                return right[index];
            }

            if (left[index].Kind == ValueKind.Number && right[index].Kind == ValueKind.Number)
            {
                sum += term(left[index].Number, right[index].Number);
            }
        }

        return Value.FromNumber(sum);
    }

    /// <summary>SERIESSUM(x, n, m, coefficients): Σ aᵢ·x^(n+i·m). A coefficient that is not a number is #VALUE!.</summary>
    private static Value SeriesSum(FunctionCall call)
    {
        if (call.TryNumber(0, out double variable, out Value error) == false || call.TryNumber(1, out double first, out error) == false || call.TryNumber(2, out double step, out error) == false)
        {
            return error;
        }

        double sum = 0;
        int position = 0;

        foreach (Value item in call.Values(3))
        {
            if (item.IsError == true)
            {
                return item;
            }

            if (item.Kind != ValueKind.Number)
            {
                return Value.ValueError;
            }

            sum += item.Number * Math.Pow(variable, first + (position * step));
            position++;
        }

        return Value.FromNumber(sum);
    }

    // ---- Random numbers ----

    /// <summary>RANDBETWEEN: a whole number from bottom (rounded up) to top (rounded down); #NUM! when there is none.</summary>
    private static Value RandBetween(FunctionCall call)
    {
        if (call.TryNumber(0, out double bottom, out Value error) == false || call.TryNumber(1, out double top, out error) == false)
        {
            return error;
        }

        bottom = Math.Ceiling(bottom);
        top = Math.Floor(top);

        if (bottom > top)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Min(top, bottom + Math.Floor(Random.Shared.NextDouble() * (top - bottom + 1))));
    }

    // ---- SUBTOTAL and AGGREGATE ----

    /// <summary>
    /// SUBTOTAL(function, ref1, …): 1–11 and 101–111 (AVERAGE, COUNT, COUNTA, MAX, MIN, PRODUCT,
    /// STDEV, STDEVP, SUM, VAR, VARP). Cells holding SUBTOTAL or AGGREGATE are skipped. Hidden rows
    /// are not tracked, so the 1xx forms behave as 1–11.
    /// </summary>
    private static Value Subtotal(FunctionCall call)
    {
        if (call.TryInteger(0, out long function, out Value error) == false)
        {
            return error;
        }

        if (function > 100)
        {
            function -= 100;
        }

        if (function < 1 || function > 11)
        {
            return Value.ValueError;
        }

        if (TryCollectItems(call, 1, call.Count, true, out List<Value> items, out error) == false)
        {
            return error;
        }

        return Statistic((int)function, items, false, 0);
    }

    /// <summary>
    /// AGGREGATE(function, options, ref1, …) for functions 1–13, and AGGREGATE(14, options, array, k)
    /// for LARGE. Options 0–3 skip nested SUBTOTAL/AGGREGATE cells; 2, 3, 6 and 7 ignore errors.
    /// Hidden rows are not tracked.
    /// </summary>
    private static Value Aggregate(FunctionCall call)
    {
        if (call.TryInteger(0, out long function, out Value error) == false || call.TryInteger(1, out long options, out error) == false)
        {
            return error;
        }

        if (function < 1 || function > 14 || options < 0 || options > 7)
        {
            return Value.ValueError;
        }

        bool ignoreErrors = options is 2 or 3 or 6 or 7;
        bool skipNested = options <= 3;
        int last = call.Count;
        double rank = 0;

        if (function == 14)
        {
            if (call.Count != 4 || call.IsMissing(3) == true)
            {
                return Value.ValueError;
            }

            if (call.TryNumber(3, out rank, out error) == false)
            {
                return error;
            }

            last = 3;
        }

        if (TryCollectItems(call, 2, last, skipNested, out List<Value> items, out error) == false)
        {
            return error;
        }

        return Statistic((int)function, items, ignoreErrors, rank);
    }

    /// <summary>The values of arguments first to last (exclusive), cell by cell; optionally without cells that hold SUBTOTAL or AGGREGATE.</summary>
    private static bool TryCollectItems(FunctionCall call, int first, int last, bool skipNested, out List<Value> items, out Value error)
    {
        items = [];
        error = Value.Blank;

        for (int index = first; index < last; index++)
        {
            if (call.IsMissing(index) == true)
            {
                continue;
            }

            Value raw = call.Raw(index);

            if (raw.IsError == true)
            {
                error = raw;

                return false;
            }

            if (raw.Kind != ValueKind.Reference)
            {
                items.AddRange(call.Evaluator.ToValueArray(raw).All());

                continue;
            }

            RangeReference range = raw.Reference!;

            for (int row = range.FirstRow; row <= range.UsedLastRow; row++)
            {
                for (int column = range.FirstColumn; column <= range.UsedLastColumn; column++)
                {
                    if (skipNested == true && IsNestedSubtotal(range.Sheet[row, column].Formula) == true)
                    {
                        continue;
                    }

                    items.Add(call.Evaluator.CellValue(range.Sheet, row, column));
                }
            }
        }

        return true;
    }

    private static bool IsNestedSubtotal(string? formula)
    {
        if (formula is null)
        {
            return false;
        }

        return formula.Contains("SUBTOTAL(", StringComparison.OrdinalIgnoreCase) == true || formula.Contains("AGGREGATE(", StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>The statistic numbered as SUBTOTAL and AGGREGATE number them, over the collected values.</summary>
    private static Value Statistic(int function, List<Value> items, bool ignoreErrors, double rank)
    {
        if (function == 2)
        {
            return Value.FromNumber(items.Count(item => item.Kind == ValueKind.Number));
        }

        if (function == 3)
        {
            return Value.FromNumber(items.Count(item => item.IsBlank == false && (ignoreErrors == false || item.IsError == false)));
        }

        List<double> numbers = [];

        foreach (Value item in items)
        {
            if (item.IsError == true)
            {
                if (ignoreErrors == true)
                {
                    continue;
                }

                return item;
            }

            if (item.Kind == ValueKind.Number)
            {
                numbers.Add(item.Number);
            }
        }

        switch (function)
        {
            case 1:
                if (numbers.Count == 0)
                {
                    return Value.Div0;
                }

                return Value.FromNumber(numbers.Average());

            case 4:
                if (numbers.Count == 0)
                {
                    return Value.FromNumber(0);
                }

                return Value.FromNumber(numbers.Max());

            case 5:
                if (numbers.Count == 0)
                {
                    return Value.FromNumber(0);
                }

                return Value.FromNumber(numbers.Min());

            case 6:
                return Value.FromNumber(ProductOf(numbers));

            case 7:
                return Deviation(numbers, true, true);

            case 8:
                return Deviation(numbers, false, true);

            case 9:
                return Value.FromNumber(numbers.Sum());

            case 10:
                return Deviation(numbers, true, false);

            case 11:
                return Deviation(numbers, false, false);

            case 12:
                return Median(numbers);

            case 13:
                return Mode(numbers);

            default:
                return Large(numbers, rank);
        }
    }

    /// <summary>Variance or standard deviation, of a sample (n−1) or a population (n); #DIV/0! with too few numbers.</summary>
    private static Value Deviation(List<double> numbers, bool sample, bool squareRoot)
    {
        int divisor = numbers.Count;

        if (sample == true)
        {
            divisor--;
        }

        if (divisor < 1)
        {
            return Value.Div0;
        }

        double mean = numbers.Average();
        double squares = numbers.Sum(item => (item - mean) * (item - mean));
        double variance = squares / divisor;

        if (squareRoot == true)
        {
            return Value.FromNumber(Math.Sqrt(variance));
        }

        return Value.FromNumber(variance);
    }

    private static Value Median(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return Value.Num;
        }

        List<double> sorted = [.. numbers.Order()];
        int middle = sorted.Count / 2;

        if (sorted.Count % 2 == 1)
        {
            return Value.FromNumber(sorted[middle]);
        }

        return Value.FromNumber((sorted[middle - 1] + sorted[middle]) / 2);
    }

    /// <summary>MODE.SNGL: the most frequent number, the first one met on a tie; #N/A when nothing repeats.</summary>
    private static Value Mode(List<double> numbers)
    {
        double best = 0;
        int bestCount = 1;

        for (int index = 0; index < numbers.Count; index++)
        {
            int count = 0;

            for (int other = index; other < numbers.Count; other++)
            {
                if (numbers[other] == numbers[index])
                {
                    count++;
                }
            }

            if (count > bestCount)
            {
                best = numbers[index];
                bestCount = count;
            }
        }

        if (bestCount < 2)
        {
            return Value.NA;
        }

        return Value.FromNumber(best);
    }

    /// <summary>LARGE: the k-th largest number, k rounded up; #NUM! when k is out of range.</summary>
    private static Value Large(List<double> numbers, double rank)
    {
        if (rank <= 0 || Math.Ceiling(rank) > numbers.Count)
        {
            return Value.Num;
        }

        List<double> sorted = [.. numbers.OrderDescending()];

        return Value.FromNumber(sorted[(int)Math.Ceiling(rank) - 1]);
    }

    // ---- Number systems ----

    /// <summary>
    /// ROMAN(number, [form]): 0–3999 (else #VALUE!), truncated. Form 0 (or TRUE) is classic; forms
    /// 1–4 (FALSE is 4) are the increasingly concise ones, as in ROMAN(499,4) = "ID".
    /// </summary>
    private static Value Roman(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false)
        {
            return error;
        }

        long form = 0;

        if (call.IsMissing(1) == false)
        {
            Value formValue = call.Scalar(1);

            if (formValue.Kind == ValueKind.Bool)
            {
                if (formValue.Bool == false)
                {
                    form = 4;
                }
            }
            else if (call.TryInteger(1, out form, out error) == false)
            {
                return error;
            }
        }

        number = Math.Truncate(number);

        if (number < 0 || number > 3999 || form < 0 || form > 4)
        {
            return Value.ValueError;
        }

        return Value.FromText(ToRoman((int)number, (int)form));
    }

    /// <summary>Roman numerals in Excel's five forms; the concise forms let a smaller letter subtract from a further one.</summary>
    private static string ToRoman(int number, int form)
    {
        StringBuilder builder = new();
        int remaining = number;
        int maxIndex = RomanValues.Length - 1;

        for (int step = 0; step <= maxIndex / 2; step++)
        {
            int position = 2 * step;
            int digit = remaining / RomanValues[position];

            if (digit % 5 == 4)
            {
                int subtractFrom = position - 2;

                if (digit == 4)
                {
                    subtractFrom = position - 1;
                }

                int steps = 0;

                while (steps < form && position < maxIndex)
                {
                    steps++;

                    if (RomanValues[subtractFrom] - RomanValues[position + 1] <= remaining)
                    {
                        position++;
                    }
                    else
                    {
                        steps = form;
                    }
                }

                builder.Append(RomanLetters[position]).Append(RomanLetters[subtractFrom]);
                remaining += RomanValues[position] - RomanValues[subtractFrom];

                continue;
            }

            if (digit > 4)
            {
                builder.Append(RomanLetters[position - 1]);
            }

            builder.Append(RomanLetters[position], digit % 5);
            remaining %= RomanValues[position];
        }

        return builder.ToString();
    }

    /// <summary>ARABIC: Roman numerals to a number, case-insensitive, with an optional leading minus; "" is 0; anything else #VALUE!.</summary>
    private static Value Arabic(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        text = text.Trim();

        if (text.Length > 255)
        {
            return Value.ValueError;
        }

        double sign = 1;

        if (text.StartsWith('-') == true)
        {
            sign = -1;
            text = text[1..];
        }

        double total = 0;

        for (int index = 0; index < text.Length; index++)
        {
            int current = RomanDigit(text[index]);

            if (current == 0)
            {
                return Value.ValueError;
            }

            int next = 0;

            if (index + 1 < text.Length)
            {
                next = RomanDigit(text[index + 1]);
            }

            if (current < next)
            {
                total -= current;
            }
            else
            {
                total += current;
            }
        }

        return Value.FromNumber(sign * total);
    }

    private static int RomanDigit(char letter)
    {
        int position = Array.IndexOf(RomanLetters, char.ToUpperInvariant(letter));

        if (position < 0)
        {
            return 0;
        }

        return RomanValues[position];
    }

    /// <summary>BASE(number, radix, [min_length]): 0 ≤ number &lt; 2^53, radix 2–36, min_length 0–255, else #NUM!.</summary>
    private static Value ToBase(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, out double radix, out error) == false || call.TryNumber(2, 0, out double minimumLength, out error) == false)
        {
            return error;
        }

        number = Math.Truncate(number);
        radix = Math.Truncate(radix);
        minimumLength = Math.Truncate(minimumLength);

        if (number < 0 || number >= ExactLimit || radix < 2 || radix > 36 || minimumLength < 0 || minimumLength > 255)
        {
            return Value.Num;
        }

        long remaining = (long)number;
        int baseValue = (int)radix;
        StringBuilder builder = new();
        builder.Insert(0, DigitCharacter((int)(remaining % baseValue)));
        remaining /= baseValue;

        while (remaining > 0)
        {
            builder.Insert(0, DigitCharacter((int)(remaining % baseValue)));
            remaining /= baseValue;
        }

        return Value.FromText(builder.ToString().PadLeft((int)minimumLength, '0'));
    }

    private static char DigitCharacter(int digit)
    {
        if (digit < 10)
        {
            return (char)('0' + digit);
        }

        return (char)('A' + digit - 10);
    }

    /// <summary>DECIMAL(text, radix): text in base 2–36 to a number, case-insensitive; a digit outside the base or a result of 2^53 or more is #NUM!.</summary>
    private static Value FromBase(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryNumber(1, out double radix, out error) == false)
        {
            return error;
        }

        radix = Math.Truncate(radix);

        if (radix < 2 || radix > 36)
        {
            return Value.Num;
        }

        if (text.Length > 255)
        {
            return Value.ValueError;
        }

        double result = 0;

        foreach (char character in text.Trim())
        {
            int digit = DigitValue(character);

            if (digit < 0 || digit >= radix)
            {
                return Value.Num;
            }

            result = (result * radix) + digit;
        }

        if (result >= ExactLimit)
        {
            return Value.Num;
        }

        return Value.FromNumber(result);
    }

    private static int DigitValue(char character)
    {
        char upper = char.ToUpperInvariant(character);

        if (upper is >= '0' and <= '9')
        {
            return upper - '0';
        }

        if (upper is >= 'A' and <= 'Z')
        {
            return upper - 'A' + 10;
        }

        return -1;
    }
}
