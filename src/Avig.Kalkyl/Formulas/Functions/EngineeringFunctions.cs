using System.Globalization;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Engineering functions: base conversion, bit operations, DELTA, GESTEP and the error function.</summary>
internal static class EngineeringFunctions
{
    private const double BitLimit = 281474976710656; // 2^48

    private const int MaxShift = 53;

    private static readonly double TwoOverSqrtPi = 2 / Math.Sqrt(Math.PI);

    /// <summary>The bases of the conversions; the value is the radix.</summary>
    private enum Radix
    {
        Binary = 2,
        Octal = 8,
        Decimal = 10,
        Hexadecimal = 16,
    }

    public static void Register(FunctionTable table)
    {
        ArgumentNullException.ThrowIfNull(table);

        table.Add("DEC2BIN", 1, 2, call => Convert(call, Radix.Decimal, Radix.Binary));
        table.Add("DEC2OCT", 1, 2, call => Convert(call, Radix.Decimal, Radix.Octal));
        table.Add("DEC2HEX", 1, 2, call => Convert(call, Radix.Decimal, Radix.Hexadecimal));
        table.Add("BIN2DEC", 1, 1, call => Convert(call, Radix.Binary, Radix.Decimal));
        table.Add("BIN2OCT", 1, 2, call => Convert(call, Radix.Binary, Radix.Octal));
        table.Add("BIN2HEX", 1, 2, call => Convert(call, Radix.Binary, Radix.Hexadecimal));
        table.Add("OCT2DEC", 1, 1, call => Convert(call, Radix.Octal, Radix.Decimal));
        table.Add("OCT2BIN", 1, 2, call => Convert(call, Radix.Octal, Radix.Binary));
        table.Add("OCT2HEX", 1, 2, call => Convert(call, Radix.Octal, Radix.Hexadecimal));
        table.Add("HEX2DEC", 1, 1, call => Convert(call, Radix.Hexadecimal, Radix.Decimal));
        table.Add("HEX2BIN", 1, 2, call => Convert(call, Radix.Hexadecimal, Radix.Binary));
        table.Add("HEX2OCT", 1, 2, call => Convert(call, Radix.Hexadecimal, Radix.Octal));
        table.Add("DELTA", 1, 2, Delta);
        table.Add("GESTEP", 1, 2, GeStep);
        table.Add("BITAND", 2, 2, call => Bitwise(call, (left, right) => left & right));
        table.Add("BITOR", 2, 2, call => Bitwise(call, (left, right) => left | right));
        table.Add("BITXOR", 2, 2, call => Bitwise(call, (left, right) => left ^ right));
        table.Add("BITLSHIFT", 2, 2, call => Shift(call, 1));
        table.Add("BITRSHIFT", 2, 2, call => Shift(call, -1));
        table.Add("ERF", 1, 2, ErfRange);
        table.Add("ERF.PRECISE", 1, 1, call => MathFunctions.Unary(call, Erf));
        table.Add("ERFC", 1, 1, call => MathFunctions.Unary(call, Erfc));
        table.Add("ERFC.PRECISE", 1, 1, call => MathFunctions.Unary(call, Erfc));
    }

    /// <summary>
    /// The base conversions. Negative numbers are 10-digit two's complement (DEC2BIN(-1) is
    /// 1111111111) and ignore places; out of the target's range, invalid digits, more than 10
    /// digits or too few places give #NUM!. Numbers and text are both accepted.
    /// </summary>
    private static Value Convert(FunctionCall call, Radix source, Radix target)
    {
        long number;
        Value error;

        if (source == Radix.Decimal)
        {
            if (call.TryNumber(0, out double value, out error) == false)
            {
                return error;
            }

            value = Math.Truncate(value);

            if (value < Minimum(target) || value > Maximum(target))
            {
                return Value.Num;
            }

            number = (long)value;
        }
        else if (TryParse(call, source, out number, out error) == false)
        {
            return error;
        }

        if (target == Radix.Decimal)
        {
            return Value.FromNumber(number);
        }

        if (number < Minimum(target) || number > Maximum(target))
        {
            return Value.Num;
        }

        if (number < 0)
        {
            long complement = number + (long)Math.Pow((int)target, 10);

            return Value.FromText(Digits(complement, target));
        }

        string digits = Digits(number, target);

        if (call.IsMissing(1) == true)
        {
            return Value.FromText(digits);
        }

        if (call.TryNumber(1, out double placesNumber, out error) == false)
        {
            return error;
        }

        double places = Math.Truncate(placesNumber);

        if (places < digits.Length || places > 10)
        {
            return Value.Num;
        }

        return Value.FromText(digits.PadLeft((int)places, '0'));
    }

    /// <summary>A binary, octal or hexadecimal argument: at most 10 digits, the tenth digit's top bit giving the sign.</summary>
    private static bool TryParse(FunctionCall call, Radix source, out long number, out Value error)
    {
        number = 0;
        Value scalar = call.Scalar(0);

        if (scalar.Kind == ValueKind.Bool)
        {
            error = Value.ValueError;

            return false;
        }

        if (call.TryText(0, out string text, out error) == false)
        {
            return false;
        }

        if (text.Length > 10)
        {
            error = Value.Num;

            return false;
        }

        int radix = (int)source;

        foreach (char character in text)
        {
            int digit = DigitValue(character);

            if (digit < 0 || digit >= radix)
            {
                error = Value.Num;

                return false;
            }

            number = (number * radix) + digit;
        }

        long full = (long)Math.Pow(radix, 10);

        if (text.Length == 10 && number >= full / 2)
        {
            number -= full;
        }

        return true;
    }

    private static int DigitValue(char character)
    {
        if (character >= '0' && character <= '9')
        {
            return character - '0';
        }

        char upper = char.ToUpperInvariant(character);

        if (upper >= 'A' && upper <= 'F')
        {
            return upper - 'A' + 10;
        }

        return -1;
    }

    private static string Digits(long number, Radix radix)
    {
        if (radix == Radix.Octal)
        {
            return System.Convert.ToString(number, 8);
        }

        if (radix == Radix.Binary)
        {
            return System.Convert.ToString(number, 2);
        }

        return number.ToString("X", CultureInfo.InvariantCulture);
    }

    private static long Minimum(Radix radix)
    {
        return radix switch
        {
            Radix.Binary => -512,
            Radix.Octal => -536870912,
            Radix.Hexadecimal => -549755813888,
            _ => long.MinValue,
        };
    }

    private static long Maximum(Radix radix)
    {
        return radix switch
        {
            Radix.Binary => 511,
            Radix.Octal => 536870911,
            Radix.Hexadecimal => 549755813887,
            _ => long.MaxValue,
        };
    }

    private static Value Delta(FunctionCall call)
    {
        if (call.TryNumber(0, out double first, out Value error) == false || call.TryNumber(1, 0, out double second, out error) == false)
        {
            return error;
        }

        if (first == second)
        {
            return Value.FromNumber(1);
        }

        return Value.FromNumber(0);
    }

    private static Value GeStep(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryNumber(1, 0, out double step, out error) == false)
        {
            return error;
        }

        if (number >= step)
        {
            return Value.FromNumber(1);
        }

        return Value.FromNumber(0);
    }

    /// <summary>The argument of a BIT function: a whole number from 0 up to 2^48 - 1, #NUM! otherwise.</summary>
    private static bool TryBits(FunctionCall call, int index, out long bits, out Value error)
    {
        bits = 0;

        if (call.TryNumber(index, out double number, out error) == false)
        {
            return false;
        }

        if (number < 0 || number >= BitLimit || number != Math.Floor(number))
        {
            error = Value.Num;

            return false;
        }

        bits = (long)number;

        return true;
    }

    private static Value Bitwise(FunctionCall call, Func<long, long, long> operation)
    {
        if (TryBits(call, 0, out long left, out Value error) == false || TryBits(call, 1, out long right, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(operation(left, right));
    }

    /// <summary>BITLSHIFT and BITRSHIFT: a negative amount shifts the other way; more than 53 places gives #NUM!.</summary>
    private static Value Shift(FunctionCall call, int direction)
    {
        if (TryBits(call, 0, out long bits, out Value error) == false || call.TryNumber(1, out double amountNumber, out error) == false)
        {
            return error;
        }

        double amount = Math.Truncate(amountNumber) * direction;

        if (Math.Abs(amount) > MaxShift)
        {
            return Value.Num;
        }

        if (amount >= 0)
        {
            return Value.FromNumber(bits * Math.Pow(2, amount));
        }

        return Value.FromNumber(bits >> (int)-amount);
    }

    /// <summary>ERF(lower) is erf(lower); ERF(lower, upper) is erf(upper) - erf(lower).</summary>
    private static Value ErfRange(FunctionCall call)
    {
        if (call.TryNumber(0, out double lower, out Value error) == false)
        {
            return error;
        }

        if (call.IsMissing(1) == true)
        {
            return Value.FromNumber(Erf(lower));
        }

        if (call.TryNumber(1, out double upper, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(Erf(upper) - Erf(lower));
    }

    /// <summary>The error function to full double precision: a series below 1.5, the complement's continued fraction above.</summary>
    internal static double Erf(double value)
    {
        double absolute = Math.Abs(value);
        double result;

        if (absolute < 1.5)
        {
            result = ErfSeries(absolute);
        }
        else
        {
            result = 1 - ErfcFraction(absolute);
        }

        if (value < 0)
        {
            return -result;
        }

        return result;
    }

    /// <summary>The complementary error function, without losing precision in the tail.</summary>
    internal static double Erfc(double value)
    {
        if (value < 0)
        {
            return 2 - Erfc(-value);
        }

        if (value < 1.5)
        {
            return 1 - ErfSeries(value);
        }

        return ErfcFraction(value);
    }

    /// <summary>erf(x) = 2/sqrt(pi) exp(-x^2) sum (2x^2)^n x / (1*3*...*(2n+1)): all terms positive, so no cancellation.</summary>
    private static double ErfSeries(double value)
    {
        double square = value * value;
        double term = value;
        double sum = value;

        for (int index = 1; index < 500; index++)
        {
            term *= 2 * square / ((2 * index) + 1);
            sum += term;

            if (term < sum * 1e-17)
            {
                break;
            }
        }

        return TwoOverSqrtPi * Math.Exp(-square) * sum;
    }

    /// <summary>erfc(x) = exp(-x^2)/sqrt(pi) * 2x / (2x^2+1 - 1*2/(2x^2+5 - 3*4/(2x^2+9 - ...))), evaluated with Lentz's method.</summary>
    private static double ErfcFraction(double value)
    {
        const double Tiny = 1e-300;
        double square = value * value;
        double fraction = (2 * square) + 1;
        double numerator = fraction;
        double denominator = 0;

        for (int index = 1; index < 5000; index++)
        {
            double coefficient = -((2.0 * index) - 1) * (2.0 * index);
            double term = (2 * square) + 1 + (4 * index);
            denominator = term + (coefficient * denominator);

            if (Math.Abs(denominator) < Tiny)
            {
                denominator = Tiny;
            }

            numerator = term + (coefficient / numerator);

            if (Math.Abs(numerator) < Tiny)
            {
                numerator = Tiny;
            }

            denominator = 1 / denominator;
            double delta = numerator * denominator;
            fraction *= delta;

            if (Math.Abs(delta - 1) < 1e-16)
            {
                break;
            }
        }

        return Math.Exp(-square) / Math.Sqrt(Math.PI) * 2 * value / fraction;
    }
}
