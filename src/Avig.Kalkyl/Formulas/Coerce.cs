using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// Excel's conversions between values. Each takes a scalar (references already dereferenced)
/// and returns the converted value or the error that stops the calculation.
/// </summary>
internal static class Coerce
{
    private static readonly string[] DateFormats =
    [
        "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy/MM/dd", "M/d/yyyy", "M/d/yyyy H:mm", "d MMM yyyy", "d MMMM yyyy", "MMM d, yyyy",
    ];

    private static readonly string[] TimeFormats = ["H:mm", "H:mm:ss", "h:mm tt", "h:mm:ss tt"];

    /// <summary>To a number: booleans are 1/0, blank is 0, text must look like a number, a percentage, a date or a time.</summary>
    public static Value ToNumber(Value value, bool uses1904)
    {
        switch (value.Kind)
        {
            case ValueKind.Number:
                return value;

            case ValueKind.Bool:
                return Value.FromNumber(value.Number);

            case ValueKind.Blank:
                return Value.FromNumber(0);

            case ValueKind.Error:
                return value;

            case ValueKind.Text:
                if (TryParseNumber(value.Text!, uses1904, out double number) == true)
                {
                    return Value.FromNumber(number);
                }

                return Value.ValueError;

            default:
                return Value.ValueError;
        }
    }

    /// <summary>To text: numbers as Excel's General format shows them, booleans as TRUE/FALSE, blank as "".</summary>
    public static Value ToText(Value value)
    {
        switch (value.Kind)
        {
            case ValueKind.Text:
            case ValueKind.Error:
                return value;

            case ValueKind.Blank:
                return Value.FromText(string.Empty);

            case ValueKind.Number:
            case ValueKind.Bool:
                return Value.FromText(value.ToString());

            default:
                return Value.ValueError;
        }
    }

    /// <summary>To a boolean: numbers are TRUE unless 0, text must be TRUE or FALSE, blank is FALSE.</summary>
    public static Value ToBool(Value value)
    {
        switch (value.Kind)
        {
            case ValueKind.Bool:
            case ValueKind.Error:
                return value;

            case ValueKind.Number:
                return Value.FromBool(value.Number != 0);

            case ValueKind.Blank:
                return Value.False;

            case ValueKind.Text when string.Equals(value.Text, "TRUE", StringComparison.OrdinalIgnoreCase) == true:
                return Value.True;

            case ValueKind.Text when string.Equals(value.Text, "FALSE", StringComparison.OrdinalIgnoreCase) == true:
                return Value.False;

            default:
                return Value.ValueError;
        }
    }

    /// <summary>
    /// Compares two scalars the way Excel's comparison operators do: numbers sort before text,
    /// text before booleans; text is compared without regard to case; blank counts as 0, "" or
    /// FALSE depending on the other side.
    /// </summary>
    public static int Compare(Value left, Value right)
    {
        if (left.IsBlank == true)
        {
            left = BlankAs(right);
        }

        if (right.IsBlank == true)
        {
            right = BlankAs(left);
        }

        int leftRank = Rank(left);
        int rightRank = Rank(right);

        if (leftRank != rightRank)
        {
            return leftRank.CompareTo(rightRank);
        }

        if (left.Kind == ValueKind.Text)
        {
            return string.Compare(left.Text, right.Text, StringComparison.OrdinalIgnoreCase);
        }

        return left.Number.CompareTo(right.Number);
    }

    public static bool TryParseNumber(string text, bool uses1904, out double number)
    {
        string trimmed = text.Trim();
        number = 0;

        if (trimmed.Length == 0)
        {
            return false;
        }

        bool percent = false;

        if (trimmed.EndsWith('%') == true)
        {
            percent = true;
            trimmed = trimmed[..^1].TrimEnd();
        }

        if (trimmed.StartsWith('$') == true)
        {
            trimmed = trimmed[1..];
        }

        if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) == true)
        {
            if (percent == true)
            {
                number /= 100;
            }

            return double.IsFinite(number);
        }

        if (percent == false && DateTime.TryParseExact(trimmed, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime date) == true)
        {
            number = Internal.ExcelDate.ToSerial(date, uses1904);

            return true;
        }

        if (percent == false && DateTime.TryParseExact(trimmed, TimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out DateTime time) == true)
        {
            number = time.TimeOfDay.TotalDays;

            return true;
        }

        return false;
    }

    private static Value BlankAs(Value other)
    {
        return other.Kind switch
        {
            ValueKind.Text => Value.FromText(string.Empty),
            ValueKind.Bool => Value.False,
            _ => Value.FromNumber(0),
        };
    }

    private static int Rank(Value value)
    {
        return value.Kind switch
        {
            ValueKind.Number => 0,
            ValueKind.Text => 1,
            ValueKind.Bool => 2,
            _ => 0,
        };
    }
}
