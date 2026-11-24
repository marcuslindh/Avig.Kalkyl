using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// Excel's conversions between values. Each takes a scalar (references already dereferenced)
/// and returns the converted value or the error that stops the calculation.
/// </summary>
internal static class Coerce
{
    /// <summary>ISO dates, which Excel reads in every culture. Everything else is read in the workbook's culture.</summary>
    private static readonly string[] IsoFormats =
    [
        "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd",
    ];

    /// <summary>To a number: booleans are 1/0, blank is 0, text must look like a number, a percentage, a date or a time.</summary>
    public static Value ToNumber(Value value, TextConversion conversion)
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
                if (TryParseNumber(value.Text!, conversion, out double number) == true)
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

    /// <summary>
    /// Text as a number the way Excel reads what is typed into a cell: a number, a percentage
    /// or an amount with the currency symbol in the workbook's culture, a date or a time in that
    /// culture's order ("31/12/2018" in en-GB, "12/31/2018" in en-US), or an ISO date in any culture.
    /// </summary>
    public static bool TryParseNumber(string text, TextConversion conversion, out double number)
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

        string currency = conversion.Culture.NumberFormat.CurrencySymbol;

        if (trimmed.StartsWith('$') == true)
        {
            trimmed = trimmed[1..].TrimStart();
        }
        else if (currency.Length > 0 && trimmed.StartsWith(currency, StringComparison.Ordinal) == true)
        {
            trimmed = trimmed[currency.Length..].TrimStart();
        }
        else if (currency.Length > 0 && trimmed.EndsWith(currency, StringComparison.Ordinal) == true)
        {
            trimmed = trimmed[..^currency.Length].TrimEnd();
        }

        if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, conversion.Culture, out number) == true)
        {
            if (percent == true)
            {
                number /= 100;
            }

            return double.IsFinite(number);
        }

        if (percent == true)
        {
            return false;
        }

        if (DateTime.TryParseExact(trimmed, IsoFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime date) == true)
        {
            number = Internal.ExcelDate.ToSerial(date, conversion.Uses1904);

            return true;
        }

        // Excel does not read ISO 8601 with T or a zone (1975-02-23T02:58:41.000Z gives #VALUE!),
        // while DateTime.TryParse does, and moves a Z time into the machine's time zone.
        if (IsTimestamp(trimmed) == true)
        {
            return false;
        }

        // NoCurrentDateDefault: a time on its own ("8:30") gets day zero rather than today, and is a fraction of a day.
        if (DateTime.TryParse(trimmed, conversion.Culture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out date) == true)
        {
            if (date.Date == DateTime.MinValue.Date)
            {
                number = date.TimeOfDay.TotalDays;

                return true;
            }

            number = Internal.ExcelDate.ToSerial(date, conversion.Uses1904);

            return number >= 0;
        }

        return false;
    }

    /// <summary>True for an ISO 8601 timestamp: a T between a digit and a digit, or a trailing Z.</summary>
    private static bool IsTimestamp(string text)
    {
        if (text.EndsWith('Z') == true || text.EndsWith('z') == true)
        {
            return true;
        }

        for (int index = 1; index < text.Length - 1; index++)
        {
            if ((text[index] == 'T' || text[index] == 't') && char.IsAsciiDigit(text[index - 1]) == true && char.IsAsciiDigit(text[index + 1]) == true)
            {
                return true;
            }
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
