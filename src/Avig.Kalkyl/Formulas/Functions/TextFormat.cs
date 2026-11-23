using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Excel's number format codes as the TEXT function applies them: digit placeholders 0 # ?,
/// decimal point, thousands separator and scaling by trailing commas, %, E+ scientific, quoted
/// and backslash literals, @, up to four sections (positive;negative;zero;text), and the date and
/// time codes with invariant (English) month and day names. Colours and conditions in brackets
/// are ignored; [$€-409] gives its currency symbol.
/// </summary>
internal static class TextFormat
{
    private enum TokenKind
    {
        Literal,
        Placeholder,
        DecimalPoint,
        Comma,
        Percent,
        Exponent,
        TextValue,
        DateCode,
        AmPm,
        Elapsed,
        General,
    }

    private readonly record struct Token(TokenKind Kind, string Text);

    /// <summary>Formats a number. False (with #VALUE!) when a date format gets a number that is not a date.</summary>
    public static bool TryFormatNumber(FunctionCall call, double number, string format, out string text, out Value error)
    {
        text = string.Empty;
        error = Value.Blank;
        List<string> sections = SplitSections(format);
        string section = sections[0];
        bool minus = false;
        double magnitude = number;

        if (sections.Count == 1)
        {
            minus = number < 0;
            magnitude = Math.Abs(number);
        }
        else if (number < 0)
        {
            section = sections[1];
            magnitude = Math.Abs(number);
        }
        else if (number == 0 && sections.Count >= 3)
        {
            section = sections[2];
        }

        List<Token> tokens = Tokenize(section);

        if (IsDate(tokens) == true)
        {
            if (number < 0)
            {
                error = Value.ValueError;

                return false;
            }

            return TryFormatDate(call, number, tokens, out text, out error);
        }

        string formatted = FormatNumber(magnitude, tokens);

        if (minus == true && HasNumberPart(tokens) == true)
        {
            formatted = "-" + formatted;
        }

        text = formatted;

        return true;
    }

    /// <summary>Formats text that is not a number: the fourth section, or a single section with @; otherwise the text as it is.</summary>
    public static string FormatText(string value, string format)
    {
        List<string> sections = SplitSections(format);
        string section;

        if (sections.Count >= 4)
        {
            section = sections[3];
        }
        else if (sections.Count == 1 && sections[0].Contains('@', StringComparison.Ordinal) == true)
        {
            section = sections[0];
        }
        else
        {
            return value;
        }

        StringBuilder builder = new();

        foreach (Token token in Tokenize(section))
        {
            if (token.Kind == TokenKind.TextValue)
            {
                builder.Append(value);
            }
            else
            {
                builder.Append(token.Text);
            }
        }

        return builder.ToString();
    }

    private static List<string> SplitSections(string format)
    {
        List<string> sections = [];
        StringBuilder current = new();
        bool quoted = false;
        bool bracket = false;

        for (int index = 0; index < format.Length; index++)
        {
            char character = format[index];

            if (quoted == false && bracket == false && character == '\\' && index + 1 < format.Length)
            {
                current.Append(character).Append(format[index + 1]);
                index++;

                continue;
            }

            if (character == '"' && bracket == false)
            {
                quoted = quoted == false;
            }
            else if (quoted == false && character == '[')
            {
                bracket = true;
            }
            else if (quoted == false && character == ']')
            {
                bracket = false;
            }
            else if (quoted == false && bracket == false && character == ';')
            {
                sections.Add(current.ToString());
                current.Clear();

                continue;
            }

            current.Append(character);
        }

        sections.Add(current.ToString());

        return sections;
    }

    private static List<Token> Tokenize(string section)
    {
        List<Token> tokens = [];
        int index = 0;

        while (index < section.Length)
        {
            char character = section[index];
            char lower = char.ToLowerInvariant(character);

            if (character == '"')
            {
                int end = section.IndexOf('"', index + 1);

                if (end < 0)
                {
                    end = section.Length;
                }

                tokens.Add(new Token(TokenKind.Literal, section[(index + 1)..end]));
                index = end + 1;

                continue;
            }

            if (character == '\\' && index + 1 < section.Length)
            {
                tokens.Add(new Token(TokenKind.Literal, section[index + 1].ToString()));
                index += 2;

                continue;
            }

            if (character == '_' && index + 1 < section.Length)
            {
                // Padding the width of the next character: a space.
                tokens.Add(new Token(TokenKind.Literal, " "));
                index += 2;

                continue;
            }

            if (character == '*' && index + 1 < section.Length)
            {
                // Repeat-to-fill has no width to fill in text.
                index += 2;

                continue;
            }

            if (character == '[')
            {
                int end = section.IndexOf(']', index + 1);

                if (end < 0)
                {
                    end = section.Length;
                }

                AddBracket(tokens, section[(index + 1)..end]);
                index = end + 1;

                continue;
            }

            if (character is '0' or '#' or '?')
            {
                tokens.Add(new Token(TokenKind.Placeholder, character.ToString()));
                index++;

                continue;
            }

            if (character == '.')
            {
                tokens.Add(new Token(TokenKind.DecimalPoint, "."));
                index++;

                continue;
            }

            if (character == ',')
            {
                tokens.Add(new Token(TokenKind.Comma, ","));
                index++;

                continue;
            }

            if (character == '%')
            {
                tokens.Add(new Token(TokenKind.Percent, "%"));
                index++;

                continue;
            }

            if (character == '@')
            {
                tokens.Add(new Token(TokenKind.TextValue, "@"));
                index++;

                continue;
            }

            if (lower == 'e' && index + 1 < section.Length && section[index + 1] is '+' or '-')
            {
                tokens.Add(new Token(TokenKind.Exponent, section[index + 1].ToString()));
                index += 2;

                continue;
            }

            if (string.Compare(section, index, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0)
            {
                tokens.Add(new Token(TokenKind.General, "General"));
                index += 7;

                continue;
            }

            if (string.Compare(section, index, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                tokens.Add(new Token(TokenKind.AmPm, "AM/PM"));
                index += 5;

                continue;
            }

            if (string.Compare(section, index, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
            {
                tokens.Add(new Token(TokenKind.AmPm, section.Substring(index, 3)));
                index += 3;

                continue;
            }

            if (lower is 'y' or 'm' or 'd' or 'h' or 's')
            {
                int end = index;

                while (end < section.Length && char.ToLowerInvariant(section[end]) == lower)
                {
                    end++;
                }

                tokens.Add(new Token(TokenKind.DateCode, new string(lower, end - index)));
                index = end;

                continue;
            }

            tokens.Add(new Token(TokenKind.Literal, character.ToString()));
            index++;
        }

        return tokens;
    }

    private static void AddBracket(List<Token> tokens, string content)
    {
        string lower = content.ToLowerInvariant();

        if (lower.Length > 0 && (lower.Trim('h').Length == 0 || lower.Trim('m').Length == 0 || lower.Trim('s').Length == 0))
        {
            tokens.Add(new Token(TokenKind.Elapsed, lower));

            return;
        }

        if (content.StartsWith('$') == true)
        {
            int dash = content.IndexOf('-', StringComparison.Ordinal);
            string symbol = content[1..];

            if (dash >= 0)
            {
                symbol = content[1..dash];
            }

            tokens.Add(new Token(TokenKind.Literal, symbol));
        }

        // Colours, conditions and locale codes do not change the text.
    }

    private static bool IsDate(List<Token> tokens)
    {
        return tokens.Any(token => token.Kind is TokenKind.DateCode or TokenKind.AmPm or TokenKind.Elapsed);
    }

    private static bool HasNumberPart(List<Token> tokens)
    {
        return tokens.Any(token => token.Kind is TokenKind.Placeholder or TokenKind.General or TokenKind.TextValue or TokenKind.DecimalPoint);
    }

    private static string FormatNumber(double value, List<Token> tokens)
    {
        int exponentIndex = tokens.FindIndex(token => token.Kind == TokenKind.Exponent);
        int lastPlaceholder = tokens.FindLastIndex(token => token.Kind == TokenKind.Placeholder);
        int numberEnd = tokens.Count;

        if (exponentIndex >= 0)
        {
            numberEnd = exponentIndex;
        }

        int decimalIndex = tokens.FindIndex(0, numberEnd, token => token.Kind == TokenKind.DecimalPoint);
        int integerEnd = numberEnd;

        if (decimalIndex >= 0)
        {
            integerEnd = decimalIndex;
        }

        // Classify commas: between integer placeholders they group thousands, right after the last
        // placeholder of the number they divide by 1000, elsewhere they are literal.
        bool grouping = false;
        int scaling = 0;
        HashSet<int> literalCommas = [];
        int firstPlaceholder = tokens.FindIndex(token => token.Kind == TokenKind.Placeholder);

        for (int index = 0; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != TokenKind.Comma)
            {
                continue;
            }

            int next = index + 1;

            while (next < tokens.Count && tokens[next].Kind == TokenKind.Comma)
            {
                next++;
            }

            bool followedByPlaceholder = next < tokens.Count && tokens[next].Kind == TokenKind.Placeholder;

            if (followedByPlaceholder == true && firstPlaceholder >= 0 && firstPlaceholder < index && index < integerEnd)
            {
                grouping = true;
            }
            else if (followedByPlaceholder == false && firstPlaceholder >= 0 && firstPlaceholder < index && (lastPlaceholder < index || index < integerEnd))
            {
                scaling++;
            }
            else if (followedByPlaceholder == false || firstPlaceholder < 0 || firstPlaceholder > index)
            {
                literalCommas.Add(index);
            }
        }

        int percents = tokens.Count(token => token.Kind == TokenKind.Percent);
        double scaled = value * Math.Pow(100, percents) / Math.Pow(1000, scaling);
        List<int> integerSlots = [];
        List<int> fractionSlots = [];
        List<int> exponentSlots = [];

        for (int index = 0; index < tokens.Count; index++)
        {
            if (tokens[index].Kind != TokenKind.Placeholder)
            {
                continue;
            }

            if (exponentIndex >= 0 && index > exponentIndex)
            {
                exponentSlots.Add(index);
            }
            else if (index < integerEnd)
            {
                integerSlots.Add(index);
            }
            else
            {
                fractionSlots.Add(index);
            }
        }

        int exponent = 0;

        if (exponentIndex >= 0)
        {
            scaled = Mantissa(scaled, Math.Max(1, integerSlots.Count), fractionSlots.Count, out exponent);
        }

        double rounded = Numeric.Round(scaled, fractionSlots.Count);
        string digits = rounded.ToString("F" + fractionSlots.Count.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        string integerDigits = digits;
        string fractionDigits = string.Empty;
        int point = digits.IndexOf('.', StringComparison.Ordinal);

        if (point >= 0)
        {
            integerDigits = digits[..point];
            fractionDigits = digits[(point + 1)..];
        }

        if (integerDigits == "0")
        {
            integerDigits = string.Empty;
        }

        Dictionary<int, string> output = [];
        FillInteger(tokens, integerSlots, integerDigits, grouping, output);
        FillFraction(tokens, fractionSlots, fractionDigits, output);

        if (exponentIndex >= 0)
        {
            string sign = string.Empty;

            if (exponent < 0)
            {
                sign = "-";
            }
            else if (tokens[exponentIndex].Text == "+")
            {
                sign = "+";
            }

            output[exponentIndex] = "E" + sign;
            FillInteger(tokens, exponentSlots, Math.Abs(exponent).ToString(CultureInfo.InvariantCulture), false, output);
        }

        StringBuilder builder = new();

        for (int index = 0; index < tokens.Count; index++)
        {
            Token token = tokens[index];

            // Integer digits with no integer placeholder to hold them go before the decimal point.
            if (index == decimalIndex && integerSlots.Count == 0)
            {
                builder.Append(integerDigits);
            }

            if (output.TryGetValue(index, out string? filled) == true)
            {
                builder.Append(filled);

                continue;
            }

            switch (token.Kind)
            {
                case TokenKind.Comma when literalCommas.Contains(index) == true:
                    builder.Append(',');

                    break;

                case TokenKind.Comma:
                    break;

                case TokenKind.General:
                    builder.Append(Value.NumberText(value));

                    break;

                case TokenKind.TextValue:
                    builder.Append(Value.NumberText(value));

                    break;

                default:
                    builder.Append(token.Text);

                    break;
            }
        }

        return builder.ToString();
    }

    /// <summary>The mantissa for scientific format; with several integer placeholders the exponent is a multiple of their count (##0.0E+0).</summary>
    private static double Mantissa(double value, int integerPlaces, int fractionPlaces, out int exponent)
    {
        exponent = 0;

        if (value == 0)
        {
            return 0;
        }

        int magnitude = (int)Math.Floor(Math.Log10(value));

        if (integerPlaces > 1)
        {
            exponent = (int)Math.Floor(magnitude / (double)integerPlaces) * integerPlaces;
        }
        else
        {
            exponent = magnitude;
        }

        double mantissa = value / Math.Pow(10, exponent);

        if (Numeric.Round(mantissa, fractionPlaces) >= Math.Pow(10, Math.Max(1, integerPlaces)))
        {
            exponent += Math.Max(1, integerPlaces);
            mantissa = value / Math.Pow(10, exponent);
        }

        return mantissa;
    }

    private static void FillInteger(List<Token> tokens, List<int> slots, string digits, bool grouping, Dictionary<int, string> output)
    {
        if (slots.Count == 0)
        {
            return;
        }

        int digitIndex = digits.Length - 1;
        int position = 0;

        for (int slot = slots.Count - 1; slot >= 0; slot--)
        {
            StringBuilder piece = new();
            char placeholder = tokens[slots[slot]].Text[0];

            if (slot == 0)
            {
                // The leftmost placeholder takes every digit that is left.
                while (digitIndex >= 0)
                {
                    piece.Insert(0, Grouped(digits[digitIndex], position, grouping));
                    digitIndex--;
                    position++;
                }

                if (piece.Length > 0)
                {
                    output[slots[slot]] = piece.ToString();

                    continue;
                }
            }
            else if (digitIndex >= 0)
            {
                output[slots[slot]] = Grouped(digits[digitIndex], position, grouping);
                digitIndex--;
                position++;

                continue;
            }

            switch (placeholder)
            {
                case '0':
                    output[slots[slot]] = Grouped('0', position, grouping);
                    position++;

                    break;

                case '?':
                    output[slots[slot]] = " ";
                    position++;

                    break;

                default:
                    output[slots[slot]] = string.Empty;

                    break;
            }
        }
    }

    private static string Grouped(char digit, int position, bool grouping)
    {
        if (grouping == true && position > 0 && position % 3 == 0)
        {
            return digit + ",";
        }

        return digit.ToString();
    }

    private static void FillFraction(List<Token> tokens, List<int> slots, string digits, Dictionary<int, string> output)
    {
        int keep = slots.Count;

        // Trailing zeros vanish under # and become spaces under ?.
        while (keep > 0 && digits[keep - 1] == '0' && tokens[slots[keep - 1]].Text[0] != '0')
        {
            keep--;
        }

        for (int index = 0; index < slots.Count; index++)
        {
            if (index < keep)
            {
                output[slots[index]] = digits[index].ToString();
            }
            else if (tokens[slots[index]].Text[0] == '?')
            {
                output[slots[index]] = " ";
            }
            else
            {
                output[slots[index]] = string.Empty;
            }
        }
    }

    private static bool TryFormatDate(FunctionCall call, double serial, List<Token> tokens, out string text, out Value error)
    {
        text = string.Empty;
        error = Value.Blank;

        // Fractional seconds: a decimal point followed by zeros after a seconds code.
        int fractionDigits = 0;

        for (int index = 0; index + 1 < tokens.Count; index++)
        {
            if (tokens[index].Kind == TokenKind.DecimalPoint && tokens[index + 1].Kind == TokenKind.Placeholder && index > 0 && IsSeconds(tokens[index - 1]) == true)
            {
                int count = 0;

                while (index + 1 + count < tokens.Count && tokens[index + 1 + count].Kind == TokenKind.Placeholder)
                {
                    count++;
                }

                fractionDigits = Math.Min(3, count);
            }
        }

        long unitsPerSecond = (long)Math.Pow(10, fractionDigits);
        long unitsPerDay = 86_400 * unitsPerSecond;
        double totalUnitsExact = Math.Round(serial * unitsPerDay, MidpointRounding.AwayFromZero);

        if (totalUnitsExact > 3e17)
        {
            error = Value.ValueError;

            return false;
        }

        long totalUnits = (long)totalUnitsExact;
        long day = totalUnits / unitsPerDay;
        long remainder = totalUnits % unitsPerDay;
        long hour = remainder / (3600 * unitsPerSecond);
        long minute = remainder / (60 * unitsPerSecond) % 60;
        long second = remainder / unitsPerSecond % 60;
        long subSecond = remainder % unitsPerSecond;

        bool needsDate = tokens.Any(token => token.Kind == TokenKind.DateCode && token.Text[0] is 'y' or 'd');
        int year = 1900;
        int month = 1;
        int dayOfMonth = 0;
        DayOfWeek weekday = DayOfWeek.Saturday;

        List<Token> resolved = ResolveMinutes(tokens);

        if (resolved.Any(token => token.Kind == TokenKind.DateCode && token.Text[0] == 'm') == true)
        {
            needsDate = true;
        }

        if (needsDate == true && (call.Uses1904 == true || day > 0))
        {
            if (DateFunctions.TryToDate(call, day, out DateTime date, out _) == false)
            {
                error = Value.ValueError;

                return false;
            }

            year = date.Year;
            month = date.Month;
            dayOfMonth = date.Day;
            weekday = date.DayOfWeek;

            if (call.Uses1904 == false && day < 61)
            {
                // Before March 1900 Excel's calendar counts the made-up 1900-02-29 and its weekdays follow it.
                weekday = (DayOfWeek)((day + 6) % 7);

                if (day == 60)
                {
                    month = 2;
                    dayOfMonth = 29;
                }
            }
        }

        bool twelveHour = tokens.Any(token => token.Kind == TokenKind.AmPm);
        StringBuilder builder = new();

        for (int index = 0; index < resolved.Count; index++)
        {
            Token token = resolved[index];

            switch (token.Kind)
            {
                case TokenKind.DateCode:
                    builder.Append(DatePart(token.Text, year, month, dayOfMonth, weekday, hour, minute, second, twelveHour));

                    break;

                case TokenKind.Elapsed:
                    long elapsed = token.Text[0] switch
                    {
                        'h' => totalUnits / (3600 * unitsPerSecond),
                        'm' => totalUnits / (60 * unitsPerSecond),
                        _ => totalUnits / unitsPerSecond,
                    };

                    builder.Append(elapsed.ToString(CultureInfo.InvariantCulture).PadLeft(token.Text.Length, '0'));

                    break;

                case TokenKind.AmPm:
                    builder.Append(AmPm(token.Text, hour));

                    break;

                case TokenKind.DecimalPoint when fractionDigits > 0 && index > 0 && IsSeconds(resolved[index - 1]) == true:
                    builder.Append('.');
                    builder.Append(subSecond.ToString(CultureInfo.InvariantCulture).PadLeft(fractionDigits, '0'));

                    while (index + 1 < resolved.Count && resolved[index + 1].Kind == TokenKind.Placeholder)
                    {
                        index++;
                    }

                    break;

                case TokenKind.TextValue:
                    builder.Append(Value.NumberText(serial));

                    break;

                default:
                    builder.Append(token.Text);

                    break;
            }
        }

        text = builder.ToString();

        return true;
    }

    private static bool IsSeconds(Token token)
    {
        return (token.Kind == TokenKind.DateCode && token.Text[0] == 's') || (token.Kind == TokenKind.Elapsed && token.Text[0] == 's');
    }

    /// <summary>m and mm mean minutes right after an hour code or right before a seconds code; otherwise the month. Minutes become "n".</summary>
    private static List<Token> ResolveMinutes(List<Token> tokens)
    {
        List<Token> resolved = [.. tokens];

        for (int index = 0; index < resolved.Count; index++)
        {
            Token token = resolved[index];

            if (token.Kind != TokenKind.DateCode || token.Text[0] != 'm' || token.Text.Length > 2)
            {
                continue;
            }

            bool afterHour = false;

            for (int back = index - 1; back >= 0; back--)
            {
                if (resolved[back].Kind is TokenKind.DateCode or TokenKind.Elapsed)
                {
                    afterHour = resolved[back].Text[0] == 'h';

                    break;
                }
            }

            bool beforeSecond = false;

            for (int forward = index + 1; forward < resolved.Count; forward++)
            {
                if (resolved[forward].Kind is TokenKind.DateCode or TokenKind.Elapsed)
                {
                    beforeSecond = resolved[forward].Text[0] == 's';

                    break;
                }
            }

            if (afterHour == true || beforeSecond == true)
            {
                resolved[index] = new Token(TokenKind.DateCode, new string('n', token.Text.Length));
            }
        }

        return resolved;
    }

    private static string DatePart(string code, int year, int month, int day, DayOfWeek weekday, long hour, long minute, long second, bool twelveHour)
    {
        DateTimeFormatInfo names = CultureInfo.InvariantCulture.DateTimeFormat;
        int length = code.Length;

        switch (code[0])
        {
            case 'y':
                if (length <= 2)
                {
                    return (year % 100).ToString("00", CultureInfo.InvariantCulture);
                }

                return year.ToString("0000", CultureInfo.InvariantCulture);

            case 'm':
                return length switch
                {
                    1 => month.ToString(CultureInfo.InvariantCulture),
                    2 => month.ToString("00", CultureInfo.InvariantCulture),
                    3 => names.GetAbbreviatedMonthName(month),
                    5 => names.GetMonthName(month)[..1],
                    _ => names.GetMonthName(month),
                };

            case 'd':
                return length switch
                {
                    1 => day.ToString(CultureInfo.InvariantCulture),
                    2 => day.ToString("00", CultureInfo.InvariantCulture),
                    3 => names.GetAbbreviatedDayName(weekday),
                    _ => names.GetDayName(weekday),
                };

            case 'h':
                long shown = hour;

                if (twelveHour == true)
                {
                    shown = hour % 12;

                    if (shown == 0)
                    {
                        shown = 12;
                    }
                }

                return Padded(shown, length);

            case 'n':
                return Padded(minute, length);

            default:
                return Padded(second, length);
        }
    }

    private static string Padded(long number, int length)
    {
        if (length >= 2)
        {
            return number.ToString("00", CultureInfo.InvariantCulture);
        }

        return number.ToString(CultureInfo.InvariantCulture);
    }

    private static string AmPm(string code, long hour)
    {
        bool morning = hour < 12;

        if (code.Length == 5)
        {
            if (morning == true)
            {
                return "AM";
            }

            return "PM";
        }

        // A/P keeps the case it is written in.
        if (morning == true)
        {
            return code[..1];
        }

        return code[2..3];
    }
}
