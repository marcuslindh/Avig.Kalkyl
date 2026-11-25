using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Text functions. Positions are 1-based; a start before the text or past its end gives #VALUE!.
/// A result longer than Excel's 32,767 characters is #VALUE!.
/// </summary>
internal static class TextFunctions
{
    private const int MaxLength = 32_767;

    /// <summary>Windows-1252 characters 128–159, which CHAR and CODE use; the rest of 1–255 is Latin-1.</summary>
    private static readonly char[] Windows1252 =
    [
        '€', '\u0081', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u008D', 'Ž', '\u008F',
        '\u0090', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', '\u009D', 'ž', 'Ÿ',
    ];

    public static void Register(FunctionTable table)
    {
        table.Add("CONCATENATE", 1, FunctionTable.Many, Concatenate);
        table.Add("CONCAT", 1, FunctionTable.Many, Concat);
        table.Add("TEXTJOIN", 3, FunctionTable.Many, TextJoin);
        table.Add("LEFT", 1, 2, call => Side(call, true));
        table.Add("RIGHT", 1, 2, call => Side(call, false));
        table.Add("MID", 3, 3, Mid);
        table.Add("LEN", 1, 1, Len);
        table.Add("LOWER", 1, 1, call => Map(call, text => text.ToLowerInvariant()));
        table.Add("UPPER", 1, 1, call => Map(call, text => text.ToUpperInvariant()));
        table.Add("PROPER", 1, 1, call => Map(call, Proper));
        table.Add("TRIM", 1, 1, call => Map(call, Trim));
        table.Add("CLEAN", 1, 1, call => Map(call, Clean));
        table.Add("SUBSTITUTE", 3, 4, Substitute);
        table.Add("REPLACE", 4, 4, Replace);
        table.Add("FIND", 2, 3, call => Find(call, false));
        table.Add("SEARCH", 2, 3, call => Find(call, true));
        table.Add("EXACT", 2, 2, Exact);
        table.Add("REPT", 2, 2, Rept);
        table.Add("TEXT", 2, 2, Text);
        table.Add("VALUE", 1, 1, ValueOf);
        table.Add("NUMBERVALUE", 1, 3, NumberValue);
        table.Add("CHAR", 1, 1, Char);
        table.Add("CODE", 1, 1, Code);
        table.Add("UNICHAR", 1, 1, UniChar);
        table.Add("UNICODE", 1, 1, UniCode);
        table.Add("FIXED", 1, 3, Fixed);
        table.Add("DOLLAR", 1, 2, Dollar);
        table.Add("T", 1, 1, TextOnly);
        table.Add("TEXTBEFORE", 2, 6, call => TextAround(call, true));
        table.Add("TEXTAFTER", 2, 6, call => TextAround(call, false));
        table.Add("VALUETOTEXT", 1, 2, ValueToText);
    }

    private static Value Result(string text)
    {
        if (text.Length > MaxLength)
        {
            return Value.ValueError;
        }

        return Value.FromText(text);
    }

    private static Value Map(FunctionCall call, Func<string, string> transform)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        return Value.FromText(transform(text));
    }

    /// <summary>CONCATENATE: each argument as one value (a range gives its intersecting cell).</summary>
    private static Value Concatenate(FunctionCall call)
    {
        StringBuilder builder = new();

        for (int index = 0; index < call.Count; index++)
        {
            if (call.TryText(index, out string text, out Value error) == false)
            {
                return error;
            }

            builder.Append(text);
        }

        return Result(builder.ToString());
    }

    /// <summary>CONCAT: like CONCATENATE, but ranges and arrays give every cell.</summary>
    private static Value Concat(FunctionCall call)
    {
        StringBuilder builder = new();

        for (int index = 0; index < call.Count; index++)
        {
            foreach (Value item in ItemValues(call, index))
            {
                Value text = Coerce.ToText(item);

                if (text.IsError == true)
                {
                    return text;
                }

                builder.Append(text.Text);
            }
        }

        return Result(builder.ToString());
    }

    /// <summary>TEXTJOIN: the delimiter can be an array whose items are used in turn.</summary>
    private static Value TextJoin(FunctionCall call)
    {
        List<string> delimiters = [];

        foreach (Value item in ItemValues(call, 0))
        {
            Value text = Coerce.ToText(item);

            if (text.IsError == true)
            {
                return text;
            }

            delimiters.Add(text.Text!);
        }

        if (delimiters.Count == 0)
        {
            delimiters.Add(string.Empty);
        }

        if (call.TryBool(1, out bool ignoreEmpty, out Value error) == false)
        {
            return error;
        }

        StringBuilder builder = new();
        int joined = 0;

        for (int index = 2; index < call.Count; index++)
        {
            foreach (Value item in ItemValues(call, index))
            {
                Value text = Coerce.ToText(item);

                if (text.IsError == true)
                {
                    return text;
                }

                if (ignoreEmpty == true && text.Text!.Length == 0)
                {
                    continue;
                }

                if (joined > 0)
                {
                    builder.Append(delimiters[(joined - 1) % delimiters.Count]);
                }

                builder.Append(text.Text);
                joined++;
            }
        }

        return Result(builder.ToString());
    }

    private static IEnumerable<Value> ItemValues(FunctionCall call, int index)
    {
        if (call.IsRangeOrArray(index) == true)
        {
            return call.Values(index);
        }

        return [call.Scalar(index)];
    }

    private static Value Side(FunctionCall call, bool left)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryInteger(1, 1, out long count, out error) == false)
        {
            return error;
        }

        if (count < 0)
        {
            return Value.ValueError;
        }

        int length = (int)Math.Min(count, text.Length);

        if (left == true)
        {
            return Value.FromText(text[..length]);
        }

        return Value.FromText(text[(text.Length - length)..]);
    }

    private static Value Mid(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryInteger(1, out long start, out error) == false || call.TryInteger(2, out long count, out error) == false)
        {
            return error;
        }

        if (start < 1 || count < 0)
        {
            return Value.ValueError;
        }

        if (start > text.Length)
        {
            return Value.FromText(string.Empty);
        }

        int first = (int)start - 1;
        int length = (int)Math.Min(count, text.Length - first);

        return Value.FromText(text.Substring(first, length));
    }

    private static Value Len(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(text.Length);
    }

    /// <summary>PROPER: a letter after anything that is not a letter becomes upper case, every other letter lower case.</summary>
    private static string Proper(string text)
    {
        StringBuilder builder = new(text.Length);
        bool previousLetter = false;

        foreach (char character in text)
        {
            if (char.IsLetter(character) == true)
            {
                if (previousLetter == true)
                {
                    builder.Append(char.ToLowerInvariant(character));
                }
                else
                {
                    builder.Append(char.ToUpperInvariant(character));
                }

                previousLetter = true;
            }
            else
            {
                builder.Append(character);
                previousLetter = false;
            }
        }

        return builder.ToString();
    }

    /// <summary>TRIM: only the space character; leading and trailing removed, inner runs made one.</summary>
    private static string Trim(string text)
    {
        StringBuilder builder = new(text.Length);
        bool pendingSpace = false;

        foreach (char character in text)
        {
            if (character == ' ')
            {
                pendingSpace = builder.Length > 0;

                continue;
            }

            if (pendingSpace == true)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(character);
        }

        return builder.ToString();
    }

    /// <summary>CLEAN: removes the control characters 0–31.</summary>
    private static string Clean(string text)
    {
        StringBuilder builder = new(text.Length);

        foreach (char character in text)
        {
            if (character >= ' ')
            {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }

    private static Value Substitute(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryText(1, out string old, out error) == false || call.TryText(2, out string replacement, out error) == false)
        {
            return error;
        }

        if (call.IsMissing(3) == true)
        {
            if (old.Length == 0)
            {
                return Value.FromText(text);
            }

            return Result(text.Replace(old, replacement, StringComparison.Ordinal));
        }

        if (call.TryInteger(3, out long instance, out error) == false)
        {
            return error;
        }

        if (instance < 1)
        {
            return Value.ValueError;
        }

        if (old.Length == 0)
        {
            return Value.FromText(text);
        }

        int position = -1;
        int from = 0;

        for (long found = 0; found < instance; found++)
        {
            position = text.IndexOf(old, from, StringComparison.Ordinal);

            if (position < 0)
            {
                return Value.FromText(text);
            }

            from = position + old.Length;
        }

        return Result(string.Concat(text.AsSpan(0, position), replacement, text.AsSpan(position + old.Length)));
    }

    private static Value Replace(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryInteger(1, out long start, out error) == false || call.TryInteger(2, out long count, out error) == false || call.TryText(3, out string replacement, out error) == false)
        {
            return error;
        }

        if (start < 1 || count < 0)
        {
            return Value.ValueError;
        }

        int first = (int)Math.Min(start - 1, text.Length);
        int length = (int)Math.Min(count, text.Length - first);

        return Result(string.Concat(text.AsSpan(0, first), replacement, text.AsSpan(first + length)));
    }

    /// <summary>FIND is case-sensitive without wildcards; SEARCH ignores case and understands * ? and ~.</summary>
    private static Value Find(FunctionCall call, bool search)
    {
        if (call.TryText(0, out string find, out Value error) == false || call.TryText(1, out string within, out error) == false || call.TryInteger(2, 1, out long start, out error) == false)
        {
            return error;
        }

        if (start < 1 || start > within.Length + 1)
        {
            return Value.ValueError;
        }

        if (find.Length == 0)
        {
            return Value.FromNumber(start);
        }

        int from = (int)start - 1;

        if (search == false)
        {
            int position = within.IndexOf(find, from, StringComparison.Ordinal);

            if (position < 0)
            {
                return Value.ValueError;
            }

            return Value.FromNumber(position + 1);
        }

        Regex pattern = new(Criteria.WildcardToRegex(find), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, TimeSpan.FromSeconds(1));
        Match match = pattern.Match(within, from);

        if (match.Success == false)
        {
            return Value.ValueError;
        }

        return Value.FromNumber(match.Index + 1);
    }

    private static Value Exact(FunctionCall call)
    {
        if (call.TryText(0, out string first, out Value error) == false || call.TryText(1, out string second, out error) == false)
        {
            return error;
        }

        return Value.FromBool(string.Equals(first, second, StringComparison.Ordinal));
    }

    private static Value Rept(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryInteger(1, out long count, out error) == false)
        {
            return error;
        }

        // Divided rather than multiplied: text.Length * count overflows for a count near long.MaxValue.
        // Excel refuses a count past int.MaxValue even for empty text.
        if (count < 0 || count > int.MaxValue || (text.Length > 0 && count > MaxLength / text.Length))
        {
            return Value.ValueError;
        }

        // Empty text repeated any number of times is empty.
        if (text.Length == 0)
        {
            return Value.FromText(string.Empty);
        }

        StringBuilder builder = new();
        builder.Insert(0, text, (int)count);

        return Value.FromText(builder.ToString());
    }

    /// <summary>TEXT: numbers (and text that reads as one) through the format code; booleans and other text stay as they are.</summary>
    private static Value Text(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (value.IsError == true)
        {
            return value;
        }

        if (call.TryText(1, out string format, out Value error) == false)
        {
            return error;
        }

        double number;

        switch (value.Kind)
        {
            case ValueKind.Number:
                number = value.Number;

                break;

            case ValueKind.Blank:
                number = 0;

                break;

            case ValueKind.Bool:
                return Value.FromText(value.ToString());

            default:
                if (Coerce.TryParseNumber(value.Text!, call.Conversion, out number) == false)
                {
                    return Result(TextFormat.FormatText(value.Text!, format));
                }

                break;
        }

        if (TextFormat.TryFormatNumber(call, number, format, out string text, out error) == false)
        {
            return error;
        }

        return Result(text);
    }

    /// <summary>VALUE: text that reads as a number, date or time; booleans and "" are #VALUE!.</summary>
    private static Value ValueOf(FunctionCall call)
    {
        Value value = call.Scalar(0);

        switch (value.Kind)
        {
            case ValueKind.Number:
            case ValueKind.Error:
                return value;

            case ValueKind.Blank:
                return Value.FromNumber(0);

            case ValueKind.Text when Coerce.TryParseNumber(value.Text!, call.Conversion, out double number) == true:
                return Value.FromNumber(number);

            default:
                return Value.ValueError;
        }
    }

    /// <summary>NUMBERVALUE: spaces are ignored, group separators removed, trailing % each divide by 100; empty text is 0.</summary>
    private static Value NumberValue(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryText(1, ".", out string decimalText, out error) == false || call.TryText(2, ",", out string groupText, out error) == false)
        {
            return error;
        }

        if (decimalText.Length == 0 || groupText.Length == 0)
        {
            return Value.ValueError;
        }

        char decimalSeparator = decimalText[0];
        char groupSeparator = groupText[0];
        string compact = string.Concat(text.Where(character => char.IsWhiteSpace(character) == false));
        int percents = 0;

        while (compact.EndsWith('%') == true)
        {
            percents++;
            compact = compact[..^1];
        }

        if (compact.Length == 0)
        {
            if (percents > 0)
            {
                return Value.ValueError;
            }

            return Value.FromNumber(0);
        }

        int decimalPosition = compact.IndexOf(decimalSeparator, StringComparison.Ordinal);

        if (decimalPosition >= 0 && (compact.IndexOf(decimalSeparator, decimalPosition + 1) >= 0 || compact.IndexOf(groupSeparator, decimalPosition + 1) >= 0))
        {
            return Value.ValueError;
        }

        StringBuilder builder = new();

        foreach (char character in compact)
        {
            if (character == decimalSeparator)
            {
                builder.Append('.');
            }
            else if (character != groupSeparator)
            {
                builder.Append(character);
            }
        }

        if (double.TryParse(builder.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == false)
        {
            return Value.ValueError;
        }

        return Value.FromNumber(number / Math.Pow(100, percents));
    }

    private static Value Char(FunctionCall call)
    {
        if (call.TryInteger(0, out long code, out Value error) == false)
        {
            return error;
        }

        if (code < 1 || code > 255)
        {
            return Value.ValueError;
        }

        if (code >= 128 && code < 160)
        {
            return Value.FromText(Windows1252[code - 128].ToString());
        }

        return Value.FromText(((char)code).ToString());
    }

    private static Value Code(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        if (text.Length == 0)
        {
            return Value.ValueError;
        }

        char first = text[0];

        if (first < 128 || (first >= 160 && first <= 255))
        {
            return Value.FromNumber(first);
        }

        int index = Array.IndexOf(Windows1252, first);

        if (index >= 0)
        {
            return Value.FromNumber(128 + index);
        }

        // Characters outside Windows-1252 read as "?".
        return Value.FromNumber(63);
    }

    private static Value UniChar(FunctionCall call)
    {
        if (call.TryInteger(0, out long code, out Value error) == false)
        {
            return error;
        }

        if (code < 1 || code > 0x10FFFF)
        {
            return Value.ValueError;
        }

        if (code >= 0xD800 && code <= 0xDFFF)
        {
            return Value.NA;
        }

        return Value.FromText(char.ConvertFromUtf32((int)code));
    }

    private static Value UniCode(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        if (text.Length == 0)
        {
            return Value.ValueError;
        }

        if (char.IsHighSurrogate(text[0]) == true && text.Length > 1 && char.IsLowSurrogate(text[1]) == true)
        {
            return Value.FromNumber(char.ConvertToUtf32(text[0], text[1]));
        }

        return Value.FromNumber(text[0]);
    }

    /// <summary>FIXED: rounded to the decimals (negative rounds left of the point), with or without thousands separators.</summary>
    private static Value Fixed(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryInteger(1, 2, out long decimals, out error) == false || call.TryBool(2, false, out bool noCommas, out error) == false)
        {
            return error;
        }

        if (decimals > 127)
        {
            return Value.ValueError;
        }

        return Value.FromText(FormatFixed(number, decimals, noCommas == false));
    }

    /// <summary>DOLLAR: like FIXED with a $ sign; negative amounts in parentheses, as en-US shows them.</summary>
    private static Value Dollar(FunctionCall call)
    {
        if (call.TryNumber(0, out double number, out Value error) == false || call.TryInteger(1, 2, out long decimals, out error) == false)
        {
            return error;
        }

        if (decimals > 127)
        {
            return Value.ValueError;
        }

        string text = FormatFixed(Math.Abs(number), decimals, true);

        if (number < 0 && Numeric.Round(number, (int)Math.Max(decimals, -15)) != 0)
        {
            return Value.FromText("($" + text + ")");
        }

        return Value.FromText("$" + text);
    }

    private static string FormatFixed(double number, long decimals, bool commas)
    {
        int digits = (int)Math.Max(decimals, -15);
        double rounded = Numeric.Round(number, digits);
        string precision = Math.Max(0, digits).ToString(CultureInfo.InvariantCulture);

        if (rounded == 0)
        {
            // No negative zero in the text.
            rounded = 0;
        }

        if (commas == true)
        {
            return rounded.ToString("N" + precision, CultureInfo.InvariantCulture);
        }

        return rounded.ToString("F" + precision, CultureInfo.InvariantCulture);
    }

    private static Value TextOnly(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (value.Kind is ValueKind.Text or ValueKind.Error)
        {
            return value;
        }

        return Value.FromText(string.Empty);
    }

    /// <summary>
    /// TEXTBEFORE and TEXTAFTER(text, delimiter, [instance], [match_mode], [match_end], [if_not_found]).
    /// The delimiter can be an array of delimiters; a negative instance counts from the end; with
    /// match_end the end (or start) of the text counts as a delimiter. Not found gives #N/A or if_not_found.
    /// </summary>
    private static Value TextAround(FunctionCall call, bool before)
    {
        if (call.TryText(0, out string text, out Value error) == false)
        {
            return error;
        }

        List<string> delimiters = [];

        foreach (Value item in ItemValues(call, 1))
        {
            Value delimiter = Coerce.ToText(item);

            if (delimiter.IsError == true)
            {
                return delimiter;
            }

            delimiters.Add(delimiter.Text!);
        }

        // A blank cell as the delimiter gives no items; it is the empty text, as Excel reads it.
        if (delimiters.Count == 0)
        {
            delimiters.Add(string.Empty);
        }

        if (call.TryInteger(2, 1, out long instance, out error) == false || call.TryInteger(3, 0, out long matchMode, out error) == false || call.TryInteger(4, 0, out long matchEnd, out error) == false)
        {
            return error;
        }

        if (instance == 0 || Math.Abs(instance) > Math.Max(1, text.Length) || matchMode is not (0 or 1) || matchEnd is not (0 or 1))
        {
            return Value.ValueError;
        }

        // An instance written out may not exceed the places where the shortest delimiter could start; Excel checks
        // this before it searches, and only when the instance is given.
        int shortest = delimiters.Min(delimiter => delimiter.Length);

        if (call.IsMissing(2) == false && shortest > 0 && Math.Abs(instance) > text.Length - shortest + 1)
        {
            return Value.ValueError;
        }

        StringComparison comparison = StringComparison.Ordinal;

        if (matchMode == 1)
        {
            comparison = StringComparison.OrdinalIgnoreCase;
        }

        if (delimiters.Any(delimiter => delimiter.Length == 0) == true)
        {
            // An empty delimiter matches at once: at the start when counting forward, at the end when counting back.
            if (instance > 0)
            {
                return Around(text, 0, 0, before);
            }

            return Around(text, text.Length, 0, before);
        }

        List<(int Position, int Length)> matches = [];
        int from = 0;

        while (from <= text.Length)
        {
            int best = -1;
            int bestLength = 0;

            foreach (string delimiter in delimiters)
            {
                int position = text.IndexOf(delimiter, from, comparison);

                if (position >= 0 && (best < 0 || position < best || (position == best && delimiter.Length > bestLength)))
                {
                    best = position;
                    bestLength = delimiter.Length;
                }
            }

            if (best < 0)
            {
                break;
            }

            matches.Add((best, bestLength));
            from = best + bestLength;
        }

        int wanted = (int)Math.Abs(instance);

        if (wanted <= matches.Count)
        {
            (int Position, int Length) chosen = matches[wanted - 1];

            if (instance < 0)
            {
                chosen = matches[matches.Count - wanted];
            }

            return Around(text, chosen.Position, chosen.Length, before);
        }

        // The end of the text counts as a delimiter only when the shortest delimiter would fit in the text.
        if (matchEnd == 1 && wanted == matches.Count + 1 && shortest <= text.Length)
        {
            if (instance > 0)
            {
                return Around(text, text.Length, 0, before);
            }

            return Around(text, 0, 0, before);
        }

        if (call.IsMissing(5) == false)
        {
            return call.Raw(5);
        }

        return Value.NA;
    }

    private static Value Around(string text, int position, int length, bool before)
    {
        if (before == true)
        {
            return Value.FromText(text[..position]);
        }

        return Value.FromText(text[(position + length)..]);
    }

    /// <summary>VALUETOTEXT: format 0 is concise (text as it is), 1 is strict (text in quotes, as a formula would write it).</summary>
    private static Value ValueToText(FunctionCall call)
    {
        Value value = call.Scalar(0);

        if (call.TryInteger(1, 0, out long format, out Value error) == false)
        {
            return error;
        }

        if (format is not (0 or 1))
        {
            return Value.ValueError;
        }

        if (value.Kind == ValueKind.Text && format == 1)
        {
            return Result("\"" + value.Text!.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"");
        }

        return Value.FromText(value.ToString());
    }
}
