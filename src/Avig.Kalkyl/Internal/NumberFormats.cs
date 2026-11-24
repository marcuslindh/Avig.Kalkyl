namespace Avig.Kalkyl.Internal;

/// <summary>
/// Excel's built-in number formats (ids 0–49) and the rule for when a format means a date.
/// In the file a date is just a number; the format is what says to show it as a date.
/// </summary>
internal static class NumberFormats
{
    /// <summary>The first id for custom formats. Lower ids are reserved for the built-in ones.</summary>
    public const int FirstCustomId = 164;

    public const string DateFormat = "yyyy-mm-dd";
    public const string DateTimeFormat = "yyyy-mm-dd hh:mm:ss";
    public const string TimeFormat = "hh:mm:ss";

    private static readonly Dictionary<int, string> BuiltIn = new()
    {
        [0] = "General",
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "mm-dd-yy",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "m/d/yy h:mm",
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    };

    private static readonly Dictionary<string, int> BuiltInIds = BuiltIn
        .Where(entry => entry.Key != 14 && entry.Key != 22)
        .ToDictionary(entry => entry.Value, entry => entry.Key, StringComparer.Ordinal);

    /// <summary>The format code for a built-in id, when the id has a known code.</summary>
    public static bool TryGetBuiltIn(int id, out string code)
    {
        if (BuiltIn.TryGetValue(id, out string? found) == true)
        {
            code = found;

            return true;
        }

        code = string.Empty;

        return false;
    }

    /// <summary>
    /// The id of a built-in code. 14 and 22 are left out: Excel shows them in the computer's
    /// regional settings, not as the code says, so a written "mm-dd-yy" gets an id of its own.
    /// </summary>
    public static bool TryGetBuiltInId(string code, out int id)
    {
        return BuiltInIds.TryGetValue(code, out id);
    }

    /// <summary>
    /// Built-in date formats. 27–36 and 50–58 are East Asian date formats without a fixed code.
    /// 46 ([h]:mm:ss) is a duration, not a point in time, and does not count.
    /// </summary>
    public static bool IsDateId(int id)
    {
        return id is (>= 14 and <= 22) or (>= 27 and <= 36) or 45 or 47 or (>= 50 and <= 58);
    }

    /// <summary>
    /// Whether a format code shows a date or a time. Only the first section (before ;) counts.
    /// Quoted text, the character after \, _ or *, and [..] (colour, locale, condition) are
    /// skipped. A d, m, y, h or s makes it a date, except when the code has [h], [m] or [s]:
    /// then it is a duration, which is read as a number.
    /// </summary>
    public static bool IsDateFormat(string code)
    {
        bool hasDatePart = false;
        int position = 0;

        while (position < code.Length)
        {
            char current = code[position];

            switch (current)
            {
                case ';':
                    return hasDatePart;

                case '"':
                    int closing = code.IndexOf('"', position + 1);

                    if (closing < 0)
                    {
                        return hasDatePart;
                    }

                    position = closing + 1;

                    continue;

                case '\\':
                case '_':
                case '*':
                    position += 2;

                    continue;

                case '[':
                    int end = code.IndexOf(']', position + 1);

                    if (end < 0)
                    {
                        return hasDatePart;
                    }

                    string inside = code[(position + 1)..end].ToUpperInvariant();

                    if (inside.Length > 0 && inside.All(letter => letter is 'H' or 'M' or 'S') == true)
                    {
                        return false;
                    }

                    position = end + 1;

                    continue;

                default:
                    if (char.ToUpperInvariant(current) is 'D' or 'M' or 'Y' or 'H' or 'S')
                    {
                        hasDatePart = true;
                    }

                    position++;

                    continue;
            }
        }

        return hasDatePart;
    }
}
