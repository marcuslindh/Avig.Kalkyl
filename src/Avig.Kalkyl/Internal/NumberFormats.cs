namespace Avig.Kalkyl.Internal;

/// <summary>
/// Excels inbyggda talformat (id 0–49) och regeln för när ett format betyder datum.
/// I filen är ett datum bara ett tal; det är formatet som säger att det ska visas som datum.
/// </summary>
internal static class NumberFormats
{
    /// <summary>Första id:t för egna format. Lägre id:n är reserverade för de inbyggda.</summary>
    public const int FirstCustomId = 164;

    public const string DateFormat = "yyyy-mm-dd";
    public const string DateTimeFormat = "yyyy-mm-dd hh:mm:ss";

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

    /// <summary>Formatkoden för ett inbyggt id, om id:t har en känd kod.</summary>
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
    /// Id:t för en kod som är inbyggd. 14 och 22 tas inte med: Excel visar dem i datorns
    /// språkinställning, inte som koden säger, så en skriven "mm-dd-yy" ska få ett eget id.
    /// </summary>
    public static bool TryGetBuiltInId(string code, out int id)
    {
        return BuiltInIds.TryGetValue(code, out id);
    }

    /// <summary>
    /// Inbyggda datumformat. 27–36 och 50–58 är östasiatiska datumformat utan fast kod.
    /// 46 ([h]:mm:ss) är en varaktighet, inte en tidpunkt, och räknas inte.
    /// </summary>
    public static bool IsDateId(int id)
    {
        return id is (>= 14 and <= 22) or (>= 27 and <= 36) or 45 or 47 or (>= 50 and <= 58);
    }

    /// <summary>
    /// Om en formatkod visar ett datum eller en tid. Bara första avsnittet (före ;) räknas.
    /// Text inom citattecken, tecken efter \, _ eller *, och [..] (färg, språk, villkor) hoppas
    /// över. Hittas d, m, y, h eller s är det ett datum, utom när koden har [h], [m] eller [s]:
    /// då är det en varaktighet, som läses som tal.
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
