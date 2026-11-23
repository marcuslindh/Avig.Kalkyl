using System.Globalization;
using System.Text;
using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Cell text can hold characters XML does not allow (for example \u0001). Excel writes them
/// as _xHHHH_, and an _x already in the text that looks like such a code is written as
/// _x005F_ so that it is not decoded by mistake.
/// </summary>
internal static class CellText
{
    public static string Encode(string text)
    {
        if (NeedsEncoding(text) == false)
        {
            return text;
        }

        StringBuilder builder = new(text.Length + 16);

        for (int index = 0; index < text.Length; index++)
        {
            char current = text[index];

            if (current == '_' && LooksLikeEscape(text, index) == true)
            {
                builder.Append("_x005F_");

                continue;
            }

            if (IsAllowed(text, index) == false)
            {
                builder.Append("_x").Append(((int)current).ToString("X4", CultureInfo.InvariantCulture)).Append('_');

                continue;
            }

            builder.Append(current);
        }

        return builder.ToString();
    }

    public static string Decode(string text)
    {
        if (text.Contains("_x", StringComparison.Ordinal) == false)
        {
            return text;
        }

        StringBuilder builder = new(text.Length);
        int index = 0;

        while (index < text.Length)
        {
            if (text[index] == '_' && LooksLikeEscape(text, index) == true)
            {
                int code = int.Parse(text.AsSpan(index + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                builder.Append((char)code);
                index += 7;

                continue;
            }

            builder.Append(text[index]);
            index++;
        }

        return builder.ToString();
    }

    private static bool NeedsEncoding(string text)
    {
        for (int index = 0; index < text.Length; index++)
        {
            if (IsAllowed(text, index) == false || (text[index] == '_' && LooksLikeEscape(text, index) == true))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsAllowed(string text, int index)
    {
        char current = text[index];

        if (char.IsHighSurrogate(current) == true)
        {
            return index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) == true;
        }

        if (char.IsLowSurrogate(current) == true)
        {
            return index > 0 && char.IsHighSurrogate(text[index - 1]) == true;
        }

        return XmlConvert.IsXmlChar(current);
    }

    private static bool LooksLikeEscape(string text, int index)
    {
        if (index + 6 >= text.Length || text[index + 1] != 'x' || text[index + 6] != '_')
        {
            return false;
        }

        for (int offset = 2; offset < 6; offset++)
        {
            if (char.IsAsciiHexDigit(text[index + offset]) == false)
            {
                return false;
            }
        }

        return true;
    }
}
