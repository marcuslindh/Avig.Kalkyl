using System.Globalization;
using System.IO.Compression;
using System.Xml;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Reading;

/// <summary>
/// Det läsaren behöver ur styles.xml: talformatet för varje cellformat (cellXfs), och om det
/// formatet är ett datum. Övrig formatering läses inte i strömläget.
/// </summary>
internal sealed class StyleTable
{
    private readonly Style[] styles;
    private readonly bool[] dates;

    private StyleTable(Style[] styles, bool[] dates)
    {
        this.styles = styles;
        this.dates = dates;
    }

    public static StyleTable Empty { get; } = new([], []);

    public bool IsDate(int index)
    {
        return index >= 0 && index < dates.Length && dates[index] == true;
    }

    public Style StyleAt(int index)
    {
        if (index >= 0 && index < styles.Length)
        {
            return styles[index];
        }

        return Style.Default;
    }

    public static async Task<StyleTable> ReadAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        Dictionary<int, string> customFormats = [];
        List<int> cellFormatIds = [];
        await using Stream stream = await entry.OpenAsync(cancellationToken);
        using XmlReader reader = XmlReader.Create(stream, XmlParts.Settings);
        bool insideCellXfs = false;
        int cellXfsDepth = -1;

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "numFmt") == true)
            {
                if (int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) == true)
                {
                    customFormats[id] = reader.GetAttribute("formatCode") ?? string.Empty;
                }

                continue;
            }

            if (XmlParts.IsElement(reader, "cellXfs") == true && reader.IsEmptyElement == false)
            {
                insideCellXfs = true;
                cellXfsDepth = reader.Depth;

                continue;
            }

            if (insideCellXfs == true && XmlParts.IsEnd(reader, cellXfsDepth) == true)
            {
                insideCellXfs = false;

                continue;
            }

            if (insideCellXfs == true && reader.Depth == cellXfsDepth + 1 && XmlParts.IsElement(reader, "xf") == true)
            {
                int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int formatId);
                cellFormatIds.Add(formatId);
            }
        }

        Style[] styles = new Style[cellFormatIds.Count];
        bool[] dates = new bool[cellFormatIds.Count];
        Dictionary<int, Style> byFormat = [];

        for (int index = 0; index < cellFormatIds.Count; index++)
        {
            int formatId = cellFormatIds[index];
            string? code = null;

            if (customFormats.TryGetValue(formatId, out string? custom) == true)
            {
                code = custom;
                dates[index] = NumberFormats.IsDateFormat(custom);
            }
            else
            {
                if (NumberFormats.TryGetBuiltIn(formatId, out string builtIn) == true)
                {
                    code = builtIn;
                }

                dates[index] = NumberFormats.IsDateId(formatId);
            }

            if (byFormat.TryGetValue(formatId, out Style? shared) == false)
            {
                shared = Style.Default;

                if (code is not null && string.Equals(code, "General", StringComparison.OrdinalIgnoreCase) == false)
                {
                    shared = new Style { NumberFormat = code };
                }

                byFormat[formatId] = shared;
            }

            styles[index] = shared;
        }

        return new StyleTable(styles, dates);
    }
}
