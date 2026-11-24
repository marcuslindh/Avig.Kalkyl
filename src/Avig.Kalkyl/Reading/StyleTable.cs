using System.Globalization;
using System.IO.Compression;
using System.Xml;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Reading;

/// <summary>
/// The cell formats (cellXfs) of a workbook turned into <see cref="Style"/> values, and whether
/// each one shows a date. Font, fill, number format and alignment are read; borders, theme
/// colours and indexed colours are not, but the document model keeps a cell's original format
/// index, so they survive as long as the cell's style is not changed.
/// </summary>
internal sealed class StyleTable
{
    private readonly Style[] styles;
    private readonly bool[] dates;

    private StyleTable(Style[] styles, bool[] dates, int fontCount, int fillCount, int maxFormatId, string baseFontName, double baseFontSize)
    {
        this.styles = styles;
        this.dates = dates;
        BaseFontName = baseFontName;
        BaseFontSize = baseFontSize;
        FontCount = fontCount;
        FillCount = fillCount;
        MaxFormatId = maxFormatId;
    }

    public static StyleTable Empty { get; } = new([], [], 0, 0, 0, "Calibri", 11);

    /// <summary>The number of cell formats in the file.</summary>
    public int CellFormatCount => styles.Length;

    /// <summary>The number of fonts in the file.</summary>
    public int FontCount { get; }

    /// <summary>The number of fills in the file.</summary>
    public int FillCount { get; }

    /// <summary>The name of the workbook's default font (the first font).</summary>
    public string BaseFontName { get; }

    /// <summary>The size of the workbook's default font.</summary>
    public double BaseFontSize { get; }

    /// <summary>The highest number format id in use, built-in or custom.</summary>
    public int MaxFormatId { get; }

    /// <summary>The number of differential formats (dxf), which conditional formatting and tables refer to by index.</summary>
    public int DifferentialFormatCount { get; private init; }

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
        List<FontInfo> fonts = [];
        List<string?> fills = [];
        List<FormatInfo> cellFormats = [];
        int differentialFormats = 0;
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        string section = string.Empty;
        int sectionDepth = -1;

        while (await reader.ReadAsync() == true)
        {
            if (section.Length > 0 && XmlParts.IsEnd(reader, sectionDepth) == true)
            {
                section = string.Empty;

                continue;
            }

            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            switch (reader.LocalName)
            {
                case "numFmt":
                    if (int.TryParse(reader.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) == true)
                    {
                        customFormats[id] = reader.GetAttribute("formatCode") ?? string.Empty;
                    }

                    break;

                case "fonts" or "fills" or "cellXfs" or "dxfs" when reader.IsEmptyElement == false:
                    section = reader.LocalName;
                    sectionDepth = reader.Depth;

                    break;

                case "font" when section == "fonts" && reader.Depth == sectionDepth + 1:
                    fonts.Add(await ReadFontAsync(reader));

                    break;

                case "fill" when section == "fills" && reader.Depth == sectionDepth + 1:
                    fills.Add(await ReadFillAsync(reader));

                    break;

                case "xf" when section == "cellXfs" && reader.Depth == sectionDepth + 1:
                    cellFormats.Add(await ReadCellFormatAsync(reader));

                    break;

                case "dxf" when section == "dxfs" && reader.Depth == sectionDepth + 1:
                    differentialFormats++;

                    break;
            }
        }

        return Build(customFormats, fonts, fills, cellFormats, differentialFormats);
    }

    private static StyleTable Build(Dictionary<int, string> customFormats, List<FontInfo> fonts, List<string?> fills, List<FormatInfo> cellFormats, int differentialFormats)
    {
        Style[] styles = new Style[cellFormats.Count];
        bool[] dates = new bool[cellFormats.Count];
        FontInfo baseFont = new(false, false, false, null, null, null);

        if (fonts.Count > 0)
        {
            baseFont = fonts[0];
        }

        int maxFormatId = 0;

        foreach (int customId in customFormats.Keys)
        {
            maxFormatId = Math.Max(maxFormatId, customId);
        }

        for (int index = 0; index < cellFormats.Count; index++)
        {
            FormatInfo format = cellFormats[index];
            maxFormatId = Math.Max(maxFormatId, format.FormatId);
            string? code = null;

            if (customFormats.TryGetValue(format.FormatId, out string? custom) == true)
            {
                code = custom;
                dates[index] = NumberFormats.IsDateFormat(custom);
            }
            else
            {
                if (NumberFormats.TryGetBuiltIn(format.FormatId, out string builtIn) == true)
                {
                    code = builtIn;
                }

                dates[index] = NumberFormats.IsDateId(format.FormatId);
            }

            if (code is not null && string.Equals(code, "General", StringComparison.OrdinalIgnoreCase) == true)
            {
                code = null;
            }

            FontInfo font = baseFont;

            if (format.FontId >= 0 && format.FontId < fonts.Count)
            {
                font = fonts[format.FontId];
            }

            string? background = null;

            if (format.FillId >= 0 && format.FillId < fills.Count)
            {
                background = fills[format.FillId];
            }

            Style style = new()
            {
                Bold = font.Bold,
                Italic = font.Italic,
                Underline = font.Underline,
                FontColor = Different(font.Color, baseFont.Color),
                FontSize = DifferentSize(font.Size, baseFont.Size),
                FontName = Different(font.Name, baseFont.Name),
                Background = background,
                NumberFormat = code,
                Horizontal = format.Horizontal,
                Vertical = format.Vertical,
                WrapText = format.Wrap,
            };

            styles[index] = Intern(style);
        }

        return new StyleTable(styles, dates, fonts.Count, fills.Count, maxFormatId, baseFont.Name ?? "Calibri", baseFont.Size ?? 11) { DifferentialFormatCount = differentialFormats };
    }

    /// <summary>The unstyled style shares the one <see cref="Style.Default"/> instance.</summary>
    private static Style Intern(Style style)
    {
        if (style.Equals(Style.Default) == true)
        {
            return Style.Default;
        }

        return style;
    }

    private static string? Different(string? value, string? baseValue)
    {
        string? result = value;

        if (string.Equals(value, baseValue, StringComparison.OrdinalIgnoreCase) == true)
        {
            result = null;
        }

        return result;
    }

    private static double? DifferentSize(double? value, double? baseValue)
    {
        double? result = value;

        if (value == baseValue)
        {
            result = null;
        }

        return result;
    }

    private static async Task<FontInfo> ReadFontAsync(XmlReader reader)
    {
        bool bold = false;
        bool italic = false;
        bool underline = false;
        string? color = null;
        double? size = null;
        string? name = null;

        if (reader.IsEmptyElement == true)
        {
            return new FontInfo(bold, italic, underline, color, size, name);
        }

        int depth = reader.Depth;

        while (await reader.ReadAsync() == true && XmlParts.IsEnd(reader, depth) == false)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            string? value = reader.GetAttribute("val");

            switch (reader.LocalName)
            {
                case "b":
                    bold = IsOn(value);

                    break;

                case "i":
                    italic = IsOn(value);

                    break;

                case "u":
                    underline = value is null || string.Equals(value, "none", StringComparison.Ordinal) == false;

                    break;

                case "sz":
                    // A size Excel cannot show (it allows 1 to 409 points) is left out, as Style would refuse it.
                    if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double points) == true && points >= 1 && points <= 409)
                    {
                        size = points;
                    }

                    break;

                case "color":
                    color = ToHexColor(reader.GetAttribute("rgb"));

                    break;

                case "name":
                    name = value;

                    break;
            }
        }

        return new FontInfo(bold, italic, underline, color, size, name);
    }

    private static async Task<string?> ReadFillAsync(XmlReader reader)
    {
        string? color = null;

        if (reader.IsEmptyElement == true)
        {
            return color;
        }

        int depth = reader.Depth;
        bool solid = false;

        while (await reader.ReadAsync() == true && XmlParts.IsEnd(reader, depth) == false)
        {
            if (XmlParts.IsElement(reader, "patternFill") == true)
            {
                solid = string.Equals(reader.GetAttribute("patternType"), "solid", StringComparison.Ordinal);
            }
            else if (XmlParts.IsElement(reader, "fgColor") == true && solid == true)
            {
                color = ToHexColor(reader.GetAttribute("rgb"));
            }
        }

        return color;
    }

    private static async Task<FormatInfo> ReadCellFormatAsync(XmlReader reader)
    {
        int formatId = ReadInt(reader, "numFmtId");
        int fontId = ReadInt(reader, "fontId");
        int fillId = ReadInt(reader, "fillId");
        HorizontalAlignment horizontal = HorizontalAlignment.General;
        VerticalAlignment vertical = VerticalAlignment.Bottom;
        bool wrap = false;

        if (reader.IsEmptyElement == false)
        {
            int depth = reader.Depth;

            while (await reader.ReadAsync() == true && XmlParts.IsEnd(reader, depth) == false)
            {
                if (XmlParts.IsElement(reader, "alignment") == false)
                {
                    continue;
                }

                horizontal = reader.GetAttribute("horizontal") switch
                {
                    "left" => HorizontalAlignment.Left,
                    "center" => HorizontalAlignment.Center,
                    "right" => HorizontalAlignment.Right,
                    "justify" => HorizontalAlignment.Justify,
                    _ => HorizontalAlignment.General,
                };

                vertical = reader.GetAttribute("vertical") switch
                {
                    "top" => VerticalAlignment.Top,
                    "center" => VerticalAlignment.Center,
                    _ => VerticalAlignment.Bottom,
                };

                wrap = reader.GetAttribute("wrapText") is "1" or "true";
            }
        }

        return new FormatInfo(formatId, fontId, fillId, horizontal, vertical, wrap);
    }

    private static int ReadInt(XmlReader reader, string name)
    {
        int.TryParse(reader.GetAttribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value);

        return value;
    }

    /// <summary>A boolean font attribute such as &lt;b/&gt; or &lt;b val="0"/&gt;: no value means on.</summary>
    private static bool IsOn(string? value)
    {
        return value is null or "1" or "true";
    }

    /// <summary>An ARGB colour such as FF1F4E79 as #1F4E79. Theme and indexed colours come back as null.</summary>
    private static string? ToHexColor(string? argb)
    {
        string? color = null;

        if (argb is not null && (argb.Length == 8 || argb.Length == 6) && argb.All(char.IsAsciiHexDigit) == true)
        {
            color = "#" + argb[^6..].ToUpperInvariant();
        }

        return color;
    }

    private readonly record struct FontInfo(bool Bold, bool Italic, bool Underline, string? Color, double? Size, string? Name);

    private readonly record struct FormatInfo(int FormatId, int FontId, int FillId, HorizontalAlignment Horizontal, VerticalAlignment Vertical, bool Wrap);
}
