using System.Globalization;
using System.Xml;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Writing;

/// <summary>
/// Samlar stilarna medan raderna skrivs och slår ihop likadana, så att varje unik stil blir
/// en enda post i styles.xml. Typsnitt, fyllningar och talformat slås ihop var för sig, så
/// två stilar som bara skiljer sig i färg delar typsnitt.
/// </summary>
internal sealed class StyleRegistry
{
    private const string DefaultFontName = "Calibri";
    private const double DefaultFontSize = 11;

    private readonly Dictionary<(Style Style, string? Format), int> cellFormats = [];
    private readonly List<CellFormat> formats = [new CellFormat(0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false)];
    private readonly Dictionary<FontKey, int> fontIds = [];
    private readonly List<FontKey> fonts = [new FontKey(false, false, false, null, DefaultFontSize, DefaultFontName)];
    private readonly Dictionary<string, int> fillIds = new(StringComparer.Ordinal);
    private readonly List<string> fills = [];
    private readonly Dictionary<string, int> customFormatIds = new(StringComparer.Ordinal);
    private readonly List<(int Id, string Code)> customFormats = [];

    private readonly Dictionary<CellFormat, int> formatIndexes = [];

    public StyleRegistry()
    {
        formatIndexes[formats[0]] = 0;
        fontIds[fonts[0]] = 0;
        cellFormats[(Style.Default, null)] = 0;
    }

    /// <summary>Antal unika cellformat, inklusive standardformatet på plats 0.</summary>
    public int Count => formats.Count;

    /// <summary>
    /// Index i cellXfs för en stil. <paramref name="defaultFormat"/> används när stilen inte
    /// har något eget talformat; det är så datum utan format får yyyy-mm-dd.
    /// </summary>
    public int GetIndex(Style style, string? defaultFormat)
    {
        string? format = style.NumberFormat ?? defaultFormat;

        if (cellFormats.TryGetValue((style, format), out int existing) == true)
        {
            return existing;
        }

        FontKey font = new(style.Bold, style.Italic, style.Underline, style.FontColor, style.FontSize ?? DefaultFontSize, style.FontName ?? DefaultFontName);

        if (fontIds.TryGetValue(font, out int fontId) == false)
        {
            fontId = fonts.Count;
            fonts.Add(font);
            fontIds[font] = fontId;
        }

        int fillId = 0;

        if (style.Background is not null)
        {
            if (fillIds.TryGetValue(style.Background, out fillId) == false)
            {
                // 0 och 1 är reserverade: Excel kräver none och gray125 först.
                fillId = fills.Count + 2;
                fills.Add(style.Background);
                fillIds[style.Background] = fillId;
            }
        }

        int formatId = GetFormatId(format);
        CellFormat cellFormat = new(fontId, fillId, formatId, style.Horizontal, style.Vertical, style.WrapText);
        if (formatIndexes.TryGetValue(cellFormat, out int index) == false)
        {
            index = formats.Count;
            formats.Add(cellFormat);
            formatIndexes[cellFormat] = index;
        }

        cellFormats[(style, format)] = index;

        return index;
    }

    public async Task WriteAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "styleSheet", Ooxml.SpreadsheetNamespace);

        if (customFormats.Count > 0)
        {
            await xml.WriteStartElementAsync(null, "numFmts", null);
            await xml.WriteAttributeStringAsync(null, "count", null, Text(customFormats.Count));

            foreach ((int id, string code) in customFormats)
            {
                await xml.WriteStartElementAsync(null, "numFmt", null);
                await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(id));
                await xml.WriteAttributeStringAsync(null, "formatCode", null, code);
                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        await WriteFontsAsync(xml);
        await WriteFillsAsync(xml);

        await xml.WriteStartElementAsync(null, "borders", null);
        await xml.WriteAttributeStringAsync(null, "count", null, "1");
        await xml.WriteStartElementAsync(null, "border", null);
        await xml.WriteElementStringAsync(null, "left", null, string.Empty);
        await xml.WriteElementStringAsync(null, "right", null, string.Empty);
        await xml.WriteElementStringAsync(null, "top", null, string.Empty);
        await xml.WriteElementStringAsync(null, "bottom", null, string.Empty);
        await xml.WriteElementStringAsync(null, "diagonal", null, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await xml.WriteStartElementAsync(null, "cellStyleXfs", null);
        await xml.WriteAttributeStringAsync(null, "count", null, "1");
        await xml.WriteStartElementAsync(null, "xf", null);
        await xml.WriteAttributeStringAsync(null, "numFmtId", null, "0");
        await xml.WriteAttributeStringAsync(null, "fontId", null, "0");
        await xml.WriteAttributeStringAsync(null, "fillId", null, "0");
        await xml.WriteAttributeStringAsync(null, "borderId", null, "0");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await WriteCellFormatsAsync(xml);

        await xml.WriteStartElementAsync(null, "cellStyles", null);
        await xml.WriteAttributeStringAsync(null, "count", null, "1");
        await xml.WriteStartElementAsync(null, "cellStyle", null);
        await xml.WriteAttributeStringAsync(null, "name", null, "Normal");
        await xml.WriteAttributeStringAsync(null, "xfId", null, "0");
        await xml.WriteAttributeStringAsync(null, "builtinId", null, "0");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private int GetFormatId(string? format)
    {
        if (format is null)
        {
            return 0;
        }

        if (NumberFormats.TryGetBuiltInId(format, out int builtIn) == true)
        {
            return builtIn;
        }

        if (customFormatIds.TryGetValue(format, out int custom) == false)
        {
            custom = NumberFormats.FirstCustomId + customFormats.Count;
            customFormats.Add((custom, format));
            customFormatIds[format] = custom;
        }

        return custom;
    }

    private async Task WriteFontsAsync(XmlWriter xml)
    {
        await xml.WriteStartElementAsync(null, "fonts", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(fonts.Count));

        foreach (FontKey font in fonts)
        {
            await xml.WriteStartElementAsync(null, "font", null);

            if (font.Bold == true)
            {
                await xml.WriteElementStringAsync(null, "b", null, string.Empty);
            }

            if (font.Italic == true)
            {
                await xml.WriteElementStringAsync(null, "i", null, string.Empty);
            }

            if (font.Underline == true)
            {
                await xml.WriteElementStringAsync(null, "u", null, string.Empty);
            }

            await xml.WriteStartElementAsync(null, "sz", null);
            await xml.WriteAttributeStringAsync(null, "val", null, font.Size.ToString("R", CultureInfo.InvariantCulture));
            await xml.WriteEndElementAsync();

            if (font.Color is not null)
            {
                await xml.WriteStartElementAsync(null, "color", null);
                await xml.WriteAttributeStringAsync(null, "rgb", null, Argb(font.Color));
                await xml.WriteEndElementAsync();
            }

            await xml.WriteStartElementAsync(null, "name", null);
            await xml.WriteAttributeStringAsync(null, "val", null, font.Name);
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync(null, "family", null);
            await xml.WriteAttributeStringAsync(null, "val", null, "2");
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private async Task WriteFillsAsync(XmlWriter xml)
    {
        await xml.WriteStartElementAsync(null, "fills", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(fills.Count + 2));

        foreach (string pattern in new[] { "none", "gray125" })
        {
            await xml.WriteStartElementAsync(null, "fill", null);
            await xml.WriteStartElementAsync(null, "patternFill", null);
            await xml.WriteAttributeStringAsync(null, "patternType", null, pattern);
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        foreach (string color in fills)
        {
            await xml.WriteStartElementAsync(null, "fill", null);
            await xml.WriteStartElementAsync(null, "patternFill", null);
            await xml.WriteAttributeStringAsync(null, "patternType", null, "solid");
            await xml.WriteStartElementAsync(null, "fgColor", null);
            await xml.WriteAttributeStringAsync(null, "rgb", null, Argb(color));
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync(null, "bgColor", null);
            await xml.WriteAttributeStringAsync(null, "indexed", null, "64");
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private async Task WriteCellFormatsAsync(XmlWriter xml)
    {
        await xml.WriteStartElementAsync(null, "cellXfs", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(formats.Count));

        foreach (CellFormat format in formats)
        {
            await xml.WriteStartElementAsync(null, "xf", null);
            await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(format.FormatId));
            await xml.WriteAttributeStringAsync(null, "fontId", null, Text(format.FontId));
            await xml.WriteAttributeStringAsync(null, "fillId", null, Text(format.FillId));
            await xml.WriteAttributeStringAsync(null, "borderId", null, "0");
            await xml.WriteAttributeStringAsync(null, "xfId", null, "0");

            if (format.FormatId != 0)
            {
                await xml.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            }

            if (format.FontId != 0)
            {
                await xml.WriteAttributeStringAsync(null, "applyFont", null, "1");
            }

            if (format.FillId != 0)
            {
                await xml.WriteAttributeStringAsync(null, "applyFill", null, "1");
            }

            bool hasAlignment = format.Horizontal != HorizontalAlignment.General || format.Vertical != VerticalAlignment.Bottom || format.Wrap == true;

            if (hasAlignment == true)
            {
                await xml.WriteAttributeStringAsync(null, "applyAlignment", null, "1");
                await xml.WriteStartElementAsync(null, "alignment", null);

                if (format.Horizontal != HorizontalAlignment.General)
                {
                    await xml.WriteAttributeStringAsync(null, "horizontal", null, HorizontalName(format.Horizontal));
                }

                if (format.Vertical != VerticalAlignment.Bottom)
                {
                    await xml.WriteAttributeStringAsync(null, "vertical", null, VerticalName(format.Vertical));
                }

                if (format.Wrap == true)
                {
                    await xml.WriteAttributeStringAsync(null, "wrapText", null, "1");
                }

                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static string HorizontalName(HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.Left => "left",
            HorizontalAlignment.Center => "center",
            HorizontalAlignment.Right => "right",
            HorizontalAlignment.Justify => "justify",
            _ => "general",
        };
    }

    private static string VerticalName(VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Center => "center",
            VerticalAlignment.Top => "top",
            _ => "bottom",
        };
    }

    private static string Argb(string color)
    {
        return "FF" + color[1..];
    }

    private static string Text(int number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }

    private readonly record struct FontKey(bool Bold, bool Italic, bool Underline, string? Color, double Size, string Name);

    private readonly record struct CellFormat(int FontId, int FillId, int FormatId, HorizontalAlignment Horizontal, VerticalAlignment Vertical, bool Wrap);
}
