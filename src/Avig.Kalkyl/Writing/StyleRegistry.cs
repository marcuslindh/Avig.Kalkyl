using System.Globalization;
using System.Xml;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Writing;

/// <summary>
/// Collects the styles while cells are written and merges equal ones, so that every unique style
/// becomes a single entry in styles.xml. Fonts, fills and number formats are merged separately,
/// so two styles that differ only in fill share a font.
/// <para>
/// For a new file the registry owns all of styles.xml. For a file that was opened, it builds on
/// the file's own styles: new entries get indexes after the existing ones, and
/// <see cref="WriteMergedAsync"/> copies the original styles.xml with the new entries appended,
/// so every cell that kept its style keeps its index.
/// </para>
/// </summary>
internal sealed class StyleRegistry
{
    private readonly string baseFontName;
    private readonly double baseFontSize;
    private readonly int fontBase;
    private readonly int fillBase;
    private readonly int formatBase;
    private readonly int differentialBase;
    private readonly Dictionary<Style, int> differentialIds = [];
    private readonly List<Style> differentialFormats = [];
    private readonly bool fresh;
    private readonly Dictionary<(Style Style, string? Format), int> styleIndexes = [];
    private readonly Dictionary<CellFormat, int> formatIndexes = [];
    private readonly List<CellFormat> formats = [];
    private readonly Dictionary<FontKey, int> fontIds = [];
    private readonly List<FontKey> fonts = [];
    private readonly Dictionary<string, int> fillIds = new(StringComparer.Ordinal);
    private readonly List<string> fills = [];
    private readonly Dictionary<string, int> customFormatIds = new(StringComparer.Ordinal);
    private readonly List<(int Id, string Code)> customFormats = [];
    private int nextCustomFormatId;

    /// <summary>A registry for a new file: Calibri 11, and the two fills Excel requires first.</summary>
    public StyleRegistry()
    {
        fresh = true;
        baseFontName = "Calibri";
        baseFontSize = 11;
        fontBase = 0;
        fillBase = 2;
        formatBase = 0;
        nextCustomFormatId = NumberFormats.FirstCustomId;
        FontKey defaultFont = new(false, false, false, null, baseFontSize, baseFontName);
        fonts.Add(defaultFont);
        fontIds[defaultFont] = 0;
        CellFormat defaultFormat = new(0, 0, 0, HorizontalAlignment.General, VerticalAlignment.Bottom, false);
        formats.Add(defaultFormat);
        formatIndexes[defaultFormat] = 0;
        styleIndexes[(Style.Default, null)] = 0;
    }

    /// <summary>A registry that adds to the styles of an opened file.</summary>
    public StyleRegistry(StyleTable original)
    {
        fresh = false;
        baseFontName = original.BaseFontName;
        baseFontSize = original.BaseFontSize;
        fontBase = original.FontCount;
        fillBase = original.FillCount;
        formatBase = original.CellFormatCount;
        differentialBase = original.DifferentialFormatCount;
        nextCustomFormatId = Math.Max(NumberFormats.FirstCustomId, original.MaxFormatId + 1);
        styleIndexes[(Style.Default, null)] = 0;
    }

    /// <summary>True when styles were added that the file does not have yet.</summary>
    public bool HasNewEntries => fresh == true || formats.Count > 0 || customFormats.Count > 0 || differentialFormats.Count > 0;

    /// <summary>
    /// The cellXfs index for a style. <paramref name="defaultFormat"/> is used when the style has
    /// no number format of its own; that is how dates without a format get yyyy-mm-dd.
    /// </summary>
    public int GetIndex(Style style, string defaultFormat)
    {
        string? format = style.NumberFormat;

        if (format is null && defaultFormat.Length > 0)
        {
            format = defaultFormat;
        }

        if (styleIndexes.TryGetValue((style, format), out int existing) == true)
        {
            return existing;
        }

        FontKey font = new(style.Bold, style.Italic, style.Underline, style.FontColor, style.FontSize ?? baseFontSize, style.FontName ?? baseFontName);

        if (fontIds.TryGetValue(font, out int fontId) == false)
        {
            fontId = fontBase + fonts.Count;
            fonts.Add(font);
            fontIds[font] = fontId;
        }

        int fillId = 0;

        if (style.Background is not null && fillIds.TryGetValue(style.Background, out fillId) == false)
        {
            fillId = fillBase + fills.Count;
            fills.Add(style.Background);
            fillIds[style.Background] = fillId;
        }

        CellFormat cellFormat = new(fontId, fillId, GetFormatId(format), style.Horizontal, style.Vertical, style.WrapText);

        if (formatIndexes.TryGetValue(cellFormat, out int index) == false)
        {
            index = formatBase + formats.Count;
            formats.Add(cellFormat);
            formatIndexes[cellFormat] = index;
        }

        styleIndexes[(style, format)] = index;

        return index;
    }

    /// <summary>
    /// The dxfs index for the style of a conditional format. Equal styles share an entry, and
    /// a file that was opened keeps its own entries in front of the new ones.
    /// </summary>
    public int GetDifferentialIndex(Style style)
    {
        if (differentialIds.TryGetValue(style, out int existing) == true)
        {
            return existing;
        }

        // The number format gets its id now, so that numFmts, which is written first, has it.
        if (style.NumberFormat is not null)
        {
            GetFormatId(style.NumberFormat);
        }

        int index = differentialBase + differentialFormats.Count;
        differentialFormats.Add(style);
        differentialIds[style] = index;

        return index;
    }

    /// <summary>Writes a complete styles.xml for a new file.</summary>
    public async Task WriteAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "styleSheet", Ooxml.SpreadsheetNamespace);

        if (customFormats.Count > 0)
        {
            await xml.WriteStartElementAsync(null, "numFmts", null);
            await xml.WriteAttributeStringAsync(null, "count", null, Text(customFormats.Count));
            await WriteCustomFormatsAsync(xml);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync(null, "fonts", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(fonts.Count));
        await WriteFontsAsync(xml);
        await xml.WriteEndElementAsync();

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

        await WriteFillsAsync(xml);
        await xml.WriteEndElementAsync();

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

        await xml.WriteStartElementAsync(null, "cellXfs", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(formats.Count));
        await WriteCellFormatsAsync(xml);
        await xml.WriteEndElementAsync();

        await xml.WriteStartElementAsync(null, "cellStyles", null);
        await xml.WriteAttributeStringAsync(null, "count", null, "1");
        await xml.WriteStartElementAsync(null, "cellStyle", null);
        await xml.WriteAttributeStringAsync(null, "name", null, "Normal");
        await xml.WriteAttributeStringAsync(null, "xfId", null, "0");
        await xml.WriteAttributeStringAsync(null, "builtinId", null, "0");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await WriteDifferentialListAsync(xml);

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>
    /// Copies an opened file's styles.xml and appends the new number formats, fonts, fills and
    /// cell formats to their lists, updating each list's count.
    /// </summary>
    public async Task WriteMergedAsync(XmlReader original, XmlWriter xml)
    {
        await original.MoveToContentAsync();
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(original.Prefix, original.LocalName, original.NamespaceURI);
        await xml.WriteAttributesAsync(original, true);
        bool numFormatsWritten = customFormats.Count == 0;
        bool differentialWritten = differentialFormats.Count == 0;

        if (original.IsEmptyElement == false)
        {
            await original.ReadAsync();

            while (XmlParts.IsEnd(original, 0) == false)
            {
                if (original.NodeType != XmlNodeType.Element || original.Depth != 1)
                {
                    await xml.WriteNodeAsync(original, true);

                    continue;
                }

                // A file without dxfs gets one where the schema puts it: after cellStyles, before these.
                if (differentialWritten == false && original.LocalName is "tableStyles" or "colors" or "extLst")
                {
                    await WriteDifferentialListAsync(xml);
                    differentialWritten = true;
                }

                switch (original.LocalName)
                {
                    case "dxfs":
                        await CopyListAsync(original, xml, differentialFormats.Count, WriteDifferentialFormatsAsync);
                        differentialWritten = true;

                        break;

                    case "numFmts":
                        await CopyListAsync(original, xml, customFormats.Count, WriteCustomFormatsAsync);
                        numFormatsWritten = true;

                        break;

                    case "fonts":
                        await WriteMissingFormatsAsync(xml, numFormatsWritten);
                        numFormatsWritten = true;
                        await CopyListAsync(original, xml, fonts.Count, WriteFontsAsync);

                        break;

                    case "fills":
                        await CopyListAsync(original, xml, fills.Count, WriteFillsAsync);

                        break;

                    case "cellXfs":
                        await CopyListAsync(original, xml, formats.Count, WriteCellFormatsAsync);

                        break;

                    default:
                        await xml.WriteNodeAsync(original, true);

                        break;
                }
            }
        }

        if (differentialWritten == false)
        {
            await WriteDifferentialListAsync(xml);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private async Task WriteMissingFormatsAsync(XmlWriter xml, bool alreadyWritten)
    {
        if (alreadyWritten == true)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, "numFmts", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(customFormats.Count));
        await WriteCustomFormatsAsync(xml);
        await xml.WriteEndElementAsync();
    }

    /// <summary>Copies a list element with its children, appends new children and adds their number to its count.</summary>
    private static async Task CopyListAsync(XmlReader original, XmlWriter xml, int added, Func<XmlWriter, Task> writeAdded)
    {
        await xml.WriteStartElementAsync(original.Prefix, original.LocalName, original.NamespaceURI);

        if (original.MoveToFirstAttribute() == true)
        {
            do
            {
                string value = original.Value;

                // The count must match the children; the original count plus the ones added.
                if (string.Equals(original.LocalName, "count", StringComparison.Ordinal) == true
                    && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) == true)
                {
                    value = Text(count + added);
                }

                await xml.WriteAttributeStringAsync(original.Prefix, original.LocalName, original.NamespaceURI, value);
            }
            while (original.MoveToNextAttribute() == true);

            original.MoveToElement();
        }

        int depth = original.Depth;

        if (original.IsEmptyElement == false)
        {
            await original.ReadAsync();

            while (XmlParts.IsEnd(original, depth) == false)
            {
                await xml.WriteNodeAsync(original, true);
            }
        }

        await original.ReadAsync();
        await writeAdded(xml);
        await xml.WriteEndElementAsync();
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
            custom = nextCustomFormatId;
            nextCustomFormatId++;
            customFormats.Add((custom, format));
            customFormatIds[format] = custom;
        }

        return custom;
    }

    private async Task WriteCustomFormatsAsync(XmlWriter xml)
    {
        foreach ((int id, string code) in customFormats)
        {
            await xml.WriteStartElementAsync(null, "numFmt", null);
            await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(id));
            await xml.WriteAttributeStringAsync(null, "formatCode", null, code);
            await xml.WriteEndElementAsync();
        }
    }

    private async Task WriteFontsAsync(XmlWriter xml)
    {
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
    }

    private async Task WriteFillsAsync(XmlWriter xml)
    {
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
    }

    private async Task WriteCellFormatsAsync(XmlWriter xml)
    {
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
    }

    private async Task WriteDifferentialListAsync(XmlWriter xml)
    {
        if (differentialFormats.Count == 0)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, "dxfs", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(differentialFormats.Count));
        await WriteDifferentialFormatsAsync(xml);
        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// The dxf entries. Children in schema order: font, numFmt, fill. A solid fill in a dxf is
    /// given as bgColor, unlike in a cell format; that is how Excel itself writes it.
    /// </summary>
    private async Task WriteDifferentialFormatsAsync(XmlWriter xml)
    {
        foreach (Style style in differentialFormats)
        {
            await xml.WriteStartElementAsync(null, "dxf", null);

            if (style.Bold == true || style.Italic == true || style.Underline == true || style.FontColor is not null)
            {
                await xml.WriteStartElementAsync(null, "font", null);

                if (style.Bold == true)
                {
                    await xml.WriteElementStringAsync(null, "b", null, string.Empty);
                }

                if (style.Italic == true)
                {
                    await xml.WriteElementStringAsync(null, "i", null, string.Empty);
                }

                if (style.Underline == true)
                {
                    await xml.WriteElementStringAsync(null, "u", null, string.Empty);
                }

                if (style.FontColor is not null)
                {
                    await xml.WriteStartElementAsync(null, "color", null);
                    await xml.WriteAttributeStringAsync(null, "rgb", null, Argb(style.FontColor));
                    await xml.WriteEndElementAsync();
                }

                await xml.WriteEndElementAsync();
            }

            if (style.NumberFormat is not null)
            {
                await xml.WriteStartElementAsync(null, "numFmt", null);
                await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(GetFormatId(style.NumberFormat)));
                await xml.WriteAttributeStringAsync(null, "formatCode", null, style.NumberFormat);
                await xml.WriteEndElementAsync();
            }

            if (style.Background is not null)
            {
                await xml.WriteStartElementAsync(null, "fill", null);
                await xml.WriteStartElementAsync(null, "patternFill", null);
                await xml.WriteStartElementAsync(null, "bgColor", null);
                await xml.WriteAttributeStringAsync(null, "rgb", null, Argb(style.Background));
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }
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
