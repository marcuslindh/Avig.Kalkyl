using System.Globalization;
using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// A &lt;col&gt; element: a run of columns sharing settings. Attributes the package does not
/// interpret (style, hidden, outline level ...) are kept as they were read.
/// </summary>
/// <param name="First">The first column, 0-based.</param>
/// <param name="Last">The last column, 0-based, inclusive.</param>
/// <param name="Width">The width, or null when the element sets none.</param>
/// <param name="OtherAttributes">Every other attribute, by name.</param>
internal sealed record ColumnRange(int First, int Last, double? Width, IReadOnlyList<KeyValuePair<string, string>> OtherAttributes);

/// <summary>Writes the parts of a worksheet that surround the cells. Shared by the streaming writer and the document model.</summary>
internal static class SheetXml
{
    /// <summary>
    /// Element order inside &lt;worksheet&gt;, from the schema. Elements must appear in this order,
    /// so inserting one means finding the first existing element that comes after it.
    /// </summary>
    public static readonly string[] WorksheetOrder =
    [
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
        "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
        "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
        "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
    ];

    /// <summary>True when the schema names the element, so that its position says where others go.</summary>
    public static bool IsKnown(string localName)
    {
        return Array.IndexOf(WorksheetOrder, localName) >= 0;
    }

    public static int OrderOf(string localName)
    {
        int index = Array.IndexOf(WorksheetOrder, localName);

        if (index < 0)
        {
            return WorksheetOrder.Length;
        }

        return index;
    }

    public static async Task WriteSheetViewsAsync(XmlWriter xml, bool selected, int frozenRows, int frozenColumns)
    {
        await xml.WriteStartElementAsync(null, "sheetViews", null);
        await xml.WriteStartElementAsync(null, "sheetView", null);

        if (selected == true)
        {
            await xml.WriteAttributeStringAsync(null, "tabSelected", null, "1");
        }

        await xml.WriteAttributeStringAsync(null, "workbookViewId", null, "0");

        if (frozenRows > 0 || frozenColumns > 0)
        {
            string activePane = "bottomRight";

            if (frozenColumns == 0)
            {
                activePane = "bottomLeft";
            }
            else if (frozenRows == 0)
            {
                activePane = "topRight";
            }

            await xml.WriteStartElementAsync(null, "pane", null);

            if (frozenColumns > 0)
            {
                await xml.WriteAttributeStringAsync(null, "xSplit", null, Text(frozenColumns));
            }

            if (frozenRows > 0)
            {
                await xml.WriteAttributeStringAsync(null, "ySplit", null, Text(frozenRows));
            }

            await xml.WriteAttributeStringAsync(null, "topLeftCell", null, new CellAddress(frozenRows, frozenColumns).ToString());
            await xml.WriteAttributeStringAsync(null, "activePane", null, activePane);
            await xml.WriteAttributeStringAsync(null, "state", null, "frozen");
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync(null, "selection", null);
            await xml.WriteAttributeStringAsync(null, "pane", null, activePane);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    public static async Task WriteColumnsAsync(XmlWriter xml, IReadOnlyList<ColumnRange> columns)
    {
        if (columns.Count == 0)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, "cols", null);

        foreach (ColumnRange column in columns)
        {
            await xml.WriteStartElementAsync(null, "col", null);
            await xml.WriteAttributeStringAsync(null, "min", null, Text(column.First + 1));
            await xml.WriteAttributeStringAsync(null, "max", null, Text(column.Last + 1));

            if (column.Width is not null)
            {
                await xml.WriteAttributeStringAsync(null, "width", null, column.Width.Value.ToString("R", CultureInfo.InvariantCulture));
                await xml.WriteAttributeStringAsync(null, "customWidth", null, "1");
            }

            foreach ((string name, string value) in column.OtherAttributes)
            {
                await xml.WriteAttributeStringAsync(null, name, null, value);
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    public static async Task WriteAutoFilterAsync(XmlWriter xml, CellAddress end)
    {
        await xml.WriteStartElementAsync(null, "autoFilter", null);
        await xml.WriteAttributeStringAsync(null, "ref", null, "A1:" + end.ToString());
        await xml.WriteEndElementAsync();
    }

    public static string Text(int number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}
