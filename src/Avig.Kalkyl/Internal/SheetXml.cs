using System.Globalization;
using System.Xml;
using System.Xml.Linq;

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

    /// <summary>The view attributes the model owns; the others on an opened sheet's first view are kept.</summary>
    private static readonly HashSet<string> ViewAttributes = new(StringComparer.Ordinal)
    {
        "tabSelected", "showGridLines", "showRowColHeaders", "rightToLeft", "zoomScale", "zoomScaleNormal", "workbookViewId", "topLeftCell",
    };

    /// <summary>
    /// The &lt;sheetViews&gt; element. For an opened sheet, <paramref name="original"/> is the file's
    /// own: the first view keeps the attributes the model does not own, and further views and
    /// extensions are kept; its panes and selections are replaced by the model's.
    /// </summary>
    public static async Task WriteSheetViewsAsync(XmlWriter xml, bool selected, int frozenRows, int frozenColumns, SheetView view, XElement? original = null)
    {
        XElement? originalView = original?.Elements().FirstOrDefault(element => element.Name.LocalName == "sheetView");
        await xml.WriteStartElementAsync(null, "sheetViews", null);
        await xml.WriteStartElementAsync(null, "sheetView", null);

        if (selected == true)
        {
            await xml.WriteAttributeStringAsync(null, "tabSelected", null, "1");
        }

        if (view.ShowGridLines == false)
        {
            await xml.WriteAttributeStringAsync(null, "showGridLines", null, "0");
        }

        if (view.ShowHeaders == false)
        {
            await xml.WriteAttributeStringAsync(null, "showRowColHeaders", null, "0");
        }

        if (view.RightToLeft == true)
        {
            await xml.WriteAttributeStringAsync(null, "rightToLeft", null, "1");
        }

        if (view.Zoom != 100)
        {
            await xml.WriteAttributeStringAsync(null, "zoomScale", null, Text(view.Zoom));
            await xml.WriteAttributeStringAsync(null, "zoomScaleNormal", null, Text(view.Zoom));
        }

        if (originalView is not null)
        {
            foreach (XAttribute attribute in originalView.Attributes())
            {
                if (attribute.IsNamespaceDeclaration == false && attribute.Name.NamespaceName.Length == 0 && ViewAttributes.Contains(attribute.Name.LocalName) == false)
                {
                    await xml.WriteAttributeStringAsync(null, attribute.Name.LocalName, null, attribute.Value);
                }
            }
        }

        await xml.WriteAttributeStringAsync(null, "workbookViewId", null, "0");
        string? activePane = null;

        if (frozenRows > 0 || frozenColumns > 0)
        {
            activePane = "bottomRight";

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
        }

        if (activePane is not null || view.ActiveCell is not null)
        {
            await xml.WriteStartElementAsync(null, "selection", null);

            if (activePane is not null)
            {
                await xml.WriteAttributeStringAsync(null, "pane", null, activePane);
            }

            if (view.ActiveCell is CellAddress active)
            {
                await xml.WriteAttributeStringAsync(null, "activeCell", null, active.ToString());
                await xml.WriteAttributeStringAsync(null, "sqref", null, active.ToString());
            }

            await xml.WriteEndElementAsync();
        }

        // pivotSelection and extLst come after the panes and selections in a sheetView.
        if (originalView is not null)
        {
            foreach (XElement child in originalView.Elements().Where(element => element.Name.LocalName is "pivotSelection" or "extLst"))
            {
                await child.WriteToAsync(xml, CancellationToken.None);
            }
        }

        await xml.WriteEndElementAsync();

        if (original is not null)
        {
            foreach (XElement other in original.Elements().Where(element => element != originalView))
            {
                await other.WriteToAsync(xml, CancellationToken.None);
            }
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// The &lt;sheetPr&gt; element with the tab's colour. For an opened sheet, <paramref name="original"/>
    /// is the file's own, whose attributes and other children are kept.
    /// </summary>
    public static async Task WriteSheetPropertiesAsync(XmlWriter xml, string? tabColor, XElement? original = null)
    {
        if (tabColor is null && original is null)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, "sheetPr", null);

        if (original is not null)
        {
            foreach (XAttribute attribute in original.Attributes().Where(attribute => attribute.IsNamespaceDeclaration == false))
            {
                if (attribute.Name.NamespaceName.Length == 0)
                {
                    await xml.WriteAttributeStringAsync(null, attribute.Name.LocalName, null, attribute.Value);
                }
                else
                {
                    await xml.WriteAttributeStringAsync(original.GetPrefixOfNamespace(attribute.Name.Namespace), attribute.Name.LocalName, attribute.Name.NamespaceName, attribute.Value);
                }
            }
        }

        // tabColor is the first child of sheetPr.
        if (tabColor is not null)
        {
            await xml.WriteStartElementAsync(null, "tabColor", null);
            await xml.WriteAttributeStringAsync(null, "rgb", null, "FF" + tabColor[1..]);
            await xml.WriteEndElementAsync();
        }

        if (original is not null)
        {
            foreach (XElement child in original.Elements().Where(element => element.Name.LocalName != "tabColor"))
            {
                await child.WriteToAsync(xml, CancellationToken.None);
            }
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// The &lt;hyperlinks&gt; element. <paramref name="externalIds"/> holds the relationship id of
    /// each link to a web address, in order. For an opened sheet, <paramref name="original"/> is
    /// the file's own, whose links are kept except where a new one is on the same cell.
    /// </summary>
    public static async Task WriteHyperlinksAsync(XmlWriter xml, IReadOnlyList<(CellAddress Cell, Hyperlink Link)> links, IReadOnlyList<string> externalIds, XElement? original = null)
    {
        HashSet<string> cells = new(links.Select(pair => pair.Cell.ToString()), StringComparer.OrdinalIgnoreCase);
        await xml.WriteStartElementAsync(null, "hyperlinks", null);

        if (original is not null)
        {
            foreach (XElement existing in original.Elements())
            {
                if (existing.Name.LocalName == "hyperlink" && existing.Attribute("ref") is XAttribute reference && cells.Contains(reference.Value) == true)
                {
                    continue;
                }

                await existing.WriteToAsync(xml, CancellationToken.None);
            }
        }

        int external = 0;

        foreach ((CellAddress cell, Hyperlink link) in links)
        {
            await xml.WriteStartElementAsync(null, "hyperlink", null);
            await xml.WriteAttributeStringAsync(null, "ref", null, cell.ToString());

            if (link.IsExternal == true)
            {
                await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, externalIds[external]);
                external++;
            }
            else
            {
                await xml.WriteAttributeStringAsync(null, "location", null, link.Location);
            }

            if (link.Tooltip is not null)
            {
                await xml.WriteAttributeStringAsync(null, "tooltip", null, link.Tooltip);
            }

            await xml.WriteEndElementAsync();
        }

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
