using System.Globalization;
using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes the parts behind a sheet's drawings: the drawing part with one anchor per chart or
/// picture, a chart part per chart, and the relationships between them. Shared by the streaming
/// writer and the document model.
/// </summary>
internal static class DrawingXml
{
    public const string SpreadsheetDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public const string DrawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public const string DrawingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    public const string ChartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    public const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    public const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
    public const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

    private const string CategoryAxisId = "500000001";
    private const string ValueAxisId = "500000002";

    /// <summary>A complete drawing part.</summary>
    public static async Task WriteDrawingAsync(XmlWriter xml, IReadOnlyList<DrawingItem> items, IReadOnlyList<string> relationshipIds, int firstShapeId)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync("xdr", "wsDr", SpreadsheetDrawingNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "a", null, DrawingMainNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await WriteAnchorsAsync(xml, items, relationshipIds, firstShapeId);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>The anchors alone, to append to the drawing part of an opened file.</summary>
    public static async Task WriteAnchorsAsync(XmlWriter xml, IReadOnlyList<DrawingItem> items, IReadOnlyList<string> relationshipIds, int firstShapeId)
    {
        for (int index = 0; index < items.Count; index++)
        {
            DrawingItem item = items[index];
            int shapeId = firstShapeId + index;
            DrawingAnchor anchor = item.Anchor;

            if (anchor.Area is CellRange area)
            {
                await xml.WriteStartElementAsync("xdr", "twoCellAnchor", SpreadsheetDrawingNamespace);
                await xml.WriteAttributeStringAsync(null, "editAs", null, "oneCell");
                await WriteCellMarkerAsync(xml, "from", area.First.Row, area.First.Column);

                // The end marker is the top-left corner of the cell after the area, so that the
                // area's last row and column are covered.
                await WriteCellMarkerAsync(xml, "to", Math.Min(area.Last.Row + 1, CellAddress.MaxRows - 1), Math.Min(area.Last.Column + 1, CellAddress.MaxColumns - 1));
            }
            else
            {
                await xml.WriteStartElementAsync("xdr", "oneCellAnchor", SpreadsheetDrawingNamespace);
                await WriteCellMarkerAsync(xml, "from", anchor.From.Row, anchor.From.Column);
                await xml.WriteStartElementAsync("xdr", "ext", SpreadsheetDrawingNamespace);
                await xml.WriteAttributeStringAsync(null, "cx", null, Text(anchor.Width));
                await xml.WriteAttributeStringAsync(null, "cy", null, Text(anchor.Height));
                await xml.WriteEndElementAsync();
            }

            if (item.Chart is not null)
            {
                await WriteChartFrameAsync(xml, item.Chart, relationshipIds[index], shapeId);
            }
            else if (item.Slicer is not null)
            {
                // The rectangle older versions show: the area at Excel's default column width and row height.
                CellRange covered = anchor.Area ?? new CellRange(anchor.From, anchor.From);
                await PivotXml.WriteSlicerFrameAsync(xml, item.Slicer, shapeId, covered.First.Column * 609600L, covered.First.Row * 190500L, covered.ColumnCount * 609600L, covered.RowCount * 190500L);
            }
            else
            {
                await WritePictureAsync(xml, item.Picture!, relationshipIds[index], shapeId, anchor);
            }

            await xml.WriteElementStringAsync("xdr", "clientData", SpreadsheetDrawingNamespace, string.Empty);
            await xml.WriteEndElementAsync();
        }
    }

    /// <summary>A .rels part.</summary>
    public static async Task WriteRelationshipsAsync(XmlWriter xml, IEnumerable<(string Id, string Type, string Target)> relationships)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
        await WriteRelationshipElementsAsync(xml, relationships);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    public static async Task WriteRelationshipElementsAsync(XmlWriter xml, IEnumerable<(string Id, string Type, string Target)> relationships)
    {
        foreach ((string id, string type, string target) in relationships)
        {
            await xml.WriteStartElementAsync(null, "Relationship", null);
            await xml.WriteAttributeStringAsync(null, "Id", null, id);
            await xml.WriteAttributeStringAsync(null, "Type", null, type);
            await xml.WriteAttributeStringAsync(null, "Target", null, target);
            await xml.WriteEndElementAsync();
        }
    }

    /// <summary>The &lt;drawing r:id&gt; element of a worksheet.</summary>
    public static async Task WriteSheetDrawingAsync(XmlWriter xml, string relationshipId)
    {
        await xml.WriteStartElementAsync(null, "drawing", null);
        await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, relationshipId);
        await xml.WriteEndElementAsync();
    }

    /// <summary>A complete chart part.</summary>
    public static async Task WriteChartAsync(XmlWriter xml, DrawingItem item)
    {
        Chart chart = item.Chart!;
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync("c", "chartSpace", ChartNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "a", null, DrawingMainNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);

        // Without this, a chart has rounded corners: the schema's default is the opposite of Excel's.
        await WriteValueAsync(xml, "roundedCorners", "0");
        await xml.WriteStartElementAsync("c", "chart", ChartNamespace);

        if (chart.Title is not null)
        {
            await WriteTitleAsync(xml, chart.Title, false);
            await WriteValueAsync(xml, "autoTitleDeleted", "0");
        }
        else
        {
            // Otherwise Excel shows the name of a lone series as the title.
            await WriteValueAsync(xml, "autoTitleDeleted", "1");
        }

        await xml.WriteStartElementAsync("c", "plotArea", ChartNamespace);
        await xml.WriteElementStringAsync("c", "layout", ChartNamespace, string.Empty);
        await WritePlotAsync(xml, chart, item.Series);
        await WriteAxesAsync(xml, chart);
        await xml.WriteEndElementAsync();

        if (chart.Legend != LegendPosition.None)
        {
            await xml.WriteStartElementAsync("c", "legend", ChartNamespace);
            await WriteValueAsync(xml, "legendPos", LegendName(chart.Legend));
            await WriteValueAsync(xml, "overlay", "0");
            await xml.WriteEndElementAsync();
        }

        await WriteValueAsync(xml, "plotVisOnly", "1");
        await WriteValueAsync(xml, "dispBlanksAs", "gap");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WritePlotAsync(XmlWriter xml, Chart chart, IReadOnlyList<ResolvedSeries> series)
    {
        string grouping = "standard";

        if (chart.Stacked == true)
        {
            grouping = "stacked";
        }

        switch (chart.Type)
        {
            case ChartType.Column or ChartType.Bar:
                await xml.WriteStartElementAsync("c", "barChart", ChartNamespace);

                if (chart.Type == ChartType.Bar)
                {
                    await WriteValueAsync(xml, "barDir", "bar");
                }
                else
                {
                    await WriteValueAsync(xml, "barDir", "col");
                }

                if (chart.Stacked == true)
                {
                    await WriteValueAsync(xml, "grouping", "stacked");
                }
                else
                {
                    await WriteValueAsync(xml, "grouping", "clustered");
                }

                await WriteValueAsync(xml, "varyColors", "0");
                await WriteSeriesAsync(xml, chart, series);
                await WriteValueAsync(xml, "gapWidth", "150");

                if (chart.Stacked == true)
                {
                    await WriteValueAsync(xml, "overlap", "100");
                }

                await WriteAxisIdsAsync(xml);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Line:
                await xml.WriteStartElementAsync("c", "lineChart", ChartNamespace);
                await WriteValueAsync(xml, "grouping", grouping);
                await WriteValueAsync(xml, "varyColors", "0");
                await WriteSeriesAsync(xml, chart, series);
                await WriteValueAsync(xml, "marker", "1");
                await WriteAxisIdsAsync(xml);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Area:
                await xml.WriteStartElementAsync("c", "areaChart", ChartNamespace);
                await WriteValueAsync(xml, "grouping", grouping);
                await WriteValueAsync(xml, "varyColors", "0");
                await WriteSeriesAsync(xml, chart, series);
                await WriteAxisIdsAsync(xml);
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Pie:
                await xml.WriteStartElementAsync("c", "pieChart", ChartNamespace);
                await WriteValueAsync(xml, "varyColors", "1");
                await WriteSeriesAsync(xml, chart, series);
                await WriteValueAsync(xml, "firstSliceAng", "0");
                await xml.WriteEndElementAsync();

                break;

            case ChartType.Doughnut:
                await xml.WriteStartElementAsync("c", "doughnutChart", ChartNamespace);
                await WriteValueAsync(xml, "varyColors", "1");
                await WriteSeriesAsync(xml, chart, series);
                await WriteValueAsync(xml, "firstSliceAng", "0");
                await WriteValueAsync(xml, "holeSize", "50");
                await xml.WriteEndElementAsync();

                break;

            default:
                await xml.WriteStartElementAsync("c", "scatterChart", ChartNamespace);
                await WriteValueAsync(xml, "scatterStyle", "lineMarker");
                await WriteValueAsync(xml, "varyColors", "0");
                await WriteSeriesAsync(xml, chart, series);
                await WriteAxisIdsAsync(xml);
                await xml.WriteEndElementAsync();

                break;
        }
    }

    /// <summary>The series, each with its children in the order its type's schema gives.</summary>
    private static async Task WriteSeriesAsync(XmlWriter xml, Chart chart, IReadOnlyList<ResolvedSeries> series)
    {
        bool pie = chart.Type is ChartType.Pie or ChartType.Doughnut;

        for (int index = 0; index < series.Count; index++)
        {
            ResolvedSeries resolved = series[index];
            await xml.WriteStartElementAsync("c", "ser", ChartNamespace);
            await WriteValueAsync(xml, "idx", Text(index));
            await WriteValueAsync(xml, "order", Text(index));

            if (resolved.Series.Name is not null)
            {
                await xml.WriteStartElementAsync("c", "tx", ChartNamespace);
                await xml.WriteElementStringAsync("c", "v", ChartNamespace, resolved.Series.Name);
                await xml.WriteEndElementAsync();
            }

            // A pie colours each slice, so one colour for the whole series would hide them.
            string? color = resolved.Series.Color;

            if (pie == true)
            {
                color = null;
            }

            switch (chart.Type)
            {
                case ChartType.Line:
                    if (color is not null)
                    {
                        await xml.WriteStartElementAsync("c", "spPr", ChartNamespace);
                        await WriteLineAsync(xml, color);
                        await xml.WriteEndElementAsync();
                        await WriteMarkerAsync(xml, color);
                    }

                    break;

                case ChartType.Scatter:
                    // Points without lines between them, as Excel's plain scatter chart.
                    await xml.WriteStartElementAsync("c", "spPr", ChartNamespace);
                    await xml.WriteStartElementAsync("a", "ln", DrawingMainNamespace);
                    await xml.WriteAttributeStringAsync(null, "w", null, "19050");
                    await xml.WriteElementStringAsync("a", "noFill", DrawingMainNamespace, string.Empty);
                    await xml.WriteEndElementAsync();
                    await xml.WriteEndElementAsync();

                    if (color is not null)
                    {
                        await WriteMarkerAsync(xml, color);
                    }

                    break;

                default:
                    if (color is not null)
                    {
                        await xml.WriteStartElementAsync("c", "spPr", ChartNamespace);
                        await WriteSolidFillAsync(xml, color);
                        await xml.WriteEndElementAsync();
                    }

                    break;
            }

            if (chart.Type is ChartType.Column or ChartType.Bar)
            {
                await WriteValueAsync(xml, "invertIfNegative", "0");
            }

            if (chart.Type == ChartType.Scatter)
            {
                if (resolved.Categories is not null)
                {
                    await WriteReferenceAsync(xml, "xVal", "numRef", resolved.Categories);
                }

                await WriteReferenceAsync(xml, "yVal", "numRef", resolved.Values);
            }
            else
            {
                if (resolved.Categories is not null)
                {
                    await WriteReferenceAsync(xml, "cat", "strRef", resolved.Categories);
                }

                await WriteReferenceAsync(xml, "val", "numRef", resolved.Values);
            }

            if (chart.Type is ChartType.Line or ChartType.Scatter)
            {
                await WriteValueAsync(xml, "smooth", "0");
            }

            await xml.WriteEndElementAsync();
        }
    }

    private static async Task WriteAxesAsync(XmlWriter xml, Chart chart)
    {
        if (chart.Type is ChartType.Pie or ChartType.Doughnut)
        {
            return;
        }

        string categoryPosition = "b";
        string valuePosition = "l";

        if (chart.Type == ChartType.Bar)
        {
            categoryPosition = "l";
            valuePosition = "b";
        }

        if (chart.Type == ChartType.Scatter)
        {
            // Two value axes: x along the bottom, y up the side.
            await WriteValueAxisAsync(xml, CategoryAxisId, ValueAxisId, "b", chart.CategoryAxisTitle, false, "midCat");
            await WriteValueAxisAsync(xml, ValueAxisId, CategoryAxisId, "l", chart.ValueAxisTitle, true, "midCat");

            return;
        }

        await xml.WriteStartElementAsync("c", "catAx", ChartNamespace);
        await WriteValueAsync(xml, "axId", CategoryAxisId);
        await WriteScalingAsync(xml);
        await WriteValueAsync(xml, "delete", "0");
        await WriteValueAsync(xml, "axPos", categoryPosition);

        if (chart.CategoryAxisTitle is not null)
        {
            await WriteTitleAsync(xml, chart.CategoryAxisTitle, false);
        }

        await WriteValueAsync(xml, "majorTickMark", "out");
        await WriteValueAsync(xml, "minorTickMark", "none");
        await WriteValueAsync(xml, "tickLblPos", "nextTo");
        await WriteValueAsync(xml, "crossAx", ValueAxisId);
        await WriteValueAsync(xml, "crosses", "autoZero");
        await WriteValueAsync(xml, "auto", "1");
        await WriteValueAsync(xml, "lblAlgn", "ctr");
        await WriteValueAsync(xml, "lblOffset", "100");
        await WriteValueAsync(xml, "noMultiLvlLbl", "0");
        await xml.WriteEndElementAsync();

        await WriteValueAxisAsync(xml, ValueAxisId, CategoryAxisId, valuePosition, chart.ValueAxisTitle, true, "between");
    }

    private static async Task WriteValueAxisAsync(XmlWriter xml, string id, string crossing, string position, string? title, bool gridlines, string crossBetween)
    {
        await xml.WriteStartElementAsync("c", "valAx", ChartNamespace);
        await WriteValueAsync(xml, "axId", id);
        await WriteScalingAsync(xml);
        await WriteValueAsync(xml, "delete", "0");
        await WriteValueAsync(xml, "axPos", position);

        if (gridlines == true)
        {
            await xml.WriteElementStringAsync("c", "majorGridlines", ChartNamespace, string.Empty);
        }

        if (title is not null)
        {
            await WriteTitleAsync(xml, title, position == "l");
        }

        await xml.WriteStartElementAsync("c", "numFmt", ChartNamespace);
        await xml.WriteAttributeStringAsync(null, "formatCode", null, "General");
        await xml.WriteAttributeStringAsync(null, "sourceLinked", null, "1");
        await xml.WriteEndElementAsync();
        await WriteValueAsync(xml, "majorTickMark", "out");
        await WriteValueAsync(xml, "minorTickMark", "none");
        await WriteValueAsync(xml, "tickLblPos", "nextTo");
        await WriteValueAsync(xml, "crossAx", crossing);
        await WriteValueAsync(xml, "crosses", "autoZero");
        await WriteValueAsync(xml, "crossBetween", crossBetween);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteScalingAsync(XmlWriter xml)
    {
        await xml.WriteStartElementAsync("c", "scaling", ChartNamespace);
        await WriteValueAsync(xml, "orientation", "minMax");
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteAxisIdsAsync(XmlWriter xml)
    {
        await WriteValueAsync(xml, "axId", CategoryAxisId);
        await WriteValueAsync(xml, "axId", ValueAxisId);
    }

    /// <summary>A title with one run of text. <paramref name="vertical"/> turns it to read upwards, as Excel does beside a vertical axis.</summary>
    private static async Task WriteTitleAsync(XmlWriter xml, string text, bool vertical)
    {
        await xml.WriteStartElementAsync("c", "title", ChartNamespace);
        await xml.WriteStartElementAsync("c", "tx", ChartNamespace);
        await xml.WriteStartElementAsync("c", "rich", ChartNamespace);
        await xml.WriteStartElementAsync("a", "bodyPr", DrawingMainNamespace);

        if (vertical == true)
        {
            await xml.WriteAttributeStringAsync(null, "rot", null, "-5400000");
            await xml.WriteAttributeStringAsync(null, "vert", null, "horz");
        }

        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "lstStyle", DrawingMainNamespace, string.Empty);
        await xml.WriteStartElementAsync("a", "p", DrawingMainNamespace);
        await xml.WriteStartElementAsync("a", "r", DrawingMainNamespace);
        await xml.WriteElementStringAsync("a", "t", DrawingMainNamespace, text);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await WriteValueAsync(xml, "overlay", "0");
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteReferenceAsync(XmlWriter xml, string element, string kind, string formula)
    {
        await xml.WriteStartElementAsync("c", element, ChartNamespace);
        await xml.WriteStartElementAsync("c", kind, ChartNamespace);
        await xml.WriteElementStringAsync("c", "f", ChartNamespace, formula);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteMarkerAsync(XmlWriter xml, string color)
    {
        await xml.WriteStartElementAsync("c", "marker", ChartNamespace);
        await WriteValueAsync(xml, "symbol", "circle");
        await WriteValueAsync(xml, "size", "5");
        await xml.WriteStartElementAsync("c", "spPr", ChartNamespace);
        await WriteSolidFillAsync(xml, color);
        await WriteLineAsync(xml, color);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteLineAsync(XmlWriter xml, string color)
    {
        await xml.WriteStartElementAsync("a", "ln", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "w", null, "28575");
        await WriteSolidFillAsync(xml, color);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteSolidFillAsync(XmlWriter xml, string color)
    {
        await xml.WriteStartElementAsync("a", "solidFill", DrawingMainNamespace);
        await xml.WriteStartElementAsync("a", "srgbClr", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "val", null, color[1..]);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteChartFrameAsync(XmlWriter xml, Chart chart, string relationshipId, int shapeId)
    {
        await xml.WriteStartElementAsync("xdr", "graphicFrame", SpreadsheetDrawingNamespace);
        await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
        await xml.WriteStartElementAsync("xdr", "nvGraphicFramePr", SpreadsheetDrawingNamespace);
        await WriteShapePropertiesAsync(xml, shapeId, "Chart " + Text(shapeId), chart.Description);
        await xml.WriteElementStringAsync("xdr", "cNvGraphicFramePr", SpreadsheetDrawingNamespace, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "xfrm", SpreadsheetDrawingNamespace);
        await WriteOffsetAndExtentAsync(xml, 0, 0);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "graphic", DrawingMainNamespace);
        await xml.WriteStartElementAsync("a", "graphicData", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "uri", null, ChartNamespace);
        await xml.WriteStartElementAsync("c", "chart", ChartNamespace);
        await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, relationshipId);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePictureAsync(XmlWriter xml, Picture picture, string relationshipId, int shapeId, DrawingAnchor anchor)
    {
        await xml.WriteStartElementAsync("xdr", "pic", SpreadsheetDrawingNamespace);
        await xml.WriteStartElementAsync("xdr", "nvPicPr", SpreadsheetDrawingNamespace);
        await WriteShapePropertiesAsync(xml, shapeId, "Picture " + Text(shapeId), picture.Description);
        await xml.WriteStartElementAsync("xdr", "cNvPicPr", SpreadsheetDrawingNamespace);
        await xml.WriteStartElementAsync("a", "picLocks", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "noChangeAspect", null, "1");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "blipFill", SpreadsheetDrawingNamespace);
        await xml.WriteStartElementAsync("a", "blip", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync("r", "embed", Ooxml.RelationshipsNamespace, relationshipId);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "stretch", DrawingMainNamespace);
        await xml.WriteElementStringAsync("a", "fillRect", DrawingMainNamespace, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "spPr", SpreadsheetDrawingNamespace);
        await xml.WriteStartElementAsync("a", "xfrm", DrawingMainNamespace);
        await WriteOffsetAndExtentAsync(xml, anchor.Width, anchor.Height);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "prstGeom", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "prst", null, "rect");
        await xml.WriteElementStringAsync("a", "avLst", DrawingMainNamespace, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteShapePropertiesAsync(XmlWriter xml, int shapeId, string name, string? description)
    {
        await xml.WriteStartElementAsync("xdr", "cNvPr", SpreadsheetDrawingNamespace);
        await xml.WriteAttributeStringAsync(null, "id", null, Text(shapeId));
        await xml.WriteAttributeStringAsync(null, "name", null, name);

        if (description is not null)
        {
            await xml.WriteAttributeStringAsync(null, "descr", null, description);
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteOffsetAndExtentAsync(XmlWriter xml, long width, long height)
    {
        await xml.WriteStartElementAsync("a", "off", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "x", null, "0");
        await xml.WriteAttributeStringAsync(null, "y", null, "0");
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "ext", DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "cx", null, Text(width));
        await xml.WriteAttributeStringAsync(null, "cy", null, Text(height));
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteCellMarkerAsync(XmlWriter xml, string name, int row, int column)
    {
        await xml.WriteStartElementAsync("xdr", name, SpreadsheetDrawingNamespace);
        await xml.WriteElementStringAsync("xdr", "col", SpreadsheetDrawingNamespace, Text(column));
        await xml.WriteElementStringAsync("xdr", "colOff", SpreadsheetDrawingNamespace, "0");
        await xml.WriteElementStringAsync("xdr", "row", SpreadsheetDrawingNamespace, Text(row));
        await xml.WriteElementStringAsync("xdr", "rowOff", SpreadsheetDrawingNamespace, "0");
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteValueAsync(XmlWriter xml, string name, string value)
    {
        await xml.WriteStartElementAsync("c", name, ChartNamespace);
        await xml.WriteAttributeStringAsync(null, "val", null, value);
        await xml.WriteEndElementAsync();
    }

    private static string LegendName(LegendPosition position)
    {
        return position switch
        {
            LegendPosition.Bottom => "b",
            LegendPosition.Top => "t",
            LegendPosition.Left => "l",
            _ => "r",
        };
    }

    private static string Text(long number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}
