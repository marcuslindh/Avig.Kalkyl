using System.Globalization;
using System.Xml;
using Avig.OpenXml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes the parts behind a sheet's drawings: the drawing part with one anchor per chart,
/// picture, shape or group, and the relationships between them. Shared by the streaming writer
/// and the document model. The anchors (xdr:twoCellAnchor, xdr:oneCellAnchor,
/// xdr:absoluteAnchor) are Excel's and are written here; what is inside them is DrawingML and is
/// written by Avig.OpenXml's <see cref="DrawingWriter"/>, in the xdr namespace. The chart parts
/// are written by <see cref="ChartXml"/>.
/// </summary>
internal static class DrawingXml
{
    public const string SpreadsheetDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public const string DrawingMainNamespace = OpcNamespaces.DrawingMain;
    public const string ChartNamespace = OpcNamespaces.DrawingChart;

    public const string DrawingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    public const string ChartRelationship = OpcNamespaces.ChartRelationship;
    public const string ImageRelationship = OpcNamespaces.ImageRelationship;
    public const string HyperlinkRelationship = OpcNamespaces.HyperlinkRelationship;

    public const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
    public const string ChartContentType = OpcNamespaces.ChartContentType;

    private const string Xdr = SpreadsheetDrawingNamespace;
    private const string Main = DrawingMainNamespace;

    /// <summary>Excel's default column width (64 pixels) and row height (20 pixels) in EMU, to estimate where a cell is.</summary>
    private const long DefaultColumnEmu = 609_600;
    private const long DefaultRowEmu = 190_500;

    /// <summary>The shape tree of a spreadsheet drawing: xdr elements, with the empty macro attributes Excel writes.</summary>
    public static ShapeTreeNamespace Tree { get; } = new("xdr", Xdr) { MacroAttributes = true };

    /// <summary>A context for one drawing part.</summary>
    public static DrawingContext NewContext(HashSet<string> usedIds, int firstShapeId)
    {
        return new DrawingContext(Tree, usedIds, firstShapeId);
    }

    /// <summary>A complete drawing part.</summary>
    public static async Task WriteDrawingAsync(XmlWriter xml, IReadOnlyList<DrawingItem> items, DrawingContext context)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync("xdr", "wsDr", Xdr);
        await xml.WriteAttributeStringAsync("xmlns", "a", null, Main);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await WriteAnchorsAsync(xml, items, context);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>The anchors alone, to append to the drawing part of an opened file.</summary>
    public static async Task WriteAnchorsAsync(XmlWriter xml, IReadOnlyList<DrawingItem> items, DrawingContext context)
    {
        foreach (DrawingItem item in items)
        {
            DrawingAnchor anchor = item.Anchor;
            DrawingBounds frame;

            if (anchor.Absolute is DrawingBounds bounds)
            {
                await xml.WriteStartElementAsync("xdr", "absoluteAnchor", Xdr);
                await xml.WriteStartElementAsync("xdr", "pos", Xdr);
                await xml.WriteAttributeStringAsync(null, "x", null, Text(bounds.Left.Emu));
                await xml.WriteAttributeStringAsync(null, "y", null, Text(bounds.Top.Emu));
                await xml.WriteEndElementAsync();
                await xml.WriteStartElementAsync("xdr", "ext", Xdr);
                await xml.WriteAttributeStringAsync(null, "cx", null, Text(bounds.Width.Emu));
                await xml.WriteAttributeStringAsync(null, "cy", null, Text(bounds.Height.Emu));
                await xml.WriteEndElementAsync();
                frame = bounds;
            }
            else if (anchor.Area is CellRange area)
            {
                await xml.WriteStartElementAsync("xdr", "twoCellAnchor", Xdr);

                // Charts, pictures and slicers keep their size when rows and columns change; shapes stretch with the cells, as in Excel.
                if (item.Chart is not null || item.Picture is not null || item.Slicer is not null)
                {
                    await xml.WriteAttributeStringAsync(null, "editAs", null, "oneCell");
                }

                await WriteCellMarkerAsync(xml, "from", area.First.Row, area.First.Column);

                // The end marker is the top-left corner of the cell after the area, so that the
                // area's last row and column are covered.
                int lastRow = Math.Min(area.Last.Row + 1, CellAddress.MaxRows - 1);
                int lastColumn = Math.Min(area.Last.Column + 1, CellAddress.MaxColumns - 1);
                await WriteCellMarkerAsync(xml, "to", lastRow, lastColumn);

                // Excel places the drawing by the markers; the frame only has to be plausible.
                frame = DrawingBounds.FromEmu(area.First.Column * DefaultColumnEmu, area.First.Row * DefaultRowEmu, (lastColumn - area.First.Column) * DefaultColumnEmu, (lastRow - area.First.Row) * DefaultRowEmu);
            }
            else
            {
                await xml.WriteStartElementAsync("xdr", "oneCellAnchor", Xdr);
                await WriteCellMarkerAsync(xml, "from", anchor.From.Row, anchor.From.Column);
                await xml.WriteStartElementAsync("xdr", "ext", Xdr);
                await xml.WriteAttributeStringAsync(null, "cx", null, Text(anchor.Width));
                await xml.WriteAttributeStringAsync(null, "cy", null, Text(anchor.Height));
                await xml.WriteEndElementAsync();
                frame = DrawingBounds.FromEmu(anchor.From.Column * DefaultColumnEmu, anchor.From.Row * DefaultRowEmu, anchor.Width, anchor.Height);
            }

            if (item.Chart is not null)
            {
                await DrawingWriter.WriteChartFrameAsync(xml, item, item.Chart.Description, context, frame);
            }
            else if (item.Picture is not null)
            {
                await DrawingWriter.WritePictureAsync(xml, item.Picture, context, frame);
            }
            else if (item.Shape is not null)
            {
                await DrawingWriter.WriteShapeAsync(xml, item.Shape, context, frame);
            }
            else if (item.Slicer is not null)
            {
                // The rectangle older versions show in its place.
                await PivotXml.WriteSlicerFrameAsync(xml, item.Slicer, context.NextShapeId(), frame.Left.Emu, frame.Top.Emu, frame.Width.Emu, frame.Height.Emu);
            }
            else
            {
                await DrawingWriter.WriteGroupAsync(xml, item.Group!, context, frame);
            }

            await xml.WriteElementStringAsync("xdr", "clientData", Xdr, string.Empty);
            await xml.WriteEndElementAsync();
        }
    }

    /// <summary>A .rels part.</summary>
    public static async Task WriteRelationshipsAsync(XmlWriter xml, IEnumerable<PartRelationship> relationships)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
        await WriteRelationshipElementsAsync(xml, relationships);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    public static async Task WriteRelationshipElementsAsync(XmlWriter xml, IEnumerable<PartRelationship> relationships)
    {
        foreach (PartRelationship relationship in relationships)
        {
            await xml.WriteStartElementAsync(null, "Relationship", null);
            await xml.WriteAttributeStringAsync(null, "Id", null, relationship.Id);
            await xml.WriteAttributeStringAsync(null, "Type", null, relationship.Type);
            await xml.WriteAttributeStringAsync(null, "Target", null, relationship.Target);

            if (relationship.External == true)
            {
                await xml.WriteAttributeStringAsync(null, "TargetMode", null, "External");
            }

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

    private static async Task WriteCellMarkerAsync(XmlWriter xml, string name, int row, int column)
    {
        await xml.WriteStartElementAsync("xdr", name, Xdr);
        await xml.WriteElementStringAsync("xdr", "col", Xdr, Text(column));
        await xml.WriteElementStringAsync("xdr", "colOff", Xdr, "0");
        await xml.WriteElementStringAsync("xdr", "row", Xdr, Text(row));
        await xml.WriteElementStringAsync("xdr", "rowOff", Xdr, "0");
        await xml.WriteEndElementAsync();
    }

    private static string Text(long number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>A relationship in a .rels part; External ones get TargetMode="External".</summary>
internal readonly record struct PartRelationship(string Id, string Type, string Target, bool External = false);
