using System.Globalization;
using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes the parts behind a sheet's drawings: the drawing part with one anchor per chart,
/// picture, shape or group, and the relationships between them. Shared by the streaming writer
/// and the document model. The chart parts themselves are written by <see cref="ChartXml"/>.
/// </summary>
internal static class DrawingXml
{
    public const string SpreadsheetDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public const string DrawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public const string SvgNamespace = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

    public const string DrawingRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    public const string ChartRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    public const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    public const string HyperlinkRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

    public const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
    public const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

    /// <summary>The extension Excel 2016 put on a:blip for the SVG behind a raster picture.</summary>
    private const string SvgExtensionUri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";

    private const string Xdr = SpreadsheetDrawingNamespace;
    private const string Main = DrawingMainNamespace;

    /// <summary>Excel's default column width (64 pixels) and row height (20 pixels) in EMU, to estimate where a cell is.</summary>
    private const long DefaultColumnEmu = 609_600;
    private const long DefaultRowEmu = 190_500;

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
            Frame frame;

            if (anchor.Absolute is DrawingBounds bounds)
            {
                await xml.WriteStartElementAsync("xdr", "absoluteAnchor", Xdr);
                await xml.WriteStartElementAsync("xdr", "pos", Xdr);
                await xml.WriteAttributeStringAsync(null, "x", null, Text(bounds.Left));
                await xml.WriteAttributeStringAsync(null, "y", null, Text(bounds.Top));
                await xml.WriteEndElementAsync();
                await xml.WriteStartElementAsync("xdr", "ext", Xdr);
                await xml.WriteAttributeStringAsync(null, "cx", null, Text(bounds.Width));
                await xml.WriteAttributeStringAsync(null, "cy", null, Text(bounds.Height));
                await xml.WriteEndElementAsync();
                frame = new Frame(bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            else if (anchor.Area is CellRange area)
            {
                await xml.WriteStartElementAsync("xdr", "twoCellAnchor", Xdr);

                // Charts and pictures keep their size when rows and columns change; shapes stretch with the cells, as in Excel.
                if (item.Chart is not null || item.Picture is not null)
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
                frame = new Frame(area.First.Column * DefaultColumnEmu, area.First.Row * DefaultRowEmu, (lastColumn - area.First.Column) * DefaultColumnEmu, (lastRow - area.First.Row) * DefaultRowEmu);
            }
            else
            {
                await xml.WriteStartElementAsync("xdr", "oneCellAnchor", Xdr);
                await WriteCellMarkerAsync(xml, "from", anchor.From.Row, anchor.From.Column);
                await xml.WriteStartElementAsync("xdr", "ext", Xdr);
                await xml.WriteAttributeStringAsync(null, "cx", null, Text(anchor.Width));
                await xml.WriteAttributeStringAsync(null, "cy", null, Text(anchor.Height));
                await xml.WriteEndElementAsync();
                frame = new Frame(anchor.From.Column * DefaultColumnEmu, anchor.From.Row * DefaultRowEmu, anchor.Width, anchor.Height);
            }

            if (item.Chart is not null)
            {
                await WriteChartFrameAsync(xml, item, context, frame);
            }
            else if (item.Picture is not null)
            {
                await WritePictureAsync(xml, item.Picture, context, frame);
            }
            else if (item.Shape is not null)
            {
                await WriteShapeAsync(xml, item.Shape, context, frame);
            }
            else
            {
                await WriteGroupAsync(xml, item.Group!, context, frame);
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

    /// <summary>An a:solidFill, a:gradFill or a:noFill.</summary>
    public static async Task WriteFillAsync(XmlWriter xml, Fill fill)
    {
        switch (fill.Kind)
        {
            case FillKind.None:
                await xml.WriteElementStringAsync("a", "noFill", Main, string.Empty);

                break;

            case FillKind.Solid:
                await xml.WriteStartElementAsync("a", "solidFill", Main);
                await WriteColorAsync(xml, fill.Color!, fill.Transparency);
                await xml.WriteEndElementAsync();

                break;

            default:
                await xml.WriteStartElementAsync("a", "gradFill", Main);
                await xml.WriteAttributeStringAsync(null, "rotWithShape", null, "1");
                await xml.WriteStartElementAsync("a", "gsLst", Main);

                foreach (GradientStop stop in fill.Stops.OrderBy(stop => stop.Position))
                {
                    await xml.WriteStartElementAsync("a", "gs", Main);
                    await xml.WriteAttributeStringAsync(null, "pos", null, Text(DrawingValues.Percent(stop.Position)));
                    await WriteColorAsync(xml, stop.Color, stop.Transparency);
                    await xml.WriteEndElementAsync();
                }

                await xml.WriteEndElementAsync();
                await xml.WriteStartElementAsync("a", "lin", Main);
                await xml.WriteAttributeStringAsync(null, "ang", null, Text(DrawingValues.Angle(fill.Angle)));
                await xml.WriteAttributeStringAsync(null, "scaled", null, "0");
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;
        }
    }

    /// <summary>An a:ln.</summary>
    public static async Task WriteOutlineAsync(XmlWriter xml, Outline outline)
    {
        await xml.WriteStartElementAsync("a", "ln", Main);

        if (outline.Color is null)
        {
            await xml.WriteElementStringAsync("a", "noFill", Main, string.Empty);
            await xml.WriteEndElementAsync();

            return;
        }

        await xml.WriteAttributeStringAsync(null, "w", null, Text(DrawingValues.Emu(outline.Width)));
        await xml.WriteStartElementAsync("a", "solidFill", Main);
        await WriteColorAsync(xml, outline.Color, outline.Transparency);
        await xml.WriteEndElementAsync();

        if (outline.Dash != LineDash.Solid)
        {
            await xml.WriteStartElementAsync("a", "prstDash", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, DashName(outline.Dash));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>An a:effectLst with the shadows and soft edges given, in the schema's order.</summary>
    public static async Task WriteEffectsAsync(XmlWriter xml, Shadow? outer, Shadow? inner, double? softEdges)
    {
        await xml.WriteStartElementAsync("a", "effectLst", Main);

        if (inner is not null)
        {
            await xml.WriteStartElementAsync("a", "innerShdw", Main);
            await WriteShadowAttributesAsync(xml, inner);
            await WriteColorAsync(xml, inner.Color, inner.Transparency);
            await xml.WriteEndElementAsync();
        }

        if (outer is not null)
        {
            await xml.WriteStartElementAsync("a", "outerShdw", Main);
            await WriteShadowAttributesAsync(xml, outer);
            await xml.WriteAttributeStringAsync(null, "algn", null, "tl");
            await xml.WriteAttributeStringAsync(null, "rotWithShape", null, "0");
            await WriteColorAsync(xml, outer.Color, outer.Transparency);
            await xml.WriteEndElementAsync();
        }

        if (softEdges is double points)
        {
            await xml.WriteStartElementAsync("a", "softEdge", Main);
            await xml.WriteAttributeStringAsync(null, "rad", null, Text(DrawingValues.Emu(points)));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// An a:rPr, a:defRPr or a:endParaRPr. Bold and italic are written either way, so that a
    /// font given without them overrides a bold default such as a title's.
    /// </summary>
    public static async Task WriteRunPropertiesAsync(XmlWriter xml, string element, TextFont? font)
    {
        await xml.WriteStartElementAsync("a", element, Main);

        if (element != "defRPr")
        {
            await xml.WriteAttributeStringAsync(null, "lang", null, "en-US");
        }

        if (font is null)
        {
            await xml.WriteEndElementAsync();

            return;
        }

        if (font.Size is double size)
        {
            await xml.WriteAttributeStringAsync(null, "sz", null, Text((long)Math.Round(size * 100)));
        }

        await xml.WriteAttributeStringAsync(null, "b", null, Flag(font.Bold));
        await xml.WriteAttributeStringAsync(null, "i", null, Flag(font.Italic));

        if (font.Color is not null)
        {
            await xml.WriteStartElementAsync("a", "solidFill", Main);
            await WriteColorAsync(xml, font.Color, 0);
            await xml.WriteEndElementAsync();
        }

        if (font.Name is not null)
        {
            await xml.WriteStartElementAsync("a", "latin", Main);
            await xml.WriteAttributeStringAsync(null, "typeface", null, font.Name);
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync("a", "cs", Main);
            await xml.WriteAttributeStringAsync(null, "typeface", null, font.Name);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>Where a drawing is, in EMU: its frame (a:off and a:ext).</summary>
    private readonly record struct Frame(long Left, long Top, long Width, long Height);

    private static async Task WriteChartFrameAsync(XmlWriter xml, DrawingItem item, DrawingContext context, Frame frame)
    {
        int shapeId = context.NextShapeId();
        string relationshipId = context.Chart(item);
        await xml.WriteStartElementAsync("xdr", "graphicFrame", Xdr);
        await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
        await xml.WriteStartElementAsync("xdr", "nvGraphicFramePr", Xdr);
        await WriteNonVisualPropertiesAsync(xml, shapeId, "Chart " + Text(shapeId - 1), item.Chart!.Description, null, context);
        await xml.WriteElementStringAsync("xdr", "cNvGraphicFramePr", Xdr, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "xfrm", Xdr);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "graphic", Main);
        await xml.WriteStartElementAsync("a", "graphicData", Main);
        await xml.WriteAttributeStringAsync(null, "uri", null, ChartNamespace);
        await xml.WriteStartElementAsync("c", "chart", ChartNamespace);
        await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, relationshipId);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePictureAsync(XmlWriter xml, Picture picture, DrawingContext context, Frame frame)
    {
        int shapeId = context.NextShapeId();
        await xml.WriteStartElementAsync("xdr", "pic", Xdr);
        await xml.WriteStartElementAsync("xdr", "nvPicPr", Xdr);
        await WriteNonVisualPropertiesAsync(xml, shapeId, "Picture " + Text(shapeId - 1), picture.Description, picture.Link, context);
        await xml.WriteStartElementAsync("xdr", "cNvPicPr", Xdr);
        await xml.WriteStartElementAsync("a", "picLocks", Main);
        await xml.WriteAttributeStringAsync(null, "noChangeAspect", null, "1");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "blipFill", Xdr);
        await xml.WriteStartElementAsync("a", "blip", Main);

        // An SVG is stored as Excel 2016 stores it: the raster fallback is the blip, and the
        // SVG an extension of it that older versions skip.
        Picture raster = picture.Fallback ?? picture;
        await xml.WriteAttributeStringAsync("r", "embed", Ooxml.RelationshipsNamespace, context.Image(raster));

        if (picture.Transparency > 0)
        {
            await xml.WriteStartElementAsync("a", "alphaModFix", Main);
            await xml.WriteAttributeStringAsync(null, "amt", null, Text(DrawingValues.Percent(1 - picture.Transparency)));
            await xml.WriteEndElementAsync();
        }

        if (picture.Fallback is not null)
        {
            await xml.WriteStartElementAsync("a", "extLst", Main);
            await xml.WriteStartElementAsync("a", "ext", Main);
            await xml.WriteAttributeStringAsync(null, "uri", null, SvgExtensionUri);
            await xml.WriteStartElementAsync("asvg", "svgBlip", SvgNamespace);
            await xml.WriteAttributeStringAsync("r", "embed", Ooxml.RelationshipsNamespace, context.Image(picture));
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "stretch", Main);
        await xml.WriteElementStringAsync("a", "fillRect", Main, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "spPr", Xdr);
        await xml.WriteStartElementAsync("a", "xfrm", Main);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        await WritePresetGeometryAsync(xml, "rect", null);

        if (picture.Shadow is not null)
        {
            await WriteEffectsAsync(xml, picture.Shadow, null, null);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteShapeAsync(XmlWriter xml, Shape shape, DrawingContext context, Frame frame)
    {
        int shapeId = context.NextShapeId();
        string name = shape.Name ?? DefaultName(shape.Kind) + " " + Text(shapeId - 1);
        bool line = shape.Kind == ShapeKind.Line;
        bool textBox = shape.Kind == ShapeKind.TextBox;

        if (line == true)
        {
            await xml.WriteStartElementAsync("xdr", "cxnSp", Xdr);
            await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
            await xml.WriteStartElementAsync("xdr", "nvCxnSpPr", Xdr);
            await WriteNonVisualPropertiesAsync(xml, shapeId, name, shape.Description, shape.Link, context);
            await xml.WriteElementStringAsync("xdr", "cNvCxnSpPr", Xdr, string.Empty);
            await xml.WriteEndElementAsync();
        }
        else
        {
            await xml.WriteStartElementAsync("xdr", "sp", Xdr);
            await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
            await xml.WriteAttributeStringAsync(null, "textlink", null, string.Empty);
            await xml.WriteStartElementAsync("xdr", "nvSpPr", Xdr);
            await WriteNonVisualPropertiesAsync(xml, shapeId, name, shape.Description, shape.Link, context);
            await xml.WriteStartElementAsync("xdr", "cNvSpPr", Xdr);

            if (textBox == true)
            {
                await xml.WriteAttributeStringAsync(null, "txBox", null, "1");
            }

            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync("xdr", "spPr", Xdr);
        await xml.WriteStartElementAsync("a", "xfrm", Main);

        if (shape.Rotation != 0)
        {
            await xml.WriteAttributeStringAsync(null, "rot", null, Text(DrawingValues.Angle(shape.Rotation)));
        }

        if (shape.FlipHorizontal == true)
        {
            await xml.WriteAttributeStringAsync(null, "flipH", null, "1");
        }

        if (shape.FlipVertical == true)
        {
            await xml.WriteAttributeStringAsync(null, "flipV", null, "1");
        }

        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteEndElementAsync();
        string? adjust = null;

        if (shape.Kind == ShapeKind.RoundedRectangle && shape.CornerRadius is double radius)
        {
            adjust = "val " + Text(DrawingValues.Percent(radius));
        }

        await WritePresetGeometryAsync(xml, PresetName(shape.Kind), adjust);
        Fill? fill = shape.Fill;
        Outline? outline = shape.Outline;

        // A text box is text alone unless it is given a fill or a line; a line has no inside.
        if (textBox == true)
        {
            fill ??= Fill.None;
            outline ??= Outline.None;
        }

        if (fill is not null && line == false)
        {
            await WriteFillAsync(xml, fill);
        }

        if (outline is not null)
        {
            await WriteOutlineAsync(xml, outline);
        }

        if (shape.Shadow is not null || shape.InnerShadow is not null || shape.SoftEdges is not null)
        {
            await WriteEffectsAsync(xml, shape.Shadow, shape.InnerShadow, shape.SoftEdges);
        }

        await xml.WriteEndElementAsync();
        await WriteStyleAsync(xml, shape.Kind);

        if (line == false)
        {
            await WriteTextBodyAsync(xml, shape.Text);
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteGroupAsync(XmlWriter xml, ShapeGroup group, DrawingContext context, Frame frame)
    {
        int shapeId = context.NextShapeId();
        await xml.WriteStartElementAsync("xdr", "grpSp", Xdr);
        await xml.WriteStartElementAsync("xdr", "nvGrpSpPr", Xdr);
        await WriteNonVisualPropertiesAsync(xml, shapeId, group.Name ?? "Group " + Text(shapeId - 1), group.Description, group.Link, context);
        await xml.WriteElementStringAsync("xdr", "cNvGrpSpPr", Xdr, string.Empty);
        await xml.WriteEndElementAsync();

        // The items' own coordinates (chOff, chExt) are mapped onto the group's place, so the
        // box around them fills it.
        long left = group.Items.Min(item => item.Bounds.Left);
        long top = group.Items.Min(item => item.Bounds.Top);
        long right = group.Items.Max(item => item.Bounds.Left + item.Bounds.Width);
        long bottom = group.Items.Max(item => item.Bounds.Top + item.Bounds.Height);
        await xml.WriteStartElementAsync("xdr", "grpSpPr", Xdr);
        await xml.WriteStartElementAsync("a", "xfrm", Main);
        await WriteOffsetAndExtentAsync(xml, frame);
        await xml.WriteStartElementAsync("a", "chOff", Main);
        await xml.WriteAttributeStringAsync(null, "x", null, Text(left));
        await xml.WriteAttributeStringAsync(null, "y", null, Text(top));
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "chExt", Main);
        await xml.WriteAttributeStringAsync(null, "cx", null, Text(Math.Max(right - left, 1)));
        await xml.WriteAttributeStringAsync(null, "cy", null, Text(Math.Max(bottom - top, 1)));
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        foreach (GroupItem item in group.Items)
        {
            Frame itemFrame = new(item.Bounds.Left, item.Bounds.Top, item.Bounds.Width, item.Bounds.Height);

            if (item.Shape is not null)
            {
                await WriteShapeAsync(xml, item.Shape, context, itemFrame);
            }
            else
            {
                await WritePictureAsync(xml, item.Picture!, context, itemFrame);
            }
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>The theme references Excel gives a new shape, so that what is left unset looks as in Excel.</summary>
    private static async Task WriteStyleAsync(XmlWriter xml, ShapeKind kind)
    {
        (string Line, string Fill, string Effect, string Font, bool Shade) style = kind switch
        {
            ShapeKind.Line => ("1", "0", "0", "tx1", false),
            ShapeKind.TextBox => ("0", "0", "0", "dk1", false),
            _ => ("2", "1", "0", "lt1", true),
        };

        await xml.WriteStartElementAsync("xdr", "style", Xdr);
        await WriteStyleReferenceAsync(xml, "lnRef", style.Line, "accent1", style.Shade);
        await WriteStyleReferenceAsync(xml, "fillRef", style.Fill, "accent1", false);
        await WriteStyleReferenceAsync(xml, "effectRef", style.Effect, "accent1", false);
        await xml.WriteStartElementAsync("a", "fontRef", Main);
        await xml.WriteAttributeStringAsync(null, "idx", null, "minor");
        await xml.WriteStartElementAsync("a", "schemeClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, style.Font);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteStyleReferenceAsync(XmlWriter xml, string element, string index, string color, bool shade)
    {
        await xml.WriteStartElementAsync("a", element, Main);
        await xml.WriteAttributeStringAsync(null, "idx", null, index);
        await xml.WriteStartElementAsync("a", "schemeClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, color);

        if (shade == true)
        {
            await xml.WriteStartElementAsync("a", "shade", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, "15000");
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>The shape's xdr:txBody: one a:p per line of the runs, each run with its own look.</summary>
    private static async Task WriteTextBodyAsync(XmlWriter xml, ShapeText? text)
    {
        await xml.WriteStartElementAsync("xdr", "txBody", Xdr);
        await xml.WriteStartElementAsync("a", "bodyPr", Main);
        await xml.WriteAttributeStringAsync(null, "vertOverflow", null, "clip");
        await xml.WriteAttributeStringAsync(null, "horzOverflow", null, "clip");

        if (text is not null)
        {
            if (text.Wrap == true)
            {
                await xml.WriteAttributeStringAsync(null, "wrap", null, "square");
            }
            else
            {
                await xml.WriteAttributeStringAsync(null, "wrap", null, "none");
            }

            await WriteInsetAsync(xml, "lIns", text.MarginLeft);
            await WriteInsetAsync(xml, "tIns", text.MarginTop);
            await WriteInsetAsync(xml, "rIns", text.MarginRight);
            await WriteInsetAsync(xml, "bIns", text.MarginBottom);
        }

        await xml.WriteAttributeStringAsync(null, "rtlCol", null, "0");
        await xml.WriteAttributeStringAsync(null, "anchor", null, AnchorName(text?.Vertical ?? VerticalAlignment.Top));
        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "lstStyle", Main, string.Empty);
        string alignment = AlignmentName(text?.Horizontal ?? HorizontalAlignment.Left);

        if (text is null)
        {
            await xml.WriteStartElementAsync("a", "p", Main);
            await WriteParagraphPropertiesAsync(xml, alignment);
            await WriteRunPropertiesAsync(xml, "endParaRPr", new TextFont { Size = 11 });
            await xml.WriteEndElementAsync();
            await xml.WriteEndElementAsync();

            return;
        }

        foreach (List<TextRun> paragraph in Paragraphs(text.Runs))
        {
            await xml.WriteStartElementAsync("a", "p", Main);
            await WriteParagraphPropertiesAsync(xml, alignment);
            TextFont? last = null;

            foreach (TextRun run in paragraph)
            {
                last = run.Font;

                if (run.Text.Length == 0)
                {
                    continue;
                }

                await xml.WriteStartElementAsync("a", "r", Main);
                await WriteRunPropertiesAsync(xml, "rPr", run.Font);
                await xml.WriteElementStringAsync("a", "t", Main, run.Text);
                await xml.WriteEndElementAsync();
            }

            // Gives an empty line its height, and the cursor its look at the end of a line.
            await WriteRunPropertiesAsync(xml, "endParaRPr", last);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>The runs split at each \n into paragraphs; a run split so keeps its font on both sides.</summary>
    private static List<List<TextRun>> Paragraphs(IReadOnlyList<TextRun> runs)
    {
        List<List<TextRun>> paragraphs = [[]];

        foreach (TextRun run in runs)
        {
            string[] lines = run.Text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');

            for (int index = 0; index < lines.Length; index++)
            {
                if (index > 0)
                {
                    paragraphs.Add([]);
                }

                paragraphs[^1].Add(new TextRun(lines[index], run.Font));
            }
        }

        return paragraphs;
    }

    private static async Task WriteParagraphPropertiesAsync(XmlWriter xml, string alignment)
    {
        await xml.WriteStartElementAsync("a", "pPr", Main);
        await xml.WriteAttributeStringAsync(null, "algn", null, alignment);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteInsetAsync(XmlWriter xml, string name, double? points)
    {
        if (points is double value)
        {
            await xml.WriteAttributeStringAsync(null, name, null, Text(DrawingValues.Emu(value)));
        }
    }

    private static async Task WritePresetGeometryAsync(XmlWriter xml, string preset, string? adjust)
    {
        await xml.WriteStartElementAsync("a", "prstGeom", Main);
        await xml.WriteAttributeStringAsync(null, "prst", null, preset);
        await xml.WriteStartElementAsync("a", "avLst", Main);

        if (adjust is not null)
        {
            await xml.WriteStartElementAsync("a", "gd", Main);
            await xml.WriteAttributeStringAsync(null, "name", null, "adj");
            await xml.WriteAttributeStringAsync(null, "fmla", null, adjust);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>An xdr:cNvPr: id, name, alternative text and the click's link.</summary>
    private static async Task WriteNonVisualPropertiesAsync(XmlWriter xml, int shapeId, string name, string? description, Hyperlink? link, DrawingContext context)
    {
        await xml.WriteStartElementAsync("xdr", "cNvPr", Xdr);
        await xml.WriteAttributeStringAsync(null, "id", null, Text(shapeId));
        await xml.WriteAttributeStringAsync(null, "name", null, name);

        if (description is not null)
        {
            await xml.WriteAttributeStringAsync(null, "descr", null, description);
        }

        if (link is not null)
        {
            await xml.WriteStartElementAsync("a", "hlinkClick", Main);
            await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, context.Link(link));

            if (link.Tooltip is not null)
            {
                await xml.WriteAttributeStringAsync(null, "tooltip", null, link.Tooltip);
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteShadowAttributesAsync(XmlWriter xml, Shadow shadow)
    {
        await xml.WriteAttributeStringAsync(null, "blurRad", null, Text(DrawingValues.Emu(shadow.Blur)));
        await xml.WriteAttributeStringAsync(null, "dist", null, Text(DrawingValues.Emu(shadow.Distance)));
        await xml.WriteAttributeStringAsync(null, "dir", null, Text(DrawingValues.Angle(shadow.Angle)));
    }

    /// <summary>An a:srgbClr, with an a:alpha when the colour is see-through.</summary>
    private static async Task WriteColorAsync(XmlWriter xml, string color, double transparency)
    {
        await xml.WriteStartElementAsync("a", "srgbClr", Main);
        await xml.WriteAttributeStringAsync(null, "val", null, color[1..]);

        if (transparency > 0)
        {
            await xml.WriteStartElementAsync("a", "alpha", Main);
            await xml.WriteAttributeStringAsync(null, "val", null, Text(DrawingValues.Percent(1 - transparency)));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteOffsetAndExtentAsync(XmlWriter xml, Frame frame)
    {
        await xml.WriteStartElementAsync("a", "off", Main);
        await xml.WriteAttributeStringAsync(null, "x", null, Text(frame.Left));
        await xml.WriteAttributeStringAsync(null, "y", null, Text(frame.Top));
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "ext", Main);
        await xml.WriteAttributeStringAsync(null, "cx", null, Text(frame.Width));
        await xml.WriteAttributeStringAsync(null, "cy", null, Text(frame.Height));
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

    private static string PresetName(ShapeKind kind)
    {
        return kind switch
        {
            ShapeKind.RoundedRectangle => "roundRect",
            ShapeKind.Ellipse => "ellipse",
            ShapeKind.Triangle => "triangle",
            ShapeKind.Diamond => "diamond",
            ShapeKind.Hexagon => "hexagon",
            ShapeKind.RightArrow => "rightArrow",
            ShapeKind.LeftArrow => "leftArrow",
            ShapeKind.UpArrow => "upArrow",
            ShapeKind.DownArrow => "downArrow",
            ShapeKind.Chevron => "chevron",
            ShapeKind.Pentagon => "homePlate",
            ShapeKind.Line => "line",
            _ => "rect",
        };
    }

    private static string DefaultName(ShapeKind kind)
    {
        return kind switch
        {
            ShapeKind.RoundedRectangle => "Rectangle: Rounded Corners",
            ShapeKind.Ellipse => "Oval",
            ShapeKind.Triangle => "Isosceles Triangle",
            ShapeKind.RightArrow => "Arrow: Right",
            ShapeKind.LeftArrow => "Arrow: Left",
            ShapeKind.UpArrow => "Arrow: Up",
            ShapeKind.DownArrow => "Arrow: Down",
            ShapeKind.Chevron => "Arrow: Chevron",
            ShapeKind.Pentagon => "Arrow: Pentagon",
            ShapeKind.Line => "Straight Connector",
            ShapeKind.TextBox => "TextBox",
            _ => kind.ToString(),
        };
    }

    private static string DashName(LineDash dash)
    {
        return dash switch
        {
            LineDash.RoundDot => "sysDot",
            LineDash.SquareDot => "sysDash",
            LineDash.Dash => "dash",
            LineDash.DashDot => "dashDot",
            LineDash.LongDash => "lgDash",
            LineDash.LongDashDot => "lgDashDot",
            LineDash.LongDashDotDot => "lgDashDotDot",
            _ => "solid",
        };
    }

    private static string AnchorName(VerticalAlignment alignment)
    {
        return alignment switch
        {
            VerticalAlignment.Center => "ctr",
            VerticalAlignment.Bottom => "b",
            _ => "t",
        };
    }

    private static string AlignmentName(HorizontalAlignment alignment)
    {
        return alignment switch
        {
            HorizontalAlignment.Center => "ctr",
            HorizontalAlignment.Right => "r",
            HorizontalAlignment.Justify => "just",
            _ => "l",
        };
    }

    private static string Flag(bool value)
    {
        if (value == true)
        {
            return "1";
        }

        return "0";
    }

    private static string Text(long number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>A relationship in a .rels part; External ones get TargetMode="External".</summary>
internal readonly record struct PartRelationship(string Id, string Type, string Target, bool External = false);

/// <summary>
/// What one drawing part refers to, collected while its XML is written: shape ids in order, and
/// a relationship id per chart, image and link. Images and links used more than once share one.
/// The parts behind them are written after the drawing, since a zip has one entry open at a time.
/// </summary>
internal sealed class DrawingContext
{
    private readonly HashSet<string> usedIds;
    private readonly Dictionary<object, string> images = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<Hyperlink, string> links = [];
    private int nextShapeId;

    public DrawingContext(HashSet<string> usedIds, int firstShapeId)
    {
        this.usedIds = usedIds;
        nextShapeId = firstShapeId;
    }

    /// <summary>The charts, images and links, in the order their ids were handed out.</summary>
    public List<DrawingTarget> Targets { get; } = [];

    public int NextShapeId()
    {
        int id = nextShapeId;
        nextShapeId++;

        return id;
    }

    public string Chart(DrawingItem item)
    {
        string id = DrawingPackage.NextId(usedIds);
        Targets.Add(new DrawingTarget(id, item, null, null));

        return id;
    }

    public string Image(Picture picture)
    {
        if (images.TryGetValue(picture.Identity, out string? existing) == true)
        {
            return existing;
        }

        string id = DrawingPackage.NextId(usedIds);
        images[picture.Identity] = id;
        Targets.Add(new DrawingTarget(id, null, picture, null));

        return id;
    }

    public string Link(Hyperlink link)
    {
        if (links.TryGetValue(link, out string? existing) == true)
        {
            return existing;
        }

        string id = DrawingPackage.NextId(usedIds);
        links[link] = id;
        Targets.Add(new DrawingTarget(id, null, null, link));

        return id;
    }
}

/// <summary>A chart, an image or a link a drawing refers to by <see cref="Id"/>.</summary>
internal sealed record DrawingTarget(string Id, DrawingItem? Chart, Picture? Image, Hyperlink? Link);
