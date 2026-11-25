namespace Avig.Kalkyl.Internal;

/// <summary>
/// A range, optionally on a named sheet, as a chart series refers to it: <c>B2:B13</c>,
/// <c>Data!B2:B13</c> or <c>'Sales 2026'!B2:B13</c>.
/// </summary>
internal readonly record struct SheetReference(string? Sheet, CellRange Range)
{
    /// <summary>Parses a reference, or throws ArgumentException: references are written by the programmer.</summary>
    public static SheetReference Parse(string text, string parameter)
    {
        string? sheet = null;
        string range = text.Trim();
        int bang = range.LastIndexOf('!');

        if (bang >= 0)
        {
            sheet = range[..bang];
            range = range[(bang + 1)..];

            if (sheet.Length >= 2 && sheet.StartsWith('\'') == true && sheet.EndsWith('\'') == true)
            {
                sheet = sheet[1..^1].Replace("''", "'", StringComparison.Ordinal);
            }

            if (XlsxWriter.ValidateName(sheet).IsFailure(out string? error, out ResultStatus _) == true)
            {
                throw new ArgumentException($"'{text}' does not name a sheet Excel accepts: {error}", parameter);
            }
        }

        if (CellRange.TryParse(range, out CellRange parsed) == false)
        {
            throw new ArgumentException($"'{text}' is not a cell range. Expected a range such as B2:B13 or 'Sales 2026'!B2:B13.", parameter);
        }

        return new SheetReference(sheet, parsed);
    }

    /// <summary>The reference as a chart stores it: <c>'Sheet'!$B$2:$B$13</c>.</summary>
    public string ToFormula(string ownSheet)
    {
        string name = Sheet ?? ownSheet;

        return "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'!" + Range.ToAbsolute();
    }
}

/// <summary>
/// Where a drawing sits: across a block of cells, at a cell with a size in EMU, or, when
/// <see cref="Absolute"/> is set, at a place on the sheet regardless of the cells.
/// </summary>
internal readonly record struct DrawingAnchor(CellAddress From, CellRange? Area, long Width, long Height)
{
    public DrawingBounds? Absolute { get; init; }

    public static DrawingAnchor At(DrawingBounds bounds)
    {
        return new DrawingAnchor(default, null, bounds.Width.Emu, bounds.Height.Emu) { Absolute = bounds };
    }

    public static DrawingAnchor Over(CellRange area)
    {
        return new DrawingAnchor(area.First, area, 0, 0);
    }
}

/// <summary>A chart, a picture, a shape or a group on a sheet; a chart comes with its chart part, the series' references already resolved to formulas.</summary>
internal sealed record DrawingItem(Chart? Chart, ChartSpace? ChartSpace, Picture? Picture, DrawingAnchor Anchor)
{
    public Shape? Shape { get; init; }

    public ShapeGroup? Group { get; init; }

    /// <summary>The name of the slicer the anchor holds, when it holds one rather than a chart, a picture, a shape or a group.</summary>
    public string? Slicer { get; init; }
}

/// <summary>
/// The conditional formats, charts and pictures added to one sheet. Shared by
/// <see cref="SheetWriter"/> and <see cref="Worksheet"/>, which check the same things.
/// </summary>
internal sealed class SheetExtras
{
    /// <summary>EMU per pixel at 96 DPI, the size Excel gives a picture inserted at 100 %.</summary>
    public const long EmuPerPixel = 9525;

    private readonly List<(CellRange Range, ConditionalFormat Format)> conditionalFormats = [];
    private readonly List<DrawingItem> drawings = [];
    private readonly List<(CellAddress Cell, Hyperlink Link)> hyperlinks = [];

    public IReadOnlyList<(CellRange Range, ConditionalFormat Format)> ConditionalFormats => conditionalFormats;

    public IReadOnlyList<DrawingItem> Drawings => drawings;

    public bool HasCharts => drawings.Exists(item => item.Chart is not null);

    /// <summary>The links added to cells, at most one per cell: a later one replaces an earlier.</summary>
    public IReadOnlyList<(CellAddress Cell, Hyperlink Link)> Hyperlinks => hyperlinks;

    /// <summary>The links that go to a web address and so need a relationship from the sheet.</summary>
    public IEnumerable<Hyperlink> ExternalHyperlinks => hyperlinks.Where(pair => pair.Link.IsExternal == true).Select(pair => pair.Link);

    public void AddHyperlink(CellAddress cell, Hyperlink link)
    {
        ArgumentNullException.ThrowIfNull(link);
        hyperlinks.RemoveAll(pair => pair.Cell == cell);

        if (hyperlinks.Count >= 65_530)
        {
            throw new InvalidOperationException("A sheet holds at most 65 530 hyperlinks.");
        }

        hyperlinks.Add((cell, link));
    }

    public void AddShape(Shape shape, DrawingAnchor anchor)
    {
        ArgumentNullException.ThrowIfNull(shape);
        drawings.Add(new DrawingItem(null, null, null, anchor) { Shape = shape });
    }

    public void AddGroup(ShapeGroup group, DrawingAnchor anchor)
    {
        ArgumentNullException.ThrowIfNull(group);
        drawings.Add(new DrawingItem(null, null, null, anchor) { Group = group });
    }

    public void AddPicture(Picture picture, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(picture);
        drawings.Add(new DrawingItem(null, null, picture, DrawingAnchor.At(bounds)));
    }

    public void AddConditionalFormat(CellRange range, ConditionalFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);
        conditionalFormats.Add((range, format));
    }

    /// <summary>Adds a chart. NotFound when a series refers to a sheet that is not in <paramref name="sheetExists"/>.</summary>
    public Result AddChart(Chart chart, CellRange area, string ownSheet, Func<string, bool> sheetExists)
    {
        return AddChart(chart, DrawingAnchor.Over(area), ownSheet, sheetExists);
    }

    /// <summary>As <see cref="AddChart(Chart, CellRange, string, Func{string, bool})"/>, at any anchor.</summary>
    public Result AddChart(Chart chart, DrawingAnchor anchor, string ownSheet, Func<string, bool> sheetExists)
    {
        ArgumentNullException.ThrowIfNull(chart);

        if (chart.Series.Count == 0)
        {
            throw new ArgumentException("The chart has no series. Set Chart.Series.", nameof(chart));
        }

        chart.Validate();

        foreach (ChartSeries series in chart.Series)
        {
            foreach (SheetReference reference in new[] { series.Values }.Concat(Optional(series.Categories)))
            {
                if (reference.Sheet is not null && string.Equals(reference.Sheet, ownSheet, StringComparison.OrdinalIgnoreCase) == false && sheetExists(reference.Sheet) == false)
                {
                    return Result.NotFound($"The chart refers to the sheet '{reference.Sheet}', which the workbook does not have.");
                }
            }
        }

        drawings.Add(new DrawingItem(chart, chart.ToChartSpace(ownSheet), null, anchor));

        return Result.Success();
    }

    /// <summary>Adds the anchor of a slicer; the slicer itself is written with the pivot tables.</summary>
    public void AddSlicer(string name, CellRange area)
    {
        drawings.Add(new DrawingItem(null, null, null, new DrawingAnchor(area.First, area, 0, 0)) { Slicer = name });
    }

    public void AddPicture(Picture picture, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(picture);
        drawings.Add(new DrawingItem(null, null, picture, DrawingAnchor.Over(area)));
    }

    public void AddPicture(Picture picture, CellAddress topLeft, double scale)
    {
        ArgumentNullException.ThrowIfNull(picture);

        if (double.IsFinite(scale) == false || scale <= 0 || scale > 100)
        {
            throw new ArgumentOutOfRangeException(nameof(scale), scale, "The scale goes from just above 0 to 100; 1 is the picture's own size.");
        }

        // ST_PositiveCoordinate ends at 27 273 042 316 900 EMU.
        long width = Math.Clamp((long)Math.Round(picture.Width * EmuPerPixel * scale), 1, 27_273_042_316_900);
        long height = Math.Clamp((long)Math.Round(picture.Height * EmuPerPixel * scale), 1, 27_273_042_316_900);
        drawings.Add(new DrawingItem(null, null, picture, new DrawingAnchor(topLeft, null, width, height)));
    }

    private static IEnumerable<SheetReference> Optional(SheetReference? reference)
    {
        if (reference is SheetReference value)
        {
            yield return value;
        }
    }
}
