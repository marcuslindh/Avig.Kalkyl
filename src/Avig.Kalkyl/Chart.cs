using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// One row or column of numbers in a chart. The references are to cells in the workbook, such as
/// <c>B2:B13</c> on the chart's own sheet or <c>'Sales 2026'!B2:B13</c> on another. Excel reads
/// the values from the cells when it opens the file, so the chart follows later changes to them.
/// How the series looks is a <see cref="ChartSpaceSeries"/> of Avig.OpenXml, which this type
/// fills in.
/// </summary>
public sealed class ChartSeries
{
    private readonly ChartSpaceSeries format;

    /// <summary>Creates a series.</summary>
    /// <param name="name">The name in the legend, or null for "Series 1" and so on.</param>
    /// <param name="values">The numbers: one row or one column of cells.</param>
    /// <param name="categories">
    /// The labels along the axis (the x values of a scatter chart): as many cells as
    /// <paramref name="values"/>, or null for 1, 2, 3 ...
    /// </param>
    /// <param name="color">The colour as <c>#RRGGBB</c>, or null for Excel's own.</param>
    public ChartSeries(string? name, string values, string? categories = null, string? color = null)
    {
        ArgumentNullException.ThrowIfNull(values);
        Values = SheetReference.Parse(values, nameof(values));

        if (Values.Range.RowCount > 1 && Values.Range.ColumnCount > 1)
        {
            throw new ArgumentException($"The values of a series ({values}) must be one row or one column.", nameof(values));
        }

        if (categories is not null)
        {
            Categories = SheetReference.Parse(categories, nameof(categories));

            if (Categories.Value.Range.RowCount * Categories.Value.Range.ColumnCount != Values.Range.RowCount * Values.Range.ColumnCount)
            {
                throw new ArgumentException($"The categories ({categories}) and the values ({values}) are not the same number of cells.", nameof(categories));
            }
        }

        if (name is not null && name.Length > 255)
        {
            throw new ArgumentException("A series name holds at most 255 characters.", nameof(name));
        }

        if (color is not null)
        {
            Color = new Style { Background = color }.Background;
        }

        if (name is not null)
        {
            DrawingValues.XmlText(name, nameof(name));
        }

        Name = name;

        // The formulas are filled in when the chart is put on a sheet, which says what B2:B13 is relative to.
        format = new ChartSpaceSeries(ChartData.Reference(values.Trim())) { Name = name, Color = Color };
    }

    /// <summary>The name in the legend.</summary>
    public string? Name { get; }

    /// <summary>The colour, as <c>#RRGGBB</c>, or null for Excel's own.</summary>
    public string? Color { get; }

    internal SheetReference Values { get; }

    internal SheetReference? Categories { get; }

    /// <summary>
    /// Draw this series as another type than the chart's, for a combination chart: columns with
    /// a line, say. Column, line and area mix; null keeps the chart's type.
    /// </summary>
    public ChartType? Type
    {
        get => format.Type;
        init => format = format with { Type = value };
    }

    /// <summary>Plot the series against a second value axis on the right, for numbers of another size.</summary>
    public bool SecondaryAxis
    {
        get => format.SecondaryAxis;
        init => format = format with { SecondaryAxis = value };
    }

    /// <summary>The inside of the columns, bars, areas or slices; wins over <see cref="Color"/>. <see cref="Fill.None"/> hides them, as the first series of a funnel.</summary>
    public Fill? Fill
    {
        get => format.Fill;
        init => format = format with { Fill = value };
    }

    /// <summary>The line of a line or scatter series, or the border of columns, bars, areas and slices; wins over <see cref="Color"/> for lines.</summary>
    public Outline? Outline
    {
        get => format.Outline;
        init => format = format with { Outline = value };
    }

    /// <summary>A shadow behind the series, or null for none.</summary>
    public Shadow? Shadow
    {
        get => format.Shadow;
        init => format = format with { Shadow = value };
    }

    /// <summary>
    /// A colour per point, as <c>#RRGGBB</c>, or null to leave that point as the series is: the
    /// slices of a pie, or one bar picked out among the rest.
    /// </summary>
    public IReadOnlyList<string?> PointColors
    {
        get => format.PointColors;
        init => format = format with { PointColors = value };
    }

    /// <summary>Curve the line through the points instead of straight segments. For line and scatter series.</summary>
    public bool Smooth
    {
        get => format.Smooth;
        init => format = format with { Smooth = value };
    }

    /// <summary>The marker on each point of a line or scatter series, or null for Excel's own.</summary>
    public MarkerStyle? Marker
    {
        get => format.Marker;
        init => format = format with { Marker = value };
    }

    /// <summary>The marker's size in points, 2 to 72, or null for Excel's 5.</summary>
    public int? MarkerSize
    {
        get => format.MarkerSize;
        init => format = format with { MarkerSize = value };
    }

    /// <summary>Labels on this series' points, or null for the chart's <see cref="Chart.DataLabels"/>.</summary>
    public DataLabels? DataLabels
    {
        get => format.DataLabels;
        init => format = format with { DataLabels = value };
    }

    /// <summary>The series as the chart part stores it, with its references as formulas on <paramref name="ownSheet"/>.</summary>
    internal ChartSpaceSeries ToChartSpace(string ownSheet)
    {
        ChartData? categories = null;

        if (Categories is SheetReference reference)
        {
            categories = ChartData.Reference(reference.ToFormula(ownSheet));
        }

        return format with { Values = ChartData.Reference(Values.ToFormula(ownSheet)), Categories = categories };
    }

    /// <summary>The series without resolved references, for the checks that do not need them.</summary>
    internal ChartSpaceSeries Format => format;
}

/// <summary>
/// A chart over cells in the workbook. Put it on a sheet with
/// <see cref="SheetWriter.AddChart(Chart, CellRange)"/> or <see cref="Worksheet.AddChart(Chart, CellRange)"/>.
/// How it looks is a <see cref="ChartSpace"/> of Avig.OpenXml, which this type fills in; the
/// series are <see cref="ChartSeries"/>, whose references are cells in the workbook.
/// </summary>
/// <example>
/// <code>
/// Chart chart = new(ChartType.Column)
/// {
///     Title = "Sales per month",
///     Series = [new ChartSeries("2026", "B2:B13", "A2:A13")],
/// };
/// sheet.AddChart(chart, "E2:L20");
/// </code>
/// </example>
public sealed record Chart
{
    private readonly IReadOnlyList<ChartSeries> series = [];
    private readonly ChartSpace space;

    /// <summary>Creates a chart.</summary>
    /// <param name="type">The kind of chart.</param>
    public Chart(ChartType type)
    {
        space = new ChartSpace(type);
    }

    /// <summary>The kind of chart.</summary>
    public ChartType Type => space.Type;

    /// <summary>The title above the chart, or null for none.</summary>
    public string? Title
    {
        get => space.Title;
        init => space = space with { Title = value };
    }

    /// <summary>The series, at least one and at most 255.</summary>
    public IReadOnlyList<ChartSeries> Series
    {
        get => series;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Count > 255)
            {
                throw new ArgumentException("A chart holds at most 255 series.", nameof(Series));
            }

            series = [.. value];
        }
    }

    /// <summary>Where the legend goes.</summary>
    public LegendPosition Legend
    {
        get => space.Legend;
        init => space = space with { Legend = value };
    }

    /// <summary>
    /// Stack the series on top of each other instead of side by side. For column, bar, line and
    /// area charts.
    /// </summary>
    public bool Stacked
    {
        get => space.Stacked;
        init => space = space with { Stacked = value };
    }

    /// <summary>The title under the category axis (the x axis of a scatter chart), or null.</summary>
    public string? CategoryAxisTitle
    {
        get => space.CategoryAxisTitle;
        init => space = space with { CategoryAxisTitle = value };
    }

    /// <summary>The title beside the value axis, or null.</summary>
    public string? ValueAxisTitle
    {
        get => space.ValueAxisTitle;
        init => space = space with { ValueAxisTitle = value };
    }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description { get; init; }

    /// <summary>The size of a doughnut's hole in percent of its diameter, 10 to 90, or null for Excel's 50.</summary>
    public int? HoleSize
    {
        get => space.HoleSize;
        init => space = space with { HoleSize = value };
    }

    /// <summary>Where a pie's or doughnut's first slice starts, in degrees clockwise from the top, 0 to 360.</summary>
    public int FirstSliceAngle
    {
        get => space.FirstSliceAngle;
        init => space = space with { FirstSliceAngle = value };
    }

    /// <summary>The space between groups of columns or bars in percent of a column's width, 0 to 500, or null for Excel's 150.</summary>
    public int? GapWidth
    {
        get => space.GapWidth;
        init => space = space with { GapWidth = value };
    }

    /// <summary>
    /// How much the columns or bars of a group overlap in percent, -100 to 100, or null for
    /// Excel's 0 (100 when stacked). 100 draws one series on top of another, as a progress bar
    /// over its background.
    /// </summary>
    public int? Overlap
    {
        get => space.Overlap;
        init => space = space with { Overlap = value };
    }

    /// <summary>
    /// Give each point its own colour, or null for Excel's default: on for pie and doughnut
    /// charts, off for the rest. Off, a doughnut's slices share the series' colour.
    /// </summary>
    public bool? VaryColors
    {
        get => space.VaryColors;
        init => space = space with { VaryColors = value };
    }

    /// <summary>Round the corners of the chart's border.</summary>
    public bool RoundedCorners
    {
        get => space.RoundedCorners;
        init => space = space with { RoundedCorners = value };
    }

    /// <summary>The chart area behind everything, or null for white; <see cref="Fill.None"/> lets the sheet or a shape show through.</summary>
    public Fill? Background
    {
        get => space.Background;
        init => space = space with { Background = value };
    }

    /// <summary>The plot area inside the axes, or null for Excel's white; see-through instead when <see cref="Background"/> is set.</summary>
    public Fill? PlotBackground
    {
        get => space.PlotBackground;
        init => space = space with { PlotBackground = value };
    }

    /// <summary>The chart's border, or null for Excel's thin grey; <see cref="Outline.None"/> removes it.</summary>
    public Outline? Border
    {
        get => space.Border;
        init => space = space with { Border = value };
    }

    /// <summary>The font of all the chart's text, or null for Excel's; the title, legend, axes and labels can override it.</summary>
    public TextFont? Font
    {
        get => space.Font;
        init => space = space with { Font = value };
    }

    /// <summary>The title's font, or null for the chart's.</summary>
    public TextFont? TitleFont
    {
        get => space.TitleFont;
        init => space = space with { TitleFont = value };
    }

    /// <summary>The legend's font, or null for the chart's.</summary>
    public TextFont? LegendFont
    {
        get => space.LegendFont;
        init => space = space with { LegendFont = value };
    }

    /// <summary>The category axis (the x axis of a scatter chart), or null for Excel's default.</summary>
    public ChartAxis? CategoryAxis
    {
        get => space.CategoryAxis;
        init => space = space with { CategoryAxis = value };
    }

    /// <summary>The value axis, or null for Excel's default.</summary>
    public ChartAxis? ValueAxis
    {
        get => space.ValueAxis;
        init => space = space with { ValueAxis = value };
    }

    /// <summary>The second value axis, on the right, of the series with <see cref="ChartSeries.SecondaryAxis"/>, or null for Excel's default.</summary>
    public ChartAxis? SecondaryValueAxis
    {
        get => space.SecondaryValueAxis;
        init => space = space with { SecondaryValueAxis = value };
    }

    /// <summary>Labels on every series' points, or null for none. A series' own <see cref="ChartSeries.DataLabels"/> wins.</summary>
    public DataLabels? DataLabels
    {
        get => space.DataLabels;
        init => space = space with { DataLabels = value };
    }

    /// <summary>The chart part for a sheet: this chart's look, and the series with their references as formulas on <paramref name="ownSheet"/>.</summary>
    internal ChartSpace ToChartSpace(string ownSheet)
    {
        return space with { Series = Series.Select(item => item.ToChartSpace(ownSheet)).ToList() };
    }

    /// <summary>Throws ArgumentException for what Excel would repair: a mix of types it does not combine, or a label position the type does not take.</summary>
    internal void Validate()
    {
        (space with { Series = Series.Select(item => item.Format).ToList() }).Validate();
    }
}
