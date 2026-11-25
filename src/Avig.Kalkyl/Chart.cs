using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>The kind of chart, as on Excel's Insert menu.</summary>
public enum ChartType
{
    /// <summary>Vertical bars.</summary>
    Column,

    /// <summary>Horizontal bars.</summary>
    Bar,

    /// <summary>Lines with markers.</summary>
    Line,

    /// <summary>Filled areas under lines.</summary>
    Area,

    /// <summary>A pie, from the first series.</summary>
    Pie,

    /// <summary>A ring, one per series.</summary>
    Doughnut,

    /// <summary>Points at x and y; the series' categories are the x values and must be numbers.</summary>
    Scatter,
}

/// <summary>Where the legend goes.</summary>
public enum LegendPosition
{
    /// <summary>No legend.</summary>
    None,

    /// <summary>To the right of the plot; Excel's default.</summary>
    Right,

    /// <summary>Below the plot.</summary>
    Bottom,

    /// <summary>Above the plot.</summary>
    Top,

    /// <summary>To the left of the plot.</summary>
    Left,
}

/// <summary>
/// One row or column of numbers in a chart. The references are to cells in the workbook, such as
/// <c>B2:B13</c> on the chart's own sheet or <c>'Sales 2026'!B2:B13</c> on another. Excel reads
/// the values from the cells when it opens the file, so the chart follows later changes to them.
/// </summary>
public sealed class ChartSeries
{
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
        get => type;
        init
        {
            if (value is ChartType given && Enum.IsDefined(given) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Type), given, "Not a chart type.");
            }

            type = value;
        }
    }

    /// <summary>Plot the series against a second value axis on the right, for numbers of another size.</summary>
    public bool SecondaryAxis { get; init; }

    /// <summary>The inside of the columns, bars, areas or slices; wins over <see cref="Color"/>. <see cref="Fill.None"/> hides them, as the first series of a funnel.</summary>
    public Fill? Fill { get; init; }

    /// <summary>The line of a line or scatter series, or the border of columns, bars, areas and slices; wins over <see cref="Color"/> for lines.</summary>
    public Outline? Outline { get; init; }

    /// <summary>A shadow behind the series, or null for none.</summary>
    public Shadow? Shadow { get; init; }

    /// <summary>
    /// A colour per point, as <c>#RRGGBB</c>, or null to leave that point as the series is: the
    /// slices of a pie, or one bar picked out among the rest.
    /// </summary>
    public IReadOnlyList<string?> PointColors
    {
        get => pointColors;
        init
        {
            ArgumentNullException.ThrowIfNull(value);
            List<string?> colors = [];

            foreach (string? color in value)
            {
                string? normalized = null;

                if (color is not null)
                {
                    normalized = DrawingValues.Color(color, nameof(PointColors));
                }

                colors.Add(normalized);
            }

            pointColors = colors;
        }
    }

    /// <summary>Curve the line through the points instead of straight segments. For line and scatter series.</summary>
    public bool Smooth { get; init; }

    /// <summary>The marker on each point of a line or scatter series, or null for Excel's own.</summary>
    public MarkerStyle? Marker
    {
        get => marker;
        init
        {
            if (value is MarkerStyle given && Enum.IsDefined(given) == false)
            {
                throw new ArgumentOutOfRangeException(nameof(Marker), given, "Not a marker style.");
            }

            marker = value;
        }
    }

    /// <summary>The marker's size in points, 2 to 72, or null for Excel's 5.</summary>
    public int? MarkerSize
    {
        get => markerSize;
        init
        {
            if (value is int size)
            {
                DrawingValues.Range(size, 2, 72, nameof(MarkerSize), "A marker size");
            }

            markerSize = value;
        }
    }

    /// <summary>Labels on this series' points, or null for the chart's <see cref="Chart.DataLabels"/>.</summary>
    public DataLabels? DataLabels { get; init; }

    private ChartType? type;
    private IReadOnlyList<string?> pointColors = [];
    private MarkerStyle? marker;
    private int? markerSize;
}

/// <summary>The marker shapes of a line or scatter series.</summary>
public enum MarkerStyle
{
    /// <summary>No marker: the line alone.</summary>
    None,

    /// <summary>A circle.</summary>
    Circle,

    /// <summary>A square.</summary>
    Square,

    /// <summary>A diamond.</summary>
    Diamond,

    /// <summary>A triangle.</summary>
    Triangle,

    /// <summary>An x.</summary>
    X,

    /// <summary>A star.</summary>
    Star,

    /// <summary>A short horizontal dash.</summary>
    Dash,

    /// <summary>A dot.</summary>
    Dot,

    /// <summary>A plus.</summary>
    Plus,
}

/// <summary>Where a data label sits by its point. Which ones a chart type takes is in <see cref="DataLabels.Position"/>.</summary>
public enum DataLabelPosition
{
    /// <summary>Excel's choice for the chart type.</summary>
    Default,

    /// <summary>On the middle of the point.</summary>
    Center,

    /// <summary>Inside the end of a column, bar or slice.</summary>
    InsideEnd,

    /// <summary>Inside the base of a column or bar.</summary>
    InsideBase,

    /// <summary>Outside the end of a column, bar or slice.</summary>
    OutsideEnd,

    /// <summary>Above a line or scatter point.</summary>
    Above,

    /// <summary>Below a line or scatter point.</summary>
    Below,

    /// <summary>Left of a line or scatter point.</summary>
    Left,

    /// <summary>Right of a line or scatter point.</summary>
    Right,

    /// <summary>Where a pie's label fits best.</summary>
    BestFit,
}

/// <summary>Labels on the points of a chart: the value, the percentage, the category or the series' name.</summary>
public sealed record DataLabels
{
    private readonly string? numberFormat;

    /// <summary>Show the value.</summary>
    public bool ShowValue { get; init; } = true;

    /// <summary>Show the share of the whole, for pie and doughnut charts.</summary>
    public bool ShowPercent { get; init; }

    /// <summary>Show the category.</summary>
    public bool ShowCategory { get; init; }

    /// <summary>Show the series' name.</summary>
    public bool ShowSeriesName { get; init; }

    /// <summary>
    /// Where the labels sit. Column and bar take <c>Center</c>, <c>InsideEnd</c>, <c>InsideBase</c>
    /// and, unless stacked, <c>OutsideEnd</c>; line and scatter take <c>Center</c>, <c>Above</c>,
    /// <c>Below</c>, <c>Left</c> and <c>Right</c>; pie takes <c>Center</c>, <c>InsideEnd</c>,
    /// <c>OutsideEnd</c> and <c>BestFit</c>; area and doughnut take <c>Default</c> only. Excel
    /// repairs a file with any other, so AddChart throws an ArgumentException for it.
    /// </summary>
    public DataLabelPosition Position { get; init; }

    /// <summary>How the labels look, or null for the chart's font.</summary>
    public TextFont? Font { get; init; }

    /// <summary>The number format of the values, such as <c>0%</c>, or null for the cells' own.</summary>
    public string? NumberFormat
    {
        get => numberFormat;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("A number format holds 1 to 255 characters.", nameof(NumberFormat));
            }

            numberFormat = value;
        }
    }
}

/// <summary>How an axis of a chart looks: shown or hidden, gridlines, its range, font and number format.</summary>
public sealed record ChartAxis
{
    private readonly double? minimum;
    private readonly double? maximum;
    private readonly double? majorUnit;
    private readonly string? numberFormat;

    /// <summary>Show the axis and its labels. Off, the chart keeps its scale but draws no axis, as in a dashboard.</summary>
    public bool Visible { get; init; } = true;

    /// <summary>Draw the major gridlines, or null for Excel's default: on for the value axis, off for the category axis.</summary>
    public bool? Gridlines { get; init; }

    /// <summary>The gridlines' colour, width and dash, or null for Excel's light grey.</summary>
    public Outline? GridlineOutline { get; init; }

    /// <summary>The axis line, or null for Excel's own; <see cref="Outline.None"/> hides the line and keeps the labels.</summary>
    public Outline? Line { get; init; }

    /// <summary>The lowest value on a value axis, or null for automatic.</summary>
    public double? Minimum
    {
        get => minimum;
        init => minimum = Finite(value, nameof(Minimum));
    }

    /// <summary>The highest value on a value axis, or null for automatic.</summary>
    public double? Maximum
    {
        get => maximum;
        init => maximum = Finite(value, nameof(Maximum));
    }

    /// <summary>The step between gridlines on a value axis, or null for automatic.</summary>
    public double? MajorUnit
    {
        get => majorUnit;
        init
        {
            if (value is double unit && (double.IsFinite(unit) == false || unit <= 0))
            {
                throw new ArgumentOutOfRangeException(nameof(MajorUnit), unit, "The major unit must be above 0.");
            }

            majorUnit = value;
        }
    }

    /// <summary>The number format of the labels, such as <c># ##0 "kr"</c>, or null for the cells' own.</summary>
    public string? NumberFormat
    {
        get => numberFormat;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("A number format holds 1 to 255 characters.", nameof(NumberFormat));
            }

            numberFormat = value;
        }
    }

    /// <summary>How the labels look, or null for the chart's font.</summary>
    public TextFont? Font { get; init; }

    private static double? Finite(double? value, string parameter)
    {
        if (value is double number && double.IsFinite(number) == false)
        {
            throw new ArgumentOutOfRangeException(parameter, number, "An axis bound must be a finite number.");
        }

        return value;
    }
}

/// <summary>
/// A chart over cells in the workbook. Put it on a sheet with
/// <see cref="SheetWriter.AddChart(Chart, CellRange)"/> or <see cref="Worksheet.AddChart(Chart, CellRange)"/>.
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
    private readonly string? title;

    /// <summary>Creates a chart.</summary>
    /// <param name="type">The kind of chart.</param>
    public Chart(ChartType type)
    {
        if (Enum.IsDefined(type) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(type), type, "Not a chart type.");
        }

        Type = type;
    }

    /// <summary>The kind of chart.</summary>
    public ChartType Type { get; }

    /// <summary>The title above the chart, or null for none.</summary>
    public string? Title
    {
        get => title;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A chart title holds at most 255 characters.", nameof(Title));
            }

            if (value is not null)
            {
                DrawingValues.XmlText(value, nameof(Title));
            }

            title = value;
        }
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
    public LegendPosition Legend { get; init; } = LegendPosition.Right;

    /// <summary>
    /// Stack the series on top of each other instead of side by side. For column, bar, line and
    /// area charts.
    /// </summary>
    public bool Stacked { get; init; }

    /// <summary>The title under the category axis (the x axis of a scatter chart), or null.</summary>
    public string? CategoryAxisTitle { get; init; }

    /// <summary>The title beside the value axis, or null.</summary>
    public string? ValueAxisTitle { get; init; }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description { get; init; }

    /// <summary>The size of a doughnut's hole in percent of its diameter, 10 to 90, or null for Excel's 50.</summary>
    public int? HoleSize
    {
        get => holeSize;
        init
        {
            if (value is int size)
            {
                DrawingValues.Range(size, 10, 90, nameof(HoleSize), "A hole size");
            }

            holeSize = value;
        }
    }

    /// <summary>Where a pie's or doughnut's first slice starts, in degrees clockwise from the top, 0 to 360.</summary>
    public int FirstSliceAngle
    {
        get => firstSliceAngle;
        init => firstSliceAngle = DrawingValues.Range(value, 0, 360, nameof(FirstSliceAngle), "The first slice's angle");
    }

    /// <summary>The space between groups of columns or bars in percent of a column's width, 0 to 500, or null for Excel's 150.</summary>
    public int? GapWidth
    {
        get => gapWidth;
        init
        {
            if (value is int gap)
            {
                DrawingValues.Range(gap, 0, 500, nameof(GapWidth), "A gap width");
            }

            gapWidth = value;
        }
    }

    /// <summary>
    /// How much the columns or bars of a group overlap in percent, -100 to 100, or null for
    /// Excel's 0 (100 when stacked). 100 draws one series on top of another, as a progress bar
    /// over its background.
    /// </summary>
    public int? Overlap
    {
        get => overlap;
        init
        {
            if (value is int given)
            {
                DrawingValues.Range(given, -100, 100, nameof(Overlap), "An overlap");
            }

            overlap = value;
        }
    }

    /// <summary>
    /// Give each point its own colour, or null for Excel's default: on for pie and doughnut
    /// charts, off for the rest. Off, a doughnut's slices share the series' colour.
    /// </summary>
    public bool? VaryColors { get; init; }

    /// <summary>Round the corners of the chart's border.</summary>
    public bool RoundedCorners { get; init; }

    /// <summary>The chart area behind everything, or null for white; <see cref="Fill.None"/> lets the sheet or a shape show through.</summary>
    public Fill? Background { get; init; }

    /// <summary>The plot area inside the axes, or null for Excel's white; see-through instead when <see cref="Background"/> is set.</summary>
    public Fill? PlotBackground { get; init; }

    /// <summary>The chart's border, or null for Excel's thin grey; <see cref="Outline.None"/> removes it.</summary>
    public Outline? Border { get; init; }

    /// <summary>The font of all the chart's text, or null for Excel's; the title, legend, axes and labels can override it.</summary>
    public TextFont? Font { get; init; }

    /// <summary>The title's font, or null for the chart's.</summary>
    public TextFont? TitleFont { get; init; }

    /// <summary>The legend's font, or null for the chart's.</summary>
    public TextFont? LegendFont { get; init; }

    /// <summary>The category axis (the x axis of a scatter chart), or null for Excel's default.</summary>
    public ChartAxis? CategoryAxis { get; init; }

    /// <summary>The value axis, or null for Excel's default.</summary>
    public ChartAxis? ValueAxis { get; init; }

    /// <summary>The second value axis, on the right, of the series with <see cref="ChartSeries.SecondaryAxis"/>, or null for Excel's default.</summary>
    public ChartAxis? SecondaryValueAxis { get; init; }

    /// <summary>Labels on every series' points, or null for none. A series' own <see cref="ChartSeries.DataLabels"/> wins.</summary>
    public DataLabels? DataLabels { get; init; }

    private int? holeSize;
    private int firstSliceAngle;
    private int? gapWidth;
    private int? overlap;

    /// <summary>The type each series is drawn as.</summary>
    internal ChartType TypeOf(ChartSeries series)
    {
        if (series.Type is ChartType own)
        {
            return own;
        }

        return Type;
    }

    /// <summary>Throws ArgumentException for what Excel would repair: a mix of types it does not combine, or a label position the type does not take.</summary>
    internal void Validate()
    {
        foreach (ChartAxis? axis in new[] { CategoryAxis, ValueAxis, SecondaryValueAxis })
        {
            if (axis is not null && axis.Minimum is double minimum && axis.Maximum is double maximum && minimum >= maximum)
            {
                throw new ArgumentException($"An axis' minimum ({minimum}) must be below its maximum ({maximum}).", nameof(ValueAxis));
            }
        }

        bool mixed = Series.Any(series => TypeOf(series) != Type || series.SecondaryAxis == true);

        if (mixed == true)
        {
            foreach (ChartSeries series in Series)
            {
                if ((TypeOf(series) is ChartType.Column or ChartType.Line or ChartType.Area) == false)
                {
                    throw new ArgumentException($"A combination chart mixes column, line and area series; {TypeOf(series)} cannot be one of them.", nameof(Series));
                }
            }

            if (Series.All(series => series.SecondaryAxis == true) == true)
            {
                throw new ArgumentException("At least one series must stay on the primary axis.", nameof(Series));
            }
        }

        foreach (ChartSeries series in Series)
        {
            DataLabels? labels = series.DataLabels ?? DataLabels;

            if (labels is null || labels.Position == DataLabelPosition.Default)
            {
                continue;
            }

            ChartType type = TypeOf(series);
            bool stacked = Stacked == true && type == Type;
            bool allowed = type switch
            {
                ChartType.Column or ChartType.Bar => labels.Position is DataLabelPosition.Center or DataLabelPosition.InsideEnd or DataLabelPosition.InsideBase
                    || (labels.Position == DataLabelPosition.OutsideEnd && stacked == false),
                ChartType.Line or ChartType.Scatter => labels.Position is DataLabelPosition.Center or DataLabelPosition.Above or DataLabelPosition.Below or DataLabelPosition.Left or DataLabelPosition.Right,
                ChartType.Pie => labels.Position is DataLabelPosition.Center or DataLabelPosition.InsideEnd or DataLabelPosition.OutsideEnd or DataLabelPosition.BestFit,
                _ => false,
            };

            if (allowed == false)
            {
                throw new ArgumentException($"A {type} series does not take data labels at {labels.Position}; Excel would repair the file. See DataLabels.Position.", nameof(DataLabels));
            }
        }
    }
}
