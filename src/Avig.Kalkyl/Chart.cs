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

        Name = name;
    }

    /// <summary>The name in the legend.</summary>
    public string? Name { get; }

    /// <summary>The colour, as <c>#RRGGBB</c>, or null for Excel's own.</summary>
    public string? Color { get; }

    internal SheetReference Values { get; }

    internal SheetReference? Categories { get; }
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
}
