namespace Avig.Kalkyl;

/// <summary>How a pivot table combines the values of a field, as in Excel's "Summarize Values By".</summary>
public enum PivotAggregation
{
    /// <summary>The sum of the numbers.</summary>
    Sum,

    /// <summary>The number of values that are not blank, text included.</summary>
    Count,

    /// <summary>The mean of the numbers.</summary>
    Average,

    /// <summary>The smallest number.</summary>
    Min,

    /// <summary>The largest number.</summary>
    Max,
}

/// <summary>How a pivot table lays out its row fields, as on Excel's Design tab, Report Layout.</summary>
public enum PivotLayout
{
    /// <summary>All row fields in one column, the inner ones indented; Excel's default.</summary>
    Compact,

    /// <summary>One column per row field, each group's subtotal on its own first row.</summary>
    Outline,

    /// <summary>One column per row field, the first item of a group on the group's row, and the subtotal on a row of its own below.</summary>
    Tabular,
}

/// <summary>
/// A field in the Values area of a pivot table: what to combine and how. The name in the table
/// is Excel's ("Sum of Sales") unless <see cref="Name"/> is set.
/// </summary>
public sealed record PivotValue
{
    private readonly string? name;

    /// <summary>Creates a value field.</summary>
    /// <param name="field">The heading of the source column.</param>
    /// <param name="aggregation">How the values are combined.</param>
    public PivotValue(string field, PivotAggregation aggregation = PivotAggregation.Sum)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(field);

        if (Enum.IsDefined(aggregation) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(aggregation), aggregation, "Not an aggregation.");
        }

        Field = field;
        Aggregation = aggregation;
    }

    /// <summary>The heading of the source column.</summary>
    public string Field { get; }

    /// <summary>How the values are combined.</summary>
    public PivotAggregation Aggregation { get; }

    /// <summary>
    /// The name in the table and in GETPIVOTDATA, such as "Sales 2026". Null for Excel's own:
    /// "Sum of Sales", "Count of Sales", "Average of Sales", "Min of Sales" or "Max of Sales".
    /// </summary>
    public string? Name
    {
        get => name;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 255))
            {
                throw new ArgumentException("The name of a value field holds 1 to 255 characters.", nameof(Name));
            }

            name = value;
        }
    }

    /// <summary>The number format of the results, such as <c>#,##0</c>, or null for General.</summary>
    public string? NumberFormat { get; init; }

    /// <summary>The name in the table: <see cref="Name"/>, or Excel's.</summary>
    internal string DisplayName
    {
        get
        {
            if (name is not null)
            {
                return name;
            }

            string verb = Aggregation switch
            {
                PivotAggregation.Count => "Count",
                PivotAggregation.Average => "Average",
                PivotAggregation.Min => "Min",
                PivotAggregation.Max => "Max",
                _ => "Sum",
            };

            return verb + " of " + Field;
        }
    }
}

/// <summary>
/// A pivot table over a block of cells with headings in its first row, or over an Excel table.
/// Put it on a sheet with <see cref="Worksheet.AddPivotTable(PivotTable, CellAddress)"/>. The
/// package calculates the table and stores it in the sheet's cells and in the pivot cache, so
/// Excel shows it right away, and marks the cache to be refreshed when the file is opened.
/// </summary>
/// <example>
/// <code>
/// PivotTable sales = new("SalesByMonth", "Data!A1:F49")
/// {
///     Rows = ["Month"],
///     Values = [new PivotValue("Sales") { NumberFormat = "#,##0" }],
/// };
/// report.AddPivotTable(sales, "A3");
/// </code>
/// </example>
public sealed record PivotTable
{
    private readonly IReadOnlyList<string> rows = [];
    private readonly IReadOnlyList<string> columns = [];
    private readonly IReadOnlyList<string> filters = [];
    private readonly IReadOnlyList<PivotValue> values = [];
    private readonly string style = "PivotStyleLight16";

    /// <summary>Creates a pivot table.</summary>
    /// <param name="name">The name, unique on its sheet; GETPIVOTDATA and slicers know the table by it.</param>
    /// <param name="source">
    /// The source: a range with the headings in its first row, such as <c>Data!A1:F49</c> or
    /// <c>'Sales 2026'!A1:F49</c> (a range without a sheet is on the pivot table's own sheet),
    /// or the name of an Excel table in an opened workbook. Rows below the sheet's last used row
    /// are left out.
    /// </param>
    public PivotTable(string name, string source)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(source);

        if (name.Length > 255)
        {
            throw new ArgumentException("A pivot table name holds at most 255 characters.", nameof(name));
        }

        Name = name;
        Source = source.Trim();
    }

    /// <summary>The name, unique on its sheet.</summary>
    public string Name { get; }

    /// <summary>The source range or table, as given.</summary>
    public string Source { get; }

    /// <summary>The fields down the side, outermost first.</summary>
    public IReadOnlyList<string> Rows
    {
        get => rows;
        init => rows = Fields(value, nameof(Rows));
    }

    /// <summary>The fields across the top, outermost first.</summary>
    public IReadOnlyList<string> Columns
    {
        get => columns;
        init => columns = Fields(value, nameof(Columns));
    }

    /// <summary>The fields in the Filters area above the table. They show (All); filter with a <see cref="Slicer"/>.</summary>
    public IReadOnlyList<string> Filters
    {
        get => filters;
        init => filters = Fields(value, nameof(Filters));
    }

    /// <summary>The values, at least one. With more than one, they go across the top after <see cref="Columns"/>.</summary>
    public IReadOnlyList<PivotValue> Values
    {
        get => values;
        init
        {
            ArgumentNullException.ThrowIfNull(value);

            if (value.Any(item => item is null) == true)
            {
                throw new ArgumentException("A value field is null.", nameof(Values));
            }

            values = [.. value];
        }
    }

    /// <summary>How the row fields are laid out.</summary>
    public PivotLayout Layout { get; init; } = PivotLayout.Compact;

    /// <summary>Show the Grand Total row at the bottom, when there are row fields.</summary>
    public bool GrandTotalRow { get; init; } = true;

    /// <summary>Show the Grand Total column at the right, when there are column fields.</summary>
    public bool GrandTotalColumn { get; init; } = true;

    /// <summary>A built-in pivot table style, such as <c>PivotStyleLight16</c> (Excel's default), <c>PivotStyleMedium9</c> or <c>PivotStyleDark2</c>.</summary>
    public string Style
    {
        get => style;
        init
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(value, nameof(Style));
            style = value;
        }
    }

    private static IReadOnlyList<string> Fields(IReadOnlyList<string> value, string parameter)
    {
        ArgumentNullException.ThrowIfNull(value, parameter);

        if (value.Any(string.IsNullOrWhiteSpace) == true)
        {
            throw new ArgumentException("A field name is empty.", parameter);
        }

        return [.. value];
    }
}

/// <summary>
/// A slicer: buttons that filter one field of one or more pivot tables on the same source. Put it
/// on a sheet with <see cref="Worksheet.AddSlicer(Slicer, CellRange)"/>.
/// </summary>
/// <example>
/// <code>
/// Slicer years = new("Year", ["SalesByMonth", "SalesByRegion"]) { Columns = 4, Selected = ["2026"] };
/// dashboard.AddSlicer(years, "B2:H4");
/// </code>
/// </example>
public sealed record Slicer
{
    private readonly int columns = 1;
    private readonly double? itemHeight;
    private readonly IReadOnlyList<string>? selected;
    private readonly string style = "SlicerStyleLight1";

    /// <summary>Creates a slicer.</summary>
    /// <param name="field">The heading of the source column to filter.</param>
    /// <param name="pivotTables">The names of the pivot tables it filters, added with <see cref="Worksheet.AddPivotTable(PivotTable, CellAddress)"/> and all on the same source.</param>
    public Slicer(string field, IReadOnlyList<string> pivotTables)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(field);
        ArgumentNullException.ThrowIfNull(pivotTables);

        if (pivotTables.Count == 0 || pivotTables.Any(string.IsNullOrWhiteSpace) == true)
        {
            throw new ArgumentException("A slicer filters at least one pivot table, named by its name.", nameof(pivotTables));
        }

        Field = field;
        PivotTables = [.. pivotTables];
    }

    /// <summary>The heading of the source column it filters.</summary>
    public string Field { get; }

    /// <summary>The names of the pivot tables it filters.</summary>
    public IReadOnlyList<string> PivotTables { get; }

    /// <summary>The name of the slicer shape, unique in the workbook; null for the field name (with a number added when taken).</summary>
    public string? Name { get; init; }

    /// <summary>The caption in the slicer's header; null for the field name.</summary>
    public string? Caption { get; init; }

    /// <summary>Show the header with the caption and the clear-filter button.</summary>
    public bool ShowCaption { get; init; } = true;

    /// <summary>A built-in slicer style: <c>SlicerStyleLight1</c> to <c>SlicerStyleLight6</c>, <c>SlicerStyleOther1</c>, <c>SlicerStyleOther2</c>, <c>SlicerStyleDark1</c> to <c>SlicerStyleDark6</c>.</summary>
    public string Style
    {
        get => style;
        init
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(value, nameof(Style));
            style = value;
        }
    }

    /// <summary>The number of columns of buttons, 1 to 20 000.</summary>
    public int Columns
    {
        get => columns;
        init
        {
            ArgumentOutOfRangeException.ThrowIfLessThan(value, 1, nameof(Columns));
            ArgumentOutOfRangeException.ThrowIfGreaterThan(value, 20_000, nameof(Columns));
            columns = value;
        }
    }

    /// <summary>The height of a button in points, or null for Excel's (19 points).</summary>
    public double? ItemHeight
    {
        get => itemHeight;
        init
        {
            if (value is double height && (double.IsFinite(height) == false || height <= 0 || height > 2000))
            {
                throw new ArgumentOutOfRangeException(nameof(ItemHeight), value, "The height of a button goes from just above 0 to 2000 points.");
            }

            itemHeight = value;
        }
    }

    /// <summary>
    /// The items selected, as they show on the buttons (2026, Jan); null or not set for all of
    /// them. The pivot tables show only the selected items.
    /// </summary>
    public IReadOnlyList<string>? Selected
    {
        get => selected;
        init
        {
            if (value is not null && value.Count == 0)
            {
                throw new ArgumentException("A slicer keeps at least one item selected; leave Selected null for all of them.", nameof(Selected));
            }

            if (value is not null && value.Any(item => item is null) == true)
            {
                throw new ArgumentException("A selected item is null.", nameof(Selected));
            }

            if (value is null)
            {
                selected = null;

                return;
            }

            selected = [.. value];
        }
    }
}

/// <summary>A pivot table in the workbook: one the file had when it was opened, or one added since.</summary>
/// <param name="Name">The pivot table's name.</param>
/// <param name="Sheet">The sheet it is on.</param>
/// <param name="Location">The cells of the table itself, without the Filters area above it.</param>
/// <param name="CacheId">The pivot cache it reads, as in <see cref="PivotCacheInfo.CacheId"/>.</param>
public sealed record PivotTableInfo(string Name, string Sheet, CellRange Location, int CacheId);

/// <summary>
/// A pivot cache: the copy of the source data that pivot tables read. Excel shows what the
/// cache holds until it is refreshed.
/// </summary>
/// <param name="CacheId">The cache's number in the workbook.</param>
/// <param name="Source">Where the data comes from: <c>'Data'!A1:F49</c>, a table or range name, or <c>external</c> for a data connection.</param>
/// <param name="Fields">The names of the cache's fields, the source's headings.</param>
/// <param name="RecordCount">The number of rows of data stored.</param>
/// <param name="RefreshOnLoad">True when Excel refreshes the cache from the source when it opens the file.</param>
public sealed record PivotCacheInfo(int CacheId, string Source, IReadOnlyList<string> Fields, int RecordCount, bool RefreshOnLoad);
