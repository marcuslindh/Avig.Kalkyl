using System.Globalization;

namespace Avig.Kalkyl.Internal;

/// <summary>One item of a pivot field, in the order the pivot table shows them.</summary>
/// <param name="Value">The item's value; empty for the subtotal item.</param>
/// <param name="SharedIndex">Its index in the cache field's shared items, or -1 for the subtotal item.</param>
/// <param name="Hidden">True when a filter or slicer hides it.</param>
/// <param name="IsMember">False for the subtotal item Excel lists last.</param>
internal readonly record struct PivotItem(CellValue Value, int SharedIndex, bool Hidden, bool IsMember);

/// <summary>What a row or column of a pivot table shows.</summary>
internal enum PivotLineKind
{
    /// <summary>Items, or a group's first row carrying its subtotal (compact and outline rows).</summary>
    Data,

    /// <summary>A group's subtotal on a line of its own ("North Total").</summary>
    Subtotal,

    /// <summary>The grand total.</summary>
    Grand,

    /// <summary>A blank line, or one the package does not know.</summary>
    Other,
}

/// <summary>
/// A row or column of a pivot table: the item of each field on the axis, outermost first, as
/// indexes into the fields' item lists (for the values field, the index of the value field).
/// </summary>
/// <param name="Kind">What the line shows.</param>
/// <param name="Members">The items, outermost first; shorter than the axis for subtotals and empty for grand totals.</param>
/// <param name="DataIndex">The value field the line is for, when the values field is on this axis; otherwise 0.</param>
/// <param name="Repeat">How many leading members are the same as on the line before, as the file stores it.</param>
internal sealed record PivotLine(PivotLineKind Kind, int[] Members, int DataIndex, int Repeat);

/// <summary>A value field of a pivot table: its name and the index of the cache field it combines.</summary>
internal sealed record PivotDataField(string Name, int Field);

/// <summary>
/// The shape of a pivot table as it lies in its sheet: where the cells are, and what each row and
/// column of it shows. GETPIVOTDATA finds its values here. Built by <see cref="PivotRender"/> for
/// a pivot table the package adds, and read from the file for one it had.
/// </summary>
internal sealed class PivotGrid
{
    /// <summary>The value of <c>x</c> in rowFields and colFields that stands for the values.</summary>
    public const int ValuesField = -2;

    public required string Name { get; init; }

    public required Worksheet Sheet { get; init; }

    public required CellRange Location { get; init; }

    public required int FirstHeaderRow { get; init; }

    public required int FirstDataRow { get; init; }

    public required int FirstDataColumn { get; init; }

    public required int CacheId { get; init; }

    /// <summary>The cache field names, which are the pivot field names.</summary>
    public required IReadOnlyList<string> FieldNames { get; init; }

    /// <summary>Per field, its items in the order shown; empty for a field without items.</summary>
    public required IReadOnlyList<IReadOnlyList<PivotItem>> Items { get; init; }

    public required IReadOnlyList<int> RowFields { get; init; }

    public required IReadOnlyList<int> ColumnFields { get; init; }

    public required IReadOnlyList<PivotLine> RowLines { get; init; }

    public required IReadOnlyList<PivotLine> ColumnLines { get; init; }

    public required IReadOnlyList<PivotDataField> DataFields { get; init; }
}

/// <summary>Where a pivot table's data comes from: a block of cells with headings in its first row.</summary>
/// <param name="Sheet">The sheet the cells are on.</param>
/// <param name="Range">The cells, headings included, with rows below the last used row left out.</param>
/// <param name="TableName">The Excel table the range is, when the source was given as one.</param>
internal sealed record PivotSource(Worksheet Sheet, CellRange Range, string? TableName)
{
    /// <summary>Pivot tables with the same key share a cache.</summary>
    public string Key
    {
        get
        {
            if (TableName is not null)
            {
                return "table:" + TableName.ToUpperInvariant();
            }

            return Sheet.Name.ToUpperInvariant() + "!" + Range.ToString();
        }
    }

    /// <summary>The source as <see cref="PivotCacheInfo.Source"/> shows it.</summary>
    public string Display
    {
        get
        {
            if (TableName is not null)
            {
                return TableName;
            }

            return "'" + Sheet.Name.Replace("'", "''", StringComparison.Ordinal) + "'!" + Range.ToString();
        }
    }
}

/// <summary>
/// The data a pivot cache holds, read from the source cells: the field names, each field's
/// distinct values (its shared items, in the order they first appear) and, per record, the index
/// of its value among them.
/// </summary>
internal sealed class PivotCacheData
{
    private PivotCacheData(IReadOnlyList<string> fieldNames, IReadOnlyList<List<CellValue>> shared, IReadOnlyList<int[]> records, IReadOnlyList<string?> formats)
    {
        FieldNames = fieldNames;
        SharedItems = shared;
        Records = records;
        Formats = formats;
    }

    public IReadOnlyList<string> FieldNames { get; }

    /// <summary>Per field, the distinct values in the order they first appear.</summary>
    public IReadOnlyList<List<CellValue>> SharedItems { get; }

    /// <summary>Per record, per field, the index of the value in <see cref="SharedItems"/>.</summary>
    public IReadOnlyList<int[]> Records { get; }

    /// <summary>Per field, the number format of its first value's cell, or null.</summary>
    public IReadOnlyList<string?> Formats { get; }

    public int FieldIndex(string name)
    {
        for (int index = 0; index < FieldNames.Count; index++)
        {
            if (string.Equals(FieldNames[index], name, StringComparison.OrdinalIgnoreCase) == true)
            {
                return index;
            }
        }

        return -1;
    }

    /// <summary>The value of a field in a record.</summary>
    public CellValue Value(int record, int field)
    {
        return SharedItems[field][Records[record][field]];
    }

    /// <summary>
    /// The shared indexes of a field in the order a new pivot table shows them: ascending, numbers
    /// first, then text (month and day names in calendar order, as Excel's custom lists sort
    /// them), then TRUE and FALSE, errors, and blanks last.
    /// </summary>
    public int[] SortedItems(int field)
    {
        List<CellValue> items = SharedItems[field];
        int[] order = [.. Enumerable.Range(0, items.Count)];
        Array.Sort(order, (left, right) =>
        {
            int compared = PivotItemOrder.Compare(items[left], items[right]);

            if (compared == 0)
            {
                return left.CompareTo(right);
            }

            return compared;
        });

        return order;
    }

    /// <summary>Reads the cache data from the source cells. Invalid when a heading is empty.</summary>
    public static Result<PivotCacheData> Read(PivotSource source)
    {
        Worksheet sheet = source.Sheet;
        CellRange range = source.Range;
        List<string> names = [];

        for (int column = range.First.Column; column <= range.Last.Column; column++)
        {
            CellValue heading = sheet[range.First.Row, column].Value;
            string name = heading.ToString().Trim();

            if (heading.IsEmpty == true || name.Length == 0)
            {
                return Result<PivotCacheData>.Invalid($"The source {source.Display} has no heading in cell {new CellAddress(range.First.Row, column)} (row {range.First.Row}, column {column}). Every column of a pivot table's source needs one.");
            }

            // Excel numbers a repeated heading: Sales, Sales2, Sales3.
            string unique = name;
            int number = 2;

            while (names.Exists(existing => string.Equals(existing, unique, StringComparison.OrdinalIgnoreCase) == true) == true)
            {
                unique = name + number.ToString(CultureInfo.InvariantCulture);
                number++;
            }

            names.Add(unique);
        }

        List<List<CellValue>> shared = [];
        List<Dictionary<ItemKey, int>> lookup = [];
        List<string?> formats = [];

        for (int field = 0; field < names.Count; field++)
        {
            shared.Add([]);
            lookup.Add([]);
            formats.Add(null);
        }

        List<int[]> records = [];

        for (int row = range.First.Row + 1; row <= range.Last.Row; row++)
        {
            int[] record = new int[names.Count];

            for (int field = 0; field < names.Count; field++)
            {
                Cell cell = sheet[row, range.First.Column + field];
                CellValue value = cell.Value;
                ItemKey key = ItemKey.Of(value);

                if (lookup[field].TryGetValue(key, out int index) == false)
                {
                    index = shared[field].Count;
                    shared[field].Add(value);
                    lookup[field][key] = index;
                }

                if (row == range.First.Row + 1)
                {
                    formats[field] = cell.Style.NumberFormat;
                }

                record[field] = index;
            }

            records.Add(record);
        }

        return new PivotCacheData(names, shared, records, formats);
    }

    /// <summary>
    /// Values the same for a pivot table: text without regard to case, as Excel groups it, and a
    /// date apart from the number it is stored as.
    /// </summary>
    private readonly record struct ItemKey(CellKind Kind, double Number, string Text)
    {
        public static ItemKey Of(CellValue value)
        {
            switch (value.Kind)
            {
                case CellKind.Number when value.TryGetNumber(out double number) == true:
                    return new ItemKey(CellKind.Number, number, string.Empty);
                case CellKind.Date when value.TryGetDate(out DateTime date) == true:
                    return new ItemKey(CellKind.Date, date.Ticks, string.Empty);
                case CellKind.Bool when value.TryGetBool(out bool flag) == true:
                    return new ItemKey(CellKind.Bool, Convert.ToDouble(flag), string.Empty);
                case CellKind.Empty:
                    return new ItemKey(CellKind.Empty, 0, string.Empty);
                default:
                    return new ItemKey(value.Kind, 0, value.ToString().ToUpperInvariant());
            }
        }
    }
}

/// <summary>The order Excel sorts pivot items in when it adds a field: ascending, with its built-in custom lists.</summary>
internal static class PivotItemOrder
{
    private static readonly string[][] CustomLists =
    [
        ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    ];

    public static int Compare(CellValue left, CellValue right)
    {
        int leftRank = Rank(left);
        int rightRank = Rank(right);

        if (leftRank != rightRank)
        {
            return leftRank.CompareTo(rightRank);
        }

        switch (leftRank)
        {
            case 0:
                return Number(left).CompareTo(Number(right));
            case 1:
                return CompareText(left.ToString(), right.ToString());
            case 2:
                left.TryGetBool(out bool leftFlag);
                right.TryGetBool(out bool rightFlag);

                return leftFlag.CompareTo(rightFlag);
            default:
                return string.CompareOrdinal(left.ToString(), right.ToString());
        }
    }

    private static int Rank(CellValue value)
    {
        return value.Kind switch
        {
            CellKind.Number or CellKind.Date => 0,
            CellKind.Text => 1,
            CellKind.Bool => 2,
            CellKind.Error => 3,
            _ => 4,
        };
    }

    private static double Number(CellValue value)
    {
        if (value.TryGetNumber(out double number) == true)
        {
            return number;
        }

        if (value.TryGetDate(out DateTime date) == true)
        {
            return ExcelDate.ToSerial(date);
        }

        return 0;
    }

    private static int CompareText(string left, string right)
    {
        foreach (string[] list in CustomLists)
        {
            int leftIndex = Array.FindIndex(list, item => string.Equals(item, left, StringComparison.OrdinalIgnoreCase) == true);
            int rightIndex = Array.FindIndex(list, item => string.Equals(item, right, StringComparison.OrdinalIgnoreCase) == true);

            if (leftIndex >= 0 && rightIndex >= 0)
            {
                return leftIndex.CompareTo(rightIndex);
            }
        }

        return string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
    }
}

/// <summary>A pivot table added with <see cref="Worksheet.AddPivotTable(PivotTable, CellAddress)"/>, and its last rendering.</summary>
internal sealed class CreatedPivot
{
    public CreatedPivot(PivotTable definition, Worksheet sheet, CellAddress topLeft, PivotSource source)
    {
        Definition = definition;
        Sheet = sheet;
        TopLeft = topLeft;
        Source = source;
    }

    public PivotTable Definition { get; }

    public Worksheet Sheet { get; }

    /// <summary>The top-left cell of the whole report, the Filters area included.</summary>
    public CellAddress TopLeft { get; }

    public PivotSource Source { get; }

    /// <summary>The last rendering, with the cells it covers.</summary>
    public PivotRendering? Rendering { get; set; }

    /// <summary>The cache number, given when the workbook is saved or rendered.</summary>
    public int CacheId { get; set; }
}

/// <summary>A slicer cache: one field of one pivot cache, the pivot tables it filters and which items are selected.</summary>
internal sealed class SlicerCacheModel
{
    public SlicerCacheModel(string name, string sourceKey, string field)
    {
        Name = name;
        SourceKey = sourceKey;
        Field = field;
    }

    /// <summary>The defined name, such as Slicer_Year.</summary>
    public string Name { get; }

    public string SourceKey { get; }

    public string Field { get; }

    public List<CreatedPivot> PivotTables { get; } = [];

    /// <summary>The items selected, as the buttons show them; null for all.</summary>
    public IReadOnlyList<string>? Selected { get; set; }
}

/// <summary>A slicer added with <see cref="Worksheet.AddSlicer(Slicer, CellRange)"/>.</summary>
internal sealed record CreatedSlicer(Slicer Definition, string Name, Worksheet Sheet, CellRange Area, SlicerCacheModel Cache);
