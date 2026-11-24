using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// A sheet in a <see cref="Workbook"/>, held in memory: read and change any cell in any order.
/// Positions are 0-based; <c>sheet[1, 2]</c> and <c>sheet["C2"]</c> are the same cell.
/// </summary>
public sealed class Worksheet
{
    private readonly SortedDictionary<int, RowEntry> rows = [];
    private List<ColumnRange> columns;
    private int frozenRows;
    private int frozenColumns;
    private bool autoFilter;
    private SheetVisibility visibility;

    internal Worksheet(Workbook workbook, string name, int index, SheetSource? source)
    {
        Workbook = workbook;
        Name = name;
        Index = index;
        Source = source;
        columns = [];

        if (source is not null)
        {
            columns = source.Columns.ToList();
            frozenRows = source.FrozenRows;
            frozenColumns = source.FrozenColumns;
            autoFilter = source.AutoFilter;
            visibility = source.Visibility;
        }
    }

    /// <summary>The workbook the sheet belongs to.</summary>
    public Workbook Workbook { get; }

    /// <summary>The name on the tab.</summary>
    public string Name { get; }

    /// <summary>The position among the worksheets, 0-based.</summary>
    public int Index { get; }

    /// <summary>Whether the tab is shown.</summary>
    public SheetVisibility Visibility
    {
        get => visibility;
        set
        {
            visibility = value;
            Workbook.MarkChanged();
        }
    }

    /// <summary>One past the last row that has a cell or row settings. 0 for an empty sheet.</summary>
    public int RowCount => rowCount;

    /// <summary>One past the last column that has a cell. 0 for an empty sheet.</summary>
    public int ColumnCount
    {
        get
        {
            if (columnCount < 0)
            {
                columnCount = CountColumns();
            }

            return columnCount;
        }
    }

    /// <summary>
    /// Kept up to date as rows are added instead of found on demand: the formula engine asks
    /// for the used area on every range it reads, and walking a SortedDictionary for it made a
    /// SUM(G2:I2) cost as much as the whole sheet (minutes for 200 000 formulas on 30 000 rows).
    /// Rows are never removed, so the count only grows.
    /// </summary>
    private int rowCount;

    /// <summary>As <see cref="rowCount"/>, but -1 when a cell in the last column was cleared and the count must be found again.</summary>
    private int columnCount;

    private int CountColumns()
    {
        int count = 0;

        foreach (RowEntry row in rows.Values)
        {
            if (row.Cells.Count > 0)
            {
                count = Math.Max(count, row.Cells.Keys.Last() + 1);
            }
        }

        return count;
    }

    private void CellAdded(int column)
    {
        if (columnCount >= 0 && column >= columnCount)
        {
            columnCount = column + 1;
        }
    }

    /// <summary>The number of rows frozen at the top.</summary>
    public int FrozenRows => frozenRows;

    /// <summary>The number of columns frozen at the left.</summary>
    public int FrozenColumns => frozenColumns;

    /// <summary>Filter buttons on the first row, covering every row and column the sheet has when it is saved.</summary>
    public bool AutoFilter
    {
        get => autoFilter;
        set
        {
            if (autoFilter != value)
            {
                autoFilter = value;
                FilterChanged = true;
                Workbook.MarkChanged();
            }
        }
    }

    /// <summary>The rows that have cells, from the top, as snapshots.</summary>
    public IEnumerable<Row> Rows
    {
        get
        {
            foreach ((int index, RowEntry row) in rows)
            {
                yield return Snapshot(index, row);
            }
        }
    }

    internal SheetSource? Source { get; }

    internal bool ViewChanged { get; private set; }

    internal bool ColumnsChanged { get; private set; }

    internal bool FilterChanged { get; private set; }

    internal IReadOnlyList<ColumnRange> Columns => columns;

    internal SortedDictionary<int, RowEntry> RowEntries => rows;

    /// <summary>The cell at a position. Setting it replaces value, formula and style; use <see cref="SetValue"/> to keep the style.</summary>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    /// <returns>The cell, or an empty cell.</returns>
    public Cell this[int row, int column]
    {
        get
        {
            Check(row, column);

            if (rows.TryGetValue(row, out RowEntry? entry) == true && entry.Cells.TryGetValue(column, out CellEntry? cell) == true)
            {
                return cell.Cell;
            }

            return Cell.Empty;
        }

        set
        {
            Check(row, column);
            Put(row, column, value);
        }
    }

    /// <summary>The cell at an address in A1 notation, such as <c>C2</c>.</summary>
    /// <param name="reference">The address.</param>
    /// <returns>The cell, or an empty cell.</returns>
    public Cell this[string reference]
    {
        get
        {
            CellAddress address = ParseAddress(reference);

            return this[address.Row, address.Column];
        }

        set
        {
            CellAddress address = ParseAddress(reference);
            this[address.Row, address.Column] = value;
        }
    }

    /// <summary>
    /// Changes a cell's value and keeps its style. A date in a cell without a date format gets
    /// <c>yyyy-mm-dd</c>, so it shows as a date. A formula in the cell is removed.
    /// </summary>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="value">The value.</param>
    public void SetValue(int row, int column, CellValue value)
    {
        Style style = this[row, column].Style;

        if (value.Kind == CellKind.Date && (style.NumberFormat is null || NumberFormats.IsDateFormat(style.NumberFormat) == false))
        {
            style = style with { NumberFormat = CellXml.DefaultFormat(value) };
        }

        Put(row, column, Cell.Of(value).WithStyle(style));
    }

    /// <summary>Puts a formula in a cell and keeps its style. The cached value is cleared until the workbook is recalculated.</summary>
    /// <remarks>Never pass text from a user; see <see cref="Cell.FromFormula(string)"/>.</remarks>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="formula">The formula, with or without a leading <c>=</c>.</param>
    public void SetFormula(int row, int column, string formula)
    {
        Style style = this[row, column].Style;
        Put(row, column, Cell.FromFormula(formula).WithStyle(style));
    }

    /// <summary>Removes a cell's value, formula and style.</summary>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    public void Clear(int row, int column)
    {
        Check(row, column);

        if (rows.TryGetValue(row, out RowEntry? entry) == true && entry.Cells.Remove(column) == true)
        {
            if (column + 1 >= columnCount)
            {
                columnCount = -1;
            }

            Workbook.MarkChanged(this, row, column);
        }
    }

    /// <summary>Sets the width of a column, in characters of the default font.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="width">The width, 0 to 255.</param>
    public void SetColumnWidth(int column, double width)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, CellAddress.MaxColumns);
        ArgumentOutOfRangeException.ThrowIfNegative(width);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(width, 255);
        List<ColumnRange> updated = [];
        bool placed = false;

        foreach (ColumnRange range in columns)
        {
            if (column < range.First || column > range.Last)
            {
                updated.Add(range);

                continue;
            }

            // Split the run so that only this column gets the new width; the rest keeps its settings.
            if (range.First < column)
            {
                updated.Add(range with { Last = column - 1 });
            }

            updated.Add(range with { First = column, Last = column, Width = width });

            if (range.Last > column)
            {
                updated.Add(range with { First = column + 1 });
            }

            placed = true;
        }

        if (placed == false)
        {
            updated.Add(new ColumnRange(column, column, width, []));
        }

        columns = updated.OrderBy(range => range.First).ToList();
        ColumnsChanged = true;
        Workbook.MarkChanged();
    }

    /// <summary>The width of a column, or null when it has the default width.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <returns>The width.</returns>
    public double? GetColumnWidth(int column)
    {
        double? width = null;

        foreach (ColumnRange range in columns)
        {
            if (column >= range.First && column <= range.Last)
            {
                width = range.Width;
            }
        }

        return width;
    }

    /// <summary>Keeps the top rows and left columns in view while scrolling. 0 and 0 unfreezes.</summary>
    /// <param name="rows">The number of rows to freeze.</param>
    /// <param name="columns">The number of columns to freeze.</param>
    public void FreezePanes(int rows, int columns = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rows);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rows, CellAddress.MaxRows);
        ArgumentOutOfRangeException.ThrowIfNegative(columns);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columns, CellAddress.MaxColumns);
        frozenRows = rows;
        frozenColumns = columns;
        ViewChanged = true;
        Workbook.MarkChanged();
    }

    /// <summary>Stores a cell read from the file, with its original XML so that it is written back unchanged.</summary>
    internal void Load(int row, int column, CellEntry entry)
    {
        GetRow(row).Cells[column] = entry;
        CellAdded(column);
    }

    internal RowEntry GetRow(int row)
    {
        if (rows.TryGetValue(row, out RowEntry? entry) == false)
        {
            entry = new RowEntry([]);
            rows[row] = entry;
            rowCount = Math.Max(rowCount, row + 1);
        }

        return entry;
    }

    /// <summary>Sets a calculated value on a formula cell, keeping its formula and style.</summary>
    internal void SetCalculated(int row, int column, CellValue value)
    {
        if (rows.TryGetValue(row, out RowEntry? entry) == false || entry.Cells.TryGetValue(column, out CellEntry? cell) == false || cell.Cell.Formula is null)
        {
            return;
        }

        if (cell.Cell.Value.Equals(value) == true)
        {
            return;
        }

        entry.Cells[column] = new CellEntry(Cell.Create(value, cell.Cell.Formula, cell.Cell.Style), null, cell.StyleIndex, cell.StyleFromFile);
    }

    private void Put(int row, int column, Cell value)
    {
        CellEntry? old = null;

        if (rows.TryGetValue(row, out RowEntry? existing) == true)
        {
            existing.Cells.TryGetValue(column, out old);
        }

        if (old is not null && old.Cell.Equals(value) == true)
        {
            return;
        }

        if (value.IsEmpty == true)
        {
            Clear(row, column);

            return;
        }

        // A cell that keeps the style it was read with keeps its index in the file, which also
        // keeps what Style does not describe, such as borders.
        bool keepStyle = old is not null && old.StyleFromFile == true && old.Cell.Style.Equals(value.Style) == true;
        int styleIndex = 0;

        if (keepStyle == true)
        {
            styleIndex = old!.StyleIndex;
        }

        GetRow(row).Cells[column] = new CellEntry(value, null, styleIndex, keepStyle);
        CellAdded(column);
        Workbook.MarkChanged(this, row, column);
    }

    private static Row Snapshot(int index, RowEntry row)
    {
        int width = 0;

        if (row.Cells.Count > 0)
        {
            width = row.Cells.Keys.Last() + 1;
        }

        Cell[] cells = new Cell[width];

        foreach ((int column, CellEntry cell) in row.Cells)
        {
            cells[column] = cell.Cell;
        }

        return new Row(index, cells);
    }

    private static CellAddress ParseAddress(string reference)
    {
        if (CellAddress.TryParse(reference, out CellAddress address) == false)
        {
            throw new ArgumentException($"'{reference}' is not a cell address. Expected a reference such as A1.", nameof(reference));
        }

        return address;
    }

    private static void Check(int row, int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(row);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(row, CellAddress.MaxRows);
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, CellAddress.MaxColumns);
    }
}

/// <summary>
/// A cell in the document model. <see cref="RawXml"/> is the cell's &lt;c&gt; element exactly as it
/// was in the file, kept while the cell is unchanged so that it is written back byte for byte.
/// </summary>
internal sealed record CellEntry(Cell Cell, string? RawXml, int StyleIndex, bool StyleFromFile);

/// <summary>A row: its cells, and the row's own attributes from the file (height, hidden, style ...).</summary>
internal sealed record RowEntry(SortedDictionary<int, CellEntry> Cells)
{
    public IReadOnlyList<(string Prefix, string LocalName, string Namespace, string Value)> Attributes { get; set; } = [];
}

/// <summary>Where an opened sheet came from, and the settings read from it.</summary>
internal sealed record SheetSource(
    string Path,
    string RelationshipId,
    int SheetId,
    string Namespace,
    IReadOnlyList<ColumnRange> Columns,
    int FrozenRows,
    int FrozenColumns,
    bool AutoFilter,
    SheetVisibility Visibility,
    bool TabSelected);
