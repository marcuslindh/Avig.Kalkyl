using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// A sheet in a <see cref="Workbook"/>, held in memory: read and change any cell in any order.
/// Positions are 0-based; <c>sheet[1, 2]</c> and <c>sheet["C2"]</c> are the same cell.
/// </summary>
public sealed class Worksheet
{
    private readonly RowTable rows = new();
    private readonly SheetExtras extras = new();
    private List<ColumnRange> columns;
    private int frozenRows;
    private int frozenColumns;
    private bool autoFilter;
    private SheetVisibility visibility;
    private SheetView view = SheetView.Default;

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
            view = source.View;
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
    /// for the used area on every range it reads, and walking the rows for it made a
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
                count = Math.Max(count, row.Cells.LastColumn + 1);
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

    internal RowTable RowEntries => rows;

    /// <summary>The stored cell at a position, without the bounds check of the indexer: the formula engine's way in.</summary>
    internal bool TryGetEntry(int row, int column, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out CellEntry? entry)
    {
        if (rows.TryGetValue(row, out RowEntry? found) == true)
        {
            return found.Cells.TryGetValue(column, out entry);
        }

        entry = null;

        return false;
    }

    /// <summary>The cell at a position. Setting it replaces value, formula and style; use <see cref="SetValue"/> to keep the style.</summary>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    /// <returns>The cell, or an empty cell.</returns>
    public Cell this[int row, int column]
    {
        get
        {
            Check(row, column);

            if (TryGetEntry(row, column, out CellEntry? cell) == true)
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

    /// <summary>
    /// Adds a conditional format to a range. Rules the file already has are kept as they are, and
    /// win over the new ones where both apply.
    /// </summary>
    /// <param name="range">The cells, for example <c>B2:B100</c>.</param>
    /// <param name="format">The rule.</param>
    public void AddConditionalFormat(CellRange range, ConditionalFormat format)
    {
        extras.AddConditionalFormat(range, format);
        Workbook.MarkChanged();
    }

    /// <summary>Adds a conditional format to a range given in A1 notation, such as <c>B2:B100</c>.</summary>
    /// <param name="range">The cells.</param>
    /// <param name="format">The rule.</param>
    public void AddConditionalFormat(string range, ConditionalFormat format)
    {
        AddConditionalFormat(SheetWriter.ParseRange(range), format);
    }

    /// <summary>
    /// Adds a chart over a block of cells. Charts and pictures the file already has are kept.
    /// The series may refer to any sheet in the workbook.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <param name="area">The cells the chart covers, for example <c>E2:L20</c>.</param>
    /// <returns>
    /// Success; NotFound when a series refers to a sheet the workbook does not have; Invalid
    /// when the file is Strict OOXML, which the package does not add drawings to.
    /// </returns>
    public Result AddChart(Chart chart, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(chart);

        if (CheckDrawable().IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result.Failure(error, status);
        }

        Result added = extras.AddChart(chart, area, Name, name => Workbook.Sheet(name).IsSuccess(out _) == true);

        if (added.IsSuccess(out _, out _) == true)
        {
            Workbook.MarkChanged();
        }

        return added;
    }

    /// <summary>Adds a chart over a block of cells given in A1 notation, such as <c>E2:L20</c>.</summary>
    /// <param name="chart">The chart.</param>
    /// <param name="area">The cells the chart covers.</param>
    /// <returns>As <see cref="AddChart(Chart, CellRange)"/>.</returns>
    public Result AddChart(Chart chart, string area)
    {
        return AddChart(chart, SheetWriter.ParseRange(area));
    }

    /// <summary>Puts a picture with its top-left corner at a cell, at its own size times <paramref name="scale"/>.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="topLeft">The cell.</param>
    /// <param name="scale">1 for the picture's own size (96 pixels to the inch), 0.5 for half.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddPicture(Picture picture, CellAddress topLeft, double scale = 1)
    {
        ArgumentNullException.ThrowIfNull(picture);
        Result drawable = CheckDrawable();

        if (drawable.IsSuccess(out _, out _) == true)
        {
            extras.AddPicture(picture, topLeft, scale);
            Workbook.MarkChanged();
        }

        return drawable;
    }

    /// <summary>Puts a picture over a block of cells, stretched to fill it.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="area">The cells.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddPicture(Picture picture, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(picture);
        Result drawable = CheckDrawable();

        if (drawable.IsSuccess(out _, out _) == true)
        {
            extras.AddPicture(picture, area);
            Workbook.MarkChanged();
        }

        return drawable;
    }

    /// <summary>
    /// Puts a picture at a cell such as <c>B2</c>, at its own size, or over a range such as
    /// <c>B2:F12</c>, stretched to fill it.
    /// </summary>
    /// <param name="picture">The picture.</param>
    /// <param name="reference">The cell or the range.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddPicture(Picture picture, string reference)
    {
        CellRange range = SheetWriter.ParseRange(reference);

        if (reference.Contains(':', StringComparison.Ordinal) == true)
        {
            return AddPicture(picture, range);
        }

        return AddPicture(picture, range.First);
    }

    /// <summary>
    /// How the sheet looks in Excel: gridlines, headers, zoom, tab colour, selected cell and
    /// direction. An opened file's view is read into it; setting it rewrites only what changed.
    /// </summary>
    public SheetView View
    {
        get => view;
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            if (string.Equals(view.TabColor, value.TabColor, StringComparison.Ordinal) == false)
            {
                TabColorChanged = true;
            }

            if ((view with { TabColor = null }).Equals(value with { TabColor = null }) == false)
            {
                ViewChanged = true;
            }

            view = value;
            Workbook.MarkChanged();
        }
    }

    /// <summary>Adds a chart at a place on the sheet, which stays put when rows and columns change size.</summary>
    /// <param name="chart">The chart.</param>
    /// <param name="bounds">The place, for example <c>DrawingBounds.FromPoints(20, 20, 360, 220)</c>.</param>
    /// <returns>As <see cref="AddChart(Chart, CellRange)"/>.</returns>
    public Result AddChart(Chart chart, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(chart);

        if (CheckDrawable().IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result.Failure(error, status);
        }

        Result added = extras.AddChart(chart, DrawingAnchor.At(bounds), Name, name => Workbook.Sheet(name).IsSuccess(out _) == true);

        if (added.IsSuccess(out _, out _) == true)
        {
            Workbook.MarkChanged();
        }

        return added;
    }

    /// <summary>Puts a picture at a place on the sheet, stretched to fill it.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="bounds">The place.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddPicture(Picture picture, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(picture);

        return AddDrawing(() => extras.AddPicture(picture, bounds));
    }

    /// <summary>Adds a shape over a block of cells; it stretches when they change size.</summary>
    /// <param name="shape">The shape.</param>
    /// <param name="area">The cells.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddShape(Shape shape, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return AddDrawing(() => extras.AddShape(shape, DrawingAnchor.Over(area)));
    }

    /// <summary>Adds a shape over a block of cells given in A1 notation, such as <c>B2:D6</c>.</summary>
    /// <param name="shape">The shape.</param>
    /// <param name="area">The cells.</param>
    /// <returns>As <see cref="AddShape(Shape, CellRange)"/>.</returns>
    public Result AddShape(Shape shape, string area)
    {
        return AddShape(shape, SheetWriter.ParseRange(area));
    }

    /// <summary>Adds a shape at a place on the sheet, which stays put when rows and columns change size.</summary>
    /// <param name="shape">The shape.</param>
    /// <param name="bounds">The place, for example <c>DrawingBounds.FromPoints(20, 20, 160, 90)</c>.</param>
    /// <returns>As <see cref="AddShape(Shape, CellRange)"/>.</returns>
    public Result AddShape(Shape shape, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return AddDrawing(() => extras.AddShape(shape, DrawingAnchor.At(bounds)));
    }

    /// <summary>Adds a group of shapes and pictures over a block of cells, scaled to fill it.</summary>
    /// <param name="group">The group.</param>
    /// <param name="area">The cells.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add drawings to.</returns>
    public Result AddGroup(ShapeGroup group, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(group);

        return AddDrawing(() => extras.AddGroup(group, DrawingAnchor.Over(area)));
    }

    /// <summary>Adds a group of shapes and pictures over a block of cells given in A1 notation.</summary>
    /// <param name="group">The group.</param>
    /// <param name="area">The cells.</param>
    /// <returns>As <see cref="AddGroup(ShapeGroup, CellRange)"/>.</returns>
    public Result AddGroup(ShapeGroup group, string area)
    {
        return AddGroup(group, SheetWriter.ParseRange(area));
    }

    /// <summary>Adds a group of shapes and pictures at a place on the sheet, scaled to fill it.</summary>
    /// <param name="group">The group.</param>
    /// <param name="bounds">The place.</param>
    /// <returns>As <see cref="AddGroup(ShapeGroup, CellRange)"/>.</returns>
    public Result AddGroup(ShapeGroup group, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(group);

        return AddDrawing(() => extras.AddGroup(group, DrawingAnchor.At(bounds)));
    }

    /// <summary>
    /// Makes a cell a link to another cell in the workbook or to a web address. The cell keeps its
    /// value and style. The file's own links are kept, except one on the same cell, which this
    /// replaces.
    /// </summary>
    /// <param name="cell">The cell.</param>
    /// <param name="link">Where a click goes.</param>
    /// <returns>Success, or Invalid when the file is Strict OOXML, which the package does not add links to.</returns>
    public Result AddHyperlink(CellAddress cell, Hyperlink link)
    {
        ArgumentNullException.ThrowIfNull(link);

        return AddDrawing(() => extras.AddHyperlink(cell, link));
    }

    /// <summary>Makes a cell given in A1 notation, such as <c>B2</c>, a link.</summary>
    /// <param name="cell">The cell.</param>
    /// <param name="link">Where a click goes.</param>
    /// <returns>As <see cref="AddHyperlink(CellAddress, Hyperlink)"/>.</returns>
    public Result AddHyperlink(string cell, Hyperlink link)
    {
        return AddHyperlink(ParseAddress(cell), link);
    }

    internal bool TabColorChanged { get; private set; }

    private Result AddDrawing(Action add)
    {
        Result drawable = CheckDrawable();

        if (drawable.IsSuccess(out _, out _) == true)
        {
            add();
            Workbook.MarkChanged();
        }

        return drawable;
    }

    /// <summary>
    /// Adds a pivot table with the top-left corner of its report at a cell (its Filters area,
    /// when it has filter fields, then a blank row and the table). The package calculates it
    /// from the source cells and writes it into this sheet's cells, so formulas and GETPIVOTDATA
    /// can read it, and again when the workbook is recalculated or saved. The file gets the pivot
    /// cache with the source data and refreshOnLoad, so Excel shows the table as it is and
    /// refreshes it when it opens the file. Pivot tables on the same source share a cache. A
    /// source whose heading is emptied after the table was added makes saving throw
    /// InvalidOperationException, as a pivot cache cannot be written without field names.
    /// </summary>
    /// <param name="pivotTable">The pivot table.</param>
    /// <param name="topLeft">The top-left cell of the report.</param>
    /// <returns>
    /// Success; NotFound when the source sheet, table or a field does not exist; Conflict when
    /// the sheet has a pivot table of that name, or the report would overlap another pivot table
    /// or its own source; Invalid when a heading of the source is empty, the table would not fit
    /// on the sheet, or the file is Strict OOXML.
    /// </returns>
    public Result AddPivotTable(PivotTable pivotTable, CellAddress topLeft)
    {
        ArgumentNullException.ThrowIfNull(pivotTable);

        if (CheckDrawable().IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result.Failure(error, status);
        }

        Result added = Workbook.Pivots.Add(this, pivotTable, topLeft);

        if (added.IsSuccess(out _, out _) == true)
        {
            Workbook.MarkChanged();
        }

        return added;
    }

    /// <summary>Adds a pivot table with the top-left corner of its report at a cell given in A1 notation, such as <c>A3</c>.</summary>
    /// <param name="pivotTable">The pivot table.</param>
    /// <param name="topLeft">The top-left cell of the report.</param>
    /// <returns>As <see cref="AddPivotTable(PivotTable, CellAddress)"/>.</returns>
    public Result AddPivotTable(PivotTable pivotTable, string topLeft)
    {
        return AddPivotTable(pivotTable, ParseAddress(topLeft));
    }

    /// <summary>
    /// Adds a slicer over a block of cells. It filters the pivot tables it names, which must have
    /// been added with <see cref="AddPivotTable(PivotTable, CellAddress)"/> and share a source;
    /// they are rendered again with the items it selects. A second slicer on the same field of
    /// the same source shares the first one's selection, as in Excel.
    /// </summary>
    /// <param name="slicer">The slicer.</param>
    /// <param name="area">The cells it covers, for example <c>B2:H4</c>.</param>
    /// <returns>
    /// Success; NotFound when a pivot table, the field or a selected item does not exist;
    /// Conflict when the name is taken or another slicer on the field selects other items;
    /// Invalid when the pivot tables have different sources or the file is Strict OOXML.
    /// </returns>
    public Result AddSlicer(Slicer slicer, CellRange area)
    {
        ArgumentNullException.ThrowIfNull(slicer);
        Result added = Workbook.Pivots.AddSlicer(this, slicer, area);

        if (added.IsSuccess(out _, out _) == true)
        {
            Workbook.MarkChanged();
        }

        return added;
    }

    /// <summary>Adds a slicer over a block of cells given in A1 notation, such as <c>B2:H4</c>.</summary>
    /// <param name="slicer">The slicer.</param>
    /// <param name="area">The cells it covers.</param>
    /// <returns>As <see cref="AddSlicer(Slicer, CellRange)"/>.</returns>
    public Result AddSlicer(Slicer slicer, string area)
    {
        return AddSlicer(slicer, SheetWriter.ParseRange(area));
    }

    internal SheetExtras Extras => extras;

    /// <summary>
    /// A Strict OOXML file has its own namespaces for drawings and charts as well. Parts written
    /// with the transitional ones would make a package that is neither, so none are added.
    /// </summary>
    internal Result CheckDrawable()
    {
        if (Source is not null && string.Equals(Source.Namespace, Ooxml.SpreadsheetNamespace, StringComparison.Ordinal) == false)
        {
            return Result.Invalid($"The sheet '{Name}' is in a Strict OOXML file. Charts and pictures can be added to transitional files only; save it as an ordinary .xlsx in Excel first.");
        }

        return Result.Success();
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
            entry = new RowEntry(new CellRow());
            rows.Add(row, entry);
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

    /// <summary>
    /// Replaces every formula that refers to another workbook or a DDE link by the value it last
    /// had, keeping the cell's style, for <see cref="Workbook.RemoveActiveContent"/>.
    /// </summary>
    internal void ReplaceExternalFormulas()
    {
        List<(RowEntry Row, int RowIndex, int Column, CellEntry Entry)> found = [];

        foreach ((int rowIndex, RowEntry row) in rows)
        {
            foreach ((int column, CellEntry cell) in row.Cells)
            {
                if (cell.Cell.Formula is string formula && ActiveContentRules.HasExternalReference(formula) == true)
                {
                    found.Add((row, rowIndex, column, cell));
                }
            }
        }

        foreach ((RowEntry row, int rowIndex, int column, CellEntry cell) in found)
        {
            row.Cells[column] = new CellEntry(Cell.Create(cell.Cell.Value, null, cell.Cell.Style), null, cell.StyleIndex, cell.StyleFromFile);
            Workbook.MarkChanged(this, rowIndex, column);
        }
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
            width = row.Cells.LastColumn + 1;
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
internal sealed record RowEntry(CellRow Cells)
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
    bool TabSelected)
{
    /// <summary>The view read from the first sheetView and the tab colour from sheetPr.</summary>
    public SheetView View { get; init; } = SheetView.Default;
}
