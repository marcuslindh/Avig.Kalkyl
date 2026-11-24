using System.Globalization;
using System.Xml;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// Writes one sheet, a row at a time from the top. Rows already written cannot be changed, which
/// is what keeps memory use flat no matter how large the file gets. Column widths and frozen
/// panes go at the top of the sheet in the file, so they must be set before the first row.
/// </summary>
public sealed class SheetWriter
{
    private readonly XmlWriter xml;
    private readonly Stream entryStream;
    private readonly Writing.StyleRegistry styles;
    private readonly Func<string, bool> sheetExists;
    private readonly SheetExtras extras = new();
    private readonly SortedDictionary<int, double> columnWidths = [];
    private int frozenRows;
    private int frozenColumns;
    private bool autoFilter;
    private bool started;
    private int lastColumn = -1;

    internal SheetWriter(string name, int index, Stream entryStream, Writing.StyleRegistry styles, Func<string, bool> sheetExists)
    {
        Name = name;
        Index = index;
        this.entryStream = entryStream;
        this.styles = styles;
        this.sheetExists = sheetExists;
        xml = XmlWriter.Create(entryStream, XlsxWriter.XmlSettings);
    }

    /// <summary>The sheet's name.</summary>
    public string Name { get; }

    /// <summary>The sheet's position in the workbook, 0-based.</summary>
    public int Index { get; }

    /// <summary>The number of rows written so far. The next row gets this index.</summary>
    public int RowCount { get; private set; }

    /// <summary>True once the sheet is finished, which happens when the next sheet is added or the workbook completes.</summary>
    public bool IsFinished { get; private set; }

    internal bool HasFormulas { get; private set; }

    /// <summary>The bottom-right corner of the autofilter, which always starts at A1. Set when the sheet is finished.</summary>
    internal CellAddress? AutoFilterEnd { get; private set; }

    internal SheetExtras Extras => extras;

    /// <summary>The relationship from the sheet to its drawing part, when it has charts or pictures.</summary>
    internal const string DrawingRelationshipId = "rId1";

    /// <summary>Sets the width of a column, in characters of the default font. Must come before the first row.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="width">The width, 0 to 255.</param>
    public void SetColumnWidth(int column, double width)
    {
        ThrowIfStarted(nameof(SetColumnWidth));
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, CellAddress.MaxColumns);
        ArgumentOutOfRangeException.ThrowIfNegative(width);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(width, 255);
        columnWidths[column] = width;
    }

    /// <summary>Keeps the top rows and left columns in view while scrolling. Must come before the first row.</summary>
    /// <param name="rows">The number of rows to freeze, for example 1 for a header row.</param>
    /// <param name="columns">The number of columns to freeze.</param>
    public void FreezePanes(int rows, int columns = 0)
    {
        ThrowIfStarted(nameof(FreezePanes));
        ArgumentOutOfRangeException.ThrowIfNegative(rows);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(rows, CellAddress.MaxRows);
        ArgumentOutOfRangeException.ThrowIfNegative(columns);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(columns, CellAddress.MaxColumns);
        frozenRows = rows;
        frozenColumns = columns;
    }

    /// <summary>
    /// Adds filter buttons to the first row, covering every column and row the sheet ends up with.
    /// Can be called at any time before the sheet is finished.
    /// </summary>
    public void EnableAutoFilter()
    {
        ThrowIfFinished();
        autoFilter = true;
    }

    /// <summary>
    /// Adds a conditional format to a range. Can be called at any time before the sheet is
    /// finished; the rule covers the cells whether they are written before or after.
    /// </summary>
    /// <param name="range">The cells, for example <c>B2:B100</c>.</param>
    /// <param name="format">The rule.</param>
    public void AddConditionalFormat(CellRange range, ConditionalFormat format)
    {
        ThrowIfFinished();
        extras.AddConditionalFormat(range, format);
    }

    /// <summary>Adds a conditional format to a range given in A1 notation, such as <c>B2:B100</c>.</summary>
    /// <param name="range">The cells.</param>
    /// <param name="format">The rule.</param>
    public void AddConditionalFormat(string range, ConditionalFormat format)
    {
        AddConditionalFormat(ParseRange(range), format);
    }

    /// <summary>
    /// Adds a chart over a block of cells. Can be called at any time before the sheet is finished.
    /// The series may refer to this sheet and to the sheets added before it.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <param name="area">The cells the chart covers, for example <c>E2:L20</c>.</param>
    /// <returns>Success, or NotFound when a series refers to a sheet the workbook does not have yet.</returns>
    public Result AddChart(Chart chart, CellRange area)
    {
        ThrowIfFinished();

        return extras.AddChart(chart, area, Name, sheetExists);
    }

    /// <summary>Adds a chart over a block of cells given in A1 notation, such as <c>E2:L20</c>.</summary>
    /// <param name="chart">The chart.</param>
    /// <param name="area">The cells the chart covers.</param>
    /// <returns>Success, or NotFound when a series refers to a sheet the workbook does not have yet.</returns>
    public Result AddChart(Chart chart, string area)
    {
        return AddChart(chart, ParseRange(area));
    }

    /// <summary>Puts a picture with its top-left corner at a cell, at its own size times <paramref name="scale"/>.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="topLeft">The cell.</param>
    /// <param name="scale">1 for the picture's own size (96 pixels to the inch), 0.5 for half.</param>
    public void AddPicture(Picture picture, CellAddress topLeft, double scale = 1)
    {
        ThrowIfFinished();
        extras.AddPicture(picture, topLeft, scale);
    }

    /// <summary>Puts a picture over a block of cells, stretched to fill it.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="area">The cells.</param>
    public void AddPicture(Picture picture, CellRange area)
    {
        ThrowIfFinished();
        extras.AddPicture(picture, area);
    }

    /// <summary>
    /// Puts a picture at a cell such as <c>B2</c>, at its own size, or over a range such as
    /// <c>B2:F12</c>, stretched to fill it.
    /// </summary>
    /// <param name="picture">The picture.</param>
    /// <param name="reference">The cell or the range.</param>
    public void AddPicture(Picture picture, string reference)
    {
        CellRange range = ParseRange(reference);

        if (reference.Contains(':', StringComparison.Ordinal) == true)
        {
            AddPicture(picture, range);
        }
        else
        {
            AddPicture(picture, range.First);
        }
    }

    internal static CellRange ParseRange(string reference)
    {
        ArgumentNullException.ThrowIfNull(reference);

        if (CellRange.TryParse(reference, out CellRange range) == false)
        {
            throw new ArgumentException($"'{reference}' is not a cell range. Expected a range such as B2:D10.", nameof(reference));
        }

        return range;
    }

    /// <summary>Writes the next row. Cell <c>i</c> in the list goes in column <c>i</c>; empty cells are skipped.</summary>
    /// <param name="cells">The cells, for example <c>["Name", 42.5, Cell.Of("Total").Bold()]</c>.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    /// <returns>A task that completes when the row is written.</returns>
    public async Task WriteRowAsync(IReadOnlyList<Cell> cells, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(cells);
        ThrowIfFinished();

        if (RowCount >= CellAddress.MaxRows)
        {
            throw new InvalidOperationException($"A sheet holds at most {CellAddress.MaxRows} rows.");
        }

        if (cells.Count > CellAddress.MaxColumns)
        {
            throw new ArgumentException($"A row holds at most {CellAddress.MaxColumns} cells.", nameof(cells));
        }

        cancellationToken.ThrowIfCancellationRequested();
        await StartAsync();
        string rowNumber = (RowCount + 1).ToString(CultureInfo.InvariantCulture);
        await xml.WriteStartElementAsync(null, "row", null);
        await xml.WriteAttributeStringAsync(null, "r", null, rowNumber);

        for (int column = 0; column < cells.Count; column++)
        {
            Cell cell = cells[column];

            if (cell.IsEmpty == true)
            {
                continue;
            }

            if (cell.Formula is not null)
            {
                HasFormulas = true;
            }

            int styleIndex = styles.GetIndex(cell.Style, CellXml.DefaultFormat(cell.Value));
            await CellXml.WriteCellAsync(xml, cell, CellAddress.ColumnName(column) + rowNumber, styleIndex, false);

            if (column > lastColumn)
            {
                lastColumn = column;
            }
        }

        await xml.WriteEndElementAsync();
        RowCount++;
    }

    internal async Task FinishAsync(CancellationToken cancellationToken)
    {
        if (IsFinished == true)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        await StartAsync();
        await xml.WriteEndElementAsync();

        if (autoFilter == true && RowCount > 0 && lastColumn >= 0)
        {
            CellAddress end = new(RowCount - 1, lastColumn);
            AutoFilterEnd = end;
            await SheetXml.WriteAutoFilterAsync(xml, end);
        }

        if (extras.ConditionalFormats.Count > 0)
        {
            await ConditionalFormatXml.WriteAsync(xml, extras.ConditionalFormats, styles, 1);
        }

        if (extras.Drawings.Count > 0)
        {
            await DrawingXml.WriteSheetDrawingAsync(xml, DrawingRelationshipId);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
        await xml.FlushAsync();
        await xml.DisposeAsync();
        await entryStream.DisposeAsync();
        IsFinished = true;
    }

    private async Task StartAsync()
    {
        if (started == true)
        {
            return;
        }

        started = true;
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "worksheet", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await SheetXml.WriteSheetViewsAsync(xml, Index == 0, frozenRows, frozenColumns);
        List<ColumnRange> ranges = columnWidths.Select(pair => new ColumnRange(pair.Key, pair.Key, pair.Value, [])).ToList();
        await SheetXml.WriteColumnsAsync(xml, ranges);
        await xml.WriteStartElementAsync(null, "sheetData", null);
    }

    private void ThrowIfStarted(string operation)
    {
        ThrowIfFinished();

        if (started == true)
        {
            throw new InvalidOperationException($"{operation} must be called before the first row is written: it goes at the top of the sheet in the file.");
        }
    }

    private void ThrowIfFinished()
    {
        if (IsFinished == true)
        {
            throw new InvalidOperationException($"The sheet '{Name}' is finished. A sheet is finished when the next sheet is added or the workbook completes.");
        }
    }
}
