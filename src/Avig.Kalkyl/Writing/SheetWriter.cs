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
    private readonly SortedDictionary<int, double> columnWidths = [];
    private int frozenRows;
    private int frozenColumns;
    private bool autoFilter;
    private bool started;
    private int lastColumn = -1;

    internal SheetWriter(string name, int index, Stream entryStream, Writing.StyleRegistry styles)
    {
        Name = name;
        Index = index;
        this.entryStream = entryStream;
        this.styles = styles;
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

    /// <summary>Nedre högra hörnet i autofiltret, som alltid börjar i A1. Sätts när bladet avslutas.</summary>
    internal CellAddress? AutoFilterEnd { get; private set; }

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

            await WriteCellAsync(cell, CellAddress.ColumnName(column) + rowNumber);

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
            await xml.WriteStartElementAsync(null, "autoFilter", null);
            await xml.WriteAttributeStringAsync(null, "ref", null, "A1:" + end.ToString());
            await xml.WriteEndElementAsync();
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
        await WriteSheetViewAsync();

        if (columnWidths.Count > 0)
        {
            await xml.WriteStartElementAsync(null, "cols", null);

            foreach ((int column, double width) in columnWidths)
            {
                string number = (column + 1).ToString(CultureInfo.InvariantCulture);
                await xml.WriteStartElementAsync(null, "col", null);
                await xml.WriteAttributeStringAsync(null, "min", null, number);
                await xml.WriteAttributeStringAsync(null, "max", null, number);
                await xml.WriteAttributeStringAsync(null, "width", null, width.ToString("R", CultureInfo.InvariantCulture));
                await xml.WriteAttributeStringAsync(null, "customWidth", null, "1");
                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync(null, "sheetData", null);
    }

    private async Task WriteSheetViewAsync()
    {
        await xml.WriteStartElementAsync(null, "sheetViews", null);
        await xml.WriteStartElementAsync(null, "sheetView", null);

        if (Index == 0)
        {
            await xml.WriteAttributeStringAsync(null, "tabSelected", null, "1");
        }

        await xml.WriteAttributeStringAsync(null, "workbookViewId", null, "0");

        if (frozenRows > 0 || frozenColumns > 0)
        {
            string activePane = "bottomRight";

            if (frozenColumns == 0)
            {
                activePane = "bottomLeft";
            }
            else if (frozenRows == 0)
            {
                activePane = "topRight";
            }

            await xml.WriteStartElementAsync(null, "pane", null);

            if (frozenColumns > 0)
            {
                await xml.WriteAttributeStringAsync(null, "xSplit", null, frozenColumns.ToString(CultureInfo.InvariantCulture));
            }

            if (frozenRows > 0)
            {
                await xml.WriteAttributeStringAsync(null, "ySplit", null, frozenRows.ToString(CultureInfo.InvariantCulture));
            }

            await xml.WriteAttributeStringAsync(null, "topLeftCell", null, new CellAddress(frozenRows, frozenColumns).ToString());
            await xml.WriteAttributeStringAsync(null, "activePane", null, activePane);
            await xml.WriteAttributeStringAsync(null, "state", null, "frozen");
            await xml.WriteEndElementAsync();
            await xml.WriteStartElementAsync(null, "selection", null);
            await xml.WriteAttributeStringAsync(null, "pane", null, activePane);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private async Task WriteCellAsync(Cell cell, string reference)
    {
        CellValue value = cell.Value;
        string? defaultFormat = null;

        if (value.TryGetDate(out DateTime date) == true)
        {
            defaultFormat = NumberFormats.DateFormat;

            if (date.TimeOfDay != TimeSpan.Zero)
            {
                defaultFormat = NumberFormats.DateTimeFormat;
            }
        }

        int styleIndex = styles.GetIndex(cell.Style, defaultFormat);
        await xml.WriteStartElementAsync(null, "c", null);
        await xml.WriteAttributeStringAsync(null, "r", null, reference);

        if (styleIndex != 0)
        {
            await xml.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture));
        }

        if (cell.Formula is not null)
        {
            HasFormulas = true;
            await WriteFormulaCellAsync(cell.Formula, value);
            await xml.WriteEndElementAsync();

            return;
        }

        switch (value.Kind)
        {
            case CellKind.Text:
                value.TryGetText(out string? text);
                await xml.WriteAttributeStringAsync(null, "t", null, "inlineStr");
                await xml.WriteStartElementAsync(null, "is", null);
                await xml.WriteStartElementAsync(null, "t", null);
                await xml.WriteAttributeStringAsync("xml", "space", null, "preserve");
                await xml.WriteStringAsync(CellText.Encode(text!));
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;

            case CellKind.Empty:
                break;

            default:
                await WriteTypeAttributeAsync(value.Kind);
                await xml.WriteElementStringAsync(null, "v", null, StoredValue(value));

                break;
        }

        await xml.WriteEndElementAsync();
    }

    private async Task WriteFormulaCellAsync(string formula, CellValue cached)
    {
        if (cached.Kind == CellKind.Text)
        {
            await xml.WriteAttributeStringAsync(null, "t", null, "str");
        }
        else
        {
            await WriteTypeAttributeAsync(cached.Kind);
        }

        await xml.WriteElementStringAsync(null, "f", null, CellText.Encode(formula));

        if (cached.IsEmpty == false)
        {
            await xml.WriteElementStringAsync(null, "v", null, StoredValue(cached));
        }
    }

    private async Task WriteTypeAttributeAsync(CellKind kind)
    {
        if (kind == CellKind.Bool)
        {
            await xml.WriteAttributeStringAsync(null, "t", null, "b");
        }
        else if (kind == CellKind.Error)
        {
            await xml.WriteAttributeStringAsync(null, "t", null, "e");
        }
    }

    private static string StoredValue(CellValue value)
    {
        switch (value.Kind)
        {
            case CellKind.Number:
                value.TryGetNumber(out double number);

                return number.ToString("R", CultureInfo.InvariantCulture);

            case CellKind.Date:
                value.TryGetDate(out DateTime date);

                return ExcelDate.ToSerial(date).ToString("R", CultureInfo.InvariantCulture);

            case CellKind.Bool:
                value.TryGetBool(out bool flag);

                if (flag == true)
                {
                    return "1";
                }

                return "0";

            case CellKind.Text:
                value.TryGetText(out string? text);

                return CellText.Encode(text!);

            case CellKind.Error:
                value.TryGetError(out string? code);

                return code!;

            default:
                return string.Empty;
        }
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
