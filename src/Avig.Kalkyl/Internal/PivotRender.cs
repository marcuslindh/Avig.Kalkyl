using System.Text;

namespace Avig.Kalkyl.Internal;

/// <summary>Where a cache field is in a pivot table.</summary>
internal enum PivotAxis
{
    None,
    Row,
    Column,
    Page,
}

/// <summary>A pivot table worked out: its shape, the cells it fills, and what its definition part needs.</summary>
internal sealed class PivotRendering
{
    public required PivotGrid Grid { get; init; }

    /// <summary>All the cells of the report, the Filters area included.</summary>
    public required CellRange Area { get; init; }

    /// <summary>The cells to write, relative to the sheet.</summary>
    public required IReadOnlyList<(int Row, int Column, Cell Cell)> Cells { get; init; }

    /// <summary>Per cache field, the axis it is on.</summary>
    public required IReadOnlyList<PivotAxis> Axes { get; init; }

    /// <summary>The cache fields in the Filters area, in order.</summary>
    public required IReadOnlyList<int> PageFields { get; init; }

    /// <summary>The value fields with the cache field each combines.</summary>
    public required IReadOnlyList<(PivotValue Value, int Field)> Values { get; init; }
}

/// <summary>
/// Works out a pivot table the way Excel lays it out: the rows and columns of items, subtotals
/// and grand totals, the headings, and the value in every cell. The captions are Excel's English
/// ones (Row Labels, Column Labels, Grand Total); Excel puts its own language's in when it refreshes.
/// </summary>
internal static class PivotRender
{
    public const string RowLabels = "Row Labels";
    public const string ColumnLabels = "Column Labels";
    public const string GrandTotal = "Grand Total";
    public const string All = "(All)";
    public const string Blank = "(blank)";

    /// <summary>
    /// Renders a pivot table. NotFound when a field is not in the source; Invalid when the table
    /// would not fit on the sheet.
    /// </summary>
    /// <param name="pivot">The pivot table.</param>
    /// <param name="data">Its cache data.</param>
    /// <param name="hidden">Per cache field, the shared indexes slicers hide.</param>
    /// <param name="slicerFields">The cache fields slicers filter, which list their items.</param>
    /// <param name="uses1904">True when dates count from 1904.</param>
    public static Result<PivotRendering> Render(CreatedPivot pivot, PivotCacheData data, IReadOnlyDictionary<int, HashSet<int>> hidden, IReadOnlySet<int> slicerFields, bool uses1904)
    {
        PivotTable definition = pivot.Definition;
        List<int[]> resolved = [];

        foreach (IReadOnlyList<string> names in new[] { definition.Rows, definition.Columns, definition.Filters, definition.Values.Select(value => value.Field).ToList() })
        {
            if (Resolve(data, names, pivot.Source).IsFailure(out string? error, out ResultStatus status, out int[]? fields) == true)
            {
                return Result<PivotRendering>.Failure(error, status);
            }

            resolved.Add(fields);
        }

        int[] rows = resolved[0];
        int[] columns = resolved[1];
        int[] pages = resolved[2];
        int[] valueIndexes = resolved[3];
        int fieldCount = data.FieldNames.Count;
        PivotAxis[] axes = new PivotAxis[fieldCount];

        foreach ((int[] fields, PivotAxis axis) in new[] { (rows, PivotAxis.Row), (columns, PivotAxis.Column), (pages, PivotAxis.Page) })
        {
            foreach (int field in fields)
            {
                axes[field] = axis;
            }
        }

        // Items: per field on an axis or under a slicer, the shared indexes in the order shown.
        List<IReadOnlyList<PivotItem>> items = [];
        List<int[]> sharedToItem = [];

        for (int field = 0; field < fieldCount; field++)
        {
            int[] map = new int[data.SharedItems[field].Count];
            Array.Fill(map, -1);
            sharedToItem.Add(map);

            if (axes[field] == PivotAxis.None && slicerFields.Contains(field) == false)
            {
                items.Add([]);

                continue;
            }

            List<PivotItem> list = [];
            hidden.TryGetValue(field, out HashSet<int>? hiddenItems);

            foreach (int shared in data.SortedItems(field))
            {
                map[shared] = list.Count;
                list.Add(new PivotItem(data.SharedItems[field][shared], shared, hiddenItems?.Contains(shared) == true, true));
            }

            list.Add(new PivotItem(CellValue.Empty, -1, false, false));
            items.Add(list);
        }

        // The records the filters let through.
        List<int> visible = [];

        for (int record = 0; record < data.Records.Count; record++)
        {
            bool shown = true;

            foreach ((int field, HashSet<int> hiddenItems) in hidden)
            {
                if (hiddenItems.Contains(data.Records[record][field]) == true)
                {
                    shown = false;

                    break;
                }
            }

            if (shown == true)
            {
                visible.Add(record);
            }
        }

        int dataCount = valueIndexes.Length;
        bool valuesOnColumns = dataCount > 1;
        List<int> columnAxis = [.. columns];

        if (valuesOnColumns == true)
        {
            columnAxis.Add(PivotGrid.ValuesField);
        }

        bool subtotalTop = definition.Layout != PivotLayout.Tabular;
        List<PivotLine> rowLines = BuildAxis(data, visible, rows, sharedToItem, false, dataCount, subtotalTop, definition.GrandTotalRow);
        List<PivotLine> columnLines = BuildAxis(data, visible, columns, sharedToItem, valuesOnColumns, dataCount, false, definition.GrandTotalColumn);

        int firstDataColumn = FirstDataColumn(definition.Layout, rows.Length, columns.Length);
        int firstHeaderRow = 1;
        int firstDataRow = 1;

        if (columns.Length == 0 && valuesOnColumns == true)
        {
            firstHeaderRow = 0;
        }
        else if (columns.Length > 0)
        {
            firstDataRow = 1 + columnAxis.Count;
        }

        int pageRows = 0;

        if (pages.Length > 0)
        {
            pageRows = pages.Length + 1;
        }

        long top = (long)pivot.TopLeft.Row + pageRows;
        long bottom = top + firstDataRow + rowLines.Count - 1;
        long right = (long)pivot.TopLeft.Column + Math.Max(firstDataColumn + columnLines.Count, Math.Min(pages.Length, 1) * 2) - 1;

        if (bottom >= CellAddress.MaxRows || right >= CellAddress.MaxColumns)
        {
            return Result<PivotRendering>.Invalid($"The pivot table '{definition.Name}' would be {bottom - top + 1:N0} rows by {firstDataColumn + columnLines.Count:N0} columns from {pivot.TopLeft} and not fit on the sheet.");
        }

        CellRange location = new(new CellAddress((int)top, pivot.TopLeft.Column), new CellAddress((int)bottom, pivot.TopLeft.Column + firstDataColumn + columnLines.Count - 1));
        CellRange area = new(pivot.TopLeft, new CellAddress((int)bottom, (int)right));

        List<PivotDataField> dataFields = [];

        for (int index = 0; index < dataCount; index++)
        {
            dataFields.Add(new PivotDataField(definition.Values[index].DisplayName, valueIndexes[index]));
        }

        PivotGrid grid = new()
        {
            Name = definition.Name,
            Sheet = pivot.Sheet,
            Location = location,
            FirstHeaderRow = firstHeaderRow,
            FirstDataRow = firstDataRow,
            FirstDataColumn = firstDataColumn,
            CacheId = pivot.CacheId,
            FieldNames = data.FieldNames,
            Items = items,
            RowFields = rows,
            ColumnFields = columnAxis,
            RowLines = rowLines,
            ColumnLines = columnLines,
            DataFields = dataFields,
        };

        List<(PivotValue Value, int Field)> values = [];

        for (int index = 0; index < dataCount; index++)
        {
            values.Add((definition.Values[index], valueIndexes[index]));
        }

        List<(int Row, int Column, Cell Cell)> cells = [];
        WritePageArea(cells, pivot.TopLeft, pages, data);
        WriteHeadings(cells, grid, definition, values);
        WriteRowLabels(cells, grid, definition, values);
        WriteValues(cells, grid, data, visible, sharedToItem, values, uses1904);

        return new PivotRendering
        {
            Grid = grid,
            Area = area,
            Cells = cells,
            Axes = axes,
            PageFields = pages,
            Values = values,
        };
    }

    /// <summary>
    /// The columns of row labels before the values. Without row fields there is one only for the
    /// captions of real column fields; the value fields alone need none.
    /// </summary>
    private static int FirstDataColumn(PivotLayout layout, int rowFields, int columnFields)
    {
        if (rowFields > 0)
        {
            if (layout == PivotLayout.Compact)
            {
                return 1;
            }

            return rowFields;
        }

        if (columnFields > 0)
        {
            return 1;
        }

        return 0;
    }

    private static Result<int[]> Resolve(PivotCacheData data, IReadOnlyList<string> names, PivotSource source)
    {
        int[] fields = new int[names.Count];

        for (int index = 0; index < names.Count; index++)
        {
            int field = data.FieldIndex(names[index].Trim());

            if (field < 0)
            {
                string known = string.Join(", ", data.FieldNames.Select(name => "'" + name + "'"));

                return Result<int[]>.NotFound($"The source {source.Display} has no column headed '{names[index]}'. Its headings are {known}.");
            }

            fields[index] = field;
        }

        return fields;
    }

    /// <summary>
    /// The lines of one axis, as Excel orders them: each item of the outer field, with the items
    /// of the inner fields under it; a group's subtotal on its first line (compact and outline
    /// rows) or after it (tabular rows, columns); the value fields innermost; the grand total last.
    /// </summary>
    private static List<PivotLine> BuildAxis(PivotCacheData data, List<int> visible, int[] fields, List<int[]> sharedToItem, bool values, int dataCount, bool subtotalTop, bool grandTotal)
    {
        AxisNode root = new();

        foreach (int record in visible)
        {
            AxisNode node = root;

            foreach (int field in fields)
            {
                int item = sharedToItem[field][data.Records[record][field]];

                if (node.Children.TryGetValue(item, out AxisNode? child) == false)
                {
                    child = new AxisNode();
                    node.Children[item] = child;
                }

                node = child;
            }
        }

        List<PivotLine> lines = [];
        List<int> prefix = [];
        AddLines(lines, root, prefix, fields.Length, values, dataCount, subtotalTop);

        if (fields.Length > 0 && grandTotal == true)
        {
            int count = 1;

            if (values == true)
            {
                count = dataCount;
            }

            for (int index = 0; index < count; index++)
            {
                lines.Add(new PivotLine(PivotLineKind.Grand, [], index, 0));
            }
        }

        return lines;
    }

    private static void AddLines(List<PivotLine> lines, AxisNode node, List<int> prefix, int depthCount, bool values, int dataCount, bool subtotalTop)
    {
        if (prefix.Count == depthCount)
        {
            if (values == true)
            {
                for (int index = 0; index < dataCount; index++)
                {
                    AddDataLine(lines, [.. prefix, index], index);
                }

                return;
            }

            AddDataLine(lines, [.. prefix], 0);

            return;
        }

        bool inner = prefix.Count + 1 < depthCount;

        foreach ((int item, AxisNode child) in node.Children)
        {
            prefix.Add(item);

            if (inner == true && subtotalTop == true)
            {
                AddDataLine(lines, [.. prefix], 0);
            }

            AddLines(lines, child, prefix, depthCount, values, dataCount, subtotalTop);

            if (inner == true && subtotalTop == false)
            {
                int count = 1;

                if (values == true)
                {
                    count = dataCount;
                }

                for (int index = 0; index < count; index++)
                {
                    lines.Add(new PivotLine(PivotLineKind.Subtotal, [.. prefix], index, 0));
                }
            }

            prefix.RemoveAt(prefix.Count - 1);
        }
    }

    private static void AddDataLine(List<PivotLine> lines, int[] members, int dataIndex)
    {
        int repeat = 0;

        if (lines.Count > 0 && lines[^1].Kind == PivotLineKind.Data)
        {
            int[] previous = lines[^1].Members;

            while (repeat < members.Length && repeat < previous.Length && previous[repeat] == members[repeat])
            {
                repeat++;
            }

            // A line repeats at most all but its last member.
            repeat = Math.Min(repeat, Math.Max(members.Length - 1, 0));
        }

        lines.Add(new PivotLine(PivotLineKind.Data, members, dataIndex, repeat));
    }

    private static void WritePageArea(List<(int Row, int Column, Cell Cell)> cells, CellAddress topLeft, int[] pages, PivotCacheData data)
    {
        for (int index = 0; index < pages.Length; index++)
        {
            cells.Add((topLeft.Row + index, topLeft.Column, Cell.Of(data.FieldNames[pages[index]])));
            cells.Add((topLeft.Row + index, topLeft.Column + 1, Cell.Of(All)));
        }
    }

    private static void WriteHeadings(List<(int Row, int Column, Cell Cell)> cells, PivotGrid grid, PivotTable definition, List<(PivotValue Value, int Field)> values)
    {
        int top = grid.Location.First.Row;
        int left = grid.Location.First.Column;
        int realColumns = grid.ColumnFields.Count(field => field != PivotGrid.ValuesField);
        int headerRow = top + grid.FirstDataRow - 1;

        // The row field captions, on the last heading row.
        if (grid.RowFields.Count > 0)
        {
            if (definition.Layout == PivotLayout.Compact)
            {
                cells.Add((headerRow, left, Cell.Of(RowLabels)));
            }
            else
            {
                for (int index = 0; index < grid.RowFields.Count; index++)
                {
                    cells.Add((headerRow, left + index, Cell.Of(grid.FieldNames[grid.RowFields[index]])));
                }
            }
        }

        if (realColumns == 0)
        {
            for (int index = 0; index < grid.ColumnLines.Count; index++)
            {
                cells.Add((headerRow, left + grid.FirstDataColumn + index, Cell.Of(values[grid.ColumnLines[index].DataIndex].Value.DisplayName)));
            }

            return;
        }

        // The first row: the value field's name, and the column field captions.
        if (values.Count == 1 && grid.RowFields.Count > 0)
        {
            cells.Add((top, left, Cell.Of(values[0].Value.DisplayName)));
        }

        if (definition.Layout == PivotLayout.Compact)
        {
            cells.Add((top, left + grid.FirstDataColumn, Cell.Of(ColumnLabels)));
        }
        else
        {
            for (int index = 0; index < realColumns; index++)
            {
                cells.Add((top, left + grid.FirstDataColumn + index, Cell.Of(grid.FieldNames[grid.ColumnFields[index]])));
            }
        }

        // One row per column field: each line's items from the first that changes.
        for (int index = 0; index < grid.ColumnLines.Count; index++)
        {
            PivotLine line = grid.ColumnLines[index];
            int column = left + grid.FirstDataColumn + index;

            switch (line.Kind)
            {
                case PivotLineKind.Data:
                    for (int depth = line.Repeat; depth < line.Members.Length; depth++)
                    {
                        cells.Add((top + 1 + depth, column, LabelCell(grid, grid.ColumnFields[depth], line.Members[depth], values)));
                    }

                    break;

                case PivotLineKind.Subtotal:
                    {
                        int depth = line.Members.Length - 1;
                        string label = Label(grid, grid.ColumnFields[depth], line.Members[depth]) + " Total";

                        if (grid.ColumnFields[^1] == PivotGrid.ValuesField)
                        {
                            label = Label(grid, grid.ColumnFields[depth], line.Members[depth]) + " " + values[line.DataIndex].Value.DisplayName;
                        }

                        cells.Add((top + 1 + depth, column, Cell.Of(label)));

                        break;
                    }

                case PivotLineKind.Grand:
                    if (grid.ColumnFields[^1] == PivotGrid.ValuesField)
                    {
                        cells.Add((top + 1, column, Cell.Of("Total " + values[line.DataIndex].Value.DisplayName)));
                    }
                    else
                    {
                        cells.Add((top + 1, column, Cell.Of(GrandTotal)));
                    }

                    break;
            }
        }
    }

    private static void WriteRowLabels(List<(int Row, int Column, Cell Cell)> cells, PivotGrid grid, PivotTable definition, List<(PivotValue Value, int Field)> values)
    {
        int left = grid.Location.First.Column;
        int realColumns = grid.ColumnFields.Count(field => field != PivotGrid.ValuesField);

        for (int index = 0; index < grid.RowLines.Count; index++)
        {
            PivotLine line = grid.RowLines[index];
            int row = grid.Location.First.Row + grid.FirstDataRow + index;

            switch (line.Kind)
            {
                case PivotLineKind.Data when grid.RowFields.Count == 0:
                    // Without row fields, the one value field's name stands beside its values.
                    if (realColumns > 0 && values.Count == 1)
                    {
                        cells.Add((row, left, Cell.Of(values[0].Value.DisplayName)));
                    }

                    break;

                case PivotLineKind.Data when definition.Layout == PivotLayout.Compact:
                    cells.Add((row, left, LabelCell(grid, grid.RowFields[line.Members.Length - 1], line.Members[^1], values)));

                    break;

                case PivotLineKind.Data:
                    for (int depth = line.Repeat; depth < line.Members.Length; depth++)
                    {
                        cells.Add((row, left + depth, LabelCell(grid, grid.RowFields[depth], line.Members[depth], values)));
                    }

                    break;

                case PivotLineKind.Subtotal:
                    {
                        int depth = line.Members.Length - 1;
                        cells.Add((row, left + depth, Cell.Of(Label(grid, grid.RowFields[depth], line.Members[depth]) + " Total")));

                        break;
                    }

                case PivotLineKind.Grand:
                    cells.Add((row, left, Cell.Of(GrandTotal)));

                    break;
            }
        }
    }

    private static void WriteValues(List<(int Row, int Column, Cell Cell)> cells, PivotGrid grid, PivotCacheData data, List<int> visible, List<int[]> sharedToItem, List<(PivotValue Value, int Field)> values, bool uses1904)
    {
        int realColumns = grid.ColumnFields.Count(field => field != PivotGrid.ValuesField);
        bool valuesOnColumns = grid.ColumnFields.Count > realColumns;
        Dictionary<string, int> rowIndex = LineIndex(grid.RowLines, false);
        Dictionary<string, int> columnIndex = LineIndex(grid.ColumnLines, valuesOnColumns);
        Dictionary<(int Row, int Column), Accumulator> totals = [];
        int[] rowMembers = new int[grid.RowFields.Count];
        int[] columnMembers = new int[realColumns];
        List<int> rowLines = [];
        List<(int Line, int Data)> columnLines = [];

        foreach (int record in visible)
        {
            for (int depth = 0; depth < rowMembers.Length; depth++)
            {
                int field = grid.RowFields[depth];
                rowMembers[depth] = sharedToItem[field][data.Records[record][field]];
            }

            for (int depth = 0; depth < columnMembers.Length; depth++)
            {
                int field = grid.ColumnFields[depth];
                columnMembers[depth] = sharedToItem[field][data.Records[record][field]];
            }

            rowLines.Clear();

            for (int length = 0; length <= rowMembers.Length; length++)
            {
                if (rowIndex.TryGetValue(Key(rowMembers, length, -1), out int line) == true)
                {
                    rowLines.Add(line);
                }
            }

            columnLines.Clear();

            for (int length = 0; length <= columnMembers.Length; length++)
            {
                if (valuesOnColumns == true)
                {
                    for (int dataIndex = 0; dataIndex < values.Count; dataIndex++)
                    {
                        if (columnIndex.TryGetValue(Key(columnMembers, length, dataIndex), out int line) == true)
                        {
                            columnLines.Add((line, dataIndex));
                        }
                    }
                }
                else if (columnIndex.TryGetValue(Key(columnMembers, length, -1), out int line) == true)
                {
                    columnLines.Add((line, 0));
                }
            }

            foreach (int row in rowLines)
            {
                foreach ((int column, int dataIndex) in columnLines)
                {
                    if (totals.TryGetValue((row, column), out Accumulator? accumulator) == false)
                    {
                        accumulator = new Accumulator();
                        totals[(row, column)] = accumulator;
                    }

                    accumulator.Add(data.Value(record, values[dataIndex].Field), uses1904);
                }
            }
        }

        foreach (((int row, int column), Accumulator accumulator) in totals.OrderBy(pair => pair.Key.Row).ThenBy(pair => pair.Key.Column))
        {
            if (accumulator.TryResult(values[grid.ColumnLines[column].DataIndex].Value.Aggregation, out CellValue result) == false)
            {
                continue;
            }

            Cell cell = Cell.Of(result);
            string? format = values[grid.ColumnLines[column].DataIndex].Value.NumberFormat;

            if (format is not null)
            {
                cell = cell.NumberFormat(format);
            }

            cells.Add((grid.Location.First.Row + grid.FirstDataRow + row, grid.Location.First.Column + grid.FirstDataColumn + column, cell));
        }
    }

    /// <summary>
    /// The lines of an axis by what they total: the real fields' items (the first n of them for a
    /// subtotal, none for the grand total), and the value field when the values are on the axis.
    /// </summary>
    private static Dictionary<string, int> LineIndex(IReadOnlyList<PivotLine> lines, bool hasValues)
    {
        Dictionary<string, int> index = new(StringComparer.Ordinal);

        for (int line = 0; line < lines.Count; line++)
        {
            PivotLine item = lines[line];

            if (item.Kind == PivotLineKind.Other)
            {
                continue;
            }

            int[] members = item.Members;
            int dataIndex = -1;

            if (hasValues == true)
            {
                dataIndex = item.DataIndex;

                if (item.Kind == PivotLineKind.Data)
                {
                    members = members[..^1];
                }
            }

            index.TryAdd(Key(members, members.Length, dataIndex), line);
        }

        return index;
    }

    private static string Key(int[] members, int length, int dataIndex)
    {
        StringBuilder builder = new();
        builder.Append(length).Append(':');

        for (int index = 0; index < length; index++)
        {
            builder.Append(members[index]).Append(',');
        }

        builder.Append('|').Append(dataIndex);

        return builder.ToString();
    }

    private static Cell LabelCell(PivotGrid grid, int field, int member, List<(PivotValue Value, int Field)> values)
    {
        if (field == PivotGrid.ValuesField)
        {
            return Cell.Of(values[member].Value.DisplayName);
        }

        CellValue value = grid.Items[field][member].Value;

        if (value.IsEmpty == true)
        {
            return Cell.Of(Blank);
        }

        return Cell.Of(value);
    }

    private static string Label(PivotGrid grid, int field, int member)
    {
        CellValue value = grid.Items[field][member].Value;

        if (value.IsEmpty == true)
        {
            return Blank;
        }

        return value.ToString();
    }

    private sealed class AxisNode
    {
        public SortedDictionary<int, AxisNode> Children { get; } = [];
    }

    /// <summary>Sums, counts and extremes of the values in one cell of a pivot table.</summary>
    private sealed class Accumulator
    {
        private double sum;
        private int count;
        private int numbers;
        private double min = double.MaxValue;
        private double max = double.MinValue;
        private string? error;

        public void Add(CellValue value, bool uses1904)
        {
            if (value.IsEmpty == true)
            {
                return;
            }

            count++;

            if (value.TryGetError(out string? code) == true)
            {
                error ??= code;

                return;
            }

            double number;

            if (value.TryGetNumber(out double plain) == true)
            {
                number = plain;
            }
            else if (value.TryGetDate(out DateTime date) == true)
            {
                number = ExcelDate.ToSerial(date, uses1904);
            }
            else
            {
                return;
            }

            numbers++;
            sum += number;
            min = Math.Min(min, number);
            max = Math.Max(max, number);
        }

        /// <summary>The cell's value; false when the cell stays blank, as Excel leaves a cell with no values.</summary>
        public bool TryResult(PivotAggregation aggregation, out CellValue result)
        {
            result = CellValue.Empty;

            if (count == 0)
            {
                return false;
            }

            if (error is not null && aggregation != PivotAggregation.Count)
            {
                result = CellValue.Error(error);

                return true;
            }

            switch (aggregation)
            {
                case PivotAggregation.Count:
                    result = CellValue.Number(count);

                    break;
                case PivotAggregation.Average when numbers == 0:
                    result = CellValue.Error("#DIV/0!");

                    break;
                case PivotAggregation.Average:
                    result = CellValue.Number(sum / numbers);

                    break;
                case PivotAggregation.Min when numbers > 0:
                    result = CellValue.Number(min);

                    break;
                case PivotAggregation.Max when numbers > 0:
                    result = CellValue.Number(max);

                    break;
                case PivotAggregation.Min or PivotAggregation.Max:
                    result = CellValue.Number(0);

                    break;
                default:
                    result = CellValue.Number(sum);

                    break;
            }

            return true;
        }
    }
}
