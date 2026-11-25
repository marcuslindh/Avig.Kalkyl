using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>GETPIVOTDATA, which reads a value out of a pivot table by the names of its fields and items.</summary>
internal static class PivotFunctions
{
    public static void Register(FunctionTable table)
    {
        // Volatile: the value comes from any cell of the pivot table, which the formula's
        // references (one cell of it) do not tell the dependency graph about.
        table.Add("GETPIVOTDATA", 2, FunctionTable.Many, GetPivotData, true);
    }

    /// <summary>
    /// GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...): the value of a pivot table's
    /// cell, found as Excel finds it. The pivot table is the one whose cells hold the reference.
    /// A field left out means its total: the subtotal of the fields given, or the grand total.
    /// #REF! when there is no such pivot table, field, item or total, as in Excel.
    /// </summary>
    private static Value GetPivotData(FunctionCall call)
    {
        if (call.Count % 2 != 0)
        {
            return Value.Ref;
        }

        if (call.TryText(0, out string dataName, out Value error) == false)
        {
            return error;
        }

        Value target = call.Raw(1);

        if (target.IsError == true)
        {
            return target;
        }

        if (target.Kind != ValueKind.Reference)
        {
            return Value.Ref;
        }

        RangeReference reference = target.Reference!;

        if (TryFindGrid(call.Evaluator.Workbook, reference, out PivotGrid? grid) == false)
        {
            return Value.Ref;
        }

        int dataIndex = FindDataField(grid, dataName.Trim());

        if (dataIndex < 0)
        {
            return Value.Ref;
        }

        int[] rowSpecified = new int[grid.RowFields.Count];
        int[] columnSpecified = new int[grid.ColumnFields.Count];
        Array.Fill(rowSpecified, -1);
        Array.Fill(columnSpecified, -1);

        for (int index = 2; index < call.Count; index += 2)
        {
            if (call.TryText(index, out string fieldName, out error) == false)
            {
                return error;
            }

            Value item = call.Scalar(index + 1);

            if (item.IsError == true)
            {
                return item;
            }

            int field = FindField(grid, fieldName.Trim());

            if (field < 0)
            {
                return Value.Ref;
            }

            int member = FindMember(grid.Items[field], item);
            int rowDepth = IndexOf(grid.RowFields, field);
            int columnDepth = IndexOf(grid.ColumnFields, field);

            if (member < 0 || (rowDepth < 0 && columnDepth < 0))
            {
                return Value.Ref;
            }

            if (rowDepth >= 0)
            {
                rowSpecified[rowDepth] = member;
            }
            else
            {
                columnSpecified[columnDepth] = member;
            }
        }

        int rowLine = FindLine(grid.RowLines, grid.RowFields, rowSpecified, dataIndex);
        int columnLine = FindLine(grid.ColumnLines, grid.ColumnFields, columnSpecified, dataIndex);

        if (rowLine < 0 || columnLine < 0)
        {
            return Value.Ref;
        }

        int row = grid.Location.First.Row + grid.FirstDataRow + rowLine;
        int column = grid.Location.First.Column + grid.FirstDataColumn + columnLine;

        if (row > grid.Location.Last.Row || column > grid.Location.Last.Column)
        {
            return Value.Ref;
        }

        Value value = call.Evaluator.CellValue(grid.Sheet, row, column);

        // A cell without data: Excel gives #REF! rather than a blank.
        if (value.IsBlank == true)
        {
            return Value.Ref;
        }

        return value;
    }

    private static bool TryFindGrid(Workbook book, RangeReference reference, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out PivotGrid? found)
    {
        foreach (PivotGrid grid in book.Pivots.Grids)
        {
            if (ReferenceEquals(grid.Sheet, reference.Sheet) == true && grid.Location.Contains(new CellAddress(reference.FirstRow, reference.FirstColumn)) == true)
            {
                found = grid;

                return true;
            }
        }

        found = null;

        return false;
    }

    /// <summary>The value field by its name ("Sum of Sales") or by the name of the field it combines ("Sales"); -1 when there is none.</summary>
    private static int FindDataField(PivotGrid grid, string name)
    {
        for (int index = 0; index < grid.DataFields.Count; index++)
        {
            if (string.Equals(grid.DataFields[index].Name.Trim(), name, StringComparison.OrdinalIgnoreCase) == true)
            {
                return index;
            }
        }

        for (int index = 0; index < grid.DataFields.Count; index++)
        {
            int field = grid.DataFields[index].Field;

            if (field >= 0 && field < grid.FieldNames.Count && string.Equals(grid.FieldNames[field], name, StringComparison.OrdinalIgnoreCase) == true)
            {
                return index;
            }
        }

        return -1;
    }

    private static int FindField(PivotGrid grid, string name)
    {
        for (int index = 0; index < grid.FieldNames.Count; index++)
        {
            if (string.Equals(grid.FieldNames[index], name, StringComparison.OrdinalIgnoreCase) == true)
            {
                return index;
            }
        }

        return -1;
    }

    /// <summary>The index of the item an argument names: the same number, or the same text as the item shows.</summary>
    private static int FindMember(IReadOnlyList<PivotItem> items, Value argument)
    {
        for (int index = 0; index < items.Count; index++)
        {
            PivotItem item = items[index];

            if (item.IsMember == false)
            {
                continue;
            }

            if (argument.Kind == ValueKind.Number && TryNumber(item.Value, out double number) == true && number == argument.Number)
            {
                return index;
            }

            string text = argument.Text ?? string.Empty;

            if (argument.Kind == ValueKind.Number)
            {
                text = Value.NumberText(argument.Number);
            }
            else if (argument.Kind == ValueKind.Bool)
            {
                text = argument.Bool.ToString().ToUpperInvariant();
            }

            if (PivotRegistry.Matches(item.Value, text) == true)
            {
                return index;
            }
        }

        return -1;
    }

    private static bool TryNumber(CellValue value, out double number)
    {
        if (value.TryGetNumber(out number) == true)
        {
            return true;
        }

        if (value.TryGetDate(out DateTime date) == true)
        {
            number = ExcelDate.ToSerial(date);

            return true;
        }

        return false;
    }

    /// <summary>
    /// The line of an axis that totals the items given: they must be the first fields of the
    /// axis, and the fields after them are totalled (a subtotal, or the grand total when none is
    /// given). -1 when the pivot table shows no such line.
    /// </summary>
    private static int FindLine(IReadOnlyList<PivotLine> lines, IReadOnlyList<int> fields, int[] specified, int dataIndex)
    {
        int valuesAt = IndexOf(fields, PivotGrid.ValuesField);
        List<int> wanted = [];
        bool ended = false;

        for (int depth = 0; depth < fields.Count; depth++)
        {
            if (depth == valuesAt)
            {
                continue;
            }

            if (specified[depth] < 0)
            {
                ended = true;

                continue;
            }

            if (ended == true)
            {
                return -1;
            }

            wanted.Add(specified[depth]);
        }

        for (int index = 0; index < lines.Count; index++)
        {
            PivotLine line = lines[index];

            if (line.Kind == PivotLineKind.Other)
            {
                continue;
            }

            List<int> members = [];
            int lineData = line.DataIndex;

            for (int depth = 0; depth < line.Members.Length; depth++)
            {
                if (depth == valuesAt)
                {
                    if (line.Kind == PivotLineKind.Data)
                    {
                        lineData = line.Members[depth];
                    }

                    continue;
                }

                members.Add(line.Members[depth]);
            }

            if (valuesAt >= 0 && lineData != dataIndex)
            {
                continue;
            }

            if (members.Count == wanted.Count && members.SequenceEqual(wanted) == true)
            {
                return index;
            }
        }

        return -1;
    }

    private static int IndexOf(IReadOnlyList<int> fields, int field)
    {
        for (int index = 0; index < fields.Count; index++)
        {
            if (fields[index] == field)
            {
                return index;
            }
        }

        return -1;
    }
}
