namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Lookup and reference functions.</summary>
internal static class LookupFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("ROW", 0, 1, call => Position(call, true));
        table.Add("COLUMN", 0, 1, call => Position(call, false));
        table.Add("INDIRECT", 1, 2, Indirect, true);
        table.Add("OFFSET", 3, 5, Offset, true);
    }

    private static Value Position(FunctionCall call, bool row)
    {
        if (call.IsMissing(0) == true)
        {
            if (row == true)
            {
                return Value.FromNumber(call.Context.Row + 1);
            }

            return Value.FromNumber(call.Context.Column + 1);
        }

        if (call.TryReference(0, out RangeReference reference, out Value error) == false)
        {
            return error;
        }

        if (row == true)
        {
            return Value.FromNumber(reference.FirstRow + 1);
        }

        return Value.FromNumber(reference.FirstColumn + 1);
    }

    /// <summary>INDIRECT(text, [a1]): a reference from text, such as "B3" or "'Q1'!A1:C9". R1C1 style is not supported (#REF!).</summary>
    private static Value Indirect(FunctionCall call)
    {
        if (call.TryText(0, out string text, out Value error) == false || call.TryBool(1, true, out bool a1, out error) == false)
        {
            return error;
        }

        if (a1 == false)
        {
            return Value.Ref;
        }

        Result<Node> parsed = FormulaParser.Parse(text);

        if (parsed.IsSuccess(out Node? node) == false || node is not ReferenceNode)
        {
            return Value.Ref;
        }

        return call.Evaluator.Evaluate(node, call.Context);
    }

    /// <summary>OFFSET(reference, rows, columns, [height], [width]): a reference moved and resized. Outside the sheet is #REF!.</summary>
    private static Value Offset(FunctionCall call)
    {
        if (call.TryReference(0, out RangeReference start, out Value error) == false)
        {
            return error;
        }

        if (call.TryInteger(1, out long rows, out error) == false
            || call.TryInteger(2, out long columns, out error) == false
            || call.TryInteger(3, start.Rows, out long height, out error) == false
            || call.TryInteger(4, start.Columns, out long width, out error) == false)
        {
            return error;
        }

        if (height == 0 || width == 0)
        {
            return Value.Ref;
        }

        long firstRow = start.FirstRow + rows;
        long firstColumn = start.FirstColumn + columns;
        long lastRow = firstRow + height - 1;
        long lastColumn = firstColumn + width - 1;

        if (height < 0)
        {
            (firstRow, lastRow) = (firstRow + height + 1, firstRow);
        }

        if (width < 0)
        {
            (firstColumn, lastColumn) = (firstColumn + width + 1, firstColumn);
        }

        if (firstRow < 0 || firstColumn < 0 || lastRow >= CellAddress.MaxRows || lastColumn >= CellAddress.MaxColumns)
        {
            return Value.Ref;
        }

        return Value.FromReference(new RangeReference(start.Sheet, (int)firstRow, (int)firstColumn, (int)lastRow, (int)lastColumn));
    }
}
