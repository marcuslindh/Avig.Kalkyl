using System.Globalization;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Lookup and reference functions. Comparisons follow <see cref="Coerce.Compare"/>, so text is
/// matched without regard to case; an exact match only finds a value of the same kind (the
/// number 1 does not match the text "1"), and approximate matches skip values of another kind.
/// </summary>
internal static class LookupFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("ROW", 0, 1, call => Position(call, true));
        table.Add("COLUMN", 0, 1, call => Position(call, false));
        table.Add("INDIRECT", 1, 2, Indirect, true);
        table.Add("OFFSET", 3, 5, Offset, true);
        table.Add("VLOOKUP", 3, 4, call => TableLookup(call, true));
        table.Add("HLOOKUP", 3, 4, call => TableLookup(call, false));
        table.Add("LOOKUP", 2, 3, Lookup);
        table.Add("MATCH", 2, 3, Match);
        table.Add("XLOOKUP", 3, 6, XLookup);
        table.Add("XMATCH", 2, 4, XMatch);
        table.Add("INDEX", 2, 4, Index);
        table.Add("CHOOSE", 2, FunctionTable.Many, Choose);
        table.Add("ROWS", 1, 1, call => Size(call, true));
        table.Add("COLUMNS", 1, 1, call => Size(call, false));
        table.Add("ADDRESS", 2, 5, Address);
        table.Add("TRANSPOSE", 1, 1, Transpose);
        table.Add("AREAS", 1, 1, Areas);
        table.Add("HYPERLINK", 1, 2, Hyperlink);
        table.Add("FORMULATEXT", 1, 1, FormulaText);
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

    /// <summary>
    /// VLOOKUP and HLOOKUP(value, table, index, [approximate]): approximate (the default) finds the
    /// last value ≤ the lookup value in a sorted first column or row; exact matching allows
    /// wildcards in text. #N/A when nothing is found, #REF! when the index is past the table.
    /// </summary>
    private static Value TableLookup(FunctionCall call, bool vertical)
    {
        Value lookup = call.Scalar(0);

        if (lookup.IsError == true)
        {
            return lookup;
        }

        if (call.TryArray(1, out ValueArray table, out Value error) == false
            || call.TryInteger(2, out long index, out error) == false
            || call.TryBool(3, true, out bool approximate, out error) == false)
        {
            return error;
        }

        long size = table.Rows;

        if (vertical == true)
        {
            size = table.Columns;
        }

        if (index < 1)
        {
            return Value.ValueError;
        }

        if (index > size)
        {
            return Value.Ref;
        }

        if (lookup.IsBlank == true)
        {
            return Value.NA;
        }

        Value[] keys = Column(table, 0);

        if (vertical == false)
        {
            keys = Row(table, 0);
        }

        int found = FindLegacy(lookup, keys, approximate, true);

        if (found < 0)
        {
            return Value.NA;
        }

        if (vertical == true)
        {
            return table[found, (int)index - 1];
        }

        return table[(int)index - 1, found];
    }

    /// <summary>
    /// LOOKUP(value, vector, [result]): an approximate match in a sorted vector. Without a result
    /// vector, the array form searches the first row or column (whichever is longer) and returns
    /// from the last.
    /// </summary>
    private static Value Lookup(FunctionCall call)
    {
        Value lookup = call.Scalar(0);

        if (lookup.IsError == true)
        {
            return lookup;
        }

        if (call.TryArray(1, out ValueArray source, out Value error) == false)
        {
            return error;
        }

        Value[] keys;
        Value[] results;

        if (call.IsMissing(2) == true)
        {
            if (source.Columns > source.Rows)
            {
                keys = Row(source, 0);
                results = Row(source, source.Rows - 1);
            }
            else
            {
                keys = Column(source, 0);
                results = Column(source, source.Columns - 1);
            }
        }
        else
        {
            if (call.TryArray(2, out ValueArray resultArray, out error) == false)
            {
                return error;
            }

            if (TryVector(source, out keys) == false || TryVector(resultArray, out results) == false)
            {
                return Value.NA;
            }
        }

        if (lookup.IsBlank == true)
        {
            return Value.NA;
        }

        int found = FindLegacy(lookup, keys, true, false);

        if (found < 0 || found >= results.Length)
        {
            return Value.NA;
        }

        return results[found];
    }

    /// <summary>MATCH(value, vector, [type]): 1 last ≤ in ascending order, 0 exact (wildcards for text), -1 last ≥ in descending order.</summary>
    private static Value Match(FunctionCall call)
    {
        Value lookup = call.Scalar(0);

        if (lookup.IsError == true)
        {
            return lookup;
        }

        if (call.TryArray(1, out ValueArray source, out Value error) == false || call.TryNumber(2, 1, out double type, out error) == false)
        {
            return error;
        }

        if (TryVector(source, out Value[] items) == false || lookup.IsBlank == true)
        {
            return Value.NA;
        }

        int found;

        if (type == 0)
        {
            found = FindLegacy(lookup, items, false, true);
        }
        else if (type > 0)
        {
            found = FindLegacy(lookup, items, true, false);
        }
        else
        {
            found = LastAtMost(lookup, items, -1);
        }

        if (found < 0)
        {
            return Value.NA;
        }

        return Value.FromNumber(found + 1);
    }

    /// <summary>
    /// XLOOKUP(value, lookup, return, [if_not_found], [match_mode], [search_mode]). Match modes:
    /// 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcards. Search modes: 1
    /// first to last, -1 last to first, 2 binary in ascending order, -2 binary in descending order.
    /// A reference to return from gives a reference back.
    /// </summary>
    private static Value XLookup(FunctionCall call)
    {
        Value lookup = call.Scalar(0);

        if (lookup.IsError == true)
        {
            return lookup;
        }

        if (call.TryArray(1, out ValueArray source, out Value error) == false
            || TryModes(call, 4, out long matchMode, out long searchMode, out error) == false)
        {
            return error;
        }

        if (TryVector(source, out Value[] items) == false)
        {
            return Value.ValueError;
        }

        Value returned = call.Raw(2);

        if (returned.IsError == true)
        {
            return returned;
        }

        bool vertical = source.Columns == 1;
        int found = Find(lookup, items, matchMode, searchMode);

        if (found < 0)
        {
            if (call.IsMissing(3) == false)
            {
                return call.Raw(3);
            }

            return Value.NA;
        }

        if (returned.Kind == ValueKind.Reference)
        {
            RangeReference range = returned.Reference!;

            if (vertical == true)
            {
                if (UsedRows(range) != items.Length)
                {
                    return Value.ValueError;
                }

                return Value.FromReference(range with { FirstRow = range.FirstRow + found, LastRow = range.FirstRow + found });
            }

            if (UsedColumns(range) != items.Length)
            {
                return Value.ValueError;
            }

            return Value.FromReference(range with { FirstColumn = range.FirstColumn + found, LastColumn = range.FirstColumn + found });
        }

        ValueArray results = call.Evaluator.ToValueArray(returned);

        if (vertical == true)
        {
            if (results.Rows != items.Length)
            {
                return Value.ValueError;
            }

            return Slice(results, found, found, 0, results.Columns - 1);
        }

        if (results.Columns != items.Length)
        {
            return Value.ValueError;
        }

        return Slice(results, 0, results.Rows - 1, found, found);
    }

    /// <summary>XMATCH(value, vector, [match_mode], [search_mode]): the position XLOOKUP would use.</summary>
    private static Value XMatch(FunctionCall call)
    {
        Value lookup = call.Scalar(0);

        if (lookup.IsError == true)
        {
            return lookup;
        }

        if (call.TryArray(1, out ValueArray source, out Value error) == false
            || TryModes(call, 2, out long matchMode, out long searchMode, out error) == false)
        {
            return error;
        }

        if (TryVector(source, out Value[] items) == false)
        {
            return Value.ValueError;
        }

        int found = Find(lookup, items, matchMode, searchMode);

        if (found < 0)
        {
            return Value.NA;
        }

        return Value.FromNumber(found + 1);
    }

    /// <summary>
    /// INDEX(array, row, [column], [area]). A reference gives a reference, an array an item or a
    /// sub-array. Row or column 0 means the whole column or row; with one row, a single number
    /// picks the column. Out of range is #REF!; only area 1 exists.
    /// </summary>
    private static Value Index(FunctionCall call)
    {
        Value source = call.Raw(0);

        if (source.IsError == true)
        {
            return source;
        }

        if (call.TryInteger(1, 0, out long row, out Value error) == false
            || call.TryInteger(2, 0, out long column, out error) == false
            || call.TryInteger(3, 1, out long area, out error) == false)
        {
            return error;
        }

        if (row < 0 || column < 0)
        {
            return Value.ValueError;
        }

        if (area != 1)
        {
            return Value.Ref;
        }

        RangeReference? reference = null;
        ValueArray? array = null;
        int rows;
        int columns;

        if (source.Kind == ValueKind.Reference)
        {
            reference = source.Reference!;
            rows = reference.Rows;
            columns = reference.Columns;
        }
        else
        {
            array = call.Evaluator.ToValueArray(source);
            rows = array.Rows;
            columns = array.Columns;
        }

        if (call.IsMissing(2) == true && rows == 1)
        {
            column = row;
            row = 1;
        }

        if (row > rows || column > columns)
        {
            return Value.Ref;
        }

        int firstRow = 0;
        int lastRow = rows - 1;
        int firstColumn = 0;
        int lastColumn = columns - 1;

        if (row > 0)
        {
            firstRow = (int)row - 1;
            lastRow = firstRow;
        }

        if (column > 0)
        {
            firstColumn = (int)column - 1;
            lastColumn = firstColumn;
        }

        if (reference is not null)
        {
            return Value.FromReference(new RangeReference(reference.Sheet, reference.FirstRow + firstRow, reference.FirstColumn + firstColumn, reference.FirstRow + lastRow, reference.FirstColumn + lastColumn));
        }

        return Slice(array!, firstRow, lastRow, firstColumn, lastColumn);
    }

    /// <summary>CHOOSE(index, value1, …): the chosen argument as it is, so a reference stays a reference.</summary>
    private static Value Choose(FunctionCall call)
    {
        if (call.IsRangeOrArray(0) == true)
        {
            return ChooseEach(call);
        }

        if (call.TryInteger(0, out long index, out Value error) == false)
        {
            return error;
        }

        if (index < 1 || index >= call.Count)
        {
            return Value.ValueError;
        }

        return call.Raw((int)index);
    }

    /// <summary>
    /// CHOOSE({1,2}, B2:B9, A2:A9): one choice per index, put together into one array, with a
    /// single row or column stretched to fit as in other array arithmetic. This is the usual way
    /// to make VLOOKUP look to the left: the columns come out in the order the indexes give.
    /// </summary>
    private static Value ChooseEach(FunctionCall call)
    {
        if (call.TryArray(0, out ValueArray indexes, out Value error) == false)
        {
            return error;
        }

        Dictionary<int, ValueArray> chosen = [];
        int rows = indexes.Rows;
        int columns = indexes.Columns;

        foreach (Value item in indexes.All())
        {
            Value number = Coerce.ToNumber(item, call.Conversion);

            if (number.IsError == true)
            {
                return number;
            }

            int index = (int)Math.Truncate(number.Number);

            if (index < 1 || index >= call.Count)
            {
                return Value.ValueError;
            }

            if (chosen.ContainsKey(index) == false)
            {
                if (call.TryArray(index, out ValueArray values, out error) == false)
                {
                    return error;
                }

                chosen[index] = values;
                rows = Math.Max(rows, values.Rows);
                columns = Math.Max(columns, values.Columns);
            }
        }

        ValueArray result = new(rows, columns);

        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                result[row, column] = Value.NA;

                if (TryStretched(indexes, row, column, out Value item) == false)
                {
                    continue;
                }

                int index = (int)Math.Truncate(Coerce.ToNumber(item, call.Conversion).Number);

                if (TryStretched(chosen[index], row, column, out Value value) == true)
                {
                    result[row, column] = value;
                }
            }
        }

        return Value.FromArray(result);
    }

    /// <summary>The item at (row, column) of an array with a single row or column stretched to the full size. False outside it, which Excel shows as #N/A.</summary>
    private static bool TryStretched(ValueArray array, int row, int column, out Value value)
    {
        int sourceRow = row;
        int sourceColumn = column;
        value = Value.Blank;

        if (array.Rows == 1)
        {
            sourceRow = 0;
        }

        if (array.Columns == 1)
        {
            sourceColumn = 0;
        }

        if (sourceRow >= array.Rows || sourceColumn >= array.Columns)
        {
            return false;
        }

        value = array[sourceRow, sourceColumn];

        return true;
    }

    private static Value Size(FunctionCall call, bool rows)
    {
        Value source = call.Raw(0);

        if (source.IsError == true)
        {
            return source;
        }

        int count = 1;

        if (source.Kind == ValueKind.Reference)
        {
            count = source.Reference!.Columns;

            if (rows == true)
            {
                count = source.Reference.Rows;
            }
        }
        else if (source.Kind == ValueKind.Array)
        {
            count = source.Array!.Columns;

            if (rows == true)
            {
                count = source.Array.Rows;
            }
        }

        return Value.FromNumber(count);
    }

    /// <summary>ADDRESS(row, column, [abs_num], [a1], [sheet]): abs_num 1 $A$1, 2 A$1, 3 $A1, 4 A1; a1 FALSE gives R1C1 text such as R[2]C3.</summary>
    private static Value Address(FunctionCall call)
    {
        if (call.TryInteger(0, out long row, out Value error) == false
            || call.TryInteger(1, out long column, out error) == false
            || call.TryInteger(2, 1, out long absolute, out error) == false
            || call.TryBool(3, true, out bool a1, out error) == false
            || call.TryText(4, string.Empty, out string sheet, out error) == false)
        {
            return error;
        }

        if (row < 1 || row > CellAddress.MaxRows || column < 1 || column > CellAddress.MaxColumns || absolute < 1 || absolute > 4)
        {
            return Value.ValueError;
        }

        bool absoluteRow = absolute is 1 or 2;
        bool absoluteColumn = absolute is 1 or 3;
        string rowText = row.ToString(CultureInfo.InvariantCulture);
        string columnText;
        string text;

        if (a1 == true)
        {
            columnText = CellAddress.ColumnName((int)column - 1);

            if (absoluteColumn == true)
            {
                columnText = "$" + columnText;
            }

            if (absoluteRow == true)
            {
                rowText = "$" + rowText;
            }

            text = columnText + rowText;
        }
        else
        {
            columnText = column.ToString(CultureInfo.InvariantCulture);

            if (absoluteRow == false)
            {
                rowText = "[" + rowText + "]";
            }

            if (absoluteColumn == false)
            {
                columnText = "[" + columnText + "]";
            }

            text = "R" + rowText + "C" + columnText;
        }

        if (call.IsMissing(4) == false)
        {
            if (sheet.Length == 0)
            {
                return Value.FromText("!" + text);
            }

            return Value.FromText(FormulaLexer.QuoteSheet(sheet) + "!" + text);
        }

        return Value.FromText(text);
    }

    private static Value Transpose(FunctionCall call)
    {
        if (call.TryArray(0, out ValueArray source, out Value error) == false)
        {
            return error;
        }

        ValueArray result = new(source.Columns, source.Rows);

        for (int row = 0; row < source.Rows; row++)
        {
            for (int column = 0; column < source.Columns; column++)
            {
                result[column, row] = source[row, column];
            }
        }

        return Value.FromArray(result);
    }

    /// <summary>AREAS(reference): the engine's references are single rectangles, so always 1.</summary>
    private static Value Areas(FunctionCall call)
    {
        if (call.TryReference(0, out _, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(1);
    }

    /// <summary>HYPERLINK(link, [name]): the cell shows the name, or the link when there is none.</summary>
    private static Value Hyperlink(FunctionCall call)
    {
        if (call.TryText(0, out string link, out Value error) == false)
        {
            return error;
        }

        if (call.IsMissing(1) == true)
        {
            return Value.FromText(link);
        }

        return call.Scalar(1);
    }

    /// <summary>FORMULATEXT(reference): the formula of the reference's first cell with its "=", or #N/A when it holds none.</summary>
    private static Value FormulaText(FunctionCall call)
    {
        if (call.TryReference(0, out RangeReference reference, out Value error) == false)
        {
            return error;
        }

        string? formula = reference.Sheet[reference.FirstRow, reference.FirstColumn].Formula;

        if (formula is null)
        {
            return Value.NA;
        }

        return Value.FromText("=" + formula);
    }

    private static bool TryModes(FunctionCall call, int index, out long matchMode, out long searchMode, out Value error)
    {
        searchMode = 1;

        if (call.TryInteger(index, 0, out matchMode, out error) == false || call.TryInteger(index + 1, 1, out searchMode, out error) == false)
        {
            return false;
        }

        if (matchMode is < -1 or > 2 || searchMode is not (1 or -1 or 2 or -2))
        {
            error = Value.ValueError;

            return false;
        }

        return true;
    }

    /// <summary>The search of MATCH, VLOOKUP, HLOOKUP and LOOKUP: binary for approximate, linear for exact.</summary>
    private static int FindLegacy(Value lookup, Value[] items, bool approximate, bool wildcards)
    {
        if (approximate == true)
        {
            return LastAtMost(lookup, items, 1);
        }

        bool hasPattern = TryPattern(lookup, wildcards, out Regex pattern);

        for (int index = 0; index < items.Length; index++)
        {
            if (IsExact(lookup, items[index], hasPattern, pattern) == true)
            {
                return index;
            }
        }

        return -1;
    }

    /// <summary>The search of XLOOKUP and XMATCH.</summary>
    private static int Find(Value lookup, Value[] items, long matchMode, long searchMode)
    {
        if (searchMode is 2 or -2)
        {
            int sign = 1;

            if (searchMode == -2)
            {
                sign = -1;
            }

            if (matchMode == -1)
            {
                if (sign == 1)
                {
                    return LastAtMost(lookup, items, sign);
                }

                return FirstAtLeast(lookup, items, sign);
            }

            if (matchMode == 1)
            {
                if (sign == 1)
                {
                    return FirstAtLeast(lookup, items, sign);
                }

                return LastAtMost(lookup, items, sign);
            }

            int candidate = FirstAtLeast(lookup, items, sign);

            if (candidate >= 0 && Coerce.Compare(items[candidate], lookup) == 0)
            {
                return candidate;
            }

            return -1;
        }

        bool hasPattern = TryPattern(lookup, matchMode == 2, out Regex pattern);
        int best = -1;

        for (int step = 0; step < items.Length; step++)
        {
            int index = step;

            if (searchMode == -1)
            {
                index = items.Length - 1 - step;
            }

            Value item = items[index];

            if (IsExact(lookup, item, hasPattern, pattern) == true)
            {
                return index;
            }

            if (matchMode is -1 or 1 && IsComparable(lookup, item) == true)
            {
                int comparison = Coerce.Compare(item, lookup) * (int)matchMode;

                // Next smaller (-1) or next larger (1): the closest on the wanted side, the first found on ties.
                if (comparison > 0 && (best < 0 || Coerce.Compare(item, items[best]) * (int)matchMode < 0))
                {
                    best = index;
                }
            }
        }

        return best;
    }

    /// <summary>
    /// A binary search for the last item ≤ the lookup value (sign 1, ascending order) or ≥ it
    /// (sign -1, descending order), among the items of the lookup value's kind.
    /// </summary>
    private static int LastAtMost(Value lookup, Value[] items, int sign)
    {
        List<int> positions = Comparable(lookup, items);
        int low = 0;
        int high = positions.Count - 1;
        int found = -1;

        while (low <= high)
        {
            int middle = low + ((high - low) / 2);

            if (sign * Coerce.Compare(items[positions[middle]], lookup) <= 0)
            {
                found = positions[middle];
                low = middle + 1;
            }
            else
            {
                high = middle - 1;
            }
        }

        return found;
    }

    /// <summary>A binary search for the first item ≥ the lookup value (sign 1) or ≤ it (sign -1).</summary>
    private static int FirstAtLeast(Value lookup, Value[] items, int sign)
    {
        List<int> positions = Comparable(lookup, items);
        int low = 0;
        int high = positions.Count - 1;
        int found = -1;

        while (low <= high)
        {
            int middle = low + ((high - low) / 2);

            if (sign * Coerce.Compare(items[positions[middle]], lookup) >= 0)
            {
                found = positions[middle];
                high = middle - 1;
            }
            else
            {
                low = middle + 1;
            }
        }

        return found;
    }

    private static List<int> Comparable(Value lookup, Value[] items)
    {
        List<int> positions = [];

        for (int index = 0; index < items.Length; index++)
        {
            if (IsComparable(lookup, items[index]) == true)
            {
                positions.Add(index);
            }
        }

        return positions;
    }

    private static bool IsComparable(Value lookup, Value item)
    {
        return item.Kind == lookup.Kind && item.Kind is ValueKind.Number or ValueKind.Text or ValueKind.Bool;
    }

    private static bool TryPattern(Value lookup, bool wildcards, out Regex pattern)
    {
        pattern = null!;

        if (wildcards == false || lookup.Kind != ValueKind.Text || Criteria.HasWildcard(lookup.Text!) == false)
        {
            return false;
        }

        pattern = new Regex("^" + Criteria.WildcardToRegex(lookup.Text!) + "$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, TimeSpan.FromSeconds(1));

        return true;
    }

    private static bool IsExact(Value lookup, Value item, bool hasPattern, Regex pattern)
    {
        if (hasPattern == true)
        {
            return item.Kind == ValueKind.Text && pattern.IsMatch(item.Text!) == true;
        }

        if (lookup.IsBlank == true)
        {
            return item.IsBlank == true;
        }

        return IsComparable(lookup, item) == true && Coerce.Compare(item, lookup) == 0;
    }

    /// <summary>The items of a one-row or one-column array. False for a wider array.</summary>
    private static bool TryVector(ValueArray array, out Value[] items)
    {
        if (array.Columns == 1)
        {
            items = Column(array, 0);

            return true;
        }

        if (array.Rows == 1)
        {
            items = Row(array, 0);

            return true;
        }

        items = [];

        return false;
    }

    private static Value[] Column(ValueArray array, int column)
    {
        Value[] items = new Value[array.Rows];

        for (int row = 0; row < array.Rows; row++)
        {
            items[row] = array[row, column];
        }

        return items;
    }

    private static Value[] Row(ValueArray array, int row)
    {
        Value[] items = new Value[array.Columns];

        for (int column = 0; column < array.Columns; column++)
        {
            items[column] = array[row, column];
        }

        return items;
    }

    /// <summary>Part of an array: a single item stays a scalar.</summary>
    private static Value Slice(ValueArray array, int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        if (firstRow == lastRow && firstColumn == lastColumn)
        {
            return array[firstRow, firstColumn];
        }

        ValueArray result = new(lastRow - firstRow + 1, lastColumn - firstColumn + 1);

        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                result[row - firstRow, column - firstColumn] = array[row, column];
            }
        }

        return Value.FromArray(result);
    }

    /// <summary>The rows of a reference as <see cref="Evaluator.ToValueArray"/> reads them: whole columns stop at the last used row.</summary>
    private static int UsedRows(RangeReference range)
    {
        if (range.LastRow < CellAddress.MaxRows - 1)
        {
            return range.Rows;
        }

        return Math.Max(0, range.UsedLastRow - range.FirstRow + 1);
    }

    private static int UsedColumns(RangeReference range)
    {
        if (range.LastColumn < CellAddress.MaxColumns - 1)
        {
            return range.Columns;
        }

        return Math.Max(0, range.UsedLastColumn - range.FirstColumn + 1);
    }
}
