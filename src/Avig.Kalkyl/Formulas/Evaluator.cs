using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Formulas;

/// <summary>Reads a formula cell that may still be waiting for its new value.</summary>
internal delegate bool PendingCellReader(Worksheet sheet, int row, int column, out Value value);

/// <summary>Where a formula is evaluated: the sheet and cell it belongs to (row and column 0-based).</summary>
internal readonly record struct EvaluationContext(Worksheet Sheet, int Row, int Column);

/// <summary>
/// Evaluates parsed formulas against a workbook. Cell values come from the document model; a
/// formula cell that the current recalculation has not reached yet is evaluated on demand
/// through <see cref="PendingCell"/>.
/// </summary>
internal sealed class Evaluator
{
    private const int MaxNameDepth = 32;

    private readonly Dictionary<(string Name, int Scope), Node?> nameCache = [];
    private int nameDepth;

    public Evaluator(Workbook workbook, DateTime now)
    {
        Workbook = workbook;
        Now = now;
    }

    public Workbook Workbook { get; }

    /// <summary>The moment NOW() and TODAY() return, fixed for one recalculation so every cell agrees.</summary>
    public DateTime Now { get; }

    public bool Uses1904 => Workbook.Uses1904Dates;

    /// <summary>Asked for a formula cell's value when it may not be calculated yet; false means "use the stored value".</summary>
    public PendingCellReader? PendingCell { get; set; }

    /// <summary>Functions that were called but do not exist; they evaluate to #NAME?.</summary>
    public HashSet<string> UnknownFunctions { get; } = new(StringComparer.Ordinal);

    public Value Evaluate(Node node, EvaluationContext context)
    {
        switch (node)
        {
            case LiteralNode literal:
                return literal.Value;

            case ReferenceNode reference:
                return Resolve(reference.Reference, context);

            case NameNode name:
                return EvaluateName(name.Name, context);

            case FunctionNode function:
                return Call(function, context);

            case UnaryNode unary:
                return Unary(unary, context);

            case PercentNode percent:
                return Map(Evaluate(percent.Operand, context), context, value => Arithmetic(value, Value.FromNumber(100), "/"));

            case BinaryNode binary:
                return Binary(binary, context);

            case ArrayNode array:
                return ToArray(array);

            case MissingNode:
                return Value.Blank;

            default:
                return Value.ValueError;
        }
    }

    /// <summary>A cell's value as the formula engine sees it: dates are serial numbers.</summary>
    public Value CellValue(Worksheet sheet, int row, int column)
    {
        Cell cell = sheet[row, column];

        if (cell.Formula is not null && PendingCell is not null && PendingCell(sheet, row, column, out Value pending) == true)
        {
            return pending;
        }

        return FromCellValue(cell.Value);
    }

    public Value FromCellValue(CellValue value)
    {
        switch (value.Kind)
        {
            case CellKind.Number:
                value.TryGetNumber(out double number);

                return Value.FromNumber(number);

            case CellKind.Text:
                value.TryGetText(out string? text);

                return Value.FromText(text!);

            case CellKind.Bool:
                value.TryGetBool(out bool flag);

                return Value.FromBool(flag);

            case CellKind.Date:
                value.TryGetDate(out DateTime date);

                return Value.FromNumber(ExcelDate.ToSerial(date, Uses1904));

            case CellKind.Error:
                value.TryGetError(out string? code);

                return Value.Error(code!);

            default:
                return Value.Blank;
        }
    }

    /// <summary>
    /// Turns a reference into a single value. A one-cell reference gives that cell. A wider one
    /// gives the cell in the formula's own row or column (Excel's implicit intersection), or
    /// #VALUE! when there is none. An array gives its first item.
    /// </summary>
    public Value Dereference(Value value, EvaluationContext context)
    {
        if (value.Kind == ValueKind.Array)
        {
            if (value.Array!.Rows > 0 && value.Array.Columns > 0)
            {
                return value.Array[0, 0];
            }

            return Value.ValueError;
        }

        if (value.Kind != ValueKind.Reference)
        {
            return value;
        }

        RangeReference range = value.Reference!;

        if (range.IsSingleCell == true)
        {
            return CellValue(range.Sheet, range.FirstRow, range.FirstColumn);
        }

        bool sameSheet = ReferenceEquals(range.Sheet, context.Sheet);

        if (sameSheet == true && range.FirstColumn == range.LastColumn && context.Row >= range.FirstRow && context.Row <= range.LastRow)
        {
            return CellValue(range.Sheet, context.Row, range.FirstColumn);
        }

        if (sameSheet == true && range.FirstRow == range.LastRow && context.Column >= range.FirstColumn && context.Column <= range.LastColumn)
        {
            return CellValue(range.Sheet, range.FirstRow, context.Column);
        }

        return Value.ValueError;
    }

    /// <summary>A reference or array as an array of values; a scalar becomes a 1×1 array. Whole columns stop at the last used row.</summary>
    public ValueArray ToValueArray(Value value)
    {
        if (value.Kind == ValueKind.Array)
        {
            return value.Array!;
        }

        if (value.Kind == ValueKind.Reference)
        {
            RangeReference range = value.Reference!;
            int rows = Math.Max(0, range.UsedLastRow - range.FirstRow + 1);
            int columns = Math.Max(0, range.UsedLastColumn - range.FirstColumn + 1);

            if (range.LastRow < CellAddress.MaxRows - 1)
            {
                rows = range.Rows;
            }

            if (range.LastColumn < CellAddress.MaxColumns - 1)
            {
                columns = range.Columns;
            }

            ValueArray array = new(rows, columns);

            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    array[row, column] = CellValue(range.Sheet, range.FirstRow + row, range.FirstColumn + column);
                }
            }

            return array;
        }

        ValueArray single = new(1, 1);
        single[0, 0] = value;

        return single;
    }

    public bool TryGetSheet(string? name, Worksheet current, out Worksheet sheet)
    {
        sheet = current;

        if (name is null)
        {
            return true;
        }

        Worksheet? found = Workbook.Sheets.FirstOrDefault(candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) == true);

        if (found is null)
        {
            return false;
        }

        sheet = found;

        return true;
    }

    /// <summary>The value of a parsed defined name's formula, or #NAME? when the name does not exist.</summary>
    public Value EvaluateName(string name, EvaluationContext context)
    {
        if (TryGetName(name, context.Sheet, out Node? node) == false)
        {
            return Value.Name;
        }

        if (nameDepth >= MaxNameDepth)
        {
            return Value.Name;
        }

        nameDepth++;

        try
        {
            return Evaluate(node, context);
        }
        finally
        {
            nameDepth--;
        }
    }

    /// <summary>A defined name's parsed formula. Names local to the sheet win over workbook names.</summary>
    public bool TryGetName(string name, Worksheet sheet, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Node? node)
    {
        node = null;

        foreach (int scope in new[] { sheet.Index, -1 })
        {
            if (nameCache.TryGetValue((name.ToUpperInvariant(), scope), out Node? cached) == false)
            {
                DefinedName? definition = Workbook.DefinedNames.Find(candidate => candidate.Scope == scope && string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase) == true);
                cached = null;

                if (definition is not null && FormulaParser.Parse(definition.Formula).IsSuccess(out Node? parsed) == true)
                {
                    cached = parsed;
                }

                nameCache[(name.ToUpperInvariant(), scope)] = cached;
            }

            if (cached is not null)
            {
                node = cached;

                return true;
            }
        }

        return false;
    }

    private Value Resolve(ReferenceToken reference, EvaluationContext context)
    {
        if (TryGetSheet(reference.Sheet, context.Sheet, out Worksheet sheet) == false)
        {
            return Value.Ref;
        }

        int firstRow = reference.FirstRow;
        int lastRow = reference.LastRow;
        int firstColumn = reference.FirstColumn;
        int lastColumn = reference.LastColumn;

        if (reference.Kind == ReferenceKind.Columns)
        {
            firstRow = 0;
            lastRow = CellAddress.MaxRows - 1;
        }
        else if (reference.Kind == ReferenceKind.Rows)
        {
            firstColumn = 0;
            lastColumn = CellAddress.MaxColumns - 1;
        }

        return Value.FromReference(new RangeReference(sheet, Math.Min(firstRow, lastRow), Math.Min(firstColumn, lastColumn), Math.Max(firstRow, lastRow), Math.Max(firstColumn, lastColumn)));
    }

    private Value Call(FunctionNode function, EvaluationContext context)
    {
        if (FunctionLibrary.TryGet(function.Name, out FunctionDefinition? definition) == false)
        {
            UnknownFunctions.Add(function.Name);

            return Value.Name;
        }

        if (function.Arguments.Count < definition.MinArguments || function.Arguments.Count > definition.MaxArguments)
        {
            return Value.ValueError;
        }

        return definition.Body(new FunctionCall(this, context, function.Arguments));
    }

    private Value Unary(UnaryNode unary, EvaluationContext context)
    {
        Value operand = Evaluate(unary.Operand, context);

        if (unary.Operator == "+")
        {
            return operand;
        }

        return Map(operand, context, value => Arithmetic(Value.FromNumber(0), value, "-"));
    }

    private Value Binary(BinaryNode binary, EvaluationContext context)
    {
        Value left = Evaluate(binary.Left, context);
        Value right = Evaluate(binary.Right, context);

        if (binary.Operator == ":")
        {
            return RangeUnion(left, right);
        }

        bool leftWide = IsWide(left);
        bool rightWide = IsWide(right);

        if (leftWide == false && rightWide == false)
        {
            return Scalar(binary.Operator, Dereference(left, context), Dereference(right, context));
        }

        // Element by element, as in an array formula: SUMPRODUCT((A1:A9>0)*B1:B9) needs it.
        ValueArray leftArray = ToValueArray(left);
        ValueArray rightArray = ToValueArray(right);
        int rows = Math.Max(leftArray.Rows, rightArray.Rows);
        int columns = Math.Max(leftArray.Columns, rightArray.Columns);
        ValueArray result = new(rows, columns);

        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                Value leftItem = Broadcast(leftArray, row, column);
                Value rightItem = Broadcast(rightArray, row, column);
                result[row, column] = Scalar(binary.Operator, leftItem, rightItem);
            }
        }

        return Value.FromArray(result);
    }

    private Value Scalar(string op, Value left, Value right)
    {
        switch (op)
        {
            case "+":
            case "-":
            case "*":
            case "/":
            case "^":
                return Arithmetic(left, right, op);

            case "&":
                Value leftText = Coerce.ToText(left);

                if (leftText.IsError == true)
                {
                    return leftText;
                }

                Value rightText = Coerce.ToText(right);

                if (rightText.IsError == true)
                {
                    return rightText;
                }

                return Value.FromText(leftText.Text + rightText.Text);

            default:
                if (left.IsError == true)
                {
                    return left;
                }

                if (right.IsError == true)
                {
                    return right;
                }

                int comparison = Coerce.Compare(left, right);

                return Value.FromBool(op switch
                {
                    "=" => comparison == 0,
                    "<>" => comparison != 0,
                    "<" => comparison < 0,
                    ">" => comparison > 0,
                    "<=" => comparison <= 0,
                    _ => comparison >= 0,
                });
        }
    }

    private Value Arithmetic(Value left, Value right, string op)
    {
        Value leftNumber = Coerce.ToNumber(left, Uses1904);

        if (leftNumber.IsError == true)
        {
            return leftNumber;
        }

        Value rightNumber = Coerce.ToNumber(right, Uses1904);

        if (rightNumber.IsError == true)
        {
            return rightNumber;
        }

        double first = leftNumber.Number;
        double second = rightNumber.Number;

        switch (op)
        {
            case "+":
                return Value.FromNumber(first + second);

            case "-":
                return Value.FromNumber(first - second);

            case "*":
                return Value.FromNumber(first * second);

            case "/":
                if (second == 0)
                {
                    return Value.Div0;
                }

                return Value.FromNumber(first / second);

            default:
                if (first == 0 && second == 0)
                {
                    return Value.Num;
                }

                if (first == 0 && second < 0)
                {
                    return Value.Div0;
                }

                return Value.FromNumber(Math.Pow(first, second));
        }
    }

    private Value Map(Value value, EvaluationContext context, Func<Value, Value> map)
    {
        if (IsWide(value) == false)
        {
            return map(Dereference(value, context));
        }

        ValueArray source = ToValueArray(value);
        ValueArray result = new(source.Rows, source.Columns);

        for (int row = 0; row < source.Rows; row++)
        {
            for (int column = 0; column < source.Columns; column++)
            {
                result[row, column] = map(source[row, column]);
            }
        }

        return Value.FromArray(result);
    }

    private static Value RangeUnion(Value left, Value right)
    {
        if (left.IsError == true)
        {
            return left;
        }

        if (right.IsError == true)
        {
            return right;
        }

        if (left.Kind != ValueKind.Reference || right.Kind != ValueKind.Reference || ReferenceEquals(left.Reference!.Sheet, right.Reference!.Sheet) == false)
        {
            return Value.ValueError;
        }

        RangeReference first = left.Reference;
        RangeReference second = right.Reference;

        return Value.FromReference(new RangeReference(
            first.Sheet,
            Math.Min(first.FirstRow, second.FirstRow),
            Math.Min(first.FirstColumn, second.FirstColumn),
            Math.Max(first.LastRow, second.LastRow),
            Math.Max(first.LastColumn, second.LastColumn)));
    }

    private static bool IsWide(Value value)
    {
        return value.Kind == ValueKind.Array || (value.Kind == ValueKind.Reference && value.Reference!.IsSingleCell == false);
    }

    private static Value Broadcast(ValueArray array, int row, int column)
    {
        int sourceRow = row;
        int sourceColumn = column;

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
            return Value.NA;
        }

        return array[sourceRow, sourceColumn];
    }

    private static Value ToArray(ArrayNode node)
    {
        ValueArray array = new(node.Rows.Count, node.Rows[0].Count);

        for (int row = 0; row < node.Rows.Count; row++)
        {
            for (int column = 0; column < node.Rows[row].Count; column++)
            {
                array[row, column] = node.Rows[row][column];
            }
        }

        return Value.FromArray(array);
    }
}

/// <summary>A defined name from the workbook. Scope is the index of the sheet it belongs to, or -1 for the whole workbook.</summary>
internal sealed record DefinedName(string Name, int Scope, string Formula);
