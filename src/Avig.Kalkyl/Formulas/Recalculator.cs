using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// Recalculates a workbook. Every formula is parsed once, its precedents (the cells and ranges
/// it reads, also through defined names) become edges in a dependency graph, and the formulas
/// that need calculating are evaluated in dependency order. A formula that the graph cannot
/// place — reached through INDIRECT or OFFSET — is evaluated on demand when it is read.
/// Formulas left over when no more can be placed are in a cycle; they get 0, as in Excel
/// without iterative calculation, and are listed in the report.
/// </summary>
internal sealed class Recalculator
{
    private readonly Workbook book;
    private readonly Evaluator evaluator;
    private readonly Dictionary<(Worksheet Sheet, int Row, int Column), FormulaCell> cells = [];
    private readonly Dictionary<Worksheet, SheetIndex> indexes = [];
    private readonly Dictionary<string, Result<Node>> parsed = new(StringComparer.Ordinal);
    private readonly List<string> circular = [];
    private readonly List<string> invalid = [];

    private Recalculator(Workbook book, DateTime now, CancellationToken cancellationToken)
    {
        this.book = book;
        evaluator = new Evaluator(book, now) { CancellationToken = cancellationToken };

        evaluator.PendingCell = TryReadPending;
    }

    public static RecalculationReport Run(Workbook book, RecalculationScope scope, DateTime now, CancellationToken cancellationToken)
    {
        Recalculator recalculator = new(book, now, cancellationToken);

        // The whole recalculation runs on this thread, so the array limit holds for all of it.
        using (ValueArray.LimitTo(book.Limits.MaxArrayCells))
        {
            return recalculator.Run(scope, cancellationToken);
        }
    }

    private RecalculationReport Run(RecalculationScope scope, CancellationToken cancellationToken)
    {
        Collect();
        Connect();
        List<FormulaCell> affected = FindAffected(scope);
        int calculated = 0;
        Queue<FormulaCell> ready = new();

        foreach (FormulaCell cell in affected)
        {
            cell.Waiting = cell.Precedents.Count(precedent => precedent.Affected == true && ReferenceEquals(precedent, cell) == false);

            if (cell.Waiting == 0)
            {
                ready.Enqueue(cell);
            }
        }

        while (ready.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            FormulaCell cell = ready.Dequeue();

            if (cell.State != CellState.Done)
            {
                Calculate(cell);
            }

            foreach (FormulaCell dependent in cell.Dependents)
            {
                if (dependent.Affected == false || ReferenceEquals(dependent, cell) == true)
                {
                    continue;
                }

                dependent.Waiting--;

                if (dependent.Waiting == 0)
                {
                    ready.Enqueue(dependent);
                }
            }
        }

        foreach (FormulaCell cell in affected)
        {
            if (cell.State != CellState.Done)
            {
                cell.Result = Value.FromNumber(0);
                cell.State = CellState.Done;
                cell.Circular = true;
            }

            if (cell.Circular == true)
            {
                circular.Add(Describe(cell));
            }

            Store(cell);
            calculated++;
        }

        book.ChangedCells.Clear();

        return new RecalculationReport(calculated, circular.Distinct().ToList(), evaluator.UnknownFunctions.Order(StringComparer.Ordinal).ToList(), invalid);
    }

    private void Collect()
    {
        foreach (Worksheet sheet in book.Sheets)
        {
            SheetIndex index = new();
            indexes[sheet] = index;

            foreach ((int row, RowEntry entry) in sheet.RowEntries)
            {
                foreach ((int column, CellEntry cellEntry) in entry.Cells)
                {
                    string? formula = cellEntry.Cell.Formula;

                    if (formula is null)
                    {
                        continue;
                    }

                    if (parsed.TryGetValue(formula, out Result<Node>? tree) == false)
                    {
                        tree = FormulaParser.Parse(formula);
                        parsed[formula] = tree;
                    }

                    FormulaCell cell = new(sheet, row, column);

                    if (tree.IsSuccess(out Node? node) == true)
                    {
                        cell.Tree = node;
                    }
                    else
                    {
                        invalid.Add(Describe(cell) + ": " + tree.Error);
                    }

                    cells[(sheet, row, column)] = cell;
                    index.Add(cell);
                }
            }
        }
    }

    private void Connect()
    {
        foreach (FormulaCell cell in cells.Values)
        {
            if (cell.Tree is null)
            {
                continue;
            }

            List<RangeReference> ranges = [];
            bool isVolatile = false;
            CollectPrecedents(cell.Tree, new EvaluationContext(cell.Sheet, cell.Row, cell.Column), ranges, ref isVolatile, 0);
            cell.Volatile = isVolatile;
            cell.Ranges = ranges;

            foreach (RangeReference range in ranges)
            {
                foreach (FormulaCell precedent in FormulaCellsIn(range))
                {
                    if (ReferenceEquals(precedent, cell) == true)
                    {
                        cell.Circular = true;

                        continue;
                    }

                    if (precedent.Dependents.Contains(cell) == false)
                    {
                        precedent.Dependents.Add(cell);
                        cell.Precedents.Add(precedent);
                    }
                }
            }
        }
    }

    private void CollectPrecedents(Node node, EvaluationContext context, List<RangeReference> ranges, ref bool isVolatile, int depth)
    {
        switch (node)
        {
            case ReferenceNode reference:
                Value resolved = evaluator.Evaluate(reference, context);

                if (resolved.Kind == ValueKind.Reference)
                {
                    ranges.Add(resolved.Reference!);
                }

                break;

            case NameNode name when depth < 16:
                if (evaluator.TryGetName(name.Name, context.Sheet, out Node? definition) == true)
                {
                    CollectPrecedents(definition, context, ranges, ref isVolatile, depth + 1);
                }

                break;

            case FunctionNode function:
                if (FunctionLibrary.TryGet(function.Name, out FunctionDefinition? known) == true && known.Volatile == true)
                {
                    isVolatile = true;
                }

                foreach (Node argument in function.Arguments)
                {
                    CollectPrecedents(argument, context, ranges, ref isVolatile, depth);
                }

                break;

            case UnaryNode unary:
                CollectPrecedents(unary.Operand, context, ranges, ref isVolatile, depth);

                break;

            case PercentNode percent:
                CollectPrecedents(percent.Operand, context, ranges, ref isVolatile, depth);

                break;

            case BinaryNode binary:
                // The left spine of 1+1+…+1 is as long as the formula has terms, so it is walked
                // in a loop rather than recursively (see Evaluator.Binary).
                Node current = binary;

                while (current is BinaryNode chain)
                {
                    CollectPrecedents(chain.Right, context, ranges, ref isVolatile, depth);

                    if (chain.Operator == ":")
                    {
                        Value span = evaluator.Evaluate(chain, context);

                        if (span.Kind == ValueKind.Reference)
                        {
                            ranges.Add(span.Reference!);
                        }
                    }

                    current = chain.Left;
                }

                CollectPrecedents(current, context, ranges, ref isVolatile, depth);

                break;
        }
    }

    private List<FormulaCell> FindAffected(RecalculationScope scope)
    {
        List<FormulaCell> affected = [];
        Queue<FormulaCell> pending = new();

        foreach (FormulaCell cell in cells.Values)
        {
            bool seed = scope == RecalculationScope.All || cell.Volatile == true || book.ChangedCells.Contains((cell.Sheet, cell.Row, cell.Column)) == true;

            if (seed == false && book.ChangedCells.Count > 0)
            {
                seed = cell.Ranges.Any(range => book.ChangedCells.Any(changed => range.Contains(changed.Sheet, changed.Row, changed.Column) == true)) == true;
            }

            if (seed == true)
            {
                cell.Affected = true;
                pending.Enqueue(cell);
            }
        }

        while (pending.Count > 0)
        {
            FormulaCell cell = pending.Dequeue();
            affected.Add(cell);

            foreach (FormulaCell dependent in cell.Dependents)
            {
                if (dependent.Affected == false)
                {
                    dependent.Affected = true;
                    pending.Enqueue(dependent);
                }
            }
        }

        return affected;
    }

    private IEnumerable<FormulaCell> FormulaCellsIn(RangeReference range)
    {
        if (indexes.TryGetValue(range.Sheet, out SheetIndex? index) == false)
        {
            return [];
        }

        return index.Within(range);
    }

    private void Calculate(FormulaCell cell)
    {
        cell.State = CellState.InProgress;
        Value result = Value.Name;

        if (cell.Tree is not null)
        {
            EvaluationContext context = new(cell.Sheet, cell.Row, cell.Column);
            result = evaluator.Dereference(evaluator.Evaluate(cell.Tree, context), context);
        }

        if (result.IsBlank == true)
        {
            result = Value.FromNumber(0);
        }

        cell.Result = result;
        cell.State = CellState.Done;
    }

    /// <summary>A formula cell read during the recalculation: its new value, calculated now if needed.</summary>
    private bool TryReadPending(Worksheet sheet, int row, int column, out Value value)
    {
        value = Value.Blank;

        if (cells.TryGetValue((sheet, row, column), out FormulaCell? cell) == false || cell.Affected == false)
        {
            return false;
        }

        switch (cell.State)
        {
            case CellState.Done:
                value = cell.Result;

                break;

            case CellState.InProgress:
                cell.Circular = true;
                value = Value.FromNumber(0);

                break;

            default:
                Calculate(cell);
                value = cell.Result;

                break;
        }

        return true;
    }

    private void Store(FormulaCell cell)
    {
        Style style = cell.Sheet[cell.Row, cell.Column].Style;
        cell.Sheet.SetCalculated(cell.Row, cell.Column, ToCellValue(cell.Result, style, book.Uses1904Dates));
    }

    /// <summary>A result as a cell value. A number in a cell with a date format becomes a date.</summary>
    internal static CellValue ToCellValue(Value result, Style style, bool uses1904)
    {
        switch (result.Kind)
        {
            case ValueKind.Number:
                if (style.NumberFormat is not null && NumberFormats.IsDateFormat(style.NumberFormat) == true && ExcelDate.TryFromSerial(result.Number, uses1904, out DateTime date) == true)
                {
                    return CellValue.Date(date);
                }

                return CellValue.Number(result.Number);

            case ValueKind.Text when result.Text!.Length > CellValue.MaxTextLength:
                return CellValue.Error("#VALUE!");

            case ValueKind.Text:
                return CellValue.Text(result.Text!);

            case ValueKind.Bool:
                return CellValue.Bool(result.Bool);

            case ValueKind.Error:
                return CellValue.Error(result.Text!);

            default:
                return CellValue.Number(0);
        }
    }

    private static string Describe(FormulaCell cell)
    {
        return FormulaLexer.QuoteSheet(cell.Sheet.Name) + "!" + new CellAddress(cell.Row, cell.Column).ToString();
    }

    private enum CellState
    {
        NotStarted,
        InProgress,
        Done,
    }

    private sealed class FormulaCell(Worksheet sheet, int row, int column)
    {
        public Worksheet Sheet { get; } = sheet;

        public int Row { get; } = row;

        public int Column { get; } = column;

        public Node? Tree { get; set; }

        public IReadOnlyList<RangeReference> Ranges { get; set; } = [];

        public List<FormulaCell> Precedents { get; } = [];

        public List<FormulaCell> Dependents { get; } = [];

        public bool Volatile { get; set; }

        public bool Affected { get; set; }

        public bool Circular { get; set; }

        public int Waiting { get; set; }

        public CellState State { get; set; }

        public Value Result { get; set; }
    }

    /// <summary>
    /// The formula cells of one sheet, by row and by column, so that the cells inside a range are
    /// found without visiting every formula. A whole column, as in SUMIF(B:B,...), spans every
    /// row; walking the rows for it made 10 000 such formulas take 9 seconds to connect. The
    /// index takes whichever way has fewer rows or columns with formulas in the range.
    /// </summary>
    private sealed class SheetIndex
    {
        private readonly Dictionary<int, List<FormulaCell>> byRow = [];
        private readonly Dictionary<int, List<FormulaCell>> byColumn = [];
        private int[]? rowKeys;
        private List<FormulaCell>[]? rowLists;
        private int[]? columnKeys;
        private List<FormulaCell>[]? columnLists;

        public void Add(FormulaCell cell)
        {
            AddTo(byRow, cell.Row, cell);
            AddTo(byColumn, cell.Column, cell);
            rowKeys = null;
            columnKeys = null;
        }

        public IEnumerable<FormulaCell> Within(RangeReference range)
        {
            if (rowKeys is null || columnKeys is null)
            {
                (rowKeys, rowLists) = Sorted(byRow);
                (columnKeys, columnLists) = Sorted(byColumn);
            }

            (int firstRowIndex, int rowCount) = Span(rowKeys, range.FirstRow, range.LastRow);
            (int firstColumnIndex, int columnCount) = Span(columnKeys, range.FirstColumn, range.LastColumn);

            if (columnCount <= rowCount)
            {
                return Collect(columnLists!, firstColumnIndex, columnCount, cell => cell.Row, range.FirstRow, range.LastRow);
            }

            return Collect(rowLists!, firstRowIndex, rowCount, cell => cell.Column, range.FirstColumn, range.LastColumn);
        }

        private static void AddTo(Dictionary<int, List<FormulaCell>> index, int key, FormulaCell cell)
        {
            if (index.TryGetValue(key, out List<FormulaCell>? list) == false)
            {
                list = [];
                index[key] = list;
            }

            list.Add(cell);
        }

        /// <summary>The keys ascending, and each key's cells sorted by the other coordinate, so that both can be searched.</summary>
        private static (int[] Keys, List<FormulaCell>[] Lists) Sorted(Dictionary<int, List<FormulaCell>> index)
        {
            int[] keys = [.. index.Keys.Order()];
            List<FormulaCell>[] lists = new List<FormulaCell>[keys.Length];

            for (int position = 0; position < keys.Length; position++)
            {
                lists[position] = index[keys[position]];
                lists[position].Sort(static (left, right) => (left.Row, left.Column).CompareTo((right.Row, right.Column)));
            }

            return (keys, lists);
        }

        /// <summary>Where the keys from <paramref name="first"/> to <paramref name="last"/> start, and how many there are.</summary>
        private static (int Start, int Count) Span(int[] keys, int first, int last)
        {
            int start = LowerBound(keys, first);
            int end = LowerBound(keys, last + 1);

            return (start, end - start);
        }

        private static int LowerBound(int[] keys, int value)
        {
            int found = Array.BinarySearch(keys, value);

            if (found < 0)
            {
                return ~found;
            }

            return found;
        }

        private static IEnumerable<FormulaCell> Collect(List<FormulaCell>[] lists, int start, int count, Func<FormulaCell, int> other, int first, int last)
        {
            for (int position = start; position < start + count; position++)
            {
                List<FormulaCell> list = lists[position];
                int low = 0;
                int high = list.Count;

                // The list is sorted by the other coordinate (row for a column, column for a row): find where the range starts.
                while (low < high)
                {
                    int middle = (low + high) / 2;

                    if (other(list[middle]) < first)
                    {
                        low = middle + 1;
                    }
                    else
                    {
                        high = middle;
                    }
                }

                for (int index = low; index < list.Count && other(list[index]) <= last; index++)
                {
                    yield return list[index];
                }
            }
        }
    }
}
