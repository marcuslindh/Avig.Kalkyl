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

    private Recalculator(Workbook book, DateTime now)
    {
        this.book = book;
        evaluator = new Evaluator(book, now);
        evaluator.PendingCell = TryReadPending;
    }

    public static RecalculationReport Run(Workbook book, RecalculationScope scope, DateTime now, CancellationToken cancellationToken)
    {
        Recalculator recalculator = new(book, now);

        return recalculator.Run(scope, cancellationToken);
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
                CollectPrecedents(binary.Left, context, ranges, ref isVolatile, depth);
                CollectPrecedents(binary.Right, context, ranges, ref isVolatile, depth);

                if (binary.Operator == ":")
                {
                    Value span = evaluator.Evaluate(binary, context);

                    if (span.Kind == ValueKind.Reference)
                    {
                        ranges.Add(span.Reference!);
                    }
                }

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

    /// <summary>The formula cells of one sheet, sorted by row, so the cells inside a range are found without visiting every formula.</summary>
    private sealed class SheetIndex
    {
        private readonly SortedDictionary<int, List<FormulaCell>> rows = [];
        private int[]? keys;

        public void Add(FormulaCell cell)
        {
            if (rows.TryGetValue(cell.Row, out List<FormulaCell>? row) == false)
            {
                row = [];
                rows[cell.Row] = row;
            }

            row.Add(cell);
            keys = null;
        }

        public IEnumerable<FormulaCell> Within(RangeReference range)
        {
            keys ??= rows.Keys.ToArray();
            int start = Array.BinarySearch(keys, range.FirstRow);

            if (start < 0)
            {
                start = ~start;
            }

            for (int index = start; index < keys.Length && keys[index] <= range.LastRow; index++)
            {
                foreach (FormulaCell cell in rows[keys[index]])
                {
                    if (cell.Column >= range.FirstColumn && cell.Column <= range.LastColumn)
                    {
                        yield return cell;
                    }
                }
            }
        }
    }
}
