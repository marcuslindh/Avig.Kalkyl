namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// A range or array argument of the *IF and *IFS functions, read cell by cell. A whole column or
/// row is only visited up to the sheet's used area; the rest is known to be blank.
/// </summary>
internal sealed class StatisticalFunctionsGrid
{
    private readonly Evaluator evaluator;
    private readonly RangeReference? reference;
    private readonly ValueArray? array;

    private StatisticalFunctionsGrid(Evaluator evaluator, RangeReference? reference, ValueArray? array)
    {
        this.evaluator = evaluator;
        this.reference = reference;
        this.array = array;

        if (reference is not null)
        {
            Rows = reference.Rows;
            Columns = reference.Columns;
            UsedRows = Rows;
            UsedColumns = Columns;

            if (reference.LastRow >= CellAddress.MaxRows - 1)
            {
                UsedRows = Math.Clamp(reference.UsedLastRow - reference.FirstRow + 1, 0, Rows);
            }

            if (reference.LastColumn >= CellAddress.MaxColumns - 1)
            {
                UsedColumns = Math.Clamp(reference.UsedLastColumn - reference.FirstColumn + 1, 0, Columns);
            }

            return;
        }

        Rows = array!.Rows;
        Columns = array.Columns;
        UsedRows = Rows;
        UsedColumns = Columns;
    }

    public int Rows { get; }

    public int Columns { get; }

    /// <summary>The rows worth visiting; beyond them every cell is blank.</summary>
    public int UsedRows { get; }

    /// <summary>The columns worth visiting; beyond them every cell is blank.</summary>
    public int UsedColumns { get; }

    /// <summary>The number of cells in the grid.</summary>
    public long CellCount => (long)Rows * Columns;

    public static bool TryCreate(FunctionCall call, int index, out StatisticalFunctionsGrid grid, out Value error)
    {
        Value raw = call.Raw(index);
        grid = null!;
        error = Value.Blank;

        if (raw.IsError == true)
        {
            error = raw;

            return false;
        }

        if (raw.Kind == ValueKind.Reference)
        {
            grid = new StatisticalFunctionsGrid(call.Evaluator, raw.Reference!, null);

            return true;
        }

        if (raw.Kind == ValueKind.Array)
        {
            grid = new StatisticalFunctionsGrid(call.Evaluator, null, raw.Array!);

            return true;
        }

        ValueArray single = new(1, 1);
        single[0, 0] = raw;
        grid = new StatisticalFunctionsGrid(call.Evaluator, null, single);

        return true;
    }

    public bool SameShape(StatisticalFunctionsGrid other)
    {
        return Rows == other.Rows && Columns == other.Columns;
    }

    /// <summary>The grid with the size of <paramref name="shape"/>, from the same top-left cell, as AVERAGEIF reads its average_range.</summary>
    public StatisticalFunctionsGrid Resized(StatisticalFunctionsGrid shape)
    {
        if (reference is null || SameShape(shape) == true)
        {
            return this;
        }

        int lastRow = Math.Min(reference.FirstRow + shape.Rows - 1, CellAddress.MaxRows - 1);
        int lastColumn = Math.Min(reference.FirstColumn + shape.Columns - 1, CellAddress.MaxColumns - 1);

        return new StatisticalFunctionsGrid(evaluator, reference with { LastRow = lastRow, LastColumn = lastColumn }, null);
    }

    /// <summary>True when the first <paramref name="usedRows"/> x <paramref name="usedColumns"/> of a range are mostly empty positions (see Evaluator.PreferStored).</summary>
    public bool IsSparse(int usedRows, int usedColumns)
    {
        return reference is not null && Evaluator.PreferStored(reference.Sheet, (long)usedRows * usedColumns);
    }

    /// <summary>The cells the sheet holds in the first <paramref name="usedRows"/> x <paramref name="usedColumns"/>, with positions counted from the top-left cell, row by row.</summary>
    public IEnumerable<(int Row, int Column, Value Value)> StoredValues(int usedRows, int usedColumns)
    {
        RangeReference range = reference!;

        foreach ((int row, int column, Value value) in evaluator.StoredValues(range.Sheet, range.FirstRow, range.FirstRow + usedRows - 1, range.FirstColumn, range.FirstColumn + usedColumns - 1))
        {
            yield return (row - range.FirstRow, column - range.FirstColumn, value);
        }
    }

    /// <summary>Stops a long walk over the grid when the evaluation is cancelled.</summary>
    public CancellationToken CancellationToken => evaluator.CancellationToken;

    /// <summary>The value at a position counted from the top-left cell; blank outside the grid.</summary>
    public Value Get(int row, int column)
    {
        if (reference is not null)
        {
            if (row >= reference.Rows || column >= reference.Columns)
            {
                return Value.Blank;
            }

            return evaluator.CellValue(reference.Sheet, reference.FirstRow + row, reference.FirstColumn + column);
        }

        if (row >= array!.Rows || column >= array.Columns)
        {
            return Value.Blank;
        }

        return array[row, column];
    }
}
