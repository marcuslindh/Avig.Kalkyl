using System.Diagnostics.CodeAnalysis;

namespace Avig.Kalkyl;

/// <summary>
/// A row read from a sheet. Columns are 0-based; a column the file has no cell for reads as
/// <see cref="Cell.Empty"/>, so <c>row[5]</c> is always safe.
/// </summary>
public sealed class Row
{
    private readonly Cell[] cells;

    internal Row(int index, Cell[] cells)
    {
        Index = index;
        this.cells = cells;
    }

    /// <summary>The row's index, 0-based. Excel shows it as <c>Index + 1</c>. Rows the file leaves out are skipped, so indexes can jump.</summary>
    public int Index { get; }

    /// <summary>One past the last column that has a cell. 0 for a row without cells.</summary>
    public int ColumnCount => cells.Length;

    /// <summary>True when no cell in the row has a value.</summary>
    public bool IsEmpty => Array.TrueForAll(cells, cell => cell.Value.IsEmpty == true);

    /// <summary>The cells from column 0 to <see cref="ColumnCount"/> − 1, with empty cells where the file has none.</summary>
    public IReadOnlyList<Cell> Cells => cells;

    /// <summary>The cell in a column, or an empty cell beyond the last one.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <returns>The cell.</returns>
    public Cell this[int column]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(column);

            if (column >= cells.Length)
            {
                return Cell.Empty;
            }

            return cells[column];
        }
    }

    /// <summary>Gets the text in a column, if the cell holds text.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="value">The text.</param>
    /// <returns>True when the cell holds text.</returns>
    public bool TryGetText(int column, [NotNullWhen(true)] out string? value)
    {
        return this[column].TryGetText(out value);
    }

    /// <summary>Gets the number in a column, if the cell holds a number.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="value">The number.</param>
    /// <returns>True when the cell holds a number.</returns>
    public bool TryGetNumber(int column, out double value)
    {
        return this[column].TryGetNumber(out value);
    }

    /// <summary>Gets the boolean in a column, if the cell holds TRUE or FALSE.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="value">The boolean.</param>
    /// <returns>True when the cell holds a boolean.</returns>
    public bool TryGetBool(int column, out bool value)
    {
        return this[column].TryGetBool(out value);
    }

    /// <summary>Gets the date in a column, if the cell holds a date or time.</summary>
    /// <param name="column">The column, 0-based.</param>
    /// <param name="value">The date.</param>
    /// <returns>True when the cell holds a date.</returns>
    public bool TryGetDate(int column, out DateTime value)
    {
        return this[column].TryGetDate(out value);
    }
}
