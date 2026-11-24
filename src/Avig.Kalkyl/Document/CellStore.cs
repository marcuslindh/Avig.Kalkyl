using System.Collections;

namespace Avig.Kalkyl;

/// <summary>
/// The rows of a sheet: found by index in constant time, and walked in order. A formula over a
/// whole column looks up every row of it, and the SortedDictionary this replaces spent about
/// 100 ns on each lookup, which made 10 000 SUMIF(B:B,...) take 20 seconds.
/// </summary>
/// <remarks>
/// Rows are never removed (a sheet only grows), so the slots only fill and <see cref="order"/>
/// only takes insertions. Rows mostly arrive from the top, as a file is read or written, which
/// appends to <see cref="order"/>.
/// </remarks>
internal sealed class RowTable : IEnumerable<KeyValuePair<int, RowEntry>>
{
    private RowEntry?[] slots = new RowEntry?[16];

    /// <summary>The indexes of the rows there are, ascending.</summary>
    private readonly List<int> order = [];

    public int Count => order.Count;

    public IEnumerable<RowEntry> Values
    {
        get
        {
            foreach (int row in order)
            {
                yield return slots[row]!;
            }
        }
    }

    public bool TryGetValue(int row, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out RowEntry? entry)
    {
        if ((uint)row < (uint)slots.Length)
        {
            entry = slots[row];

            return entry is not null;
        }

        entry = null;

        return false;
    }

    /// <summary>Adds a row that is not there yet.</summary>
    public void Add(int row, RowEntry entry)
    {
        if (row >= slots.Length)
        {
            int size = slots.Length;

            while (size <= row)
            {
                size *= 2;
            }

            Array.Resize(ref slots, Math.Min(size, CellAddress.MaxRows));
        }

        slots[row] = entry;

        if (order.Count == 0 || row > order[^1])
        {
            order.Add(row);
        }
        else
        {
            order.Insert(~order.BinarySearch(row), row);
        }
    }

    /// <summary>The rows from <paramref name="firstRow"/> to <paramref name="lastRow"/>, in order, without passing the rows before them.</summary>
    public IEnumerable<KeyValuePair<int, RowEntry>> Between(int firstRow, int lastRow)
    {
        int start = order.BinarySearch(firstRow);

        if (start < 0)
        {
            start = ~start;
        }

        for (int index = start; index < order.Count; index++)
        {
            int row = order[index];

            if (row > lastRow)
            {
                yield break;
            }

            yield return new KeyValuePair<int, RowEntry>(row, slots[row]!);
        }
    }

    /// <summary>How many rows there are from <paramref name="firstRow"/> to <paramref name="lastRow"/>.</summary>
    public int CountBetween(int firstRow, int lastRow)
    {
        int start = order.BinarySearch(firstRow);
        int end = order.BinarySearch(lastRow);

        if (start < 0)
        {
            start = ~start;
        }

        if (end < 0)
        {
            end = ~end - 1;
        }

        return Math.Max(0, end - start + 1);
    }

    public IEnumerator<KeyValuePair<int, RowEntry>> GetEnumerator()
    {
        foreach (int row in order)
        {
            yield return new KeyValuePair<int, RowEntry>(row, slots[row]!);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// The cells of a row, sorted by column in two arrays and found by binary search. A row holds
/// a handful to a few hundred cells, mostly added from the left, so an addition is usually an
/// append, and a lookup touches one or two cache lines instead of a tree of objects.
/// </summary>
internal sealed class CellRow : IEnumerable<KeyValuePair<int, CellEntry>>
{
    private int[] columns = [];
    private CellEntry[] entries = [];
    private int count;

    public int Count => count;

    /// <summary>The last column with a cell; only when <see cref="Count"/> is above 0.</summary>
    public int LastColumn => columns[count - 1];

    public IEnumerable<CellEntry> Values
    {
        get
        {
            for (int index = 0; index < count; index++)
            {
                yield return entries[index];
            }
        }
    }

    public CellEntry this[int column]
    {
        set
        {
            int index = Find(column);

            if (index >= 0)
            {
                entries[index] = value;

                return;
            }

            Insert(~index, column, value);
        }
    }

    public bool TryGetValue(int column, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out CellEntry? entry)
    {
        int index = Find(column);

        if (index >= 0)
        {
            entry = entries[index];

            return true;
        }

        entry = null;

        return false;
    }

    public bool Remove(int column)
    {
        int index = Find(column);

        if (index < 0)
        {
            return false;
        }

        count--;
        Array.Copy(columns, index + 1, columns, index, count - index);
        Array.Copy(entries, index + 1, entries, index, count - index);
        entries[count] = null!;

        return true;
    }

    /// <summary>The cells from <paramref name="firstColumn"/> to <paramref name="lastColumn"/>, in order.</summary>
    public IEnumerable<KeyValuePair<int, CellEntry>> Between(int firstColumn, int lastColumn)
    {
        int start = Find(firstColumn);

        if (start < 0)
        {
            start = ~start;
        }

        for (int index = start; index < count && columns[index] <= lastColumn; index++)
        {
            yield return new KeyValuePair<int, CellEntry>(columns[index], entries[index]);
        }
    }

    public IEnumerator<KeyValuePair<int, CellEntry>> GetEnumerator()
    {
        for (int index = 0; index < count; index++)
        {
            yield return new KeyValuePair<int, CellEntry>(columns[index], entries[index]);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>The index of the column, or the bitwise complement of where it would go.</summary>
    private int Find(int column)
    {
        if (count > 0 && columns[count - 1] < column)
        {
            return ~count;
        }

        return Array.BinarySearch(columns, 0, count, column);
    }

    private void Insert(int index, int column, CellEntry entry)
    {
        if (count == columns.Length)
        {
            int size = Math.Max(4, count * 2);
            Array.Resize(ref columns, size);
            Array.Resize(ref entries, size);
        }

        Array.Copy(columns, index, columns, index + 1, count - index);
        Array.Copy(entries, index, entries, index + 1, count - index);
        columns[index] = column;
        entries[index] = entry;
        count++;
    }
}
