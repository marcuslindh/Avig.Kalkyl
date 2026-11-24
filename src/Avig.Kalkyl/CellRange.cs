namespace Avig.Kalkyl;

/// <summary>
/// A rectangle of cells, 0-based and inclusive: <c>new CellRange(1, 1, 9, 1)</c> is the area Excel
/// calls B2:B10. <see cref="Parse"/> and <see cref="ToString"/> convert to and from A1 notation.
/// </summary>
public readonly record struct CellRange
{
    /// <summary>Creates a range from its top-left and bottom-right cells.</summary>
    /// <param name="first">The top-left cell.</param>
    /// <param name="last">The bottom-right cell; not above or left of <paramref name="first"/>.</param>
    public CellRange(CellAddress first, CellAddress last)
    {
        if (last.Row < first.Row || last.Column < first.Column)
        {
            throw new ArgumentException($"The range {first}:{last} ends above or to the left of where it starts.", nameof(last));
        }

        First = first;
        Last = last;
    }

    /// <summary>Creates a range from row and column numbers, 0-based and inclusive.</summary>
    /// <param name="firstRow">The top row.</param>
    /// <param name="firstColumn">The left column.</param>
    /// <param name="lastRow">The bottom row.</param>
    /// <param name="lastColumn">The right column.</param>
    public CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        : this(new CellAddress(firstRow, firstColumn), new CellAddress(lastRow, lastColumn))
    {
    }

    /// <summary>The top-left cell.</summary>
    public CellAddress First { get; }

    /// <summary>The bottom-right cell.</summary>
    public CellAddress Last { get; }

    /// <summary>The number of rows.</summary>
    public int RowCount => Last.Row - First.Row + 1;

    /// <summary>The number of columns.</summary>
    public int ColumnCount => Last.Column - First.Column + 1;

    /// <summary>Parses a range in A1 notation, such as <c>B2:D10</c>, <c>$B$2:$D$10</c> or a single cell, <c>B2</c>.</summary>
    /// <param name="reference">The range.</param>
    /// <returns>The range, or Invalid.</returns>
    public static Result<CellRange> Parse(string reference)
    {
        if (TryParse(reference, out CellRange range) == true)
        {
            return range;
        }

        return Result<CellRange>.Invalid($"'{reference}' is not a cell range. Expected a range such as B2:D10.");
    }

    /// <summary>Parses a range in A1 notation.</summary>
    /// <param name="reference">The range.</param>
    /// <param name="range">The range, or the default when parsing fails.</param>
    /// <returns>True when the text was a range.</returns>
    public static bool TryParse(string? reference, out CellRange range)
    {
        range = default;

        if (string.IsNullOrEmpty(reference) == true)
        {
            return false;
        }

        int colon = reference.IndexOf(':', StringComparison.Ordinal);

        if (colon < 0)
        {
            if (CellAddress.TryParse(reference, out CellAddress single) == false)
            {
                return false;
            }

            range = new CellRange(single, single);

            return true;
        }

        if (CellAddress.TryParse(reference[..colon], out CellAddress first) == false || CellAddress.TryParse(reference[(colon + 1)..], out CellAddress last) == false)
        {
            return false;
        }

        // Excel accepts the corners in any order and means the same rectangle.
        range = new CellRange(
            new CellAddress(Math.Min(first.Row, last.Row), Math.Min(first.Column, last.Column)),
            new CellAddress(Math.Max(first.Row, last.Row), Math.Max(first.Column, last.Column)));

        return true;
    }

    /// <summary>True when the cell is inside the range.</summary>
    /// <param name="address">The cell.</param>
    /// <returns>True when it is inside.</returns>
    public bool Contains(CellAddress address)
    {
        return address.Row >= First.Row && address.Row <= Last.Row && address.Column >= First.Column && address.Column <= Last.Column;
    }

    /// <summary>The range in A1 notation, such as <c>B2:D10</c>, or <c>B2</c> for a single cell.</summary>
    /// <returns>The text.</returns>
    public override string ToString()
    {
        if (First == Last)
        {
            return First.ToString();
        }

        return First.ToString() + ":" + Last.ToString();
    }

    /// <summary>The range with dollar signs, such as <c>$B$2:$D$10</c>, as charts and defined names use it.</summary>
    internal string ToAbsolute()
    {
        string first = "$" + CellAddress.ColumnName(First.Column) + "$" + (First.Row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        string last = "$" + CellAddress.ColumnName(Last.Column) + "$" + (Last.Row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

        if (First == Last)
        {
            return first;
        }

        return first + ":" + last;
    }
}
