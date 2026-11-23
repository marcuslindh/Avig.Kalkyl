using System.Globalization;
using System.Text;

namespace Avig.Kalkyl;

/// <summary>
/// A cell position, 0-based: <c>new CellAddress(1, 2)</c> is the cell Excel calls C2.
/// <see cref="Parse"/> and <see cref="ToString"/> convert to and from Excel's A1 notation.
/// </summary>
public readonly record struct CellAddress
{
    /// <summary>The number of rows in a sheet.</summary>
    public const int MaxRows = 1_048_576;

    /// <summary>The number of columns in a sheet (A to XFD).</summary>
    public const int MaxColumns = 16_384;

    /// <summary>Creates an address.</summary>
    /// <param name="row">The row, 0-based.</param>
    /// <param name="column">The column, 0-based.</param>
    public CellAddress(int row, int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(row);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(row, MaxRows);
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, MaxColumns);
        Row = row;
        Column = column;
    }

    /// <summary>The row, 0-based. Excel shows it as <c>Row + 1</c>.</summary>
    public int Row { get; }

    /// <summary>The column, 0-based. 0 is A.</summary>
    public int Column { get; }

    /// <summary>Parses an address in A1 notation, such as <c>C2</c> or <c>$C$2</c>.</summary>
    /// <param name="reference">The address.</param>
    /// <returns>The address, or Invalid.</returns>
    public static Result<CellAddress> Parse(string reference)
    {
        if (TryParse(reference, out CellAddress address) == true)
        {
            return address;
        }

        return Result<CellAddress>.Invalid($"'{reference}' is not a cell address. Expected a reference such as A1 or $C$2.");
    }

    /// <summary>Parses an address in A1 notation.</summary>
    /// <param name="reference">The address.</param>
    /// <param name="address">The address, or the default when parsing fails.</param>
    /// <returns>True when the text was an address.</returns>
    public static bool TryParse(string? reference, out CellAddress address)
    {
        address = default;

        if (string.IsNullOrEmpty(reference) == true)
        {
            return false;
        }

        ReadOnlySpan<char> text = reference.AsSpan();
        int position = 0;

        if (position < text.Length && text[position] == '$')
        {
            position++;
        }

        int letterStart = position;

        while (position < text.Length && char.IsAsciiLetter(text[position]) == true)
        {
            position++;
        }

        if (TryParseColumn(text[letterStart..position], out int column) == false)
        {
            return false;
        }

        if (position < text.Length && text[position] == '$')
        {
            position++;
        }

        ReadOnlySpan<char> digits = text[position..];

        if (digits.IsEmpty || digits[0] == '0' || int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out int rowNumber) == false)
        {
            return false;
        }

        if (rowNumber > MaxRows)
        {
            return false;
        }

        address = new CellAddress(rowNumber - 1, column);

        return true;
    }

    /// <summary>Converts column letters to a 0-based index: A is 0, Z is 25, AA is 26.</summary>
    /// <param name="letters">The letters, in either case.</param>
    /// <param name="column">The index.</param>
    /// <returns>True when the letters name a column that exists (A to XFD).</returns>
    public static bool TryParseColumn(ReadOnlySpan<char> letters, out int column)
    {
        column = -1;

        if (letters.IsEmpty || letters.Length > 3)
        {
            return false;
        }

        int value = 0;

        foreach (char letter in letters)
        {
            if (char.IsAsciiLetter(letter) == false)
            {
                return false;
            }

            value = (value * 26) + (char.ToUpperInvariant(letter) - 'A' + 1);
        }

        if (value > MaxColumns)
        {
            return false;
        }

        column = value - 1;

        return true;
    }

    /// <summary>Converts a 0-based column index to letters: 0 is A, 26 is AA.</summary>
    /// <param name="column">The index.</param>
    /// <returns>The letters.</returns>
    public static string ColumnName(int column)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(column);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(column, MaxColumns);
        Span<char> buffer = stackalloc char[3];
        int position = buffer.Length;
        int remaining = column + 1;

        while (remaining > 0)
        {
            remaining--;
            position--;
            buffer[position] = (char)('A' + (remaining % 26));
            remaining /= 26;
        }

        return new string(buffer[position..]);
    }

    /// <summary>The address in A1 notation, for example <c>C2</c>.</summary>
    /// <returns>The address.</returns>
    public override string ToString()
    {
        StringBuilder builder = new(10);
        builder.Append(ColumnName(Column));
        builder.Append((Row + 1).ToString(CultureInfo.InvariantCulture));

        return builder.ToString();
    }
}
