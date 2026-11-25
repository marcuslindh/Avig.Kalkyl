using System.ComponentModel;
using System.Globalization;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// A row whose cells are matched to a mapping's columns. The generated code calls a
/// <c>TryRead</c> per property; when one fails, <see cref="Failure{T}"/> carries the reason with
/// the cell's address.
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public sealed class MappedRow
{
    private readonly IReadOnlyList<KalkylColumnInfo> columns;
    private readonly int[] sheetColumns;
    private string problem = "The row could not be read.";

    internal MappedRow(Row row, IReadOnlyList<KalkylColumnInfo> columns, int[] sheetColumns, string sheetName)
    {
        Row = row;
        this.columns = columns;
        this.sheetColumns = sheetColumns;
        SheetName = sheetName;
    }

    /// <summary>The row as read.</summary>
    public Row Row { get; }

    /// <summary>The sheet the row comes from.</summary>
    public string SheetName { get; }

    /// <summary>The failure for the last <c>TryRead</c> that returned false.</summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <returns>Invalid, with the address and the problem.</returns>
    public Result<T> Failure<T>()
    {
        return Result<T>.Invalid(problem);
    }

    /// <summary>The cell for a mapping column, or an empty cell when the sheet lacks the column.</summary>
    /// <param name="column">The mapping column.</param>
    /// <returns>The cell.</returns>
    public Cell CellFor(int column)
    {
        int sheetColumn = sheetColumns[column];

        if (sheetColumn < 0)
        {
            return Cell.Empty;
        }

        return Row[sheetColumn];
    }

    /// <summary>Reads text. Numbers, dates and booleans are turned into invariant text.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The text, or null for an empty optional cell.</param>
    /// <returns>False when a required cell is empty.</returns>
    public bool TryRead(int column, out string? value)
    {
        value = null;
        Cell cell = CellFor(column);

        if (cell.Value.IsEmpty == true)
        {
            return Optional(column);
        }

        value = cell.Value.ToString();

        return true;
    }

    /// <summary>Reads a whole number.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a whole number in range.</returns>
    public bool TryRead(int column, out int value)
    {
        value = 0;

        if (TryReadInteger(column, int.MinValue, int.MaxValue, out long number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a whole number");
        }

        value = (int)number;

        return true;
    }

    /// <summary>Reads a whole number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a whole number in range.</returns>
    public bool TryRead(int column, out int? value)
    {
        value = null;

        if (TryReadInteger(column, int.MinValue, int.MaxValue, out long number, out bool empty) == false)
        {
            return Fail(column, empty, "a whole number");
        }

        if (empty == false)
        {
            value = (int)number;
        }

        return true;
    }

    /// <summary>Reads a whole number.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a whole number in range.</returns>
    public bool TryRead(int column, out long value)
    {
        value = 0;

        if (TryReadInteger(column, long.MinValue, long.MaxValue, out long number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a whole number");
        }

        value = number;

        return true;
    }

    /// <summary>Reads a whole number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a whole number in range.</returns>
    public bool TryRead(int column, out long? value)
    {
        value = null;

        if (TryReadInteger(column, long.MinValue, long.MaxValue, out long number, out bool empty) == false)
        {
            return Fail(column, empty, "a whole number");
        }

        if (empty == false)
        {
            value = number;
        }

        return true;
    }

    /// <summary>Reads a whole number.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a whole number in range.</returns>
    public bool TryRead(int column, out short value)
    {
        value = 0;

        if (TryReadInteger(column, short.MinValue, short.MaxValue, out long number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a whole number from -32768 to 32767");
        }

        value = (short)number;

        return true;
    }

    /// <summary>Reads a whole number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a whole number in range.</returns>
    public bool TryRead(int column, out short? value)
    {
        value = null;

        if (TryReadInteger(column, short.MinValue, short.MaxValue, out long number, out bool empty) == false)
        {
            return Fail(column, empty, "a whole number from -32768 to 32767");
        }

        if (empty == false)
        {
            value = (short)number;
        }

        return true;
    }

    /// <summary>Reads a whole number from 0 to 255.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a whole number in range.</returns>
    public bool TryRead(int column, out byte value)
    {
        value = 0;

        if (TryReadInteger(column, byte.MinValue, byte.MaxValue, out long number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a whole number from 0 to 255");
        }

        value = (byte)number;

        return true;
    }

    /// <summary>Reads a whole number from 0 to 255, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a whole number in range.</returns>
    public bool TryRead(int column, out byte? value)
    {
        value = null;

        if (TryReadInteger(column, byte.MinValue, byte.MaxValue, out long number, out bool empty) == false)
        {
            return Fail(column, empty, "a whole number from 0 to 255");
        }

        if (empty == false)
        {
            value = (byte)number;
        }

        return true;
    }

    /// <summary>Reads a number.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a number.</returns>
    public bool TryRead(int column, out double value)
    {
        value = 0;

        if (TryReadDouble(column, out double number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a number");
        }

        value = number;

        return true;
    }

    /// <summary>Reads a number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a number.</returns>
    public bool TryRead(int column, out double? value)
    {
        value = null;

        if (TryReadDouble(column, out double number, out bool empty) == false)
        {
            return Fail(column, empty, "a number");
        }

        if (empty == false)
        {
            value = number;
        }

        return true;
    }

    /// <summary>Reads a number.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a number.</returns>
    public bool TryRead(int column, out float value)
    {
        value = 0;

        if (TryReadDouble(column, out double number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a number");
        }

        value = (float)number;

        return true;
    }

    /// <summary>Reads a number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a number.</returns>
    public bool TryRead(int column, out float? value)
    {
        value = null;

        if (TryReadDouble(column, out double number, out bool empty) == false)
        {
            return Fail(column, empty, "a number");
        }

        if (empty == false)
        {
            value = (float)number;
        }

        return true;
    }

    /// <summary>Reads a number. Excel stores doubles, so about 15 significant digits come through.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is empty or not a number.</returns>
    public bool TryRead(int column, out decimal value)
    {
        value = 0;

        if (TryReadDecimal(column, out decimal number, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a number");
        }

        value = number;

        return true;
    }

    /// <summary>Reads a number, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The number.</param>
    /// <returns>False when the cell is not a number.</returns>
    public bool TryRead(int column, out decimal? value)
    {
        value = null;

        if (TryReadDecimal(column, out decimal number, out bool empty) == false)
        {
            return Fail(column, empty, "a number");
        }

        if (empty == false)
        {
            value = number;
        }

        return true;
    }

    /// <summary>Reads TRUE or FALSE. Also accepts 1/0 and the words true/false, yes/no, ja/nej, sant/falskt.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The boolean.</param>
    /// <returns>False when the cell is empty or not a boolean.</returns>
    public bool TryRead(int column, out bool value)
    {
        value = false;

        if (TryReadBool(column, out bool flag, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "TRUE or FALSE");
        }

        value = flag;

        return true;
    }

    /// <summary>Reads TRUE or FALSE, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The boolean.</param>
    /// <returns>False when the cell is not a boolean.</returns>
    public bool TryRead(int column, out bool? value)
    {
        value = null;

        if (TryReadBool(column, out bool flag, out bool empty) == false)
        {
            return Fail(column, empty, "TRUE or FALSE");
        }

        if (empty == false)
        {
            value = flag;
        }

        return true;
    }

    /// <summary>Reads a date. Accepts date cells, Excel serial numbers and ISO 8601 text.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The date.</param>
    /// <returns>False when the cell is empty or not a date.</returns>
    public bool TryRead(int column, out DateTime value)
    {
        value = default;

        if (TryReadDate(column, out DateTime date, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a date");
        }

        value = date;

        return true;
    }

    /// <summary>Reads a date, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The date.</param>
    /// <returns>False when the cell is not a date.</returns>
    public bool TryRead(int column, out DateTime? value)
    {
        value = null;

        if (TryReadDate(column, out DateTime date, out bool empty) == false)
        {
            return Fail(column, empty, "a date");
        }

        if (empty == false)
        {
            value = date;
        }

        return true;
    }

    /// <summary>Reads a date and drops the time of day.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The date.</param>
    /// <returns>False when the cell is empty or not a date.</returns>
    public bool TryRead(int column, out DateOnly value)
    {
        value = default;

        if (TryReadDate(column, out DateTime date, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a date");
        }

        value = DateOnly.FromDateTime(date);

        return true;
    }

    /// <summary>Reads a date and drops the time of day, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The date.</param>
    /// <returns>False when the cell is not a date.</returns>
    public bool TryRead(int column, out DateOnly? value)
    {
        value = null;

        if (TryReadDate(column, out DateTime date, out bool empty) == false)
        {
            return Fail(column, empty, "a date");
        }

        if (empty == false)
        {
            value = DateOnly.FromDateTime(date);
        }

        return true;
    }

    /// <summary>Reads a time of day.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The time.</param>
    /// <returns>False when the cell is empty or not a time.</returns>
    public bool TryRead(int column, out TimeOnly value)
    {
        value = default;

        if (TryReadDate(column, out DateTime date, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a time");
        }

        value = TimeOnly.FromDateTime(date);

        return true;
    }

    /// <summary>Reads a time of day, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The time.</param>
    /// <returns>False when the cell is not a time.</returns>
    public bool TryRead(int column, out TimeOnly? value)
    {
        value = null;

        if (TryReadDate(column, out DateTime date, out bool empty) == false)
        {
            return Fail(column, empty, "a time");
        }

        if (empty == false)
        {
            value = TimeOnly.FromDateTime(date);
        }

        return true;
    }

    /// <summary>Reads a GUID from text.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The GUID.</param>
    /// <returns>False when the cell is empty or not a GUID.</returns>
    public bool TryRead(int column, out Guid value)
    {
        value = Guid.Empty;

        if (TryReadGuid(column, out Guid guid, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "a GUID");
        }

        value = guid;

        return true;
    }

    /// <summary>Reads a GUID from text, or null for an empty cell.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The GUID.</param>
    /// <returns>False when the cell is not a GUID.</returns>
    public bool TryRead(int column, out Guid? value)
    {
        value = null;

        if (TryReadGuid(column, out Guid guid, out bool empty) == false)
        {
            return Fail(column, empty, "a GUID");
        }

        if (empty == false)
        {
            value = guid;
        }

        return true;
    }

    /// <summary>Reads the value as it is.</summary>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The value.</param>
    /// <returns>False when a required cell is empty.</returns>
    public bool TryRead(int column, out CellValue value)
    {
        value = CellFor(column).Value;

        if (value.IsEmpty == true)
        {
            return Optional(column);
        }

        return true;
    }

    /// <summary>Reads an enum value from its name (any case) or its number.</summary>
    /// <typeparam name="TEnum">The enum.</typeparam>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The value.</param>
    /// <returns>False when the cell is empty or names no value of the enum.</returns>
    public bool TryReadEnum<TEnum>(int column, out TEnum value)
        where TEnum : struct, Enum
    {
        value = default;

        if (TryReadEnumCore(column, out TEnum parsed, out bool empty) == false || empty == true)
        {
            return Fail(column, empty, "one of " + string.Join(", ", Enum.GetNames<TEnum>()));
        }

        value = parsed;

        return true;
    }

    /// <summary>Reads an enum value from its name or number, or null for an empty cell.</summary>
    /// <typeparam name="TEnum">The enum.</typeparam>
    /// <param name="column">The mapping column.</param>
    /// <param name="value">The value.</param>
    /// <returns>False when the cell names no value of the enum.</returns>
    public bool TryReadNullableEnum<TEnum>(int column, out TEnum? value)
        where TEnum : struct, Enum
    {
        value = null;

        if (TryReadEnumCore(column, out TEnum parsed, out bool empty) == false)
        {
            return Fail(column, empty, "one of " + string.Join(", ", Enum.GetNames<TEnum>()));
        }

        if (empty == false)
        {
            value = parsed;
        }

        return true;
    }

    private bool TryReadEnumCore<TEnum>(int column, out TEnum value, out bool empty)
        where TEnum : struct, Enum
    {
        value = default;
        CellValue cell = CellFor(column).Value;
        empty = cell.IsEmpty;

        if (empty == true)
        {
            return true;
        }

        if (cell.TryGetNumber(out double number) == true)
        {
            string name = number.ToString(CultureInfo.InvariantCulture);

            return Enum.TryParse(name, out value) && Enum.IsDefined(value);
        }

        if (cell.TryGetText(out string? text) == false)
        {
            return false;
        }

        string trimmed = text.Trim();

        // Blank text is an empty cell, as it is for numbers and dates.
        if (trimmed.Length == 0)
        {
            empty = true;

            return true;
        }

        return Enum.TryParse(trimmed, true, out value) && Enum.IsDefined(value);
    }

    private bool TryReadInteger(int column, long minimum, long maximum, out long value, out bool empty)
    {
        value = 0;

        if (TryReadDouble(column, out double number, out empty) == false)
        {
            return false;
        }

        if (empty == true)
        {
            return true;
        }

        if (Math.Floor(number) != number || number < minimum || number > maximum)
        {
            return false;
        }

        value = (long)number;

        return true;
    }

    private bool TryReadDouble(int column, out double value, out bool empty)
    {
        value = 0;
        CellValue cell = CellFor(column).Value;
        empty = cell.IsEmpty;

        if (empty == true || cell.TryGetNumber(out value) == true)
        {
            return true;
        }

        if (cell.TryGetText(out string? text) == true)
        {
            string trimmed = text.Trim();

            if (trimmed.Length == 0)
            {
                empty = true;

                return true;
            }

            return TextNumbers.TryParse(trimmed, out value);
        }

        return false;
    }

    private bool TryReadDecimal(int column, out decimal value, out bool empty)
    {
        value = 0;

        if (TryReadDouble(column, out double number, out empty) == false)
        {
            return false;
        }

        if (empty == true)
        {
            return true;
        }

        if (number > (double)decimal.MaxValue || number < (double)decimal.MinValue)
        {
            return false;
        }

        // Through text with 15 digits, as Excel shows them, so that 0.1 becomes 0.1m and not 0.1000000000000000055m.
        value = decimal.Parse(number.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture);

        return true;
    }

    private bool TryReadBool(int column, out bool value, out bool empty)
    {
        value = false;
        CellValue cell = CellFor(column).Value;
        empty = cell.IsEmpty;

        if (empty == true || cell.TryGetBool(out value) == true)
        {
            return true;
        }

        if (cell.TryGetNumber(out double number) == true)
        {
            value = number != 0;

            return number is 0 or 1;
        }

        if (cell.TryGetText(out string? text) == true)
        {
            switch (text.Trim().ToUpperInvariant())
            {
                case "TRUE" or "YES" or "JA" or "SANT" or "1" or "X":
                    value = true;

                    return true;

                case "FALSE" or "NO" or "NEJ" or "FALSKT" or "0":
                    return true;

                case "":
                    empty = true;

                    return true;
            }
        }

        return false;
    }

    private bool TryReadDate(int column, out DateTime value, out bool empty)
    {
        value = default;
        CellValue cell = CellFor(column).Value;
        empty = cell.IsEmpty;

        if (empty == true || cell.TryGetDate(out value) == true)
        {
            return true;
        }

        if (cell.TryGetNumber(out double serial) == true)
        {
            return ExcelDate.TryFromSerial(serial, false, out value);
        }

        if (cell.TryGetText(out string? text) == true)
        {
            string trimmed = text.Trim();

            if (trimmed.Length == 0)
            {
                empty = true;

                return true;
            }

            return DateTime.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out value);
        }

        return false;
    }

    private bool TryReadGuid(int column, out Guid value, out bool empty)
    {
        value = Guid.Empty;
        CellValue cell = CellFor(column).Value;
        empty = cell.IsEmpty;

        if (empty == true)
        {
            return true;
        }

        if (cell.TryGetText(out string? text) == false)
        {
            return false;
        }

        string trimmed = text.Trim();

        // Blank text is an empty cell, as it is for numbers and dates.
        if (trimmed.Length == 0)
        {
            empty = true;

            return true;
        }

        return Guid.TryParse(trimmed, out value);
    }

    private bool Optional(int column)
    {
        if (columns[column].Required == false)
        {
            return true;
        }

        problem = Describe(column) + " is empty, but a value is required.";

        return false;
    }

    private bool Fail(int column, bool empty, string expected)
    {
        if (empty == true)
        {
            return Optional(column);
        }

        problem = $"{Describe(column)} holds '{CellFor(column).Value}', which is not {expected}.";

        return false;
    }

    private string Describe(int column)
    {
        string header = columns[column].Header;
        int sheetColumn = sheetColumns[column];

        if (sheetColumn < 0)
        {
            return $"Row {Row.Index + 1} (index {Row.Index}) in '{SheetName}' has no column '{header}', so it";
        }

        CellAddress address = new(Row.Index, sheetColumn);

        return $"Cell {address} (row {Row.Index}, column {sheetColumn}) in '{SheetName}', column '{header}',";
    }
}
