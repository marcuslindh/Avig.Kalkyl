using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Avig.Kalkyl;

/// <summary>The kind of value a cell holds.</summary>
public enum CellKind
{
    /// <summary>The cell holds nothing.</summary>
    Empty,

    /// <summary>The cell holds text.</summary>
    Text,

    /// <summary>The cell holds a number.</summary>
    Number,

    /// <summary>The cell holds TRUE or FALSE.</summary>
    Bool,

    /// <summary>The cell holds a date or time. In the file it is a number with a date format.</summary>
    Date,

    /// <summary>The cell holds an error such as <c>#N/A</c> or <c>#DIV/0!</c>.</summary>
    Error,
}

/// <summary>
/// The value of a cell: one of text, number, boolean, date, error or nothing.
/// Read it with the <c>TryGet…</c> methods, or switch on <see cref="Kind"/>.
/// </summary>
public readonly struct CellValue : IEquatable<CellValue>
{
    /// <summary>The longest text a cell can hold, set by Excel.</summary>
    public const int MaxTextLength = 32767;

    private static readonly DateTime EarliestDate = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);

    private static readonly HashSet<string> ErrorCodes = new(StringComparer.Ordinal)
    {
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!",
    };

    private readonly string? text;
    private readonly double number;
    private readonly DateTime date;

    private CellValue(CellKind kind, string? text, double number, DateTime date)
    {
        Kind = kind;
        this.text = text;
        this.number = number;
        this.date = date;
    }

    /// <summary>What the cell holds.</summary>
    public CellKind Kind { get; }

    /// <summary>A value that holds nothing. The same as <c>default</c>.</summary>
    public static CellValue Empty => default;

    /// <summary>True when the cell holds nothing.</summary>
    public bool IsEmpty => Kind == CellKind.Empty;

    /// <summary>Creates a text value.</summary>
    /// <param name="value">The text, at most <see cref="MaxTextLength"/> characters.</param>
    /// <returns>The value.</returns>
    public static CellValue Text(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(value.Length, MaxTextLength);

        return new CellValue(CellKind.Text, value, 0, default);
    }

    /// <summary>Creates a number value.</summary>
    /// <param name="value">A finite number. Excel has no NaN or infinity.</param>
    /// <returns>The value.</returns>
    public static CellValue Number(double value)
    {
        if (double.IsFinite(value) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Excel cannot store NaN or infinity.");
        }

        return new CellValue(CellKind.Number, null, value, default);
    }

    /// <summary>Creates a boolean value.</summary>
    /// <param name="value">TRUE or FALSE.</param>
    /// <returns>The value.</returns>
    public static CellValue Bool(bool value)
    {
        double stored = 0;

        if (value == true)
        {
            stored = 1;
        }

        return new CellValue(CellKind.Bool, null, stored, default);
    }

    /// <summary>Creates a date value.</summary>
    /// <param name="value">
    /// A date from 1899-12-31 to 9999-12-31. 1899-12-31 is Excel's day zero, which is what a
    /// time without a date is stored as.
    /// </param>
    /// <returns>The value.</returns>
    public static CellValue Date(DateTime value)
    {
        if (value < EarliestDate)
        {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Excel cannot store dates before 1900. Store the date as text instead.");
        }

        return new CellValue(CellKind.Date, null, 0, DateTime.SpecifyKind(value, DateTimeKind.Unspecified));
    }

    /// <summary>Creates a date value.</summary>
    /// <param name="value">A date from 1900-01-01 to 9999-12-31.</param>
    /// <returns>The value.</returns>
    public static CellValue Date(DateOnly value)
    {
        return Date(value.ToDateTime(TimeOnly.MinValue));
    }

    /// <summary>Creates an error value.</summary>
    /// <param name="code">An Excel error code such as <c>#N/A</c>, <c>#DIV/0!</c> or <c>#NAME?</c>.</param>
    /// <returns>The value.</returns>
    public static CellValue Error(string code)
    {
        ArgumentNullException.ThrowIfNull(code);

        if (ErrorCodes.Contains(code) == false)
        {
            throw new ArgumentException($"'{code}' is not an Excel error code.", nameof(code));
        }

        return new CellValue(CellKind.Error, code, 0, default);
    }

    /// <summary>Gets the text, if the cell holds text.</summary>
    /// <param name="value">The text, or null when the cell holds something else.</param>
    /// <returns>True when the cell holds text.</returns>
    public bool TryGetText([NotNullWhen(true)] out string? value)
    {
        if (Kind == CellKind.Text)
        {
            value = text!;

            return true;
        }

        value = null;

        return false;
    }

    /// <summary>Gets the number, if the cell holds a number. A date is not a number here.</summary>
    /// <param name="value">The number, or 0 when the cell holds something else.</param>
    /// <returns>True when the cell holds a number.</returns>
    public bool TryGetNumber(out double value)
    {
        if (Kind == CellKind.Number)
        {
            value = number;

            return true;
        }

        value = 0;

        return false;
    }

    /// <summary>Gets the boolean, if the cell holds TRUE or FALSE.</summary>
    /// <param name="value">The boolean, or false when the cell holds something else.</param>
    /// <returns>True when the cell holds a boolean.</returns>
    public bool TryGetBool(out bool value)
    {
        if (Kind == CellKind.Bool)
        {
            value = number != 0;

            return true;
        }

        value = false;

        return false;
    }

    /// <summary>Gets the date, if the cell holds a date or time.</summary>
    /// <param name="value">
    /// The date, or <see cref="DateTime.MinValue"/> when the cell holds something else. A time
    /// without a date comes back on 1899-12-31; use <see cref="DateTime.TimeOfDay"/>.
    /// </param>
    /// <returns>True when the cell holds a date.</returns>
    public bool TryGetDate(out DateTime value)
    {
        if (Kind == CellKind.Date)
        {
            value = date;

            return true;
        }

        value = DateTime.MinValue;

        return false;
    }

    /// <summary>Gets the error code, if the cell holds an error.</summary>
    /// <param name="code">The code, such as <c>#N/A</c>, or null when the cell holds something else.</param>
    /// <returns>True when the cell holds an error.</returns>
    public bool TryGetError([NotNullWhen(true)] out string? code)
    {
        if (Kind == CellKind.Error)
        {
            code = text!;

            return true;
        }

        code = null;

        return false;
    }

    /// <summary>The value as invariant text: numbers with a dot, dates as ISO 8601, booleans as TRUE/FALSE.</summary>
    /// <returns>The text; empty for an empty cell.</returns>
    public override string ToString()
    {
        return Kind switch
        {
            CellKind.Text => text!,
            CellKind.Number => number.ToString("R", CultureInfo.InvariantCulture),
            CellKind.Bool when number != 0 => "TRUE",
            CellKind.Bool => "FALSE",
            CellKind.Date => date.ToString("yyyy-MM-ddTHH:mm:ss.FFF", CultureInfo.InvariantCulture),
            CellKind.Error => text!,
            _ => string.Empty,
        };
    }

    /// <inheritdoc/>
    public bool Equals(CellValue other)
    {
        return Kind == other.Kind
            && string.Equals(text, other.text, StringComparison.Ordinal)
            && number.Equals(other.number)
            && date.Equals(other.date);
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj)
    {
        return obj is CellValue other && Equals(other);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return HashCode.Combine(Kind, text, number, date);
    }

    /// <summary>Compares two values.</summary>
    /// <param name="left">The first value.</param>
    /// <param name="right">The second value.</param>
    /// <returns>True when they are equal.</returns>
    public static bool operator ==(CellValue left, CellValue right)
    {
        return left.Equals(right);
    }

    /// <summary>Compares two values.</summary>
    /// <param name="left">The first value.</param>
    /// <param name="right">The second value.</param>
    /// <returns>True when they differ.</returns>
    public static bool operator !=(CellValue left, CellValue right)
    {
        return left.Equals(right) == false;
    }

    /// <summary>Text, or an empty value for null.</summary>
    /// <param name="value">The text.</param>
    public static implicit operator CellValue(string? value)
    {
        if (value is null)
        {
            return Empty;
        }

        return Text(value);
    }

    /// <summary>A number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator CellValue(double value)
    {
        return Number(value);
    }

    /// <summary>A number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator CellValue(int value)
    {
        return Number(value);
    }

    /// <summary>A number. Excel stores numbers as double, so more than about 15 significant digits are lost.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator CellValue(long value)
    {
        return Number(value);
    }

    /// <summary>A number. Excel stores numbers as double, so more than about 15 significant digits are lost.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator CellValue(decimal value)
    {
        return Number((double)value);
    }

    /// <summary>A boolean.</summary>
    /// <param name="value">The boolean.</param>
    public static implicit operator CellValue(bool value)
    {
        return Bool(value);
    }

    /// <summary>A date.</summary>
    /// <param name="value">The date.</param>
    public static implicit operator CellValue(DateTime value)
    {
        return Date(value);
    }

    /// <summary>A date.</summary>
    /// <param name="value">The date.</param>
    public static implicit operator CellValue(DateOnly value)
    {
        return Date(value);
    }
}
