using System.ComponentModel;

namespace Avig.Kalkyl;

/// <summary>Conversions from property values to cells, used by the generated code.</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class KalkylCells
{
    private static readonly DateTime DayZero = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);

    /// <summary>Applies a number format when there is one.</summary>
    /// <param name="cell">The cell.</param>
    /// <param name="numberFormat">The format, or null.</param>
    /// <returns>The cell, formatted.</returns>
    public static Cell Format(Cell cell, string? numberFormat)
    {
        if (numberFormat is null || cell.Value.IsEmpty == true)
        {
            return cell;
        }

        return cell.NumberFormat(numberFormat);
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(int? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(long? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(short? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return (int)value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(byte? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return (int)value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(double? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(float? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        // Through text, so that 0.1f becomes 0.1 and not 0.100000001490116.
        return double.Parse(value.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture), System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(decimal? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(bool? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(DateTime? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A cell from a value, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(DateOnly? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value;
    }

    /// <summary>A time cell, shown as hh:mm:ss, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(TimeOnly? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return Cell.Of(DayZero.Add(value.Value.ToTimeSpan())).NumberFormat("hh:mm:ss");
    }

    /// <summary>A text cell, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(Guid? value)
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value.ToString("D");
    }

    /// <summary>A cell holding the value as it is.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(CellValue value)
    {
        return value;
    }

    /// <summary>A text cell, or an empty cell for null.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell From(string? value)
    {
        return value;
    }

    /// <summary>A text cell with the enum value's name, or an empty cell for null.</summary>
    /// <typeparam name="TEnum">The enum.</typeparam>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell FromEnum<TEnum>(TEnum? value)
        where TEnum : struct, Enum
    {
        if (value is null)
        {
            return Cell.Empty;
        }

        return value.Value.ToString();
    }
}
