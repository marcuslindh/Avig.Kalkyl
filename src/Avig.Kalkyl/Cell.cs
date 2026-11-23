using System.Diagnostics.CodeAnalysis;

namespace Avig.Kalkyl;

/// <summary>
/// A cell: a value, an optional formula and a style. Immutable; the styling methods return a
/// new cell, so they chain: <c>Cell.Of("Total").Bold().Background("#EEEEEE")</c>.
/// Any value converts to a cell implicitly, so a row can be written as
/// <c>["Name", 42.5, DateTime.Today]</c>.
/// </summary>
public readonly struct Cell : IEquatable<Cell>
{
    private readonly Style? style;

    private Cell(CellValue value, string? formula, Style? style)
    {
        Value = value;
        Formula = formula;
        this.style = style;
    }

    /// <summary>A cell with no value, formula or style.</summary>
    public static Cell Empty => default;

    /// <summary>The value. For a formula cell this is the last calculated result, if any.</summary>
    public CellValue Value { get; }

    /// <summary>The formula without a leading <c>=</c>, or null when the cell holds a plain value.</summary>
    public string? Formula { get; }

    /// <summary>
    /// How the cell looks. When reading, only <see cref="Kalkyl.Style.NumberFormat"/> is filled in.
    /// </summary>
    public Style Style => style ?? Style.Default;

    /// <summary>The kind of <see cref="Value"/>.</summary>
    public CellKind Kind => Value.Kind;

    /// <summary>True when the cell has no value, no formula and the default style.</summary>
    public bool IsEmpty => Value.IsEmpty && Formula is null && Style.IsDefault;

    /// <summary>Creates a cell holding a value.</summary>
    /// <param name="value">The value.</param>
    /// <returns>The cell.</returns>
    public static Cell Of(CellValue value)
    {
        return new Cell(value, null, null);
    }

    /// <summary>Creates a formula cell.</summary>
    /// <param name="formula">The formula, with or without a leading <c>=</c>, for example <c>SUM(B2:B9)</c>.</param>
    /// <returns>The cell. Excel calculates it when the file is opened.</returns>
    public static Cell FromFormula(string formula)
    {
        return FromFormula(formula, CellValue.Empty);
    }

    /// <summary>Creates a formula cell with a result already filled in.</summary>
    /// <param name="formula">The formula, with or without a leading <c>=</c>.</param>
    /// <param name="cachedValue">The result to store, which programs that do not calculate will show.</param>
    /// <returns>The cell.</returns>
    public static Cell FromFormula(string formula, CellValue cachedValue)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);
        string text = formula.Trim();

        if (text.StartsWith('=') == true)
        {
            text = text[1..];
        }

        return new Cell(cachedValue, text, null);
    }

    /// <summary>Internal factory for the reader, which knows all three parts at once.</summary>
    internal static Cell Create(CellValue value, string? formula, Style? style)
    {
        return new Cell(value, formula, style);
    }

    /// <summary>Gets the text, if the cell holds text.</summary>
    /// <param name="value">The text.</param>
    /// <returns>True when the cell holds text.</returns>
    public bool TryGetText([NotNullWhen(true)] out string? value)
    {
        return Value.TryGetText(out value);
    }

    /// <summary>Gets the number, if the cell holds a number.</summary>
    /// <param name="value">The number.</param>
    /// <returns>True when the cell holds a number.</returns>
    public bool TryGetNumber(out double value)
    {
        return Value.TryGetNumber(out value);
    }

    /// <summary>Gets the boolean, if the cell holds TRUE or FALSE.</summary>
    /// <param name="value">The boolean.</param>
    /// <returns>True when the cell holds a boolean.</returns>
    public bool TryGetBool(out bool value)
    {
        return Value.TryGetBool(out value);
    }

    /// <summary>Gets the date, if the cell holds a date or time.</summary>
    /// <param name="value">The date.</param>
    /// <returns>True when the cell holds a date.</returns>
    public bool TryGetDate(out DateTime value)
    {
        return Value.TryGetDate(out value);
    }

    /// <summary>Gets the error code, if the cell holds an error.</summary>
    /// <param name="code">The error code.</param>
    /// <returns>True when the cell holds an error.</returns>
    public bool TryGetError([NotNullWhen(true)] out string? code)
    {
        return Value.TryGetError(out code);
    }

    /// <summary>Replaces the whole style.</summary>
    /// <param name="newStyle">The style.</param>
    /// <returns>A new cell.</returns>
    public Cell WithStyle(Style newStyle)
    {
        ArgumentNullException.ThrowIfNull(newStyle);

        return new Cell(Value, Formula, newStyle);
    }

    /// <summary>Makes the text bold.</summary>
    /// <returns>A new cell.</returns>
    public Cell Bold()
    {
        return WithStyle(Style with { Bold = true });
    }

    /// <summary>Makes the text italic.</summary>
    /// <returns>A new cell.</returns>
    public Cell Italic()
    {
        return WithStyle(Style with { Italic = true });
    }

    /// <summary>Underlines the text.</summary>
    /// <returns>A new cell.</returns>
    public Cell Underline()
    {
        return WithStyle(Style with { Underline = true });
    }

    /// <summary>Sets the text colour.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <returns>A new cell.</returns>
    public Cell FontColor(string color)
    {
        return WithStyle(Style with { FontColor = color });
    }

    /// <summary>Sets the font size.</summary>
    /// <param name="points">The size in points.</param>
    /// <returns>A new cell.</returns>
    public Cell FontSize(double points)
    {
        return WithStyle(Style with { FontSize = points });
    }

    /// <summary>Sets the font.</summary>
    /// <param name="name">The font name, for example <c>Arial</c>.</param>
    /// <returns>A new cell.</returns>
    public Cell Font(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

        return WithStyle(Style with { FontName = name });
    }

    /// <summary>Fills the cell with a colour.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <returns>A new cell.</returns>
    public Cell Background(string color)
    {
        return WithStyle(Style with { Background = color });
    }

    /// <summary>Sets the number format.</summary>
    /// <param name="format">A format code such as <c>0.00</c> or <c>yyyy-mm-dd</c>.</param>
    /// <returns>A new cell.</returns>
    public Cell NumberFormat(string format)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(format);

        return WithStyle(Style with { NumberFormat = format });
    }

    /// <summary>Sets the horizontal alignment.</summary>
    /// <param name="alignment">The alignment.</param>
    /// <returns>A new cell.</returns>
    public Cell Align(HorizontalAlignment alignment)
    {
        return WithStyle(Style with { Horizontal = alignment });
    }

    /// <summary>Sets the vertical alignment.</summary>
    /// <param name="alignment">The alignment.</param>
    /// <returns>A new cell.</returns>
    public Cell Align(VerticalAlignment alignment)
    {
        return WithStyle(Style with { Vertical = alignment });
    }

    /// <summary>Wraps long text onto several lines.</summary>
    /// <returns>A new cell.</returns>
    public Cell Wrap()
    {
        return WithStyle(Style with { WrapText = true });
    }

    /// <inheritdoc/>
    public override string ToString()
    {
        if (Formula is not null)
        {
            return "=" + Formula;
        }

        return Value.ToString();
    }

    /// <inheritdoc/>
    public bool Equals(Cell other)
    {
        return Value.Equals(other.Value)
            && string.Equals(Formula, other.Formula, StringComparison.Ordinal)
            && Style.Equals(other.Style);
    }

    /// <inheritdoc/>
    public override bool Equals(object? obj)
    {
        return obj is Cell other && Equals(other);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return HashCode.Combine(Value, Formula, Style);
    }

    /// <summary>Compares two cells.</summary>
    /// <param name="left">The first cell.</param>
    /// <param name="right">The second cell.</param>
    /// <returns>True when they are equal.</returns>
    public static bool operator ==(Cell left, Cell right)
    {
        return left.Equals(right);
    }

    /// <summary>Compares two cells.</summary>
    /// <param name="left">The first cell.</param>
    /// <param name="right">The second cell.</param>
    /// <returns>True when they differ.</returns>
    public static bool operator !=(Cell left, Cell right)
    {
        return left.Equals(right) == false;
    }

    /// <summary>A cell holding the value.</summary>
    /// <param name="value">The value.</param>
    public static implicit operator Cell(CellValue value)
    {
        return Of(value);
    }

    /// <summary>A cell holding text, or an empty cell for null.</summary>
    /// <param name="value">The text.</param>
    public static implicit operator Cell(string? value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator Cell(double value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator Cell(int value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator Cell(long value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a number.</summary>
    /// <param name="value">The number.</param>
    public static implicit operator Cell(decimal value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a boolean.</summary>
    /// <param name="value">The boolean.</param>
    public static implicit operator Cell(bool value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a date.</summary>
    /// <param name="value">The date.</param>
    public static implicit operator Cell(DateTime value)
    {
        return Of(value);
    }

    /// <summary>A cell holding a date.</summary>
    /// <param name="value">The date.</param>
    public static implicit operator Cell(DateOnly value)
    {
        return Of(value);
    }
}
