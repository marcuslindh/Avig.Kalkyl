using System.Globalization;

namespace Avig.Kalkyl;

/// <summary>How a cell's value is compared in <see cref="ConditionalFormat.WhenCell(ComparisonOperator, double, Style)"/>.</summary>
public enum ComparisonOperator
{
    /// <summary>Equal to.</summary>
    Equal,

    /// <summary>Not equal to.</summary>
    NotEqual,

    /// <summary>Greater than.</summary>
    GreaterThan,

    /// <summary>Greater than or equal to.</summary>
    GreaterThanOrEqual,

    /// <summary>Less than.</summary>
    LessThan,

    /// <summary>Less than or equal to.</summary>
    LessThanOrEqual,
}

/// <summary>The icons of <see cref="ConditionalFormat.IconSet"/>, as Excel names them on the Icon Sets menu.</summary>
public enum IconSetKind
{
    /// <summary>Three coloured arrows.</summary>
    ThreeArrows,

    /// <summary>Three grey arrows.</summary>
    ThreeArrowsGray,

    /// <summary>Three flags.</summary>
    ThreeFlags,

    /// <summary>Three traffic lights without a rim; Excel's default.</summary>
    ThreeTrafficLights,

    /// <summary>Three traffic lights with a rim.</summary>
    ThreeTrafficLightsRimmed,

    /// <summary>Three signs.</summary>
    ThreeSigns,

    /// <summary>Three symbols in circles.</summary>
    ThreeSymbols,

    /// <summary>Three symbols without circles.</summary>
    ThreeSymbolsUncircled,

    /// <summary>Four coloured arrows.</summary>
    FourArrows,

    /// <summary>Four grey arrows.</summary>
    FourArrowsGray,

    /// <summary>Four circles from red to black.</summary>
    FourRedToBlack,

    /// <summary>Four rating bars.</summary>
    FourRating,

    /// <summary>Four traffic lights.</summary>
    FourTrafficLights,

    /// <summary>Five coloured arrows.</summary>
    FiveArrows,

    /// <summary>Five grey arrows.</summary>
    FiveArrowsGray,

    /// <summary>Five rating bars.</summary>
    FiveRating,

    /// <summary>Five quarters of a circle.</summary>
    FiveQuarters,
}

/// <summary>
/// A rule that changes how cells look depending on their values, as on Excel's Conditional
/// Formatting menu. Create one with the static methods and add it to a range with
/// <see cref="SheetWriter.AddConditionalFormat(CellRange, ConditionalFormat)"/> or
/// <see cref="Worksheet.AddConditionalFormat(CellRange, ConditionalFormat)"/>.
/// <para>
/// The rules that apply a <see cref="Style"/> can use its bold, italic, underline, font colour,
/// fill and number format. Excel does not let a rule change the font, its size or the alignment,
/// so a style that sets those is refused.
/// </para>
/// <para>
/// Formulas are written as Excel stores them: in English, with commas between arguments, and
/// without the leading <c>=</c>. A relative reference is relative to the top-left cell of the
/// range, so <c>$B2&gt;$C2</c> on B2:B100 compares each row with its own C cell. Never pass text
/// from a user as a formula.
/// </para>
/// </summary>
/// <example>
/// <code>
/// sheet.AddConditionalFormat("B2:B100", ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, new Style { FontColor = "#C00000" }));
/// sheet.AddConditionalFormat("C2:C100", ConditionalFormat.ColorScale("#F8696B", "#FFEB84", "#63BE7B"));
/// sheet.AddConditionalFormat("D2:D100", ConditionalFormat.DataBar("#638EC6"));
/// </code>
/// </example>
public sealed class ConditionalFormat
{
    private ConditionalFormat(ConditionalFormatKind kind, Style? style)
    {
        Kind = kind;
        Style = style;
    }

    /// <summary>The style the rule applies, or null for colour scales, data bars and icon sets.</summary>
    public Style? Style { get; }

    internal ConditionalFormatKind Kind { get; }

    internal ComparisonOperator Operator { get; private init; }

    internal bool Negated { get; private init; }

    internal IReadOnlyList<string> Arguments { get; private init; } = [];

    internal IReadOnlyList<string> Colors { get; private init; } = [];

    internal IconSetKind Icons { get; private init; }

    internal bool Reverse { get; private init; }

    internal bool ShowValue { get; private init; } = true;

    internal int Rank { get; private init; }

    internal bool Percent { get; private init; }

    internal bool FromBottom { get; private init; }

    internal string? Text { get; private init; }

    /// <summary>Applies a style when the cell's value compares with a number as given.</summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="value">The number.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenCell(ComparisonOperator comparison, double value, Style style)
    {
        return WhenCell(comparison, Number(value), style);
    }

    /// <summary>
    /// Applies a style when the cell's value compares with a formula's result as given, for
    /// example <c>$F$1</c>, <c>AVERAGE($B$2:$B$100)</c> or <c>"Done"</c> (text in quotes).
    /// </summary>
    /// <param name="comparison">The comparison.</param>
    /// <param name="formula">The formula, without a leading <c>=</c>.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenCell(ComparisonOperator comparison, string formula, Style style)
    {
        if (Enum.IsDefined(comparison) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(comparison), comparison, "Not a comparison.");
        }

        return new ConditionalFormat(ConditionalFormatKind.CellIs, CheckStyle(style)) { Operator = comparison, Arguments = [CheckFormula(formula, nameof(formula))] };
    }

    /// <summary>Applies a style when the cell's value is between two numbers, both included.</summary>
    /// <param name="low">The lower number.</param>
    /// <param name="high">The higher number.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenBetween(double low, double high, Style style)
    {
        return WhenBetween(Number(low), Number(high), style);
    }

    /// <summary>Applies a style when the cell's value is between the results of two formulas, both included.</summary>
    /// <param name="low">The formula for the lower bound.</param>
    /// <param name="high">The formula for the upper bound.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenBetween(string low, string high, Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.CellIs, CheckStyle(style)) { Arguments = [CheckFormula(low, nameof(low)), CheckFormula(high, nameof(high))] };
    }

    /// <summary>Applies a style when the cell's value is outside two numbers.</summary>
    /// <param name="low">The lower number.</param>
    /// <param name="high">The higher number.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenNotBetween(double low, double high, Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.CellIs, CheckStyle(style)) { Negated = true, Arguments = [Number(low), Number(high)] };
    }

    /// <summary>
    /// Applies a style when a formula is true, for example <c>$C2&lt;TODAY()</c> to mark whole rows
    /// whose date has passed.
    /// </summary>
    /// <param name="formula">The formula, without a leading <c>=</c>.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenFormula(string formula, Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.Expression, CheckStyle(style)) { Arguments = [CheckFormula(formula, nameof(formula))] };
    }

    /// <summary>Applies a style when the cell's text contains a text, without regard to case.</summary>
    /// <param name="text">The text to look for.</param>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat WhenTextContains(string text, Style style)
    {
        ArgumentException.ThrowIfNullOrEmpty(text);

        if (text.Length > 255)
        {
            throw new ArgumentException("Excel compares at most 255 characters in a rule.", nameof(text));
        }

        return new ConditionalFormat(ConditionalFormatKind.ContainsText, CheckStyle(style)) { Text = text };
    }

    /// <summary>Applies a style to the highest values in the range.</summary>
    /// <param name="rank">How many values, or what percentage of them when <paramref name="percent"/> is true: 1 to 1000.</param>
    /// <param name="style">The style.</param>
    /// <param name="percent">True to count a percentage of the cells instead.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat Top(int rank, Style style, bool percent = false)
    {
        return Ranked(rank, style, percent, false);
    }

    /// <summary>Applies a style to the lowest values in the range.</summary>
    /// <param name="rank">How many values, or what percentage of them when <paramref name="percent"/> is true: 1 to 1000.</param>
    /// <param name="style">The style.</param>
    /// <param name="percent">True to count a percentage of the cells instead.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat Bottom(int rank, Style style, bool percent = false)
    {
        return Ranked(rank, style, percent, true);
    }

    /// <summary>Applies a style to the values above the range's average.</summary>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat AboveAverage(Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.AboveAverage, CheckStyle(style));
    }

    /// <summary>Applies a style to the values below the range's average.</summary>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat BelowAverage(Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.AboveAverage, CheckStyle(style)) { FromBottom = true };
    }

    /// <summary>Applies a style to values that occur more than once in the range.</summary>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat Duplicates(Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.Duplicates, CheckStyle(style));
    }

    /// <summary>Applies a style to values that occur only once in the range.</summary>
    /// <param name="style">The style.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat Unique(Style style)
    {
        return new ConditionalFormat(ConditionalFormatKind.Unique, CheckStyle(style));
    }

    /// <summary>Colours each cell on a scale from the lowest value to the highest.</summary>
    /// <param name="lowest">The colour of the lowest value, as <c>#RRGGBB</c>.</param>
    /// <param name="highest">The colour of the highest value.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat ColorScale(string lowest, string highest)
    {
        return new ConditionalFormat(ConditionalFormatKind.ColorScale, null) { Colors = [Color(lowest, nameof(lowest)), Color(highest, nameof(highest))] };
    }

    /// <summary>Colours each cell on a scale from the lowest value through the median to the highest.</summary>
    /// <param name="lowest">The colour of the lowest value, as <c>#RRGGBB</c>.</param>
    /// <param name="middle">The colour of the median.</param>
    /// <param name="highest">The colour of the highest value.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat ColorScale(string lowest, string middle, string highest)
    {
        return new ConditionalFormat(ConditionalFormatKind.ColorScale, null) { Colors = [Color(lowest, nameof(lowest)), Color(middle, nameof(middle)), Color(highest, nameof(highest))] };
    }

    /// <summary>Draws a bar in each cell whose length follows the value.</summary>
    /// <param name="color">The colour of the bars, as <c>#RRGGBB</c>.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat DataBar(string color)
    {
        return new ConditionalFormat(ConditionalFormatKind.DataBar, null) { Colors = [Color(color, nameof(color))] };
    }

    /// <summary>Shows an icon in each cell, chosen by which equal part of the range's spread the value falls in.</summary>
    /// <param name="icons">The icons.</param>
    /// <param name="reverse">Reverse the order, so the first icon goes with the highest values.</param>
    /// <param name="showValue">False to show only the icon.</param>
    /// <returns>The rule.</returns>
    public static ConditionalFormat IconSet(IconSetKind icons, bool reverse = false, bool showValue = true)
    {
        if (Enum.IsDefined(icons) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(icons), icons, "Not an icon set.");
        }

        return new ConditionalFormat(ConditionalFormatKind.IconSet, null) { Icons = icons, Reverse = reverse, ShowValue = showValue };
    }

    private static ConditionalFormat Ranked(int rank, Style style, bool percent, bool bottom)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(rank, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(rank, 1000);

        if (percent == true && rank > 100)
        {
            throw new ArgumentOutOfRangeException(nameof(rank), rank, "A percentage goes from 1 to 100.");
        }

        return new ConditionalFormat(ConditionalFormatKind.Top, CheckStyle(style)) { Rank = rank, Percent = percent, FromBottom = bottom };
    }

    private static Style CheckStyle(Style style)
    {
        ArgumentNullException.ThrowIfNull(style);

        if (style.FontName is not null || style.FontSize is not null)
        {
            throw new ArgumentException("Excel does not let conditional formatting change the font or its size. Use bold, italic, underline, font colour, fill or number format.", nameof(style));
        }

        if (style.Horizontal != HorizontalAlignment.General || style.Vertical != VerticalAlignment.Bottom || style.WrapText == true)
        {
            throw new ArgumentException("Excel does not let conditional formatting change the alignment.", nameof(style));
        }

        if (style.IsDefault == true)
        {
            throw new ArgumentException("The style changes nothing. Set at least one of bold, italic, underline, font colour, fill or number format.", nameof(style));
        }

        return style;
    }

    private static string CheckFormula(string formula, string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formula, name);
        string text = formula.Trim();

        if (text.StartsWith('=') == true)
        {
            text = text[1..];
        }

        if (Avig.Kalkyl.Formulas.FormulaParser.Parse(text).IsFailure(out string? error, out ResultStatus _) == true)
        {
            throw new ArgumentException($"'{formula}' is not a formula: {error}", name);
        }

        return text;
    }

    private static string Color(string color, string name)
    {
        ArgumentNullException.ThrowIfNull(color, name);

        // Style checks colours the same way; reusing it keeps one rule for what a colour is.
        return new Style { FontColor = color }.FontColor!;
    }

    private static string Number(double value)
    {
        if (double.IsFinite(value) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(value), value, "Excel cannot store infinity or NaN.");
        }

        return value.ToString("R", CultureInfo.InvariantCulture);
    }
}

/// <summary>The cfRule types the package writes.</summary>
internal enum ConditionalFormatKind
{
    CellIs,
    Expression,
    ContainsText,
    Top,
    AboveAverage,
    Duplicates,
    Unique,
    ColorScale,
    DataBar,
    IconSet,
}
