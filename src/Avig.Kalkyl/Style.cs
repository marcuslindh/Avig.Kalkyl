using System.Globalization;

namespace Avig.Kalkyl;

/// <summary>Horizontal alignment of a cell's content.</summary>
public enum HorizontalAlignment
{
    /// <summary>Excel's default: text to the left, numbers to the right.</summary>
    General,

    /// <summary>Aligned to the left edge.</summary>
    Left,

    /// <summary>Centred.</summary>
    Center,

    /// <summary>Aligned to the right edge.</summary>
    Right,

    /// <summary>Spread over the full width.</summary>
    Justify,
}

/// <summary>Vertical alignment of a cell's content.</summary>
public enum VerticalAlignment
{
    /// <summary>Excel's default.</summary>
    Bottom,

    /// <summary>Centred.</summary>
    Center,

    /// <summary>Aligned to the top edge.</summary>
    Top,
}

/// <summary>
/// How a cell looks. Immutable and compared by value: two cells with equal styles share one
/// entry in the file, so styling every cell separately never runs into Excel's limit of about
/// 64 000 styles.
/// </summary>
public sealed record Style
{
    private readonly string? fontColor;
    private readonly string? background;
    private readonly double? fontSize;

    /// <summary>The style of an unstyled cell.</summary>
    public static Style Default { get; } = new();

    /// <summary>Bold text.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic text.</summary>
    public bool Italic { get; init; }

    /// <summary>Underlined text.</summary>
    public bool Underline { get; init; }

    /// <summary>Text colour as <c>#RRGGBB</c>, or null for automatic.</summary>
    public string? FontColor
    {
        get => fontColor;
        init => fontColor = NormalizeColor(value, nameof(FontColor));
    }

    /// <summary>Font size in points, or null for the default (11).</summary>
    public double? FontSize
    {
        get => fontSize;
        init
        {
            if (value is not null && (value < 1 || value > 409))
            {
                throw new ArgumentOutOfRangeException(nameof(FontSize), value, "Excel font sizes go from 1 to 409 points.");
            }

            fontSize = value;
        }
    }

    /// <summary>Font name, or null for the default (Calibri).</summary>
    public string? FontName { get; init; }

    /// <summary>Fill colour as <c>#RRGGBB</c>, or null for no fill.</summary>
    public string? Background
    {
        get => background;
        init => background = NormalizeColor(value, nameof(Background));
    }

    /// <summary>
    /// Number format code such as <c>0.00</c>, <c># ##0</c> or <c>yyyy-mm-dd</c>, or null for
    /// General. Dates written without a format get <c>yyyy-mm-dd</c>, or
    /// <c>yyyy-mm-dd hh:mm:ss</c> when they have a time of day.
    /// </summary>
    public string? NumberFormat { get; init; }

    /// <summary>Horizontal alignment.</summary>
    public HorizontalAlignment Horizontal { get; init; }

    /// <summary>Vertical alignment.</summary>
    public VerticalAlignment Vertical { get; init; }

    /// <summary>Wrap long text onto several lines.</summary>
    public bool WrapText { get; init; }

    /// <summary>True when the style is the same as <see cref="Default"/>.</summary>
    public bool IsDefault => Equals(Default);

    private static string? NormalizeColor(string? value, string name)
    {
        string? normalized = null;

        if (value is not null)
        {
            normalized = NormalizeHex(value, name);
        }

        return normalized;
    }

    private static string NormalizeHex(string value, string name)
    {
        string hex = value;

        if (hex.StartsWith('#') == true)
        {
            hex = hex[1..];
        }

        if (hex.Length != 6 || int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _) == false)
        {
            throw new ArgumentException($"'{value}' is not a colour. Expected #RRGGBB.", name);
        }

        return "#" + hex.ToUpperInvariant();
    }
}
