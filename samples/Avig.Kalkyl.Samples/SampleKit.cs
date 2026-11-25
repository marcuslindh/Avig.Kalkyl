using System.Text;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// The small helpers the samples share: adding a sheet, turning a failed <see cref="Result"/>
/// into an exception (a sample that fails is a bug in the sample), SVG icons and the colours
/// of the light theme.
/// </summary>
public static class SampleKit
{
    /// <summary>Dark blue for headings and header rows.</summary>
    public const string Navy = "#1E2A44";

    /// <summary>The main accent.</summary>
    public const string Blue = "#2F6FED";

    /// <summary>A lighter tint of the accent, for a second series.</summary>
    public const string LightBlue = "#9DB9F6";

    /// <summary>A second accent for highlights and "good".</summary>
    public const string Teal = "#14A38B";

    /// <summary>Warnings and a third series.</summary>
    public const string Amber = "#F59E0B";

    /// <summary>"Bad": over budget, late.</summary>
    public const string Red = "#DC3545";

    /// <summary>Body text.</summary>
    public const string Ink = "#1F2937";

    /// <summary>Secondary text, axis labels.</summary>
    public const string Grey = "#6B7280";

    /// <summary>Hairlines and gridlines.</summary>
    public const string Line = "#E5E7EB";

    /// <summary>The page behind cards.</summary>
    public const string Page = "#F4F6FA";

    /// <summary>Every other row of a table.</summary>
    public const string Band = "#F7F9FC";

    /// <summary>The font every styled cell and drawing uses.</summary>
    public const string FontName = "Segoe UI";

    /// <summary>
    /// A number as text for a shape, formatted as the pictures show the cells (Swedish
    /// separators, which is how Excel on the machine that took them displays numbers).
    /// </summary>
    public static string Swedish(double value, string format)
    {
        return value.ToString(format, System.Globalization.CultureInfo.GetCultureInfo("sv-SE"));
    }

    /// <summary>The height of a row Excel has not been told otherwise about, in points.</summary>
    public const double RowPoints = 15;

    /// <summary>
    /// The width in points that Excel draws a column of the given width (in characters of the
    /// default font, Calibri 11, whose digits are 7 pixels wide) at 100 % zoom. Used to put a
    /// drawing exactly over cells with <see cref="DrawingBounds"/>.
    /// </summary>
    public static double ColumnPoints(double width)
    {
        return Math.Floor((width * 7) + (18.0 / 256 * 7)) * 0.75;
    }

    /// <summary>Adds a sheet, or throws when the name is refused.</summary>
    public static Worksheet AddSheet(Workbook book, string name)
    {
        Result<Worksheet> added = book.AddSheet(name);

        if (added.IsFailure(out string? error, out _, out Worksheet? sheet) == true)
        {
            throw new InvalidOperationException(error);
        }

        return sheet;
    }

    /// <summary>Adds a sheet to a streaming writer, or throws when the name is refused.</summary>
    public static async Task<SheetWriter> AddSheetAsync(XlsxWriter writer, string name)
    {
        Result<SheetWriter> added = await writer.AddSheetAsync(name);

        if (added.IsFailure(out string? error, out _, out SheetWriter? sheet) == true)
        {
            throw new InvalidOperationException(error);
        }

        return sheet;
    }

    /// <summary>Throws when a call the sample relies on failed.</summary>
    public static void Check(Result result)
    {
        if (result.IsFailure(out string? error, out _) == true)
        {
            throw new InvalidOperationException(error);
        }
    }

    /// <summary>An SVG picture from its markup.</summary>
    public static Picture Svg(string svg, string description)
    {
        Result<Picture> picture = Picture.FromSvg(Encoding.UTF8.GetBytes(svg));

        if (picture.IsFailure(out string? error, out _, out Picture? made) == true)
        {
            throw new InvalidOperationException(error);
        }

        return made with { Description = description };
    }

    /// <summary>A line icon in a 24 by 24 box, drawn with one colour.</summary>
    public static Picture Icon(string paths, string color, string description)
    {
        string svg = $"""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{paths}</svg>""";

        return Svg(svg, description);
    }

    /// <summary>A text box with one run of text.</summary>
    public static Shape Label(string text, double size, string color, bool bold = false, HorizontalAlignment horizontal = HorizontalAlignment.Left)
    {
        TextFont font = new() { Name = FontName, Size = size, Color = color, Bold = bold };

        return new Shape(ShapeKind.TextBox)
        {
            Text = new ShapeText(text, font)
            {
                Horizontal = horizontal,
                Vertical = VerticalAlignment.Center,
                MarginLeft = 0,
                MarginRight = 0,
                MarginTop = 0,
                MarginBottom = 0,
            },
        };
    }

    /// <summary>A white card with rounded corners and a soft shadow, for the light theme.</summary>
    public static Shape Card()
    {
        return new Shape(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid("#FFFFFF"),
            Outline = new Outline(Line, 0.75),
            Shadow = new Shadow { Blur = 8, Distance = 2, Angle = 90, Transparency = 0.88 },
            CornerRadius = 0.06,
        };
    }

    /// <summary>A cell styled with the samples' font.</summary>
    public static Cell Text(string text, double size = 10, string color = Ink, bool bold = false)
    {
        Cell cell = Cell.Of(text).Font(FontName).FontSize(size).FontColor(color);

        if (bold == true)
        {
            cell = cell.Bold();
        }

        return cell;
    }

    /// <summary>A number in the samples' font with a number format.</summary>
    public static Cell Number(double value, string format, string color = Ink, bool bold = false)
    {
        Cell cell = Cell.Of(value).Font(FontName).FontSize(10).FontColor(color).NumberFormat(format);

        if (bold == true)
        {
            cell = cell.Bold();
        }

        return cell;
    }

    /// <summary>A formula in the samples' font with a number format.</summary>
    public static Cell Formula(string formula, string format, string color = Ink, bool bold = false)
    {
        Cell cell = Cell.FromFormula(formula).Font(FontName).FontSize(10).FontColor(color).NumberFormat(format);

        if (bold == true)
        {
            cell = cell.Bold();
        }

        return cell;
    }

    /// <summary>Gives every cell of a block the same background, so it reads as one panel.</summary>
    public static void Paint(Worksheet sheet, CellRange range, string color)
    {
        for (int row = range.First.Row; row <= range.Last.Row; row++)
        {
            for (int column = range.First.Column; column <= range.Last.Column; column++)
            {
                Cell cell = sheet[row, column];
                sheet[row, column] = cell.WithStyle(cell.Style with { Background = color });
            }
        }
    }
}
