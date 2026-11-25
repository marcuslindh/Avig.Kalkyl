using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// How the inside of a shape, a chart area or a chart series is painted: <see cref="None"/>, a
/// <see cref="Solid(string, double)"/> colour or a linear <see cref="Gradient(double, GradientStop[])"/>.
/// </summary>
public sealed record Fill
{
    private Fill(FillKind kind, string? color, double transparency, double angle, IReadOnlyList<GradientStop> stops)
    {
        Kind = kind;
        Color = color;
        Transparency = transparency;
        Angle = angle;
        Stops = stops;
    }

    /// <summary>Nothing: the shape or area is see-through.</summary>
    public static Fill None { get; } = new(FillKind.None, null, 0, 0, []);

    internal FillKind Kind { get; }

    internal string? Color { get; }

    internal double Transparency { get; }

    internal double Angle { get; }

    internal IReadOnlyList<GradientStop> Stops { get; }

    /// <summary>One colour.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="transparency">0 for opaque to 1 for invisible.</param>
    /// <returns>The fill.</returns>
    public static Fill Solid(string color, double transparency = 0)
    {
        return new Fill(FillKind.Solid, DrawingValues.Color(color, nameof(color)), DrawingValues.Transparency(transparency, nameof(transparency)), 0, []);
    }

    /// <summary>A linear gradient through two or more colours.</summary>
    /// <param name="angle">
    /// The direction in degrees, clockwise: 0 runs from left to right, 90 from top to bottom and
    /// 45 from the top-left corner to the bottom-right.
    /// </param>
    /// <param name="stops">The colours and where they sit, at least two.</param>
    /// <returns>The fill.</returns>
    public static Fill Gradient(double angle, params GradientStop[] stops)
    {
        ArgumentNullException.ThrowIfNull(stops);
        DrawingValues.Range(angle, -360, 360, nameof(angle), "The angle");

        if (stops.Length < 2 || stops.Length > 10)
        {
            throw new ArgumentException("A gradient has 2 to 10 stops.", nameof(stops));
        }

        foreach (GradientStop stop in stops)
        {
            ArgumentNullException.ThrowIfNull(stop, nameof(stops));
        }

        return new Fill(FillKind.Gradient, null, 0, angle, [.. stops]);
    }

    /// <summary>A linear gradient from one colour to another.</summary>
    /// <param name="from">The colour at the start, as <c>#RRGGBB</c>.</param>
    /// <param name="to">The colour at the end.</param>
    /// <param name="angle">The direction in degrees, as in <see cref="Gradient(double, GradientStop[])"/>; 90 runs from top to bottom.</param>
    /// <returns>The fill.</returns>
    public static Fill Gradient(string from, string to, double angle = 90)
    {
        return Gradient(angle, new GradientStop(0, from), new GradientStop(1, to));
    }
}

/// <summary>The kinds of <see cref="Fill"/>.</summary>
internal enum FillKind
{
    None,
    Solid,
    Gradient,
}

/// <summary>A colour at a position in a <see cref="Fill.Gradient(double, GradientStop[])"/>.</summary>
public sealed record GradientStop
{
    /// <summary>Creates a stop.</summary>
    /// <param name="position">Where the colour sits, from 0 (the start) to 1 (the end).</param>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="transparency">0 for opaque to 1 for invisible.</param>
    public GradientStop(double position, string color, double transparency = 0)
    {
        Position = DrawingValues.Range(position, 0, 1, nameof(position), "A stop's position");
        Color = DrawingValues.Color(color, nameof(color));
        Transparency = DrawingValues.Transparency(transparency, nameof(transparency));
    }

    /// <summary>Where the colour sits, from 0 to 1.</summary>
    public double Position { get; }

    /// <summary>The colour as <c>#RRGGBB</c>.</summary>
    public string Color { get; }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency { get; }
}

/// <summary>The dash patterns of an <see cref="Outline"/>, as on Excel's Dash type menu.</summary>
public enum LineDash
{
    /// <summary>A solid line.</summary>
    Solid,

    /// <summary>Round dots.</summary>
    RoundDot,

    /// <summary>Square dots.</summary>
    SquareDot,

    /// <summary>Dashes.</summary>
    Dash,

    /// <summary>Dash, dot.</summary>
    DashDot,

    /// <summary>Long dashes.</summary>
    LongDash,

    /// <summary>Long dash, dot.</summary>
    LongDashDot,

    /// <summary>Long dash, dot, dot.</summary>
    LongDashDotDot,
}

/// <summary>
/// The line around a shape, the line of a line shape or a chart series, or a chart's border:
/// a colour, a width and a dash pattern, or <see cref="None"/>.
/// </summary>
public sealed record Outline
{
    private readonly double transparency;

    private Outline()
    {
    }

    /// <summary>Creates a line.</summary>
    /// <param name="color">The colour as <c>#RRGGBB</c>.</param>
    /// <param name="width">The width in points, 0 to 1584; Excel's default is 0.75.</param>
    public Outline(string color, double width = 0.75)
    {
        Color = DrawingValues.Color(color, nameof(color));
        Width = DrawingValues.Range(width, 0, 1584, nameof(width), "A line's width");
    }

    /// <summary>No line.</summary>
    public static Outline None { get; } = new();

    /// <summary>The colour as <c>#RRGGBB</c>, or null for <see cref="None"/>.</summary>
    public string? Color { get; }

    /// <summary>The width in points.</summary>
    public double Width { get; }

    /// <summary>The dash pattern.</summary>
    public LineDash Dash { get; init; }

    /// <summary>0 for opaque to 1 for invisible.</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }
}

/// <summary>
/// A shadow behind a shape, a picture or a chart series (or inside a shape, as
/// <see cref="Shape.InnerShadow"/>). The defaults are Excel's "Offset: Bottom Right".
/// </summary>
public sealed record Shadow
{
    private readonly string color = "#000000";
    private readonly double transparency = 0.6;
    private readonly double blur = 4;
    private readonly double distance = 3;
    private readonly double angle = 45;

    /// <summary>The colour as <c>#RRGGBB</c>; black unless set.</summary>
    public string Color
    {
        get => color;
        init => color = DrawingValues.Color(value, nameof(Color));
    }

    /// <summary>0 for opaque to 1 for invisible; 0.6 unless set.</summary>
    public double Transparency
    {
        get => transparency;
        init => transparency = DrawingValues.Transparency(value, nameof(Transparency));
    }

    /// <summary>How soft the edge is, in points, 0 to 100; 4 unless set.</summary>
    public double Blur
    {
        get => blur;
        init => blur = DrawingValues.Range(value, 0, 100, nameof(Blur), "A shadow's blur");
    }

    /// <summary>How far the shadow falls from the shape, in points, 0 to 200; 3 unless set.</summary>
    public double Distance
    {
        get => distance;
        init => distance = DrawingValues.Range(value, 0, 200, nameof(Distance), "A shadow's distance");
    }

    /// <summary>The direction it falls in, in degrees clockwise from the right: 45 is down to the right, 90 straight down.</summary>
    public double Angle
    {
        get => angle;
        init => angle = DrawingValues.Range(value, -360, 360, nameof(Angle), "A shadow's angle");
    }
}

/// <summary>How text in a shape, a chart title, a legend, an axis or a data label looks. Every property left unset is Excel's own.</summary>
public sealed record TextFont
{
    private readonly double? size;
    private readonly string? color;
    private readonly string? name;

    /// <summary>The typeface, such as <c>Segoe UI</c>, or null for the theme's.</summary>
    public string? Name
    {
        get => name;
        init
        {
            if (value is not null && (value.Length == 0 || value.Length > 31))
            {
                throw new ArgumentException("A font name holds 1 to 31 characters.", nameof(Name));
            }

            name = value;
        }
    }

    /// <summary>The size in points, 1 to 400, or null for the default (11 in a shape, 10 in a chart).</summary>
    public double? Size
    {
        get => size;
        init
        {
            if (value is double points)
            {
                DrawingValues.Range(points, 1, 400, nameof(Size), "A font size");
            }

            size = value;
        }
    }

    /// <summary>Bold.</summary>
    public bool Bold { get; init; }

    /// <summary>Italic.</summary>
    public bool Italic { get; init; }

    /// <summary>The colour as <c>#RRGGBB</c>, or null for the default.</summary>
    public string? Color
    {
        get => color;
        init
        {
            string? normalized = null;

            if (value is not null)
            {
                normalized = DrawingValues.Color(value, nameof(Color));
            }

            color = normalized;
        }
    }
}

/// <summary>A piece of text with one look. A <c>\n</c> in the text starts a new paragraph.</summary>
public sealed record TextRun
{
    /// <summary>Creates a run.</summary>
    /// <param name="text">The text.</param>
    /// <param name="font">How it looks, or null for the shape's default.</param>
    public TextRun(string text, TextFont? font = null)
    {
        ArgumentNullException.ThrowIfNull(text);

        if (text.Length > CellValue.MaxTextLength)
        {
            throw new ArgumentException($"A text run holds at most {CellValue.MaxTextLength} characters.", nameof(text));
        }

        Text = text;
        Font = font;
    }

    /// <summary>The text.</summary>
    public string Text { get; }

    /// <summary>How it looks, or null for the shape's default.</summary>
    public TextFont? Font { get; }
}

/// <summary>The text inside a <see cref="Shape"/>: one or more runs, aligned and wrapped in the shape.</summary>
public sealed record ShapeText
{
    private readonly double? marginLeft;
    private readonly double? marginTop;
    private readonly double? marginRight;
    private readonly double? marginBottom;

    /// <summary>Text with one look.</summary>
    /// <param name="text">The text; <c>\n</c> starts a new paragraph.</param>
    /// <param name="font">How it looks, or null for the shape's default.</param>
    public ShapeText(string text, TextFont? font = null)
        : this([new TextRun(text, font)])
    {
    }

    /// <summary>Text made of runs, such as a word in another colour.</summary>
    /// <param name="runs">The runs, at least one.</param>
    public ShapeText(IReadOnlyList<TextRun> runs)
    {
        ArgumentNullException.ThrowIfNull(runs);

        if (runs.Count == 0)
        {
            throw new ArgumentException("A shape's text has at least one run.", nameof(runs));
        }

        foreach (TextRun run in runs)
        {
            ArgumentNullException.ThrowIfNull(run, nameof(runs));
        }

        Runs = [.. runs];
    }

    /// <summary>The runs.</summary>
    public IReadOnlyList<TextRun> Runs { get; }

    /// <summary>Where the lines sit across the shape; <see cref="HorizontalAlignment.General"/> is the left.</summary>
    public HorizontalAlignment Horizontal { get; init; } = HorizontalAlignment.Center;

    /// <summary>Where the text sits from top to bottom.</summary>
    public VerticalAlignment Vertical { get; init; } = VerticalAlignment.Center;

    /// <summary>Break lines at the shape's edge. Off, a line runs on past it.</summary>
    public bool Wrap { get; init; } = true;

    /// <summary>The space between the shape's left edge and the text, in points, or null for Excel's 7.2.</summary>
    public double? MarginLeft
    {
        get => marginLeft;
        init => marginLeft = Margin(value, nameof(MarginLeft));
    }

    /// <summary>The space above the text, in points, or null for Excel's 3.6.</summary>
    public double? MarginTop
    {
        get => marginTop;
        init => marginTop = Margin(value, nameof(MarginTop));
    }

    /// <summary>The space right of the text, in points, or null for Excel's 7.2.</summary>
    public double? MarginRight
    {
        get => marginRight;
        init => marginRight = Margin(value, nameof(MarginRight));
    }

    /// <summary>The space below the text, in points, or null for Excel's 3.6.</summary>
    public double? MarginBottom
    {
        get => marginBottom;
        init => marginBottom = Margin(value, nameof(MarginBottom));
    }

    private static double? Margin(double? value, string parameter)
    {
        if (value is double points)
        {
            DrawingValues.Range(points, 0, 1000, parameter, "A margin");
        }

        return value;
    }
}

/// <summary>
/// Where a click goes: a cell in the workbook (<see cref="ToCell(string, string)"/>) or a web
/// address (<see cref="ToUrl(string)"/>). Put one on a shape, a group, a picture or a cell.
/// </summary>
public sealed record Hyperlink
{
    private readonly string? tooltip;

    private Hyperlink(string? url, string? sheet, CellAddress cell)
    {
        Url = url;
        Sheet = sheet;
        Cell = cell;
    }

    /// <summary>The web address, or null for a link inside the workbook.</summary>
    public string? Url { get; }

    /// <summary>The sheet a link inside the workbook goes to, or null for a web address.</summary>
    public string? Sheet { get; }

    /// <summary>The cell a link inside the workbook goes to.</summary>
    public CellAddress Cell { get; }

    /// <summary>The text Excel shows when the pointer rests on the link, or null.</summary>
    public string? Tooltip
    {
        get => tooltip;
        init
        {
            if (value is not null && value.Length > 255)
            {
                throw new ArgumentException("A tooltip holds at most 255 characters.", nameof(Tooltip));
            }

            tooltip = value;
        }
    }

    internal bool IsExternal => Url is not null;

    /// <summary>The sheet and cell as Excel stores them: <c>'Sales 2026'!A1</c>.</summary>
    internal string Location => "'" + Sheet!.Replace("'", "''", StringComparison.Ordinal) + "'!" + Cell.ToString();

    /// <summary>A link to a cell in the workbook. The sheet does not have to exist yet.</summary>
    /// <param name="sheet">The sheet's name.</param>
    /// <param name="cell">The cell.</param>
    /// <returns>The link.</returns>
    public static Hyperlink ToCell(string sheet, CellAddress cell)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        if (XlsxWriter.ValidateName(sheet).IsFailure(out string? error, out ResultStatus _) == true)
        {
            throw new ArgumentException(error, nameof(sheet));
        }

        return new Hyperlink(null, sheet, cell);
    }

    /// <summary>A link to a cell in the workbook, such as <c>Hyperlink.ToCell("Sales", "A1")</c>.</summary>
    /// <param name="sheet">The sheet's name.</param>
    /// <param name="cell">The cell in A1 notation.</param>
    /// <returns>The link.</returns>
    public static Hyperlink ToCell(string sheet, string cell = "A1")
    {
        ArgumentNullException.ThrowIfNull(cell);

        if (CellAddress.TryParse(cell, out CellAddress address) == false)
        {
            throw new ArgumentException($"'{cell}' is not a cell address. Expected a reference such as A1.", nameof(cell));
        }

        return ToCell(sheet, address);
    }

    /// <summary>A link to a web page or a mail address.</summary>
    /// <param name="url">An absolute address: <c>https://</c>, <c>http://</c> or <c>mailto:</c>.</param>
    /// <returns>The link.</returns>
    public static Hyperlink ToUrl(string url)
    {
        ArgumentNullException.ThrowIfNull(url);

        if (url.Length > 2079 || Uri.TryCreate(url, UriKind.Absolute, out Uri? parsed) == false
            || (parsed.Scheme != Uri.UriSchemeHttps && parsed.Scheme != Uri.UriSchemeHttp && parsed.Scheme != Uri.UriSchemeMailto))
        {
            throw new ArgumentException($"'{url}' is not an absolute https, http or mailto address of at most 2079 characters.", nameof(url));
        }

        return new Hyperlink(url, null, default);
    }
}

/// <summary>
/// A place on a sheet in points (or EMU), from the sheet's top-left corner. A drawing placed so
/// stays where it is when rows and columns change size, as a dashboard's cards should.
/// </summary>
public readonly record struct DrawingBounds
{
    private DrawingBounds(long left, long top, long width, long height)
    {
        if (left < 0 || top < 0 || width < 0 || height < 0 || left > DrawingValues.MaxCoordinate || top > DrawingValues.MaxCoordinate
            || width > DrawingValues.MaxCoordinate || height > DrawingValues.MaxCoordinate)
        {
            throw new ArgumentOutOfRangeException(nameof(left), "A drawing's position and size must be 0 or more and inside the sheet.");
        }

        Left = left;
        Top = top;
        Width = width;
        Height = height;
    }

    /// <summary>The distance from the sheet's left edge, in EMU (12 700 to the point).</summary>
    public long Left { get; }

    /// <summary>The distance from the sheet's top edge, in EMU.</summary>
    public long Top { get; }

    /// <summary>The width in EMU.</summary>
    public long Width { get; }

    /// <summary>The height in EMU.</summary>
    public long Height { get; }

    /// <summary>A place in points (72 to the inch). Excel's default column is 48 points wide and its default row 15 high.</summary>
    /// <param name="left">From the left edge.</param>
    /// <param name="top">From the top edge.</param>
    /// <param name="width">The width.</param>
    /// <param name="height">The height.</param>
    /// <returns>The bounds.</returns>
    public static DrawingBounds FromPoints(double left, double top, double width, double height)
    {
        foreach (double value in (ReadOnlySpan<double>)[left, top, width, height])
        {
            if (double.IsFinite(value) == false || value < 0 || value > 2_000_000)
            {
                throw new ArgumentOutOfRangeException(nameof(left), value, "A drawing's position and size go from 0 to 2 000 000 points.");
            }
        }

        return new DrawingBounds(DrawingValues.Emu(left), DrawingValues.Emu(top), DrawingValues.Emu(width), DrawingValues.Emu(height));
    }

    /// <summary>A place in EMU, as DrawingML stores it: 914 400 to the inch, 12 700 to the point.</summary>
    /// <param name="left">From the left edge.</param>
    /// <param name="top">From the top edge.</param>
    /// <param name="width">The width.</param>
    /// <param name="height">The height.</param>
    /// <returns>The bounds.</returns>
    public static DrawingBounds FromEmu(long left, long top, long width, long height)
    {
        return new DrawingBounds(left, top, width, height);
    }
}

/// <summary>The shapes on Excel's Insert &gt; Shapes menu that <see cref="Shape"/> draws.</summary>
public enum ShapeKind
{
    /// <summary>A rectangle.</summary>
    Rectangle,

    /// <summary>A rectangle with rounded corners; <see cref="Shape.CornerRadius"/> sets how round.</summary>
    RoundedRectangle,

    /// <summary>An ellipse, or a circle in a square.</summary>
    Ellipse,

    /// <summary>An isosceles triangle pointing up.</summary>
    Triangle,

    /// <summary>A diamond.</summary>
    Diamond,

    /// <summary>A hexagon.</summary>
    Hexagon,

    /// <summary>A block arrow pointing right.</summary>
    RightArrow,

    /// <summary>A block arrow pointing left.</summary>
    LeftArrow,

    /// <summary>A block arrow pointing up.</summary>
    UpArrow,

    /// <summary>A block arrow pointing down.</summary>
    DownArrow,

    /// <summary>A chevron pointing right.</summary>
    Chevron,

    /// <summary>A pentagon pointing right, as a step in a process.</summary>
    Pentagon,

    /// <summary>A straight line from the top-left corner to the bottom-right (flip it for the other diagonal); only its <see cref="Shape.Outline"/> shows.</summary>
    Line,

    /// <summary>A text box: a rectangle without fill or line, for text alone.</summary>
    TextBox,
}

/// <summary>
/// A shape to put on a sheet with <c>AddShape</c> on <see cref="SheetWriter"/> or
/// <see cref="Worksheet"/>: a preset geometry with a fill, an outline, shadows, text and a link.
/// Everything left unset looks as a shape just inserted in Excel does.
/// </summary>
/// <example>
/// <code>
/// Shape card = new(ShapeKind.RoundedRectangle)
/// {
///     Fill = Fill.Gradient("#2384F9", "#605776", 45),
///     Outline = Outline.None,
///     Shadow = new Shadow(),
///     Text = new ShapeText("Projects", new TextFont { Size = 14, Bold = true, Color = "#FFFFFF" }),
///     Link = Hyperlink.ToCell("Projects", "A1"),
/// };
/// sheet.AddShape(card, DrawingBounds.FromPoints(20, 20, 160, 90));
/// </code>
/// </example>
public sealed record Shape
{
    private readonly double? cornerRadius;
    private readonly double rotation;
    private readonly double? softEdges;
    private readonly string? name;
    private readonly string? description;

    /// <summary>Creates a shape.</summary>
    /// <param name="kind">The geometry.</param>
    public Shape(ShapeKind kind)
    {
        if (Enum.IsDefined(kind) == false)
        {
            throw new ArgumentOutOfRangeException(nameof(kind), kind, "Not a shape kind.");
        }

        Kind = kind;
    }

    /// <summary>The geometry.</summary>
    public ShapeKind Kind { get; }

    /// <summary>
    /// How round a <see cref="ShapeKind.RoundedRectangle"/>'s corners are, from 0 (square) to 0.5
    /// (half the shorter side, a pill), or null for Excel's 0.1667.
    /// </summary>
    public double? CornerRadius
    {
        get => cornerRadius;
        init
        {
            if (value is double radius)
            {
                DrawingValues.Range(radius, 0, 0.5, nameof(CornerRadius), "A corner radius");
            }

            cornerRadius = value;
        }
    }

    /// <summary>The inside, or null for Excel's default (the theme's first accent colour; none for a text box).</summary>
    public Fill? Fill { get; init; }

    /// <summary>The line around it, or null for Excel's default (a darker accent; none for a text box).</summary>
    public Outline? Outline { get; init; }

    /// <summary>A shadow behind it, or null for none.</summary>
    public Shadow? Shadow { get; init; }

    /// <summary>A shadow inside its edge, or null for none.</summary>
    public Shadow? InnerShadow { get; init; }

    /// <summary>Softened edges, in points from 0 to 100, or null for sharp ones.</summary>
    public double? SoftEdges
    {
        get => softEdges;
        init
        {
            if (value is double points)
            {
                DrawingValues.Range(points, 0, 100, nameof(SoftEdges), "Soft edges");
            }

            softEdges = value;
        }
    }

    /// <summary>The text inside, or null for none. A <see cref="ShapeKind.Line"/> has no text.</summary>
    public ShapeText? Text { get; init; }

    /// <summary>The turn in degrees, clockwise, -360 to 360.</summary>
    public double Rotation
    {
        get => rotation;
        init => rotation = DrawingValues.Range(value, -360, 360, nameof(Rotation), "A rotation");
    }

    /// <summary>Mirror it left to right; for a line, it then runs from the top-right corner to the bottom-left.</summary>
    public bool FlipHorizontal { get; init; }

    /// <summary>Mirror it top to bottom.</summary>
    public bool FlipVertical { get; init; }

    /// <summary>Where a click on it goes, or null.</summary>
    public Hyperlink? Link { get; init; }

    /// <summary>The name in Excel's Selection Pane, or null for "Shape 2" and so on.</summary>
    public string? Name
    {
        get => name;
        init => name = DrawingText(value, nameof(Name));
    }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description
    {
        get => description;
        init => description = DrawingText(value, nameof(Description));
    }

    internal static string? DrawingText(string? value, string parameter)
    {
        if (value is not null && value.Length > 1000)
        {
            throw new ArgumentException("A drawing's name or alternative text holds at most 1000 characters.", parameter);
        }

        return value;
    }
}

/// <summary>
/// Shapes and pictures that move and resize as one, with one link: a dashboard's navigation
/// button made of an icon and a label, for example. Each item has its place inside the group;
/// the group is then put on the sheet with <c>AddGroup</c>, and its items are scaled to fit.
/// </summary>
public sealed record ShapeGroup
{
    private readonly string? name;
    private readonly string? description;

    /// <summary>Creates a group.</summary>
    /// <param name="items">The shapes and pictures, at least one, drawn in this order (the last on top).</param>
    public ShapeGroup(IReadOnlyList<GroupItem> items)
    {
        ArgumentNullException.ThrowIfNull(items);

        if (items.Count == 0 || items.Count > 1000)
        {
            throw new ArgumentException("A group holds 1 to 1000 items.", nameof(items));
        }

        foreach (GroupItem item in items)
        {
            ArgumentNullException.ThrowIfNull(item, nameof(items));
        }

        Items = [.. items];
    }

    /// <summary>The shapes and pictures.</summary>
    public IReadOnlyList<GroupItem> Items { get; }

    /// <summary>Where a click on the group goes, or null. Links on the items themselves are kept too.</summary>
    public Hyperlink? Link { get; init; }

    /// <summary>The name in Excel's Selection Pane, or null for "Group 2" and so on.</summary>
    public string? Name
    {
        get => name;
        init => name = Shape.DrawingText(value, nameof(Name));
    }

    /// <summary>The alternative text that screen readers read, or null.</summary>
    public string? Description
    {
        get => description;
        init => description = Shape.DrawingText(value, nameof(Description));
    }
}

/// <summary>A shape or a picture in a <see cref="ShapeGroup"/>, with its place inside the group.</summary>
public sealed record GroupItem
{
    private GroupItem(Shape? shape, Picture? picture, DrawingBounds bounds)
    {
        Shape = shape;
        Picture = picture;
        Bounds = bounds;
    }

    /// <summary>The shape, or null when the item is a picture.</summary>
    public Shape? Shape { get; }

    /// <summary>The picture, or null when the item is a shape.</summary>
    public Picture? Picture { get; }

    /// <summary>The place inside the group, measured from the group's top-left corner.</summary>
    public DrawingBounds Bounds { get; }

    /// <summary>A shape at a place inside the group.</summary>
    /// <param name="shape">The shape.</param>
    /// <param name="bounds">The place, from the group's top-left corner.</param>
    /// <returns>The item.</returns>
    public static GroupItem Of(Shape shape, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(shape);

        return new GroupItem(shape, null, bounds);
    }

    /// <summary>A picture at a place inside the group.</summary>
    /// <param name="picture">The picture.</param>
    /// <param name="bounds">The place, from the group's top-left corner.</param>
    /// <returns>The item.</returns>
    public static GroupItem Of(Picture picture, DrawingBounds bounds)
    {
        ArgumentNullException.ThrowIfNull(picture);

        return new GroupItem(null, picture, bounds);
    }
}
