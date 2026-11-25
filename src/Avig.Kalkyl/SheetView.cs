using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl;

/// <summary>
/// How a sheet looks in Excel's window: gridlines, row and column headers, zoom, the tab's
/// colour, the selected cell and the direction. Set it with <see cref="SheetWriter.SetView(SheetView)"/>
/// or <see cref="Worksheet.View"/>; every property left unset is Excel's default.
/// </summary>
/// <example>
/// <code>
/// sheet.SetView(new SheetView { ShowGridLines = false, ShowHeaders = false, Zoom = 85, TabColor = "#1F4E79" });
/// </code>
/// </example>
public sealed record SheetView
{
    private readonly int zoom = 100;
    private readonly string? tabColor;

    /// <summary>Excel's default view.</summary>
    public static SheetView Default { get; } = new();

    /// <summary>Show the gridlines between cells. Off, a dashboard's background shows clean.</summary>
    public bool ShowGridLines { get; init; } = true;

    /// <summary>Show the row numbers and column letters.</summary>
    public bool ShowHeaders { get; init; } = true;

    /// <summary>The zoom in percent, 10 to 400.</summary>
    public int Zoom
    {
        get => zoom;
        init => zoom = DrawingValues.Range(value, 10, 400, nameof(Zoom), "The zoom");
    }

    /// <summary>The colour of the sheet's tab as <c>#RRGGBB</c>, or null for none.</summary>
    public string? TabColor
    {
        get => tabColor;
        init
        {
            string? normalized = null;

            if (value is not null)
            {
                normalized = DrawingValues.Color(value, nameof(TabColor));
            }

            tabColor = normalized;
        }
    }

    /// <summary>The selected cell when the file opens, or null for A1 (or the first cell after frozen panes).</summary>
    public CellAddress? ActiveCell { get; init; }

    /// <summary>Lay the sheet out from right to left, with column A at the right, as for Arabic or Hebrew.</summary>
    public bool RightToLeft { get; init; }
}
