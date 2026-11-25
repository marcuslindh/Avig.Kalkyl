using System.Text;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// A project dashboard built in code only, in the style of Excel's dashboard templates: a dark
/// sheet without gridlines or headers, a navigation column of rounded buttons with SVG icons
/// that link to other sheets, four KPI cards with gradients and shadows, a column chart with one
/// highlighted month and a smooth line, a doughnut gauge, a funnel of stacked bars and an area
/// chart. Built by both writers from one list of drawings. Also compiled into
/// tools/Avig.Kalkyl.ExcelCheck, which writes it for Excel to open.
/// </summary>
internal static class DashboardWorkbook
{
    public const string Dashboard = "Dashboard";
    public const string Data = "Data";

    /// <summary>The sheets the navigation buttons go to.</summary>
    public static readonly string[] Pages = ["Projects", "Tasks", "Budget", Data];

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    private static readonly double[] Completed = [42, 55, 61, 48, 70, 66, 74, 92, 63, 58, 71, 80];
    private static readonly double[] Planned = [50, 52, 60, 58, 64, 68, 72, 78, 70, 66, 72, 76];
    private static readonly (string Stage, double Count)[] Funnel = [("Won", 120), ("Negotiation", 260), ("Proposal", 420), ("Qualified", 610), ("Leads", 900)];
    private static readonly double[] Visitors = [320, 410, 380, 520, 610, 580, 690, 740, 720, 860];

    private const string Panel = "#1F2940";
    private const string Muted = "#AEB8CC";

    /// <summary>The dashboard in the document model, its sheets in reading order.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet dashboard = Added(book, Dashboard);

        foreach (string page in Pages[..^1])
        {
            Worksheet sheet = Added(book, page);
            sheet["A1"] = Cell.Of(page).Bold().FontSize(16);
            sheet["A3"] = Cell.Of("Back to the dashboard").FontColor("#0563C1").Underline();
            Check(sheet.AddHyperlink("A3", Hyperlink.ToCell(Dashboard, "A1") with { Tooltip = "Dashboard" }));
            sheet.View = new SheetView { TabColor = "#3A4A6B" };
        }

        Worksheet data = Added(book, Data);

        foreach ((int row, List<Cell> cells) in DataRows())
        {
            for (int column = 0; column < cells.Count; column++)
            {
                if (cells[column].IsEmpty == false)
                {
                    data[row, column] = cells[column];
                }
            }
        }

        Check(data.AddHyperlink("O1", Hyperlink.ToUrl("https://learn.microsoft.com/office/")));
        data["O1"] = Cell.Of("Office documentation").FontColor("#0563C1").Underline();
        dashboard.View = View();

        foreach ((object drawing, DrawingBounds bounds) in Drawings())
        {
            Result added = drawing switch
            {
                Shape shape => dashboard.AddShape(shape, bounds),
                ShapeGroup group => dashboard.AddGroup(group, bounds),
                Chart chart => dashboard.AddChart(chart, bounds),
                _ => dashboard.AddPicture((Picture)drawing, bounds),
            };

            Check(added);
        }

        return book;
    }

    /// <summary>
    /// The same dashboard from the streaming writer. Data comes first, since a chart can refer
    /// only to the sheets written before it, so the file opens on Data.
    /// </summary>
    public static async Task WriteAsync(XlsxWriter writer)
    {
        SheetWriter data = Sheet(await writer.AddSheetAsync(Data));
        data.SetColumnWidth(0, 12);
        int next = 0;

        foreach ((int row, List<Cell> cells) in DataRows())
        {
            while (next < row)
            {
                await data.WriteRowAsync([]);
                next++;
            }

            if (row == 0)
            {
                data.AddHyperlink("O1", Hyperlink.ToUrl("https://learn.microsoft.com/office/"));
                cells.AddRange(Enumerable.Repeat(Cell.Empty, 14 - cells.Count));
                cells.Add(Cell.Of("Office documentation").FontColor("#0563C1").Underline());
            }

            await data.WriteRowAsync(cells);
            next++;
        }

        SheetWriter dashboard = Sheet(await writer.AddSheetAsync(Dashboard));
        dashboard.SetView(View());

        foreach ((object drawing, DrawingBounds bounds) in Drawings())
        {
            switch (drawing)
            {
                case Shape shape:
                    dashboard.AddShape(shape, bounds);

                    break;

                case ShapeGroup group:
                    dashboard.AddGroup(group, bounds);

                    break;

                case Chart chart:
                    Check(dashboard.AddChart(chart, bounds));

                    break;

                default:
                    dashboard.AddPicture((Picture)drawing, bounds);

                    break;
            }
        }

        foreach (string page in Pages[..^1])
        {
            SheetWriter sheet = Sheet(await writer.AddSheetAsync(page));
            sheet.SetView(new SheetView { TabColor = "#3A4A6B" });
            sheet.AddHyperlink("A3", Hyperlink.ToCell(Dashboard, "A1") with { Tooltip = "Dashboard" });
            await sheet.WriteRowAsync([Cell.Of(page).Bold().FontSize(16)]);
            await sheet.WriteRowAsync([]);
            await sheet.WriteRowAsync([Cell.Of("Back to the dashboard").FontColor("#0563C1").Underline()]);
        }
    }

    private static SheetView View()
    {
        return new SheetView { ShowGridLines = false, ShowHeaders = false, Zoom = 90, TabColor = "#2384F9", ActiveCell = new CellAddress(0, 0) };
    }

    /// <summary>The data sheet: months in A:C, the funnel in E:G, the gauge in I:J and visitors in L:M.</summary>
    private static List<(int Row, List<Cell> Cells)> DataRows()
    {
        List<List<Cell>> rows = [];

        for (int index = 0; index < 13; index++)
        {
            rows.Add([.. Enumerable.Repeat(Cell.Empty, 13)]);
        }

        string[] headers = ["Month", "Completed", "Planned", "", "Stage", "Spacer", "Count", "", "Status", "Share", "", "Week", "Visitors"];

        for (int column = 0; column < headers.Length; column++)
        {
            if (headers[column].Length > 0)
            {
                rows[0][column] = Cell.Of(headers[column]).Bold();
            }
        }

        for (int month = 0; month < Months.Length; month++)
        {
            rows[month + 1][0] = Months[month];
            rows[month + 1][1] = Completed[month];
            rows[month + 1][2] = Planned[month];
        }

        double widest = Funnel.Max(stage => stage.Count);

        for (int stage = 0; stage < Funnel.Length; stage++)
        {
            // The spacer pushes each bar to the middle: half of what it lacks of the widest.
            rows[stage + 1][4] = Funnel[stage].Stage;
            rows[stage + 1][5] = (widest - Funnel[stage].Count) / 2;
            rows[stage + 1][6] = Funnel[stage].Count;
        }

        rows[1][8] = "Done";
        rows[1][9] = Cell.Of(0.72).NumberFormat("0%");
        rows[2][8] = "Remaining";
        rows[2][9] = Cell.Of(0.28).NumberFormat("0%");

        for (int week = 0; week < Visitors.Length; week++)
        {
            rows[week + 1][11] = "W" + (week + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            rows[week + 1][12] = Visitors[week];
        }

        return rows.Select((cells, row) => (row, cells)).ToList();
    }

    /// <summary>Every drawing on the dashboard with its place, from the back to the front.</summary>
    private static List<(object Drawing, DrawingBounds Bounds)> Drawings()
    {
        List<(object Drawing, DrawingBounds Bounds)> drawings = [];

        void Add(object drawing, double left, double top, double width, double height)
        {
            drawings.Add((drawing, DrawingBounds.FromPoints(left, top, width, height)));
        }

        TextFont white = new() { Color = "#FFFFFF", Size = 11 };

        // The page and the navigation column.
        Add(new Shape(ShapeKind.Rectangle) { Fill = Fill.Solid("#141A2A"), Outline = Outline.None, Name = "Background" }, 0, 0, 1200, 730);
        Add(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Gradient("#232D47", "#161D2E"), Outline = Outline.None, Shadow = new Shadow(), CornerRadius = 0.05, Name = "Navigation" }, 10, 10, 165, 710);
        Add(new Shape(ShapeKind.TextBox)
        {
            Text = new ShapeText([new TextRun("PROJECT\n", new TextFont { Bold = true, Size = 13, Color = "#FFFFFF" }), new TextRun("DASHBOARD", new TextFont { Bold = true, Size = 13, Color = "#8BFFF7" })])
            {
                Horizontal = TextAlignment.Left,
                Vertical = TextAnchor.Center,
            },
            Description = "Project dashboard",
        }, 22, 22, 145, 50);
        Add(new Shape(ShapeKind.Line) { Outline = new Outline("#3A4A6B", 1) }, 24, 80, 136, 0);

        string[] icons = [FolderIcon, ChecklistIcon, WalletIcon, GridIcon];

        for (int index = 0; index < Pages.Length; index++)
        {
            Fill button = Fill.Solid("#232D45");

            if (index == 0)
            {
                button = Fill.Gradient(0, new GradientStop(0, "#2384F9"), new GradientStop(1, "#605776"));
            }

            Picture icon = Svg(icons[index]) with { Description = Pages[index] + " icon" };

            ShapeGroup group = new(
            [
                GroupItem.Of(new Shape(ShapeKind.RoundedRectangle) { Fill = button, Outline = Outline.None, CornerRadius = 0.25 }, DrawingBounds.FromPoints(0, 0, 145, 40)),
                GroupItem.Of(icon, DrawingBounds.FromPoints(12, 9, 22, 22)),
                GroupItem.Of(new Shape(ShapeKind.TextBox) { Text = new ShapeText(Pages[index], white) { Horizontal = TextAlignment.Left, Vertical = TextAnchor.Center } }, DrawingBounds.FromPoints(42, 0, 100, 40)),
            ])
            {
                Link = Hyperlink.ToCell(Pages[index], "A1") with { Tooltip = "Go to " + Pages[index] },
                Name = "Go to " + Pages[index],
            };

            Add(group, 20, 100 + index * 52, 145, 40);
        }

        Add(Svg(LogoIcon) with { Link = Hyperlink.ToUrl("https://learn.microsoft.com/office/"), Description = "Help", Transparency = 0.2 }, 70, 650, 44, 44);

        // The KPI cards.
        (string Title, string Value, string From, string To)[] cards =
        [
            ("Projects", "24", "#2384F9", "#605776"),
            ("Tasks done", "780", "#0BA9A5", "#046A90"),
            ("Budget used", "62 %", "#E1541C", "#9C2F6B"),
            ("Team", "18", "#7B61FF", "#3A2C8F"),
        ];

        for (int index = 0; index < cards.Length; index++)
        {
            double left = 190 + index * 250;
            Shape card = new(ShapeKind.RoundedRectangle)
            {
                Fill = Fill.Gradient(45, new GradientStop(0, cards[index].From), new GradientStop(1, cards[index].To)),
                Outline = Outline.None,
                Shadow = new Shadow { Blur = 6, Distance = 4 },
                CornerRadius = 0.12,
                Text = new ShapeText([new TextRun(cards[index].Title + "\n", new TextFont { Size = 11, Color = "#E6ECF5" }), new TextRun(cards[index].Value, new TextFont { Size = 26, Bold = true, Color = "#FFFFFF" })])
                {
                    Horizontal = TextAlignment.Left,
                    MarginLeft = 16,
                },
                Name = cards[index].Title + " card",
            };

            if (index == 3)
            {
                card = card with { InnerShadow = new Shadow { Blur = 10, Distance = 0, Transparency = 0.7 } };
            }

            Add(card, left, 20, 235, 95);
            Add(new Shape(ShapeKind.UpArrow) { Fill = Fill.Solid("#8BFFF7"), Outline = Outline.None, Description = "Up" }, left + 195, 52, 16, 24);
        }

        // Cards behind the charts.
        (double Left, double Top, double Width, double Height)[] panels = [(190, 135, 610, 300), (810, 135, 365, 300), (190, 450, 410, 270), (610, 450, 565, 270)];

        foreach ((double left, double top, double width, double height) in panels)
        {
            Add(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Solid(Panel), Outline = Outline.None, Shadow = new Shadow(), CornerRadius = 0.04 }, left, top, width, height);
        }

        TextFont axisFont = new() { Color = Muted, Size = 9 };

        ChartAxis valueAxis = new() { Gridlines = true, GridlineOutline = new Outline("#2E3850", 0.5), Line = Outline.None, Font = axisFont };

        ChartAxis categoryAxis = new() { Line = new Outline("#3A4A6B"), Font = axisFont };

        // Tasks per month: columns with August picked out, and the plan as a smooth line.
        string?[] highlight = new string?[12];
        highlight[7] = "#00C5AE";
        Add(new Chart(ChartType.Column)
        {
            Title = "Tasks completed per month",
            TitleFont = new TextFont { Color = "#FFFFFF", Size = 12, Bold = true },
            Series =
            [
                new ChartSeries("Completed", "Data!B2:B13", "Data!A2:A13", "#3A4A6B") { PointColors = highlight, DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, Font = new TextFont { Color = "#FFFFFF", Size = 8 } } },
                new ChartSeries("Planned", "Data!C2:C13", "Data!A2:A13") { Type = ChartType.Line, Smooth = true, Marker = MarkerStyle.None, Outline = new Outline("#E1541C", 2.25), Shadow = new Shadow { Angle = 90 } },
            ],
            GapWidth = 60,
            Legend = LegendPosition.Top,
            LegendFont = axisFont,
            Background = Fill.None,
            Border = Outline.None,
            Font = axisFont,
            ValueAxis = valueAxis with { Minimum = 0, Maximum = 100, MajorUnit = 25 },
            CategoryAxis = categoryAxis,
            Description = "Tasks completed per month, with the plan as a line",
        }, 200, 142, 590, 285);

        // The gauge: a doughnut with a thick ring and the share in the middle.
        Add(new Chart(ChartType.Doughnut)
        {
            Title = "Done",
            TitleFont = new TextFont { Color = "#FFFFFF", Size = 12, Bold = true },
            Series = [new ChartSeries("Share", "Data!J2:J3", "Data!I2:I3") { PointColors = ["#00C5AE", "#2E3850"], Outline = Outline.None }],
            HoleSize = 72,
            FirstSliceAngle = 0,
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            Description = "72 % of the tasks are done",
        }, 820, 142, 345, 285);
        Add(new Shape(ShapeKind.TextBox) { Text = new ShapeText("72 %", new TextFont { Size = 30, Bold = true, Color = "#FFFFFF" }) }, 932, 275, 120, 50);

        // The funnel: an invisible spacer and the counts, stacked, so each bar sits in the middle.
        Add(new Chart(ChartType.Bar)
        {
            Title = "Sales funnel",
            TitleFont = new TextFont { Color = "#FFFFFF", Size = 12, Bold = true },
            Stacked = true,
            Series =
            [
                new ChartSeries("Spacer", "Data!F2:F6", "Data!E2:E6") { Fill = Fill.None },
                new ChartSeries("Count", "Data!G2:G6", "Data!E2:E6", "#2384F9")
                {
                    PointColors = ["#00C5AE", "#0BA9A5", "#1E8FD8", "#2384F9", "#605776"],
                    DataLabels = new DataLabels { Position = DataLabelPosition.Center, Font = new TextFont { Color = "#FFFFFF", Size = 9, Bold = true } },
                },
            ],
            GapWidth = 15,
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            ValueAxis = new ChartAxis { Visible = false, Gridlines = false },
            CategoryAxis = new ChartAxis { Line = Outline.None, Font = axisFont },
            Description = "Leads, qualified, proposal, negotiation and won deals",
        }, 200, 457, 390, 255);

        // Visitors: a gradient area under a smooth line.
        Add(new Chart(ChartType.Area)
        {
            Title = "Visitors per week",
            TitleFont = new TextFont { Color = "#FFFFFF", Size = 12, Bold = true },
            Series =
            [
                new ChartSeries("Visitors", "Data!M2:M11", "Data!L2:L11") { Fill = Fill.Gradient(90, new GradientStop(0, "#2384F9", 0.5), new GradientStop(1, "#2384F9", 1)), Outline = Outline.None },
                new ChartSeries("Trend", "Data!M2:M11", "Data!L2:L11") { Type = ChartType.Line, Smooth = true, Marker = MarkerStyle.Circle, MarkerSize = 6, Outline = new Outline("#8BFFF7", 2) },
            ],
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            Font = axisFont,
            ValueAxis = valueAxis with { Gridlines = false, Visible = false },
            CategoryAxis = categoryAxis,
            Description = "Visitors per week",
        }, 620, 457, 545, 255);

        return drawings;
    }

    private static Picture Svg(string svg)
    {
        Result<Picture> picture = Picture.FromSvg(Encoding.UTF8.GetBytes(svg));

        if (picture.IsFailure(out string? error, out _, out Picture? made) == true)
        {
            throw new InvalidOperationException(error);
        }

        return made;
    }

    private const string FolderIcon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96"><path fill="#8BFFF7" d="M2 6a2 2 0 0 1 2-2h5l2 2h9a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2z"/></svg>""";
    private const string ChecklistIcon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96"><rect x="3" y="2" width="18" height="20" rx="2" fill="none" stroke="#8BFFF7" stroke-width="2"/><path d="M7 8l2 2 4-4M7 15l2 2 4-4" fill="none" stroke="#8BFFF7" stroke-width="2"/></svg>""";
    private const string WalletIcon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96"><rect x="2" y="5" width="20" height="15" rx="2" fill="#8BFFF7"/><circle cx="17" cy="12.5" r="1.8" fill="#1F2940"/></svg>""";
    private const string GridIcon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96"><rect x="3" y="3" width="8" height="8" rx="1" fill="#8BFFF7"/><rect x="13" y="3" width="8" height="8" rx="1" fill="#8BFFF7"/><rect x="3" y="13" width="8" height="8" rx="1" fill="#8BFFF7"/><rect x="13" y="13" width="8" height="8" rx="1" fill="#8BFFF7"/></svg>""";
    private const string LogoIcon = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96"><circle cx="12" cy="12" r="10" fill="#2384F9"/><path d="M9 9a3 3 0 1 1 4 2.8c-.7.3-1 .8-1 1.5V14" fill="none" stroke="#FFFFFF" stroke-width="2"/><circle cx="12" cy="17.5" r="1.2" fill="#FFFFFF"/></svg>""";

    private static Worksheet Added(Workbook book, string name)
    {
        Result<Worksheet> added = book.AddSheet(name);

        if (added.IsFailure(out string? error, out _, out Worksheet? sheet) == true)
        {
            throw new InvalidOperationException(error);
        }

        return sheet;
    }

    private static SheetWriter Sheet(Result<SheetWriter> added)
    {
        if (added.IsFailure(out string? error, out _, out SheetWriter? sheet) == true)
        {
            throw new InvalidOperationException(error);
        }

        return sheet;
    }

    private static void Check(Result result)
    {
        if (result.IsFailure(out string? error, out _) == true)
        {
            throw new InvalidOperationException(error);
        }
    }
}
