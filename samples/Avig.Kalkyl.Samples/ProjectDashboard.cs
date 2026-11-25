using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A dark project portfolio dashboard in the style of Excel's dashboard templates, made of
/// drawings only: a navigation column of buttons with SVG icons that link to the other sheets,
/// four KPI cards with gradients, icons and the change since last month, a column chart with
/// one month picked out and the plan as a smooth line, a doughnut gauge with the share in the
/// middle, a funnel of stacked bars and an area chart under a smooth line. The charts read the
/// Data sheet; the Projects sheet is a table with conditional formatting.
/// </summary>
public static class ProjectDashboard
{
    /// <summary>The sheet with the dashboard.</summary>
    public const string Dashboard = "Dashboard";

    /// <summary>The cells the gallery picture covers: 25 standard columns and 49 rows, 1200 by 735 points.</summary>
    public const string PictureRange = "A1:Y49";

    private const string Data = "Data";
    private const string Projects = "Projects";
    private const string Page = "#0F1524";
    private const string Panel = "#1A2236";
    private const string Muted = "#9AA6BF";
    private const string Accent = "#5EEAD4";

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    private static readonly double[] Completed = [42, 55, 61, 48, 70, 66, 74, 92, 63, 58, 71, 80];
    private static readonly double[] Planned = [50, 52, 60, 58, 64, 68, 72, 78, 70, 66, 72, 76];
    private static readonly (string Stage, double Count)[] Funnel = [("Delivered", 120), ("In review", 260), ("In progress", 420), ("Planned", 610), ("Backlog", 900)];
    private static readonly double[] Hours = [320, 410, 380, 520, 610, 580, 690, 740, 720, 860, 820, 905];

    private static readonly (string Name, string Icon)[] Pages =
    [
        ("Dashboard", """<rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/>"""),
        ("Projects", """<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>"""),
        ("Tasks", """<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 9l2 2 4-4M8 16h8"/>"""),
        ("Budget", """<rect x="3" y="6" width="18" height="13" rx="2"/><path d="M3 10h18M16 15h2"/>"""),
        ("Team", """<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0"/><circle cx="17" cy="9" r="2.5"/><path d="M16 14.5a5 5 0 0 1 5.5 5.5"/>"""),
        (Data, """<ellipse cx="12" cy="5.5" rx="8" ry="2.5"/><path d="M4 5.5v13c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5v-13M4 12c0 1.4 3.6 2.5 8 2.5s8-1.1 8-2.5"/>"""),
    ];

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet dashboard = AddSheet(book, Dashboard);
        dashboard.View = new SheetView { ShowGridLines = false, ShowHeaders = false, TabColor = "#2384F9", ActiveCell = new CellAddress(0, 0) };

        WriteProjects(AddSheet(book, Projects));

        foreach (string page in (string[])["Tasks", "Budget", "Team"])
        {
            Worksheet sheet = AddSheet(book, page);
            sheet.View = new SheetView { ShowGridLines = false, TabColor = "#3A4A6B" };

            sheet["B2"] = Text(page, 18, Navy, bold: true);
            sheet["B3"] = Text("A page the dashboard's menu links to.", 10, Grey);
            sheet["B5"] = Text("Back to the dashboard", 10, "#0563C1").Underline();
            Check(sheet.AddHyperlink("B5", Hyperlink.ToCell(Dashboard, "A1") with { Tooltip = "Dashboard" }));
        }

        WriteData(AddSheet(book, Data));
        DrawPage(dashboard);
        DrawCards(dashboard);
        DrawCharts(dashboard);

        return book;
    }

    private static void WriteProjects(Worksheet sheet)
    {
        sheet.View = new SheetView { ShowGridLines = false, TabColor = "#3A4A6B" };

        double[] widths = [2, 28, 16, 13, 11, 13, 13];

        for (int column = 0; column < widths.Length; column++)
        {
            sheet.SetColumnWidth(column, widths[column]);
        }

        sheet["B2"] = Text("Projects", 18, Navy, bold: true);
        sheet["B3"] = Text("Back to the dashboard", 10, "#0563C1").Underline();
        Check(sheet.AddHyperlink("B3", Hyperlink.ToCell(Dashboard, "A1") with { Tooltip = "Dashboard" }));
        string[] headers = ["Project", "Lead", "Status", "Done", "Budget", "Spent"];

        for (int column = 0; column < headers.Length; column++)
        {
            Cell header = Text(headers[column], 10, "#FFFFFF", bold: true).Background(Navy);

            if (column >= 3)
            {
                header = header.Align(HorizontalAlignment.Right);
            }

            sheet[4, column + 1] = header;
        }

        (string Name, string Lead, string Status, double Done, double Budget, double Spent)[] projects =
        [
            ("Website relaunch", "Sara Kim", "On track", 0.62, 850_000, 498_000),
            ("Customer portal", "Erik Sund", "At risk", 0.41, 1_200_000, 655_000),
            ("Warehouse scanners", "Lina Moberg", "On track", 0.88, 420_000, 371_000),
            ("Data platform", "Jonas Berg", "Late", 0.35, 2_100_000, 1_240_000),
            ("Mobile app 2.0", "Anna Lind", "On track", 0.57, 960_000, 530_000),
            ("Office move", "Omar Haddad", "Done", 1, 300_000, 287_000),
            ("Security review", "Mia Ek", "At risk", 0.48, 260_000, 171_000),
            ("Partner API", "Karl Nyström", "On track", 0.71, 540_000, 362_000),
        ];

        for (int index = 0; index < projects.Length; index++)
        {
            int row = 5 + index;
            sheet[row, 1] = Text(projects[index].Name, 10, Ink, bold: true);
            sheet[row, 2] = Text(projects[index].Lead, 10, Grey);
            sheet[row, 3] = Text(projects[index].Status);
            sheet[row, 4] = Number(projects[index].Done, "0%");
            sheet[row, 5] = Number(projects[index].Budget, "#,##0");
            sheet[row, 6] = Number(projects[index].Spent, "#,##0");

            if (index % 2 == 1)
            {
                Paint(sheet, new CellRange(row, 1, row, 6), Band);
            }
        }

        CellRange status = new(5, 3, 12, 3);
        sheet.AddConditionalFormat(status, ConditionalFormat.WhenTextContains("On track", new Style { FontColor = "#067647", Bold = true }));
        sheet.AddConditionalFormat(status, ConditionalFormat.WhenTextContains("At risk", new Style { FontColor = "#B45309", Bold = true }));
        sheet.AddConditionalFormat(status, ConditionalFormat.WhenTextContains("Late", new Style { FontColor = Red, Bold = true }));
        sheet.AddConditionalFormat(new CellRange(5, 4, 12, 4), ConditionalFormat.DataBar("#5EA2FF"));
        sheet.FreezePanes(5);
    }

    private static void WriteData(Worksheet data)
    {
        string[] headers = ["Month", "Completed", "Planned", "", "Stage", "Spacer", "Count", "", "Status", "Share", "", "Month", "Hours"];

        for (int column = 0; column < headers.Length; column++)
        {
            if (headers[column].Length > 0)
            {
                data[0, column] = Cell.Of(headers[column]).Bold();
            }
        }

        for (int month = 0; month < Months.Length; month++)
        {
            data[month + 1, 0] = Months[month];
            data[month + 1, 1] = Completed[month];
            data[month + 1, 2] = Planned[month];
            data[month + 1, 11] = Months[month];
            data[month + 1, 12] = Hours[month];
        }

        double widest = Funnel.Max(stage => stage.Count);

        for (int stage = 0; stage < Funnel.Length; stage++)
        {
            // The spacer pushes each bar to the middle: half of what it lacks of the widest.
            data[stage + 1, 4] = Funnel[stage].Stage;
            data[stage + 1, 5] = (widest - Funnel[stage].Count) / 2;
            data[stage + 1, 6] = Funnel[stage].Count;
        }

        data[1, 8] = "Done";
        data[1, 9] = Cell.Of(0.72).NumberFormat("0%");
        data[2, 8] = "Remaining";
        data[2, 9] = Cell.Of(0.28).NumberFormat("0%");
        data.SetColumnWidth(4, 12);
    }

    /// <summary>The page, the navigation column with its buttons, and the header.</summary>
    private static void DrawPage(Worksheet sheet)
    {
        Check(sheet.AddShape(new Shape(ShapeKind.Rectangle) { Fill = Fill.Solid(Page), Outline = Outline.None, Name = "Background" }, DrawingBounds.FromPoints(0, 0, 1212, 747)));
        Check(sheet.AddShape(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Gradient("#1E2740", "#121829"), Outline = Outline.None, Shadow = new Shadow { Blur = 10, Distance = 3 }, CornerRadius = 0.04, Name = "Navigation" }, DrawingBounds.FromPoints(12, 12, 170, 711)));

        Shape brand = new(ShapeKind.TextBox)
        {
            Text = new ShapeText([new TextRun("PORTFOLIO\n", new TextFont { Name = FontName, Bold = true, Size = 13, Color = "#FFFFFF" }), new TextRun("DASHBOARD", new TextFont { Name = FontName, Bold = true, Size = 13, Color = Accent })])
            {
                Horizontal = TextAlignment.Left,
                Vertical = TextAnchor.Center,
                MarginLeft = 0,
            },
            Description = "Portfolio dashboard",
        };

        Check(sheet.AddShape(brand, DrawingBounds.FromPoints(28, 26, 140, 44)));
        Check(sheet.AddShape(new Shape(ShapeKind.Line) { Outline = new Outline("#2E3A57", 1) }, DrawingBounds.FromPoints(28, 82, 138, 0)));

        for (int index = 0; index < Pages.Length; index++)
        {
            bool active = index == 0;
            Fill fill = Fill.None;
            string iconColor = Muted;
            string textColor = "#C9D2E3";

            if (active == true)
            {
                fill = Fill.Gradient(0, new GradientStop(0, "#2384F9"), new GradientStop(1, "#6D5BD0"));
                iconColor = "#FFFFFF";
                textColor = "#FFFFFF";
            }

            ShapeGroup button = new(
            [
                GroupItem.Of(new Shape(ShapeKind.RoundedRectangle) { Fill = fill, Outline = Outline.None, CornerRadius = 0.25 }, DrawingBounds.FromPoints(0, 0, 146, 36)),
                GroupItem.Of(Icon(Pages[index].Icon, iconColor, Pages[index].Name), DrawingBounds.FromPoints(12, 9, 18, 18)),
                GroupItem.Of(Label(Pages[index].Name, 10.5, textColor, bold: active), DrawingBounds.FromPoints(40, 0, 100, 36)),
            ])
            {
                Link = Hyperlink.ToCell(Pages[index].Name, "A1") with { Tooltip = "Go to " + Pages[index].Name },
                Name = "Go to " + Pages[index].Name,
            };

            Check(sheet.AddGroup(button, DrawingBounds.FromPoints(24, 100 + (index * 46), 146, 36)));
        }

        // A help card at the foot of the menu, with a link out.
        Shape help = new(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid("#FFFFFF", 0.94),
            Outline = new Outline("#2E3A57", 0.75),
            CornerRadius = 0.12,
            Text = new ShapeText([new TextRun("Need help?\n", new TextFont { Name = FontName, Size = 10, Bold = true, Color = "#FFFFFF" }), new TextRun("Open the documentation", new TextFont { Name = FontName, Size = 8.5, Color = Accent })])
            {
                Horizontal = TextAlignment.Left,
                MarginLeft = 12,
            },
            Link = Hyperlink.ToUrl("https://learn.microsoft.com/office/") with { Tooltip = "Documentation" },
        };

        Check(sheet.AddShape(help, DrawingBounds.FromPoints(24, 652, 146, 54)));

        Shape title = new(ShapeKind.TextBox)
        {
            Text = new ShapeText([new TextRun("Portfolio overview\n", new TextFont { Name = FontName, Size = 18, Bold = true, Color = "#FFFFFF" }), new TextRun("All active projects, updated 25 November 2026", new TextFont { Name = FontName, Size = 9, Color = Muted })])
            {
                Horizontal = TextAlignment.Left,
                MarginLeft = 0,
            },
        };

        Check(sheet.AddShape(title, DrawingBounds.FromPoints(200, 14, 500, 46)));
        Shape period = new(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid(Panel),
            Outline = new Outline("#2E3A57", 0.75),
            CornerRadius = 0.5,
            Text = new ShapeText("Q4 2026  ▾", new TextFont { Name = FontName, Size = 9.5, Bold = true, Color = "#FFFFFF" }),
        };

        Check(sheet.AddShape(period, DrawingBounds.FromPoints(1086, 22, 100, 28)));
    }

    private static void DrawCards(Worksheet sheet)
    {
        (string Title, string Value, string Change, string From, string To, string Icon)[] cards =
        [
            ("Active projects", "24", "▲ 3 since October", "#2384F9", "#5B4BC4", Pages[1].Icon),
            ("Tasks done", "780", "▲ 12 % on plan", "#0BA9A5", "#05638A", """<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.8 2.8L16.5 9"/>"""),
            ("Budget used", "62 %", "▼ 4 points under plan", "#E1541C", "#9C2F6B", Pages[3].Icon),
            ("People on projects", "18", "▲ 2 new this month", "#7B61FF", "#3A2C8F", Pages[4].Icon),
        ];

        for (int index = 0; index < cards.Length; index++)
        {
            double left = 200 + (index * 250);
            Shape card = new(ShapeKind.RoundedRectangle)
            {
                Fill = Fill.Gradient(45, new GradientStop(0, cards[index].From), new GradientStop(1, cards[index].To)),
                Outline = Outline.None,
                Shadow = new Shadow { Blur = 8, Distance = 4, Transparency = 0.6 },
                CornerRadius = 0.1,
                Text = new ShapeText(
                [
                    new TextRun(cards[index].Title + "\n", new TextFont { Name = FontName, Size = 10, Color = "#E6ECF5" }),
                    new TextRun(cards[index].Value + "\n", new TextFont { Name = FontName, Size = 26, Bold = true, Color = "#FFFFFF" }),
                    new TextRun(cards[index].Change, new TextFont { Name = FontName, Size = 8.5, Color = "#E6ECF5" }),
                ])
                {
                    Horizontal = TextAlignment.Left,
                    MarginLeft = 16,
                },
                Name = cards[index].Title + " card",
                Description = cards[index].Title + ": " + cards[index].Value,
            };

            Check(sheet.AddShape(card, DrawingBounds.FromPoints(left, 70, 236, 102)));
            Check(sheet.AddShape(new Shape(ShapeKind.Ellipse) { Fill = Fill.Solid("#FFFFFF", 0.8), Outline = Outline.None }, DrawingBounds.FromPoints(left + 182, 84, 38, 38)));
            Check(sheet.AddPicture(Icon(cards[index].Icon, "#FFFFFF", cards[index].Title), DrawingBounds.FromPoints(left + 191, 93, 20, 20)));
        }
    }

    private static void DrawCharts(Worksheet sheet)
    {
        (double Left, double Top, double Width, double Height)[] panels = [(200, 186, 610, 270), (824, 186, 362, 270), (200, 470, 420, 253), (634, 470, 552, 253)];

        foreach ((double left, double top, double width, double height) in panels)
        {
            Check(sheet.AddShape(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Solid(Panel), Outline = new Outline("#253049", 0.75), CornerRadius = 0.04 }, DrawingBounds.FromPoints(left, top, width, height)));
        }

        TextFont axisFont = new() { Name = FontName, Color = Muted, Size = 9 };

        TextFont titleFont = new() { Name = FontName, Color = "#FFFFFF", Size = 12, Bold = true };

        ChartAxis valueAxis = new() { Gridlines = true, GridlineOutline = new Outline("#2A3550", 0.5), Line = Outline.None, Font = axisFont };

        ChartAxis categoryAxis = new() { Line = new Outline("#34405E"), Font = axisFont };

        // Tasks per month: columns with August picked out, the plan as a smooth line.
        string?[] highlight = new string?[12];
        highlight[7] = "#2DD4BF";
        Chart tasks = new(ChartType.Column)
        {
            Title = "Tasks completed per month",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("Completed", $"{Data}!$B$2:$B$13", $"{Data}!$A$2:$A$13", "#3B4A6E") { PointColors = highlight, DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, Font = new TextFont { Name = FontName, Color = "#FFFFFF", Size = 8 } } },
                new ChartSeries("Planned", $"{Data}!$C$2:$C$13", $"{Data}!$A$2:$A$13") { Type = ChartType.Line, Smooth = true, Marker = MarkerStyle.None, Outline = new Outline("#F97316", 2.25), Shadow = new Shadow { Angle = 90, Transparency = 0.5 } },
            ],
            GapWidth = 55,
            Legend = LegendPosition.Top,
            LegendFont = axisFont,
            Background = Fill.None,
            Border = Outline.None,
            Font = axisFont,
            ValueAxis = valueAxis with { Minimum = 0, Maximum = 100, MajorUnit = 25 },
            CategoryAxis = categoryAxis,
            Description = "Tasks completed per month, with the plan as a line",
        };

        Check(sheet.AddChart(tasks, DrawingBounds.FromPoints(210, 194, 590, 254)));

        // The gauge: a doughnut with a thick ring and the share in the middle.
        Chart gauge = new(ChartType.Doughnut)
        {
            Title = "Portfolio progress",
            TitleFont = titleFont,
            Series = [new ChartSeries("Share", $"{Data}!$J$2:$J$3", $"{Data}!$I$2:$I$3") { PointColors = ["#2DD4BF", "#2A3550"], Outline = Outline.None }],
            HoleSize = 74,
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            Description = "72 % of the planned work is done",
        };

        Check(sheet.AddChart(gauge, DrawingBounds.FromPoints(834, 194, 342, 228)));
        Check(sheet.AddShape(Label("72 %", 30, "#FFFFFF", bold: true, TextAlignment.Center), DrawingBounds.FromPoints(945, 300, 120, 40)));
        Check(sheet.AddShape(Label("of planned work done", 9, Muted, horizontal: TextAlignment.Center), DrawingBounds.FromPoints(925, 338, 160, 16)));
        Shape legend = new(ShapeKind.TextBox)
        {
            Text = new ShapeText(
            [
                new TextRun("●  ", new TextFont { Name = FontName, Size = 10, Color = "#2DD4BF" }),
                new TextRun("Done 72 %        ", new TextFont { Name = FontName, Size = 9, Color = Muted }),
                new TextRun("●  ", new TextFont { Name = FontName, Size = 10, Color = "#3B4A6E" }),
                new TextRun("Remaining 28 %", new TextFont { Name = FontName, Size = 9, Color = Muted }),
            ]),
        };

        Check(sheet.AddShape(legend, DrawingBounds.FromPoints(860, 424, 290, 20)));

        // The funnel: an invisible spacer and the counts, stacked, so each bar sits in the middle.
        Chart funnel = new(ChartType.Bar)
        {
            Title = "Work items by stage",
            TitleFont = titleFont,
            Stacked = true,
            Series =
            [
                new ChartSeries("Spacer", $"{Data}!$F$2:$F$6", $"{Data}!$E$2:$E$6") { Fill = Fill.None },
                new ChartSeries("Count", $"{Data}!$G$2:$G$6", $"{Data}!$E$2:$E$6", "#2384F9")
                {
                    PointColors = ["#2DD4BF", "#14B8A6", "#1E8FD8", "#2384F9", "#6D5BD0"],
                    DataLabels = new DataLabels { Position = DataLabelPosition.Center, Font = new TextFont { Name = FontName, Color = "#FFFFFF", Size = 9, Bold = true } },
                },
            ],
            GapWidth = 18,
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            ValueAxis = new ChartAxis { Visible = false, Gridlines = false },
            CategoryAxis = new ChartAxis { Line = Outline.None, Font = axisFont },
            Description = "Backlog, planned, in progress, in review and delivered work items",
        };

        Check(sheet.AddChart(funnel, DrawingBounds.FromPoints(210, 478, 400, 237)));

        // Hours logged: a gradient area under a smooth line.
        Chart hours = new(ChartType.Area)
        {
            Title = "Hours logged per month",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("Hours", $"{Data}!$M$2:$M$13", $"{Data}!$L$2:$L$13") { Fill = Fill.Gradient(90, new GradientStop(0, "#2384F9", 0.45), new GradientStop(1, "#2384F9", 1)), Outline = Outline.None },
                new ChartSeries("Trend", $"{Data}!$M$2:$M$13", $"{Data}!$L$2:$L$13") { Type = ChartType.Line, Smooth = true, Marker = MarkerStyle.Circle, MarkerSize = 6, Outline = new Outline(Accent, 2) },
            ],
            Legend = LegendPosition.None,
            Background = Fill.None,
            Border = Outline.None,
            Font = axisFont,
            ValueAxis = valueAxis with { NumberFormat = "#,##0" },
            CategoryAxis = categoryAxis,
            Description = "Hours logged per month",
        };

        Check(sheet.AddChart(hours, DrawingBounds.FromPoints(644, 478, 532, 237)));
    }
}
