using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A scorecard for customers and staff: four doughnut gauges with the figure in the middle,
/// a table per team with traffic lights, arrows, colour scales and data bars, and the net
/// promoter score over the year against its target. The drawings are placed in points, so the
/// rows are kept at their standard height and the headline is a text box.
/// </summary>
public static class Scorecard
{
    /// <summary>The sheet with the scorecard.</summary>
    public const string Sheet = "Scorecard";

    /// <summary>The cells the gallery picture covers.</summary>
    public const string PictureRange = "A1:U29";

    private const string Data = "Data";
    private const double BlockWidth = 4;
    private const double ColumnWidth = 8.3;
    private const double GapWidth = 2;

    private static readonly (string Title, double Value, string Shown, string Note, string Color)[] Gauges =
    [
        ("Customer satisfaction", 0.87, "87 %", "Target 85 %, up 3 points", Teal),
        ("Net promoter score", 0.71, "+42", "Scale -100 to +100, target +40", Blue),
        ("Staff engagement (eNPS)", 0.655, "+31", "Up from +24 in Q2", "#7C3AED"),
        ("Staff retention", 0.94, "94 %", "12-month rolling, target 92 %", Amber),
    ];

    private static readonly (string Team, double Satisfaction, double Promoter, double Response, double Engagement, double Sick, double Turnover)[] Teams =
    [
        ("Support North", 0.91, 48, 2.1, 38, 0.031, 0.06),
        ("Support South", 0.84, 35, 3.4, 22, 0.052, 0.11),
        ("Onboarding", 0.89, 51, 1.6, 44, 0.024, 0.04),
        ("Field service", 0.78, 19, 6.2, 12, 0.067, 0.15),
        ("Key accounts", 0.93, 57, 1.2, 41, 0.028, 0.05),
        ("Webshop", 0.86, 40, 2.8, 30, 0.039, 0.08),
        ("Billing", 0.81, 27, 4.1, 18, 0.046, 0.09),
        ("Partners", 0.88, 44, 2.4, 35, 0.033, 0.07),
    ];

    private static readonly double[] Promoters = [31, 33, 30, 35, 36, 38, 37, 39, 42, 41, 43, 45];

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = AddSheet(book, Sheet);
        Worksheet data = AddSheet(book, Data);
        sheet.View = new SheetView { ShowGridLines = false, ShowHeaders = false, TabColor = "#7C3AED", ActiveCell = new CellAddress(0, 0) };

        sheet.SetColumnWidth(0, GapWidth);

        for (int column = 1; column <= 19; column++)
        {
            sheet.SetColumnWidth(column, ColumnWidth);
        }

        foreach (int gap in (int[])[5, 10, 15, 20])
        {
            sheet.SetColumnWidth(gap, GapWidth);
        }

        Paint(sheet, new CellRange(0, 0, 2, 20), Navy);
        Paint(sheet, new CellRange(3, 0, 28, 20), Page);
        Shape headline = new(ShapeKind.TextBox)
        {
            Text = new ShapeText(
            [
                new TextRun("Customer and staff scorecard\n", new TextFont { Name = FontName, Size = 16, Bold = true, Color = "#FFFFFF" }),
                new TextRun("Q3 2026, all teams. Green is on target, amber close, red needs attention.", new TextFont { Name = FontName, Size = 9, Color = "#C7D2FE" }),
            ])
            {
                Horizontal = HorizontalAlignment.Left,
                MarginLeft = 0,
            },
        };

        Check(sheet.AddShape(headline, DrawingBounds.FromPoints(ColumnPoints(GapWidth), 2, 600, 41)));
        WriteData(data);
        AddGauges(sheet);
        AddTeams(sheet);
        AddTrend(sheet);

        return book;
    }

    private static void WriteData(Worksheet data)
    {
        data["A1"] = Text("Gauge", 10, Navy, bold: true);
        data["B1"] = Text("Value", 10, Navy, bold: true);
        data["C1"] = Text("Rest", 10, Navy, bold: true);

        for (int index = 0; index < Gauges.Length; index++)
        {
            data[index + 1, 0] = Gauges[index].Title;
            data[index + 1, 1] = Cell.Of(Gauges[index].Value).NumberFormat("0.0%");
            data[index + 1, 2] = Cell.FromFormula($"1-B{index + 2}").NumberFormat("0.0%");
        }

        data["E1"] = Text("Month", 10, Navy, bold: true);
        data["F1"] = Text("NPS", 10, Navy, bold: true);
        data["G1"] = Text("Target", 10, Navy, bold: true);
        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        for (int month = 0; month < months.Length; month++)
        {
            data[month + 1, 4] = months[month];
            data[month + 1, 5] = Promoters[month];
            data[month + 1, 6] = 40;
        }

        data.SetColumnWidth(0, 24);
    }

    /// <summary>The left edge of block 0 to 3 (four columns and a gap each), in points.</summary>
    private static double BlockLeft(int block)
    {
        return ColumnPoints(GapWidth) + (block * ((BlockWidth * ColumnPoints(ColumnWidth)) + ColumnPoints(GapWidth)));
    }

    private static void AddGauges(Worksheet sheet)
    {
        double blockWidth = BlockWidth * ColumnPoints(ColumnWidth);

        for (int index = 0; index < Gauges.Length; index++)
        {
            int column = 1 + (index * 5);
            Paint(sheet, new CellRange(4, column, 14, column + 3), "#FFFFFF");
            sheet[4, column] = Text(Gauges[index].Title, 9, Grey, bold: true).Background("#FFFFFF");
            sheet[14, column] = Text(Gauges[index].Note, 8, Grey).Background("#FFFFFF");

            double left = BlockLeft(index);
            double size = 118;
            double gaugeLeft = left + ((blockWidth - size) / 2);
            double gaugeTop = (5 * RowPoints) + 6;

            Chart gauge = new(ChartType.Doughnut)
            {
                Series =
                [
                    new ChartSeries(Gauges[index].Title, $"{Data}!$B${index + 2}:$C${index + 2}")
                    {
                        PointColors = [Gauges[index].Color, "#E5E7EB"],
                        Outline = Outline.None,
                    },
                ],
                HoleSize = 74,
                Legend = LegendPosition.None,
                Background = Fill.None,
                Border = Outline.None,
                Description = Gauges[index].Title + ": " + Gauges[index].Shown,
            };

            Check(sheet.AddChart(gauge, DrawingBounds.FromPoints(gaugeLeft, gaugeTop, size, size)));
            Shape figure = Label(Gauges[index].Shown, 20, Navy, bold: true, HorizontalAlignment.Center);
            Check(sheet.AddShape(figure, DrawingBounds.FromPoints(gaugeLeft, gaugeTop + (size / 2) - 15, size, 30)));
        }
    }

    private static void AddTeams(Worksheet sheet)
    {
        int top = 16;
        Paint(sheet, new CellRange(top, 1, top + Teams.Length + 3, 9), "#FFFFFF");
        sheet[top, 1] = Text("Teams", 11, Navy, bold: true).Background("#FFFFFF");

        // Name in B (running on into the empty C), then one column per measure; F is the gap.
        (int Column, string Heading)[] headings = [(1, "Team"), (3, "CSAT"), (4, "NPS"), (6, "Response"), (7, "eNPS"), (8, "Sick leave"), (9, "Turnover")];

        foreach ((int column, string heading) in headings)
        {
            Cell cell = Text(heading, 8, Grey, bold: true).Background("#FFFFFF");

            if (column > 1)
            {
                cell = cell.Align(HorizontalAlignment.Right);
            }

            sheet[top + 1, column] = cell;
        }

        for (int index = 0; index < Teams.Length; index++)
        {
            int row = top + 2 + index;
            string background = "#FFFFFF";

            if (index % 2 == 1)
            {
                background = Band;
            }

            sheet[row, 1] = Text(Teams[index].Team);
            sheet[row, 3] = Number(Teams[index].Satisfaction, "0%");
            sheet[row, 4] = Number(Teams[index].Promoter, "+0;-0;0");
            sheet[row, 6] = Number(Teams[index].Response, "0.0\" h\"");
            sheet[row, 7] = Number(Teams[index].Engagement, "+0;-0;0");
            sheet[row, 8] = Number(Teams[index].Sick, "0.0%");
            sheet[row, 9] = Number(Teams[index].Turnover, "0%");
            Paint(sheet, new CellRange(row, 1, row, 9), background);
        }

        int first = top + 2;
        int last = top + 1 + Teams.Length;
        sheet.AddConditionalFormat(new CellRange(first, 3, last, 3), ConditionalFormat.IconSet(IconSetKind.ThreeTrafficLights));
        sheet.AddConditionalFormat(new CellRange(first, 4, last, 4), ConditionalFormat.IconSet(IconSetKind.ThreeArrows));
        sheet.AddConditionalFormat(new CellRange(first, 6, last, 6), ConditionalFormat.ColorScale("#CDEBD9", "#FFF3CD", "#F8D7DA"));
        sheet.AddConditionalFormat(new CellRange(first, 7, last, 7), ConditionalFormat.DataBar("#C4B5FD"));
        sheet.AddConditionalFormat(new CellRange(first, 8, last, 8), ConditionalFormat.IconSet(IconSetKind.ThreeSymbols, reverse: true));
        sheet.AddConditionalFormat(new CellRange(first, 9, last, 9), ConditionalFormat.ColorScale("#CDEBD9", "#FFF3CD", "#F8D7DA"));
        sheet[last + 1, 1] = Text("Response: median hours to first answer. Sick leave and turnover: share of staff.", 8, Grey).Background("#FFFFFF");
    }

    private static void AddTrend(Worksheet sheet)
    {
        TextFont font = new() { Name = FontName, Size = 9, Color = Grey };

        Chart trend = new(ChartType.Line)
        {
            Title = "Net promoter score, 2026",
            TitleFont = new TextFont { Name = FontName, Size = 11, Bold = true, Color = Navy },
            Series =
            [
                new ChartSeries("NPS", $"{Data}!$F$2:$F$13", $"{Data}!$E$2:$E$13")
                {
                    Type = ChartType.Area,
                    Fill = Fill.Gradient(90, new GradientStop(0, Blue, 0.6), new GradientStop(1, Blue, 0.95)),
                    Outline = Outline.None,
                },
                new ChartSeries("NPS", $"{Data}!$F$2:$F$13", $"{Data}!$E$2:$E$13")
                {
                    Smooth = true,
                    Marker = MarkerStyle.Circle,
                    MarkerSize = 5,
                    Outline = new Outline(Blue, 2.25),
                    DataLabels = new DataLabels { Position = DataLabelPosition.Above, Font = new TextFont { Name = FontName, Size = 8, Color = Navy } },
                },
                new ChartSeries("Target", $"{Data}!$G$2:$G$13", $"{Data}!$E$2:$E$13")
                {
                    Marker = MarkerStyle.None,
                    Outline = new Outline(Amber, 1.5) { Dash = LineDash.Dash },
                },
            ],
            Legend = LegendPosition.None,
            Font = font,
            Background = Fill.Solid("#FFFFFF"),
            Border = Outline.None,
            ValueAxis = new ChartAxis { Minimum = 20, Maximum = 50, MajorUnit = 10, Line = Outline.None, GridlineOutline = new Outline(Line, 0.5), Font = font },
            CategoryAxis = new ChartAxis { Line = new Outline("#D1D5DB"), Font = font },
            Description = "Net promoter score per month in 2026 against the target of 40",
        };

        Check(sheet.AddChart(trend, new CellRange(16, 11, 16 + Teams.Length + 3, 19)));
    }
}
