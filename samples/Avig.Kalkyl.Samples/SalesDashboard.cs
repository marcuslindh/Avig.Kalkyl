using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A sales dashboard driven by pivot tables. The order lines are on Data; Pivots has three
/// pivot tables (by month, region and product, each with the years side by side); Region and
/// Product slicers on the dashboard filter all three. The key figures and the chart data are
/// GETPIVOTDATA formulas, so they follow the slicers when Excel recalculates, and the charts
/// read those cells.
/// </summary>
public static class SalesDashboard
{
    /// <summary>The sheet with the dashboard.</summary>
    public const string Dashboard = "Dashboard";

    /// <summary>The cells the gallery picture covers.</summary>
    public const string PictureRange = "A1:U47";

    private const string Data = "Data";
    private const string Pivots = "Pivots";
    private const string ChartData = "Chart data";
    private const string Money = "#,##0";

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    private static readonly string[] Regions = ["North", "South", "East", "West"];
    private static readonly string[] Products = ["Bikes", "E-bikes", "Accessories", "Service"];

    // Average price per unit, SEK, and how the months and regions differ.
    private static readonly double[] Prices = [8_900, 21_500, 450, 1_200];
    private static readonly double[] Season = [0.55, 0.6, 0.85, 1.15, 1.35, 1.4, 1.3, 1.2, 1.0, 0.8, 0.6, 0.75];
    private static readonly double[] RegionSize = [0.7, 1.35, 1.0, 1.15];
    private static readonly double[] ProductVolume = [42, 14, 260, 95];

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet dashboard = AddSheet(book, Dashboard);
        Worksheet chartData = AddSheet(book, ChartData);
        Worksheet pivots = AddSheet(book, Pivots);
        Worksheet data = AddSheet(book, Data);
        int rows = WriteData(data);
        string source = $"{Data}!A1:F{rows + 1}";

        // Each pivot table has the years across, so the dashboard can compare them.
        PivotTable byMonth = new("SalesByMonth", source)
        {
            Rows = ["Month"],
            Columns = ["Year"],
            Values = [new PivotValue("Sales") { NumberFormat = Money }],
            GrandTotalColumn = false,
            Style = "PivotStyleLight16",
        };

        PivotTable byRegion = new("SalesByRegion", source)
        {
            Rows = ["Region"],
            Columns = ["Year"],
            Values = [new PivotValue("Sales") { NumberFormat = Money }],
            GrandTotalColumn = false,
            Style = "PivotStyleLight16",
        };

        PivotTable byProduct = new("SalesByProduct", source)
        {
            Rows = ["Product"],
            Columns = ["Year"],
            Values = [new PivotValue("Sales") { NumberFormat = Money }, new PivotValue("Units")],
            GrandTotalColumn = false,
            Style = "PivotStyleLight16",
        };

        Check(pivots.AddPivotTable(byMonth, "A3"));
        Check(pivots.AddPivotTable(byRegion, "F3"));
        Check(pivots.AddPivotTable(byProduct, "K3"));
        pivots["A1"] = Text("The dashboard reads these pivot tables with GETPIVOTDATA.", 10, Grey);

        WriteChartData(chartData);
        DrawDashboard(dashboard);

        return book;
    }

    /// <summary>Two years of monthly sales per region and product, from a fixed seed.</summary>
    private static int WriteData(Worksheet data)
    {
        string[] headings = ["Year", "Month", "Region", "Product", "Sales", "Units"];

        for (int column = 0; column < headings.Length; column++)
        {
            data[0, column] = Text(headings[column], 10, "#FFFFFF", bold: true).Background(Navy);
        }

        Random random = new(7);
        int row = 1;

        foreach (int year in (int[])[2025, 2026])
        {
            double growth = 1;

            if (year == 2026)
            {
                growth = 1.12;
            }

            for (int month = 0; month < Months.Length; month++)
            {
                for (int region = 0; region < Regions.Length; region++)
                {
                    for (int product = 0; product < Products.Length; product++)
                    {
                        double units = Math.Round(ProductVolume[product] * Season[month] * RegionSize[region] * growth * (0.85 + (random.NextDouble() * 0.3)));
                        data[row, 0] = year;
                        data[row, 1] = Months[month];
                        data[row, 2] = Regions[region];
                        data[row, 3] = Products[product];
                        data[row, 4] = Cell.Of(units * Prices[product]).NumberFormat(Money);
                        data[row, 5] = units;
                        row++;
                    }
                }
            }
        }

        data.FreezePanes(1);
        data.AutoFilter = true;
        data.SetColumnWidth(3, 12);
        data.SetColumnWidth(4, 12);

        return row - 1;
    }

    /// <summary>The cells the charts draw: GETPIVOTDATA formulas over the pivot tables.</summary>
    private static void WriteChartData(Worksheet sheet)
    {
        sheet["A1"] = Text("Month", 10, Navy, bold: true);
        sheet["B1"] = Text("2025", 10, Navy, bold: true);
        sheet["C1"] = Text("2026", 10, Navy, bold: true);

        for (int month = 0; month < Months.Length; month++)
        {
            int row = month + 2;
            sheet[month + 1, 0] = Months[month];
            sheet[month + 1, 1] = Pivot($"\"Sum of Sales\",{Pivots}!$A$3,\"Month\",$A{row},\"Year\",2025");
            sheet[month + 1, 2] = Pivot($"\"Sum of Sales\",{Pivots}!$A$3,\"Month\",$A{row},\"Year\",2026");
        }

        sheet["E1"] = Text("Region", 10, Navy, bold: true);
        sheet["F1"] = Text("2025", 10, Navy, bold: true);
        sheet["G1"] = Text("2026", 10, Navy, bold: true);

        for (int region = 0; region < Regions.Length; region++)
        {
            int row = region + 2;
            sheet[region + 1, 4] = Regions[region];
            sheet[region + 1, 5] = Pivot($"\"Sum of Sales\",{Pivots}!$F$3,\"Region\",$E{row},\"Year\",2025");
            sheet[region + 1, 6] = Pivot($"\"Sum of Sales\",{Pivots}!$F$3,\"Region\",$E{row},\"Year\",2026");
        }

        sheet["I1"] = Text("Product", 10, Navy, bold: true);
        sheet["J1"] = Text("2026", 10, Navy, bold: true);

        for (int product = 0; product < Products.Length; product++)
        {
            int row = product + 2;
            sheet[product + 1, 8] = Products[product];
            sheet[product + 1, 9] = Pivot($"\"Sum of Sales\",{Pivots}!$K$3,\"Product\",$I{row},\"Year\",2026");
        }

        sheet.SetColumnWidth(5, 12);
        sheet.SetColumnWidth(6, 12);
        sheet.SetColumnWidth(9, 12);
    }

    /// <summary>A GETPIVOTDATA formula that gives 0 for an item the slicers filter away.</summary>
    private static Cell Pivot(string arguments)
    {
        return Cell.FromFormula($"IFERROR(GETPIVOTDATA({arguments}),0)").NumberFormat(Money);
    }

    private static void DrawDashboard(Worksheet sheet)
    {
        sheet.View = new SheetView { ShowGridLines = false, ShowHeaders = false, Zoom = 100, TabColor = Blue, ActiveCell = new CellAddress(0, 0) };

        // Columns: a margin, then four blocks of four columns with narrow gaps between them.
        sheet.SetColumnWidth(0, 2);

        for (int column = 1; column <= 19; column++)
        {
            sheet.SetColumnWidth(column, 8.3);
        }

        foreach (int gap in (int[])[5, 10, 15])
        {
            sheet.SetColumnWidth(gap, 2);
        }

        sheet.SetColumnWidth(20, 2);
        sheet.SetColumnWidth(18, 10);
        Paint(sheet, new CellRange(0, 0, 2, 20), Navy);
        Paint(sheet, new CellRange(3, 0, 46, 20), Page);
        sheet["B2"] = Text("Sales dashboard", 16, "#FFFFFF", bold: true).Background(Navy);
        sheet["B3"] = Text("Cycles and e-bikes, SEK. Pick a region or a product; the figures follow.", 9, "#C7D2FE").Background(Navy);

        // The slicers filter all three pivot tables, and with them every GETPIVOTDATA below.
        string[] tables = ["SalesByMonth", "SalesByRegion", "SalesByProduct"];
        Check(sheet.AddSlicer(new Slicer("Region", tables) { Columns = 4, Style = "SlicerStyleLight1", Caption = "Region" }, "B5:J8"));
        Check(sheet.AddSlicer(new Slicer("Product", tables) { Columns = 4, Style = "SlicerStyleLight1", Caption = "Product" }, "L5:T8"));

        AddKeyFigures(sheet);
        AddCharts(sheet);
    }

    private static void AddKeyFigures(Worksheet sheet)
    {
        string month = $"{Pivots}!$A$3";
        string product = $"{Pivots}!$K$3";
        string sales2026 = $"GETPIVOTDATA(\"Sum of Sales\",{month},\"Year\",2026)";
        string sales2025 = $"GETPIVOTDATA(\"Sum of Sales\",{month},\"Year\",2025)";
        string units2026 = $"GETPIVOTDATA(\"Sum of Units\",{product},\"Year\",2026)";
        string units2025 = $"GETPIVOTDATA(\"Sum of Units\",{product},\"Year\",2025)";

        // Text formulas (FIXED gives the reader's own separators), so a large figure runs on
        // into the empty cells beside it instead of turning into ###.
        string Growth(string ratio)
        {
            return $"IF({ratio}>=0,\"+\",\"\")&FIXED(({ratio})*100,1)&\" % vs 2025\"";
        }

        string best = $"MAX('{ChartData}'!$C$2:$C$13)";

        (string Title, string Value, string Note, int Column)[] figures =
        [
            ("Sales 2026", $"FIXED({sales2026}/1000000,1)&\" M SEK\"", Growth($"{sales2026}/{sales2025}-1"), 1),
            ("Units sold", $"FIXED({units2026},0)", Growth($"{units2026}/{units2025}-1"), 6),
            ("Average price", $"FIXED({sales2026}/{units2026},0)&\" SEK\"", Growth($"({sales2026}/{units2026})/({sales2025}/{units2025})-1"), 11),
            ("Best month 2026", $"INDEX('{ChartData}'!$A$2:$A$13,MATCH({best},'{ChartData}'!$C$2:$C$13,0))", $"FIXED({best}/1000000,2)&\" M SEK in sales\"", 16),
        ];

        foreach ((string title, string value, string note, int column) in figures)
        {
            Paint(sheet, new CellRange(9, column, 12, column + 3), "#FFFFFF");
            sheet[9, column] = Text(title, 9, Grey).Background("#FFFFFF");
            sheet[10, column] = Cell.FromFormula("IFERROR(" + value + ",\"–\")").Font(FontName).FontSize(20).Bold().FontColor(Navy).Background("#FFFFFF");
            sheet[11, column] = Cell.FromFormula("IFERROR(" + note + ",\"\")").Font(FontName).FontSize(9).FontColor(Teal).Background("#FFFFFF");
        }

        sheet.AddConditionalFormat(new CellRange(11, 1, 11, 19), ConditionalFormat.WhenTextContains("-", new Style { FontColor = Red }));
    }

    private static void AddCharts(Worksheet sheet)
    {
        TextFont font = new() { Name = FontName, Size = 9, Color = Grey };

        TextFont titleFont = new() { Name = FontName, Size = 11, Bold = true, Color = Navy };

        ChartAxis valueAxis = new() { Line = Outline.None, GridlineOutline = new Outline(Line, 0.5), NumberFormat = "#,##0,\" k\"", Font = font };

        ChartAxis categoryAxis = new() { Line = new Outline("#D1D5DB"), Font = font };

        string data = $"'{ChartData}'!";

        Chart months = new(ChartType.Column)
        {
            Title = "Sales per month, 2026 against 2025",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("2026", data + "$C$2:$C$13", data + "$A$2:$A$13", Blue),
                new ChartSeries("2025", data + "$B$2:$B$13", data + "$A$2:$A$13")
                {
                    Type = ChartType.Line,
                    Smooth = true,
                    Marker = MarkerStyle.Circle,
                    MarkerSize = 5,
                    Outline = new Outline("#94A3B8", 2) { Dash = LineDash.Dash },
                },
            ],
            GapWidth = 55,
            Legend = LegendPosition.Top,
            LegendFont = font,
            Font = font,
            Background = Fill.Solid("#FFFFFF"),
            Border = Outline.None,
            ValueAxis = valueAxis,
            CategoryAxis = categoryAxis,
            Description = "Sales per month in 2026 as columns, 2025 as a dashed line",
        };

        Check(sheet.AddChart(months, "B15:O31"));

        Chart products = new(ChartType.Doughnut)
        {
            Title = "Sales by product, 2026",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("2026", data + "$J$2:$J$5", data + "$I$2:$I$5")
                {
                    PointColors = [Navy, Blue, LightBlue, Teal],
                    Outline = new Outline("#FFFFFF", 1.5),
                    DataLabels = new DataLabels { ShowValue = false, ShowPercent = true, NumberFormat = "0%", Font = new TextFont { Name = FontName, Size = 8, Bold = true, Color = "#FFFFFF" } },
                },
            ],
            HoleSize = 55,
            Legend = LegendPosition.Bottom,
            LegendFont = font,
            Background = Fill.Solid("#FFFFFF"),
            Border = Outline.None,
            Description = "Share of 2026 sales per product",
        };

        Check(sheet.AddChart(products, "Q15:T31"));

        Chart regions = new(ChartType.Bar)
        {
            Title = "Sales by region",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("2025", data + "$F$2:$F$5", data + "$E$2:$E$5", LightBlue),
                new ChartSeries("2026", data + "$G$2:$G$5", data + "$E$2:$E$5", Navy)
                {
                    DataLabels = new DataLabels { Position = DataLabelPosition.OutsideEnd, NumberFormat = "#,##0,\" k\"", Font = font },
                },
            ],
            GapWidth = 45,
            Legend = LegendPosition.Top,
            LegendFont = font,
            Font = font,
            Background = Fill.Solid("#FFFFFF"),
            Border = Outline.None,
            ValueAxis = valueAxis with { Visible = false, Gridlines = false },
            CategoryAxis = categoryAxis with { Line = Outline.None },
            Description = "Sales per region, 2025 and 2026",
        };

        Check(sheet.AddChart(regions, "B33:I46"));

        // A table of the regions next to the chart: the same cells, with the growth in green or red.
        Paint(sheet, new CellRange(32, 11, 45, 19), "#FFFFFF");
        sheet[33, 11] = Text("Region", 9, Grey, bold: true).Background("#FFFFFF");
        sheet[33, 13] = Text("2025, k", 9, Grey, bold: true).Background("#FFFFFF").Align(HorizontalAlignment.Right);
        sheet[33, 14] = Text("2026, k", 9, Grey, bold: true).Background("#FFFFFF").Align(HorizontalAlignment.Right);
        sheet[33, 18] = Text("Growth", 9, Grey, bold: true).Background("#FFFFFF").Align(HorizontalAlignment.Right);

        for (int region = 0; region < Regions.Length; region++)
        {
            int row = 35 + (region * 2);
            int source = region + 2;
            sheet[row, 11] = Text(Regions[region], 10, Ink, bold: true).Background("#FFFFFF");
            sheet[row, 13] = Formula($"{data}$F${source}", "#,##0,", Grey).Background("#FFFFFF");
            sheet[row, 14] = Formula($"{data}$G${source}", "#,##0,", Ink, bold: true).Background("#FFFFFF");
            sheet[row, 18] = Formula($"IFERROR({data}$G${source}/{data}$F${source}-1,0)", "+0.0%;-0.0%", Ink).Background("#FFFFFF");
        }

        sheet[43, 11] = Text("Growth against 2025 per region, from the same GETPIVOTDATA cells.", 8, Grey).Background("#FFFFFF");
        CellRange growth = new(35, 18, 41, 18);
        sheet.AddConditionalFormat(growth, ConditionalFormat.WhenCell(ComparisonOperator.GreaterThanOrEqual, 0, new Style { FontColor = "#067647", Bold = true }));
        sheet.AddConditionalFormat(growth, ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, new Style { FontColor = Red, Bold = true }));
    }
}
