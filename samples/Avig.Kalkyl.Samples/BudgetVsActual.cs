using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// Budget against actual costs per department: a table with the variance as formulas, red and
/// green conditional formatting on the variance, data bars on the actual costs, a total row,
/// three key figures as cards, budget and actual side by side in a bar chart, and the share of
/// the budget each department has used, red where it is over.
/// </summary>
public static class BudgetVsActual
{
    /// <summary>The sheet with the report.</summary>
    public const string Report = "Budget";

    /// <summary>The cells the gallery picture covers.</summary>
    public const string PictureRange = "A1:R39";

    private static readonly (string Department, double Budget, double Actual)[] Departments =
    [
        ("Sales", 4200, 4385),
        ("Marketing", 2600, 2310),
        ("Product", 1900, 1965),
        ("Engineering", 6800, 6540),
        ("Support", 1450, 1520),
        ("Operations", 2300, 2185),
        ("Finance", 980, 955),
        ("People", 860, 905),
        ("IT", 1700, 1890),
        ("Legal", 540, 470),
        ("Facilities", 1250, 1215),
        ("Travel", 720, 840),
    ];

    private const int HeaderRow = 9;
    private const int ChartDataColumn = 19;
    private const string Thousands = "#,##0";
    private const string Signed = "+#,##0;-#,##0;0";
    private const string SignedPercent = "+0.0%;-0.0%;0.0%";

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = AddSheet(book, Report);
        sheet.View = new SheetView { ShowGridLines = false, TabColor = Blue, ActiveCell = new CellAddress(HeaderRow + 1, 1) };

        sheet.SetColumnWidth(0, 2);
        sheet.SetColumnWidth(1, 16);

        for (int column = 2; column <= 5; column++)
        {
            sheet.SetColumnWidth(column, 11);
        }

        sheet.SetColumnWidth(6, 3);
        sheet.SetColumnWidth(17, 2);

        sheet["B2"] = Text("Budget vs actual, Q3 2026", 18, Navy, bold: true);
        sheet["B3"] = Text("Operating costs per department, SEK thousands", 10, Grey);

        string[] headers = ["Department", "Budget", "Actual", "Variance", "Variance %"];

        for (int column = 0; column < headers.Length; column++)
        {
            Cell header = Text(headers[column], 10, "#FFFFFF", bold: true).Background(Navy);

            if (column > 0)
            {
                header = header.Align(HorizontalAlignment.Right);
            }

            sheet[HeaderRow, column + 1] = header;
        }

        int first = HeaderRow + 1;

        for (int index = 0; index < Departments.Length; index++)
        {
            int row = first + index;
            int excelRow = row + 1;
            sheet[row, 1] = Text(Departments[index].Department);
            sheet[row, 2] = Number(Departments[index].Budget, Thousands);
            sheet[row, 3] = Number(Departments[index].Actual, Thousands);
            sheet[row, 4] = Formula($"C{excelRow}-D{excelRow}", Signed);
            sheet[row, 5] = Formula($"E{excelRow}/C{excelRow}", SignedPercent);

            if (index % 2 == 1)
            {
                Paint(sheet, new CellRange(row, 1, row, 5), Band);
            }

            // What the second chart draws: the share used, and the budget as a line at 100 %.
            sheet[row, ChartDataColumn] = Formula($"D{excelRow}/C{excelRow}", "0%", Grey);
            sheet[row, ChartDataColumn + 1] = Number(1, "0%", Grey);
        }

        sheet[HeaderRow, ChartDataColumn] = Text("Used", 9, Grey, bold: true);
        sheet[HeaderRow, ChartDataColumn + 1] = Text("Budget", 9, Grey, bold: true);

        int last = first + Departments.Length - 1;
        int total = last + 1;
        sheet[total, 1] = Text("Total", 10, Navy, bold: true);
        sheet[total, 2] = Formula($"SUM(C{first + 1}:C{last + 1})", Thousands, Navy, bold: true);
        sheet[total, 3] = Formula($"SUM(D{first + 1}:D{last + 1})", Thousands, Navy, bold: true);
        sheet[total, 4] = Formula($"C{total + 1}-D{total + 1}", Signed, Navy, bold: true);
        sheet[total, 5] = Formula($"E{total + 1}/C{total + 1}", SignedPercent, Navy, bold: true);
        Paint(sheet, new CellRange(total, 1, total, 5), "#E8EEFB");

        // Under budget is good: green; over budget is red.
        Style over = new() { FontColor = "#B42318", Bold = true };

        Style under = new() { FontColor = "#067647" };

        CellRange variance = new(first, 4, last, 5);
        sheet.AddConditionalFormat(variance, ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, over));
        sheet.AddConditionalFormat(variance, ConditionalFormat.WhenCell(ComparisonOperator.GreaterThan, 0, under));
        sheet.AddConditionalFormat(new CellRange(first, 3, last, 3), ConditionalFormat.DataBar(LightBlue));

        AddCards(sheet);
        AddCharts(sheet, first, last);

        return book;
    }

    private static void AddCards(Worksheet sheet)
    {
        double budget = Departments.Sum(department => department.Budget);
        double actual = Departments.Sum(department => department.Actual);
        double difference = budget - actual;
        int over = Departments.Count(department => department.Actual > department.Budget);

        (string Title, string Value, string Note, string Color, string Area)[] cards =
        [
            ("Actual costs", Swedish(actual, "#,##0"), "of " + Swedish(budget, "#,##0") + " budgeted", Blue, "B5:B8"),
            ("Under budget", Swedish(difference, "+#,##0;-#,##0"), Swedish(difference / budget, "+0.0%;-0.0%") + " of the budget", Teal, "C5:D8"),
            ("Over budget", Swedish(over, "0") + " of " + Swedish(Departments.Length, "0"), "departments", Red, "E5:F8"),
        ];

        foreach ((string title, string value, string note, string color, string area) in cards)
        {
            Shape card = Card() with
            {
                Text = new ShapeText(
                [
                    new TextRun(title + "\n", new TextFont { Name = FontName, Size = 9, Color = Grey }),
                    new TextRun(value + "\n", new TextFont { Name = FontName, Size = 18, Bold = true, Color = color }),
                    new TextRun(note, new TextFont { Name = FontName, Size = 8, Color = Grey }),
                ])
                {
                    Horizontal = TextAlignment.Left,
                    MarginLeft = 10,
                },
                Description = title + ": " + value,
            };

            Check(sheet.AddShape(card, area));
        }
    }

    private static void AddCharts(Worksheet sheet, int first, int last)
    {
        string categories = $"'{Report}'!$B${first + 1}:$B${last + 1}";
        string used = CellAddress.ColumnName(ChartDataColumn);
        string target = CellAddress.ColumnName(ChartDataColumn + 1);
        TextFont font = new() { Name = FontName, Size = 9, Color = Grey };

        TextFont titleFont = new() { Name = FontName, Size = 12, Bold = true, Color = Navy };

        ChartAxis valueAxis = new() { Line = Outline.None, GridlineOutline = new Outline(Line, 0.5), NumberFormat = Thousands, Font = font };

        ChartAxis categoryAxis = new() { Line = new Outline("#D1D5DB", 0.75), Font = font };

        Chart comparison = new(ChartType.Bar)
        {
            Title = "Budget and actual per department",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("Budget", $"'{Report}'!$C${first + 1}:$C${last + 1}", categories, LightBlue),
                new ChartSeries("Actual", $"'{Report}'!$D${first + 1}:$D${last + 1}", categories, Navy),
            ],
            GapWidth = 50,
            Overlap = -10,
            Legend = LegendPosition.Top,
            LegendFont = font,
            Font = font,
            Border = Outline.None,
            ValueAxis = valueAxis,
            CategoryAxis = categoryAxis,
            Description = "Budget and actual costs per department",
        };

        Check(sheet.AddChart(comparison, "H2:Q23"));

        // Teal where the department is within its budget, red where it is over.
        string?[] colors = new string?[Departments.Length];

        for (int index = 0; index < Departments.Length; index++)
        {
            colors[index] = Teal;

            if (Departments[index].Actual > Departments[index].Budget)
            {
                colors[index] = Red;
            }
        }

        Chart share = new(ChartType.Column)
        {
            Title = "Share of the budget used",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("Used", $"'{Report}'!${used}${first + 1}:${used}${last + 1}", categories, Teal)
                {
                    PointColors = colors,
                    DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, NumberFormat = "0%", Font = new TextFont { Name = FontName, Size = 8, Bold = true, Color = "#FFFFFF" } },
                },
                new ChartSeries("Budget", $"'{Report}'!${target}${first + 1}:${target}${last + 1}", categories)
                {
                    Type = ChartType.Line,
                    Marker = MarkerStyle.None,
                    Outline = new Outline(Navy, 1.5) { Dash = LineDash.Dash },
                },
            ],
            GapWidth = 35,
            Legend = LegendPosition.None,
            Font = font,
            Border = Outline.None,
            ValueAxis = valueAxis with { Minimum = 0.8, Maximum = 1.2, MajorUnit = 0.1, NumberFormat = "0%" },
            CategoryAxis = categoryAxis,
            Description = "Share of the budget used per department, red where it is over 100 %",
        };

        Check(sheet.AddChart(share, "B25:Q39"));
    }
}
