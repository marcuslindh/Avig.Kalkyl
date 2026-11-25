using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A profit and loss statement for a year: twelve months, the year, the budget and the
/// difference, with subtotals as formulas (gross profit, EBITDA, EBIT, profit before tax, net
/// profit), costs in parentheses, a colour scale on the gross margin, data bars on revenue,
/// arrows against the budget, frozen headings and a chart of revenue and margin under the table.
/// </summary>
public static class ProfitAndLoss
{
    /// <summary>The sheet with the statement.</summary>
    public const string Report = "P&L";

    /// <summary>The cells the gallery picture covers.</summary>
    public const string PictureRange = "A1:S53";

    private static readonly string[] Months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    // Seasonal shape of the business: slow summer, strong autumn.
    private static readonly double[] Season = [0.86, 0.9, 1.02, 0.98, 1.04, 0.95, 0.72, 0.84, 1.08, 1.12, 1.16, 1.33];

    private const string Amount = "#,##0;(#,##0);\"–\"";
    private const string Percent = "0.0%";
    private const string SignedPercent = "+0.0%;-0.0%;0.0%";
    private const int HeaderRow = 4;
    private const int FirstMonthColumn = 2;
    private const int GapColumn = 14;
    private const int YearColumn = 15;
    private const int BudgetColumn = 16;
    private const int VersusColumn = 17;

    /// <summary>A line of the statement: a heading, an account with values, a subtotal formula, or a gap.</summary>
    private enum LineKind
    {
        Heading,
        Account,
        Subtotal,
        Calculated,
        Ratio,
        Result,
        Gap,
    }

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = AddSheet(book, Report);
        sheet.View = new SheetView { ShowGridLines = false, TabColor = Navy, ActiveCell = new CellAddress(HeaderRow + 1, FirstMonthColumn) };

        sheet.SetColumnWidth(0, 2);
        sheet.SetColumnWidth(1, 26);

        for (int column = FirstMonthColumn; column < GapColumn; column++)
        {
            sheet.SetColumnWidth(column, 7.6);
        }

        sheet.SetColumnWidth(GapColumn, 1);

        sheet.SetColumnWidth(YearColumn, 9.5);
        sheet.SetColumnWidth(BudgetColumn, 9.5);
        sheet.SetColumnWidth(VersusColumn, 9.5);
        sheet.SetColumnWidth(VersusColumn + 1, 2);
        sheet.FreezePanes(HeaderRow + 1, 2);

        sheet["B2"] = Text("Profit and loss 2026", 18, Navy, bold: true);
        sheet["B3"] = Text("Nordljus AB, SEK thousands. Costs in parentheses.", 10, Grey);
        WriteHeader(sheet);

        Random random = new(2026);
        Dictionary<string, int> rows = [];
        int row = HeaderRow + 1;

        foreach ((string label, LineKind kind, double yearly, double budget, string formula) in Lines())
        {
            int excelRow = row + 1;
            rows[label] = excelRow;
            string background = "#FFFFFF";
            string color = Ink;
            bool bold = false;

            switch (kind)
            {
                case LineKind.Heading:
                    sheet[row, 1] = Text(label.ToUpperInvariant(), 9, Blue, bold: true);

                    break;

                case LineKind.Account:
                    sheet[row, 1] = Text("   " + label);

                    for (int month = 0; month < Months.Length; month++)
                    {
                        double noise = 1 + ((random.NextDouble() - 0.5) * 0.08);
                        sheet[row, FirstMonthColumn + month] = Number(Math.Round(yearly / 12 * Season[month] * noise), Amount);
                    }

                    sheet[row, BudgetColumn] = Number(budget, Amount, Grey);

                    break;

                case LineKind.Subtotal:
                case LineKind.Calculated:
                case LineKind.Result:
                    if (kind != LineKind.Calculated)
                    {
                        bold = true;
                        color = Navy;
                        background = "#EEF3FC";
                    }

                    if (kind == LineKind.Result)
                    {
                        color = "#FFFFFF";
                        background = Navy;
                    }

                    sheet[row, 1] = Text(label, 10, color, bold);

                    foreach (int column in Columns(includeYear: false))
                    {
                        string letter = CellAddress.ColumnName(column);
                        sheet[row, column] = Formula(Substitute(formula, letter, rows), Amount, color, bold);
                    }

                    break;

                case LineKind.Ratio:
                    sheet[row, 1] = Text("   " + label, 9, Grey);

                    foreach (int column in Columns(includeYear: true))
                    {
                        string letter = CellAddress.ColumnName(column);
                        sheet[row, column] = Formula(Substitute(formula, letter, rows), Percent, Grey).Italic();
                    }

                    break;

                default:
                    break;
            }

            if (kind == LineKind.Account || kind == LineKind.Subtotal || kind == LineKind.Calculated || kind == LineKind.Result)
            {
                string first = CellAddress.ColumnName(FirstMonthColumn);
                string last = CellAddress.ColumnName(GapColumn - 1);

                // The year is the sum of the months on every line, so the column reads as one formula.
                sheet[row, YearColumn] = Formula($"SUM({first}{excelRow}:{last}{excelRow})", Amount, color, bold);

                sheet[row, VersusColumn] = Formula($"IFERROR(({Column(YearColumn)}{excelRow}-{Column(BudgetColumn)}{excelRow})/ABS({Column(BudgetColumn)}{excelRow}),0)", SignedPercent, color, bold);
            }

            if (background != "#FFFFFF")
            {
                Paint(sheet, new CellRange(row, 1, row, VersusColumn), background);
            }

            row++;
        }

        AddRules(sheet, rows);
        AddCharts(sheet, rows, row + 1);

        return book;
    }

    /// <summary>The columns a subtotal or ratio formula goes in: the months and the budget, and the year for a ratio.</summary>
    private static List<int> Columns(bool includeYear)
    {
        List<int> columns = [.. Enumerable.Range(FirstMonthColumn, Months.Length)];

        if (includeYear == true)
        {
            columns.Add(YearColumn);
        }

        columns.Add(BudgetColumn);

        return columns;
    }

    private static void WriteHeader(Worksheet sheet)
    {
        sheet[HeaderRow, 1] = Text("SEK thousands", 9, "#FFFFFF", bold: true).Background(Navy);

        for (int month = 0; month < Months.Length; month++)
        {
            sheet[HeaderRow, FirstMonthColumn + month] = Text(Months[month], 9, "#FFFFFF", bold: true).Background(Navy).Align(HorizontalAlignment.Right);
        }

        sheet[HeaderRow, YearColumn] = Text("Year", 9, "#FFFFFF", bold: true).Background(Blue).Align(HorizontalAlignment.Right);
        sheet[HeaderRow, BudgetColumn] = Text("Budget", 9, "#FFFFFF", bold: true).Background(Navy).Align(HorizontalAlignment.Right);
        sheet[HeaderRow, VersusColumn] = Text("vs budget", 9, "#FFFFFF", bold: true).Background(Navy).Align(HorizontalAlignment.Right);
        sheet[HeaderRow, GapColumn] = Cell.Empty.Background(Navy);
    }

    /// <summary>The statement from the top: label, kind, the year's amount, the budget, and the formula of a subtotal.</summary>
    private static List<(string Label, LineKind Kind, double Yearly, double Budget, string Formula)> Lines()
    {
        return
        [
            ("Revenue", LineKind.Heading, 0, 0, ""),
            ("Products", LineKind.Account, 38_400, 37_000, ""),
            ("Services", LineKind.Account, 16_900, 18_000, ""),
            ("Subscriptions", LineKind.Account, 9_800, 9_000, ""),
            ("Net revenue", LineKind.Subtotal, 0, 0, "SUM({Products}:{Subscriptions})"),
            ("Cost of goods sold", LineKind.Account, -24_300, -23_500, ""),
            ("Gross profit", LineKind.Subtotal, 0, 0, "{Net revenue}+{Cost of goods sold}"),
            ("", LineKind.Gap, 0, 0, ""),
            ("Operating expenses", LineKind.Heading, 0, 0, ""),
            ("Personnel", LineKind.Account, -21_600, -21_000, ""),
            ("Premises", LineKind.Account, -3_000, -3_000, ""),
            ("Marketing", LineKind.Account, -3_900, -4_200, ""),
            ("IT and software", LineKind.Account, -2_250, -2_000, ""),
            ("Other external costs", LineKind.Account, -1_700, -1_800, ""),
            ("Operating expenses", LineKind.Subtotal, 0, 0, "SUM({Personnel}:{Other external costs})"),
            ("", LineKind.Gap, 0, 0, ""),
            ("EBITDA", LineKind.Subtotal, 0, 0, "{Gross profit}+{Operating expenses}"),
            ("Depreciation", LineKind.Account, -1_450, -1_400, ""),
            ("EBIT", LineKind.Subtotal, 0, 0, "{EBITDA}+{Depreciation}"),
            ("Net financial items", LineKind.Account, -380, -400, ""),
            ("Profit before tax", LineKind.Subtotal, 0, 0, "{EBIT}+{Net financial items}"),
            ("Tax", LineKind.Account, -1_375, -1_380, ""),
            ("Net profit", LineKind.Result, 0, 0, "{Profit before tax}+{Tax}"),
            ("", LineKind.Gap, 0, 0, ""),
            ("Key ratios", LineKind.Heading, 0, 0, ""),
            ("Gross margin", LineKind.Ratio, 0, 0, "IFERROR({Gross profit}/{Net revenue},0)"),
            ("EBIT margin", LineKind.Ratio, 0, 0, "IFERROR({EBIT}/{Net revenue},0)"),
            ("Net margin", LineKind.Ratio, 0, 0, "IFERROR({Net profit}/{Net revenue},0)"),
        ];
    }

    /// <summary>Puts the column letter and row numbers into a subtotal formula: {Label} becomes C9.</summary>
    private static string Substitute(string formula, string letter, Dictionary<string, int> rows)
    {
        string result = formula;

        foreach ((string label, int row) in rows.OrderByDescending(pair => pair.Key.Length))
        {
            if (label.Length > 0)
            {
                result = result.Replace("{" + label + "}", letter + row.ToString(System.Globalization.CultureInfo.InvariantCulture), StringComparison.Ordinal);
            }
        }

        return result;
    }

    private static string Column(int column)
    {
        return CellAddress.ColumnName(column);
    }

    private static void AddRules(Worksheet sheet, Dictionary<string, int> rows)
    {
        int margin = rows["Gross margin"] - 1;
        int ebitMargin = rows["EBIT margin"] - 1;
        sheet.AddConditionalFormat(new CellRange(margin, FirstMonthColumn, margin, GapColumn - 1), ConditionalFormat.ColorScale("#F8D7DA", "#FFF3CD", "#CDEBD9"));
        sheet.AddConditionalFormat(new CellRange(ebitMargin, FirstMonthColumn, ebitMargin, GapColumn - 1), ConditionalFormat.ColorScale("#F8D7DA", "#FFF3CD", "#CDEBD9"));
        sheet.AddConditionalFormat(new CellRange(rows["Products"] - 1, YearColumn, rows["Subscriptions"] - 1, YearColumn), ConditionalFormat.DataBar(LightBlue));

        // Better or worse than the budget: costs are negative, so a positive difference is good on every line.
        CellRange versus = new(HeaderRow + 1, VersusColumn, rows["Net profit"] - 1, VersusColumn);
        sheet.AddConditionalFormat(versus, ConditionalFormat.IconSet(IconSetKind.ThreeArrows));
        sheet.AddConditionalFormat(versus, ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, new Style { FontColor = "#B42318" }));
    }

    private static void AddCharts(Worksheet sheet, Dictionary<string, int> rows, int top)
    {
        string months = $"'{Report}'!$C${HeaderRow + 1}:$N${HeaderRow + 1}";
        TextFont font = new() { Name = FontName, Size = 9, Color = Grey };

        TextFont titleFont = new() { Name = FontName, Size = 12, Bold = true, Color = Navy };

        string Row(string label)
        {
            int row = rows[label];

            return $"'{Report}'!$C${row}:$N${row}";
        }

        Chart revenue = new(ChartType.Column)
        {
            Title = "Revenue per month and EBIT margin",
            TitleFont = titleFont,
            Stacked = true,
            Series =
            [
                new ChartSeries("Products", Row("Products"), months, Navy),
                new ChartSeries("Services", Row("Services"), months, Blue),
                new ChartSeries("Subscriptions", Row("Subscriptions"), months, LightBlue),
                new ChartSeries("EBIT margin", Row("EBIT margin"), months)
                {
                    Type = ChartType.Line,
                    SecondaryAxis = true,
                    Smooth = true,
                    Marker = MarkerStyle.Circle,
                    MarkerSize = 5,
                    Outline = new Outline(Amber, 2.25),
                },
            ],
            GapWidth = 45,
            Legend = LegendPosition.Top,
            LegendFont = font,
            Font = font,
            Border = Outline.None,
            ValueAxis = new ChartAxis { Line = Outline.None, GridlineOutline = new Outline(Line, 0.5), NumberFormat = "#,##0", Font = font },
            SecondaryValueAxis = new ChartAxis { Line = Outline.None, NumberFormat = "0%", Minimum = 0, Font = font },
            CategoryAxis = new ChartAxis { Line = new Outline("#D1D5DB"), Font = font },
            Description = "Revenue per month by stream, with the EBIT margin as a line",
        };

        Check(sheet.AddChart(revenue, new CellRange(top, 1, top + 17, 10)));

        int costs = rows["Personnel"];
        Chart split = new(ChartType.Doughnut)
        {
            Title = "Operating expenses",
            TitleFont = titleFont,
            Series =
            [
                new ChartSeries("Year", $"'{Report}'!$P${costs}:$P${costs + 4}", $"'{Report}'!$B${costs}:$B${costs + 4}")
                {
                    PointColors = [Navy, Blue, LightBlue, Teal, Amber],
                    Outline = new Outline("#FFFFFF", 1.5),
                },
            ],
            HoleSize = 62,
            Legend = LegendPosition.Right,
            LegendFont = font,
            Border = Outline.None,
            Description = "Operating expenses for the year by kind",
        };

        Check(sheet.AddChart(split, new CellRange(top, 11, top + 17, VersusColumn)));
    }
}
