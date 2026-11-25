using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A project plan with a Gantt chart drawn in shapes: a task table (owner, dates, progress with
/// data bars) and beside it a timeline of narrow day columns, with weekends shaded, a bar per
/// task that fills with its progress, a summary bar per phase, diamonds for milestones and a
/// line for today. The bars are placed in points over the day columns, so they stay where they
/// are when a row grows.
/// </summary>
public static class ProjectPlan
{
    /// <summary>The sheet with the plan.</summary>
    public const string Plan = "Plan";

    private const int Days = 84;
    private const int FirstDayColumn = 7;
    private const int MonthRow = 3;
    private const int HeaderRow = 4;
    private const int FirstTaskRow = 5;
    private const double DayWidth = 8.0 / 7;

    private static readonly DateOnly Start = new(2026, 8, 31);
    private static readonly DateOnly Today = new(2026, 10, 15);
    private static readonly double[] TableWidths = [2, 26, 11, 8, 8, 8, 2];

    /// <summary>The cells the gallery picture covers: the table, the timeline and a margin.</summary>
    public static string PictureRange { get; } = "A1:" + CellAddress.ColumnName(FirstDayColumn + Days) + (FirstTaskRow + Tasks().Count + 3).ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>A row of the plan: a phase, a task, or a milestone (no duration).</summary>
    private sealed record PlanItem(string Name, string Owner, int Day, int Duration, double Progress, bool IsPhase, string Color);

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = AddSheet(book, Plan);
        sheet.View = new SheetView { ShowGridLines = false, TabColor = Teal, ActiveCell = new CellAddress(FirstTaskRow, 1) };

        for (int column = 0; column < TableWidths.Length; column++)
        {
            sheet.SetColumnWidth(column, TableWidths[column]);
        }

        for (int day = 0; day < Days; day++)
        {
            sheet.SetColumnWidth(FirstDayColumn + day, DayWidth);
        }

        sheet.SetColumnWidth(FirstDayColumn + Days, 2);
        sheet.FreezePanes(FirstTaskRow);

        List<PlanItem> tasks = Tasks();
        int lastRow = FirstTaskRow + tasks.Count - 1;
        WriteTable(sheet, tasks);
        WriteTimeline(sheet, lastRow);
        DrawBars(sheet, tasks, lastRow);

        Shape title = new(ShapeKind.TextBox)
        {
            Text = new ShapeText(
            [
                new TextRun("Website relaunch\n", new TextFont { Name = FontName, Size = 18, Bold = true, Color = Navy }),
                new TextRun("Project plan, autumn 2026. Bars fill with progress; diamonds are milestones.", new TextFont { Name = FontName, Size = 9, Color = Grey }),
            ])
            {
                Horizontal = HorizontalAlignment.Left,
                Vertical = VerticalAlignment.Top,
                MarginLeft = 0,
                MarginTop = 0,
            },
        };

        Check(sheet.AddShape(title, DrawingBounds.FromPoints(ColumnPoints(2), 4, 520, 40)));

        return book;
    }

    /// <summary>The plan: four phases with their tasks and milestones, days counted from the start.</summary>
    private static List<PlanItem> Tasks()
    {
        const string discovery = "#14A38B";
        const string design = "#2F6FED";
        const string build = "#7C3AED";
        const string launch = "#E8590C";

        return
        [
            new("Discovery", "", 0, 14, 1, true, discovery),
            new("Stakeholder interviews", "Anna L.", 0, 7, 1, false, discovery),
            new("Requirements and scope", "Erik S.", 5, 9, 1, false, discovery),
            new("Scope approved", "", 14, 0, 1, false, discovery),
            new("Design", "", 14, 26, 0.75, true, design),
            new("Information architecture", "Sara K.", 14, 8, 1, false, design),
            new("Visual design", "Sara K.", 19, 14, 0.85, false, design),
            new("Prototype and user tests", "Jonas B.", 28, 12, 0.4, false, design),
            new("Design sign-off", "", 40, 0, 0, false, design),
            new("Build", "", 35, 31, 0.1, true, build),
            new("Front end", "Erik S.", 35, 25, 0.25, false, build),
            new("CMS and integrations", "Lina M.", 38, 24, 0.12, false, build),
            new("Content migration", "Anna L.", 50, 16, 0, false, build),
            new("Launch", "", 62, 16, 0, true, launch),
            new("Testing and fixes", "Jonas B.", 62, 12, 0, false, launch),
            new("Training", "Lina M.", 70, 5, 0, false, launch),
            new("Go live", "", 77, 0, 0, false, launch),
        ];
    }

    private static void WriteTable(Worksheet sheet, List<PlanItem> tasks)
    {
        string[] headers = ["Task", "Owner", "Start", "End", "Done"];

        for (int column = 0; column < headers.Length; column++)
        {
            Cell header = Text(headers[column], 9, "#FFFFFF", bold: true).Background(Navy);

            if (column >= 2)
            {
                header = header.Align(HorizontalAlignment.Right);
            }

            sheet[HeaderRow, column + 1] = header;
        }

        sheet[HeaderRow, 6] = Cell.Empty.Background(Navy);

        for (int index = 0; index < tasks.Count; index++)
        {
            PlanItem task = tasks[index];
            int row = FirstTaskRow + index;
            DateOnly first = Start.AddDays(task.Day);
            DateOnly last = Start.AddDays(task.Day + Math.Max(task.Duration, 1) - 1);

            if (task.IsPhase == true)
            {
                sheet[row, 1] = Text(task.Name, 10, Navy, bold: true);
                Paint(sheet, new CellRange(row, 1, row, 6), "#EEF3FC");
            }
            else if (task.Duration == 0)
            {
                sheet[row, 1] = Text("   ◆ " + task.Name, 10, task.Color);
            }
            else
            {
                sheet[row, 1] = Text("   " + task.Name);
                sheet[row, 2] = Text(task.Owner, 9, Grey);
            }

            Cell start = Cell.Of(first).Font(FontName).FontSize(9).FontColor(Grey).NumberFormat("d mmm");
            Cell end = Cell.Of(last).Font(FontName).FontSize(9).FontColor(Grey).NumberFormat("d mmm");
            Cell done = Number(task.Progress, "0%", Ink);

            if (task.IsPhase == true)
            {
                start = start.Background("#EEF3FC");
                end = end.Background("#EEF3FC");
                done = done.Bold().Background("#EEF3FC");
            }

            sheet[row, 3] = start;

            if (task.Duration > 0)
            {
                sheet[row, 4] = end;
                sheet[row, 5] = done;
            }
        }

        sheet.AddConditionalFormat(new CellRange(FirstTaskRow, 5, FirstTaskRow + tasks.Count - 1, 5), ConditionalFormat.DataBar("#9FD8CB"));
    }

    private static void WriteTimeline(Worksheet sheet, int lastRow)
    {
        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        for (int day = 0; day < Days; day++)
        {
            DateOnly date = Start.AddDays(day);
            int column = FirstDayColumn + day;
            sheet[HeaderRow, column] = Cell.Empty.Background(Navy);

            if (date.DayOfWeek == DayOfWeek.Monday)
            {
                int week = System.Globalization.ISOWeek.GetWeekOfYear(date.ToDateTime(TimeOnly.MinValue));
                sheet[HeaderRow, column] = Text("W" + week.ToString(System.Globalization.CultureInfo.InvariantCulture), 8, "#FFFFFF", bold: true).Background(Navy);
            }

            // The month at the first of each month, and at the start unless a new month is days away.
            if (date.Day == 1 || (day == 0 && date.Day <= 20))
            {
                sheet[MonthRow, column] = Text(months[date.Month - 1] + " " + date.Year.ToString(System.Globalization.CultureInfo.InvariantCulture), 9, Navy, bold: true);
            }

            if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
            {
                Paint(sheet, new CellRange(FirstTaskRow, column, lastRow, column), "#F1F3F7");
            }
        }
    }

    /// <summary>The left edge of a day column, in points from the sheet's left edge.</summary>
    private static double DayLeft(double day)
    {
        double left = TableWidths.Sum(ColumnPoints);

        return left + (day * ColumnPoints(DayWidth));
    }

    private static double RowTop(int row)
    {
        return row * RowPoints;
    }

    private static void DrawBars(Worksheet sheet, List<PlanItem> tasks, int lastRow)
    {
        for (int index = 0; index < tasks.Count; index++)
        {
            PlanItem task = tasks[index];
            double top = RowTop(FirstTaskRow + index);
            double left = DayLeft(task.Day);
            double width = DayLeft(task.Day + task.Duration) - left;

            if (task.IsPhase == true)
            {
                // A thin dark summary bar over the phase.
                Shape summary = new(ShapeKind.Rectangle) { Fill = Fill.Solid(Navy), Outline = Outline.None, Name = task.Name + " phase", Description = task.Name };

                Check(sheet.AddShape(summary, DrawingBounds.FromPoints(left, top + 5, width, 5)));

                continue;
            }

            if (task.Duration == 0)
            {
                Shape milestone = new(ShapeKind.Diamond) { Fill = Fill.Solid(task.Color), Outline = new Outline("#FFFFFF", 1), Description = "Milestone: " + task.Name };

                Check(sheet.AddShape(milestone, DrawingBounds.FromPoints(left - 6, top + 1.5, 12, 12)));
                Shape label = Label(task.Name, 8, task.Color, bold: true);
                Check(sheet.AddShape(label, DrawingBounds.FromPoints(left + 9, top, 120, RowPoints)));

                continue;
            }

            Shape planned = new(ShapeKind.RoundedRectangle) { Fill = Fill.Solid(task.Color, 0.7), Outline = Outline.None, CornerRadius = 0.3, Description = task.Name };

            Check(sheet.AddShape(planned, DrawingBounds.FromPoints(left, top + 3, width, 9)));

            if (task.Progress > 0)
            {
                Shape done = new(ShapeKind.RoundedRectangle) { Fill = Fill.Solid(task.Color), Outline = Outline.None, CornerRadius = 0.3, Description = task.Name + " done" };

                Check(sheet.AddShape(done, DrawingBounds.FromPoints(left, top + 3, width * task.Progress, 9)));
            }
        }

        // Today: a red line from the header to the last task, with its label above.
        double today = DayLeft(Today.DayNumber - Start.DayNumber);
        double lineTop = RowTop(HeaderRow);
        Shape line = new(ShapeKind.Rectangle) { Fill = Fill.Solid(Red), Outline = Outline.None, Description = "Today" };

        Check(sheet.AddShape(line, DrawingBounds.FromPoints(today - 0.75, lineTop, 1.5, RowTop(lastRow + 1) - lineTop)));
        Shape flag = new(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid(Red),
            Outline = Outline.None,
            CornerRadius = 0.3,
            Text = new ShapeText("Today", new TextFont { Name = FontName, Size = 8, Bold = true, Color = "#FFFFFF" }) { MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0 },
        };

        Check(sheet.AddShape(flag, DrawingBounds.FromPoints(today - 18, RowTop(MonthRow) + 1, 36, 13)));
    }
}
