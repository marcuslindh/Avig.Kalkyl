using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Avig.OpenXml;

namespace Avig.Kalkyl.Tests.Readme;

#region example: typed-row-type
public enum CustomerLevel
{
    Basic,
    Gold,
}

[KalkylRow]
public partial class Customer
{
    [KalkylColumn("Name", Width = 30, Order = 0)]
    public required string Name { get; init; }

    [KalkylColumn(NumberFormat = "# ##0.00")]
    public decimal Balance { get; init; }

    public DateOnly? Since { get; init; }

    public CustomerLevel Level { get; init; }

    [KalkylIgnore]
    public string Notes { get; set; } = "";
}
#endregion

/// <summary>
/// The examples in README.md, run as tests. Each README code block is preceded by
/// <c>&lt;!-- example: name --&gt;</c> and must be exactly the code of <c>#region example: name</c>
/// here; <see cref="Every_README_example_is_code_that_runs_here"/> checks that. The examples use
/// relative paths, so each test runs in a folder of its own, and the collection runs alone,
/// since the current directory belongs to the whole process.
/// </summary>
[Collection(ReadmeGroup.Name)]
public sealed class ReadmeExamples : IDisposable
{
    private readonly string previousDirectory = Environment.CurrentDirectory;
    private readonly string folder = Path.Combine(Path.GetTempPath(), "avig-readme-" + Guid.NewGuid().ToString("N"));

    public ReadmeExamples()
    {
        Directory.CreateDirectory(folder);
        Environment.CurrentDirectory = folder;
    }

    public void Dispose()
    {
        Environment.CurrentDirectory = previousDirectory;
        Directory.Delete(folder, true);
    }

    [Fact]
    public async Task Result_pattern()
    {
        await CreateBudgetAsync("budget.xlsx");

        #region example: result-pattern
        Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

        // One call gives either the error and its status, or the value.
        if (opened.IsFailure(out string? error, out ResultStatus status, out Workbook? book) == true)
        {
            Console.WriteLine($"{status}: {error}");   // NotFound: 'budget.xlsx' does not exist.

            return;
        }

        // When only success matters, or only the value:
        bool saved = (await book.SaveAsync("copy.xlsx")).IsSuccess(out _, out _);

        if (book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == true)
        {
            Console.WriteLine($"{sheet.Name}, saved: {saved}");
        }
        #endregion

        Assert.True(saved);
        Assert.True((await Workbook.OpenAsync("missing.xlsx")).IsFailure(out _, out ResultStatus missing));
        Assert.Equal(ResultStatus.NotFound, missing);
    }

    [Fact]
    public async Task Typed_write_and_read()
    {
        #region example: typed-write-read
        List<Customer> customers =
        [
            new() { Name = "Ada Lovelace", Balance = 1234.5m, Since = new DateOnly(2024, 3, 1), Level = CustomerLevel.Gold },
            new() { Name = "Alan Turing", Balance = -20m, Level = CustomerLevel.Basic },
        ];

        Result written = await Xlsx.WriteAsync("customers.xlsx", customers, "Customers");

        await foreach (Result<Customer> row in Xlsx.ReadAsync<Customer>("customers.xlsx", "Customers"))
        {
            if (row.IsFailure(out string? error, out _, out Customer? customer) == true)
            {
                // Cell B3 (row 2, column 1) in 'Customers', column 'Balance', holds 'abc', which is not a number.
                Console.WriteLine(error);

                continue;
            }

            Console.WriteLine($"{customer.Name}: {customer.Balance} ({customer.Level})");
        }
        #endregion

        Assert.True(written.IsSuccess(out _, out _));
        List<Customer> read = [];

        await foreach (Result<Customer> row in Xlsx.ReadAsync<Customer>("customers.xlsx"))
        {
            Assert.True(row.IsSuccess(out Customer? customer));
            read.Add(customer);
        }

        Assert.Equal(["Ada Lovelace", "Alan Turing"], read.Select(customer => customer.Name));
        Assert.Equal(new DateOnly(2024, 3, 1), read[0].Since);
        Assert.Null(read[1].Since);
        Assert.Equal(CustomerLevel.Gold, read[0].Level);
    }

    [Fact]
    public async Task Typed_lists_in_a_streamed_file()
    {
        #region example: typed-streaming
        if ((await XlsxWriter.CreateAsync("export.xlsx")).IsSuccess(out XlsxWriter? writer) == true)
        {
            await using (writer)
            {
                // A sheet with settings of its own, then the typed rows: a header row and one row per item.
                if ((await writer.AddSheetAsync("Customers")).IsSuccess(out SheetWriter? sheet) == true)
                {
                    sheet.FreezePanes(1);
                    sheet.EnableAutoFilter();
                    await sheet.WriteAsync(LoadCustomers());
                }

                // An IAsyncEnumerable is written as it arrives, for example straight from a database.
                if ((await writer.AddSheetAsync("Stream")).IsSuccess(out SheetWriter? streamed) == true)
                {
                    await streamed.WriteAsync(LoadCustomersAsync());
                }
            }
        }

        if ((await XlsxReader.OpenAsync("export.xlsx")).IsSuccess(out XlsxReader? reader) == true)
        {
            await using (reader)
            {
                if (reader.Sheet("Stream").IsSuccess(out SheetReader? sheet) == true)
                {
                    await foreach (Result<Customer> row in sheet.ReadAsync<Customer>())
                    {
                        if (row.IsSuccess(out Customer? customer) == true)
                        {
                            Console.WriteLine(customer.Name);
                        }
                    }
                }
            }
        }
        #endregion

        Assert.Equal(3, await CountRowsAsync("export.xlsx", "Customers"));
        Assert.Equal(3, await CountRowsAsync("export.xlsx", "Stream"));
    }

    [Fact]
    public void Typed_map()
    {
        #region example: typed-map
        KalkylMap<Customer> map = Customer.KalkylMap;

        foreach (KalkylColumnInfo column in map.Columns)
        {
            Console.WriteLine($"{column.Header}: required {column.Required}, format {column.NumberFormat}, width {column.Width}");
        }

        IReadOnlyList<Cell> cells = map.ToCells(new Customer { Name = "Ada", Balance = 10m });
        #endregion

        Assert.Equal(["Name", "Balance", "Since", "Level"], map.Columns.Select(column => column.Header));
        Assert.True(map.Columns[0].Required);
        Assert.False(map.Columns[2].Required);
        Assert.Equal((CellValue)"Ada", cells[0].Value);
    }

    [Fact]
    public async Task Streamed_report()
    {
        #region example: stream-write
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync("report.xlsx");

        if (created.IsFailure(out string? error, out _, out XlsxWriter? writer) == true)
        {
            Console.WriteLine(error);   // the folder is missing, the file is open in Excel ...

            return;
        }

        await using (writer)
        {
            // A name Excel refuses (too long, one of : \ / ? * [ ], already taken) is Invalid or Conflict.
            if ((await writer.AddSheetAsync("Sales")).IsFailure(out string? refused, out _, out SheetWriter? sheet) == true)
            {
                Console.WriteLine(refused);

                return;
            }

            // Widths and frozen panes go at the top of the sheet in the file: set them before the first row.
            sheet.SetColumnWidth(0, 24);
            sheet.SetColumnWidth(2, 14);
            sheet.FreezePanes(1);
            sheet.EnableAutoFilter();

            await sheet.WriteRowAsync([Cell.Of("Product").Bold().Background("#DDEBF7"), Cell.Of("Units").Bold(), Cell.Of("Sold on").Bold(), Cell.Of("Paid").Bold()]);

            for (int index = 0; index < 10_000; index++)
            {
                await sheet.WriteRowAsync(["Product " + index, index % 17, new DateTime(2026, 1, 1).AddDays(index % 365), index % 3 != 0]);
            }

            int last = sheet.RowCount;
            await sheet.WriteRowAsync([Cell.Of("Total").Bold(), Cell.FromFormula($"SUM(B2:B{last})").Bold()]);

            // Adding the next sheet finishes this one; it cannot be written to afterwards.
            if ((await writer.AddSheetAsync("Notes")).IsSuccess(out SheetWriter? notes) == true)
            {
                await notes.WriteRowAsync(["Written by Avig.Kalkyl", DateTime.UtcNow]);
            }

            Console.WriteLine($"{writer.Sheets.Count} sheets, '{sheet.Name}' (index {sheet.Index}) finished: {sheet.IsFinished}");
        }
        #endregion

        Assert.Equal(10_002, await CountRowsAsync("report.xlsx", "Sales"));
    }

    [Fact]
    public async Task Streamed_to_a_stream()
    {
        MemoryStream response = new();
        await WriteToResponseAsync(response);
        response.Position = 0;
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(response, false);

        Assert.True(opened.IsSuccess(out XlsxReader? reader));
        await reader.DisposeAsync();
    }

    [Fact]
    public async Task Streamed_reading()
    {
        await WriteSampleAsync("big.xlsx");

        #region example: stream-read
        Result<XlsxReader> opened = await XlsxReader.OpenAsync("big.xlsx");

        if (opened.IsFailure(out string? error, out _, out XlsxReader? reader) == true)
        {
            Console.WriteLine(error);

            return;
        }

        await using (reader)
        {
            foreach (SheetInfo info in reader.Sheets)
            {
                Console.WriteLine($"{info.Index}: {info.Name} ({info.Visibility})");
            }

            Console.WriteLine($"Dates count from 1904: {reader.Uses1904Dates}");

            if (reader.Sheet("Sales").IsSuccess(out SheetReader? sheet) == true)   // or reader.Sheet(0)
            {
                await foreach (Row row in sheet.ReadRowsAsync())
                {
                    if (row.Index == 0 || row.IsEmpty == true)
                    {
                        continue;   // the header, and rows that are there but hold nothing
                    }

                    if (row.TryGetText(0, out string? product) == true
                        && row.TryGetNumber(1, out double units) == true
                        && row.TryGetDate(2, out DateTime soldOn) == true
                        && row.TryGetBool(3, out bool paid) == true)
                    {
                        Console.WriteLine($"{product}: {units} on {soldOn:yyyy-MM-dd}, paid {paid}");
                    }

                    Cell fifth = row[4];   // beyond the last cell: an empty cell
                    Console.WriteLine($"{row.ColumnCount} columns, {row.Cells.Count} cells; E is empty: {fifth.IsEmpty}");
                }
            }
        }
        #endregion
    }

    [Fact]
    public async Task Uploaded_files()
    {
        await WriteSampleAsync("upload.xlsx");

        await using (FileStream file = File.OpenRead("upload.xlsx"))
        {
            Assert.Equal("ok", await HandleUploadAsync(file));
        }

        await using MemoryStream notExcel = new("not a workbook"u8.ToArray());
        Assert.Equal("Upload an .xlsx file.", await HandleUploadAsync(notExcel));
    }

    [Fact]
    public void Values()
    {
        #region example: values
        CellValue text = "Ada";                          // implicit from string, double, int, long, decimal, bool, DateTime, DateOnly
        CellValue number = CellValue.Number(42.5);
        CellValue date = new DateOnly(2026, 11, 24);
        CellValue error = CellValue.Error("#N/A");
        CellValue nothing = CellValue.Empty;

        foreach (CellValue value in new[] { text, number, date, error, nothing, CellValue.Bool(true), CellValue.Text("x"), CellValue.Date(DateTime.UtcNow.Date) })
        {
            string kind = value.Kind switch
            {
                CellKind.Text => "text",
                CellKind.Number => "number",
                CellKind.Bool => "true or false",
                CellKind.Date => "date",
                CellKind.Error => "error",
                _ => "empty",
            };

            Console.WriteLine($"{value} is {kind}");
        }

        if (number.TryGetNumber(out double amount) == true && date.TryGetDate(out DateTime day) == true && error.TryGetError(out string? code) == true)
        {
            Console.WriteLine($"{amount}, {day:d}, {code}");
        }

        if (text.TryGetText(out string? name) == true && CellValue.Bool(false).TryGetBool(out bool flag) == true)
        {
            Console.WriteLine($"{name} {flag}");
        }

        Console.WriteLine($"A cell holds at most {CellValue.MaxTextLength} characters; empty: {nothing.IsEmpty}");
        #endregion

        Assert.Equal(CellKind.Date, date.Kind);
        Assert.True(error.TryGetError(out string? errorCode));
        Assert.Equal("#N/A", errorCode);
    }

    [Fact]
    public void Cells_and_styles()
    {
        #region example: styles
        Cell plain = "Total";                                   // implicit, like CellValue
        Cell styled = Cell.Of(1234.5)
            .Bold()
            .Italic()
            .Underline()
            .Font("Arial")
            .FontSize(12)
            .FontColor("#C00000")
            .Background("#FFF2CC")
            .NumberFormat("# ##0.00 \"kr\"")
            .Align(HorizontalAlignment.Right)
            .Align(VerticalAlignment.Top)
            .Wrap();

        // Style is a record: build it once, vary it with `with`, and give it to any cell.
        Style heading = new() { Bold = true, FontSize = 14, Background = "#DDEBF7" };

        Cell title = Cell.Of("Report").WithStyle(heading);
        Cell subtitle = Cell.Of("2026").WithStyle(heading with { FontSize = 12 });

        Cell formula = Cell.FromFormula("=SUM(B2:B9)");                      // calculated by Excel, or by Workbook.RecalculateAsync
        Cell cached = Cell.FromFormula("SUM(B2:B9)", CellValue.Number(99));   // with the result that programs that do not calculate show

        Console.WriteLine($"{styled.Style.Bold} {styled.Style.NumberFormat} {plain.Style.IsDefault} {Style.Default == plain.Style}");
        Console.WriteLine($"{formula.Formula} {cached.Value} {Cell.Empty.IsEmpty} {title.Kind} {subtitle.Style.FontSize}");

        if (styled.TryGetNumber(out double value) == true)
        {
            Console.WriteLine(value);
        }
        #endregion

        Assert.Equal("SUM(B2:B9)", formula.Formula);
        Assert.Equal(12, subtitle.Style.FontSize);
        Assert.True(styled.Style.WrapText);
    }

    [Fact]
    public void Addresses()
    {
        #region example: addresses
        CellAddress c2 = new(1, 2);                            // row 1, column 2: the cell Excel calls C2
        Console.WriteLine(c2);                                 // C2

        if (CellAddress.Parse("$AA$10").IsSuccess(out CellAddress parsed) == true)
        {
            Console.WriteLine($"row {parsed.Row}, column {parsed.Column}");   // row 9, column 26
        }

        bool valid = CellAddress.TryParse("XFE1", out _);      // false: past the last column
        string name = CellAddress.ColumnName(27);              // AB

        if (CellAddress.TryParseColumn("AB", out int column) == true)
        {
            Console.WriteLine($"{valid} {name} {column} {CellAddress.MaxRows} x {CellAddress.MaxColumns}");   // False AB 27 1048576 x 16384
        }

        CellRange range = new(1, 1, 9, 3);                     // B2:D10

        if (CellRange.TryParse("D10:B2", out CellRange same) == true)   // corners in any order
        {
            Console.WriteLine($"{range} {range == same} {range.RowCount}x{range.ColumnCount} {range.Contains(c2)} {range.First}-{range.Last}");
        }
        #endregion

        Assert.Equal("C2", c2.ToString());
        Assert.Equal(26, parsed.Column);
        Assert.False(valid);
        Assert.Equal("AB", name);
        Assert.Equal(27, column);
        Assert.Equal(range, same);
    }

    [Fact]
    public async Task Creating_a_workbook()
    {
        #region example: document-create
        Workbook book = Workbook.Create();

        if (book.AddSheet("Budget").IsFailure(out string? error, out _, out Worksheet? budget) == true)
        {
            Console.WriteLine(error);

            return;
        }

        budget["A1"] = Cell.Of("Month").Bold();
        budget["B1"] = Cell.Of("Amount").Bold();

        for (int month = 1; month <= 12; month++)
        {
            budget[month, 0] = new DateOnly(2026, month, 1);   // row month, column A
            budget[month, 1] = month * 1000;
        }

        budget["A14"] = Cell.Of("Total").Bold();
        budget.SetFormula(13, 1, "SUM(B2:B13)");
        budget.SetColumnWidth(0, 14);
        budget.FreezePanes(1);
        budget.AutoFilter = true;

        if (book.AddSheet("Settings").IsSuccess(out Worksheet? settings) == true)
        {
            settings["A1"] = 0.25;
            settings.Visibility = SheetVisibility.Hidden;
        }

        RecalculationReport report = await book.RecalculateAsync();
        Console.WriteLine($"{report.Calculated} formulas, total {budget["B14"].Value}");

        Result saved = await book.SaveAsync("budget.xlsx");
        #endregion

        Assert.True(saved.IsSuccess(out _, out _));
        Assert.Equal((CellValue)78000, budget["B14"].Value);
    }

    [Fact]
    public async Task Editing_a_workbook()
    {
        await CreateBudgetAsync("budget.xlsx");

        #region example: document-edit
        Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

        if (opened.IsFailure(out string? error, out _, out Workbook? book) == true)
        {
            Console.WriteLine(error);

            return;
        }

        if (book.Sheet("Budget").IsFailure(out string? missing, out _, out Worksheet? sheet) == true)
        {
            Console.WriteLine(missing);   // NotFound

            return;
        }

        sheet.SetValue(4, 1, 125_000);                        // B5: the value changes, the cell keeps its style
        sheet["C5"] = Cell.FromFormula("B5*1.25").Bold();      // replaces value, formula and style
        sheet.Clear(5, 1);                                     // B6: value, formula and style gone

        Console.WriteLine($"{sheet.RowCount} rows, {sheet.ColumnCount} columns, A is {sheet.GetColumnWidth(0)} wide, {sheet.FrozenRows} row and {sheet.FrozenColumns} columns frozen, filter {sheet.AutoFilter}");

        foreach (Row row in sheet.Rows)
        {
            Console.WriteLine($"{row.Index}: {string.Join(" | ", row.Cells.Select(cell => cell.Value.ToString()))}");
        }

        foreach (Worksheet each in book.Sheets)
        {
            Console.WriteLine($"{each.Index} {each.Name} {each.Visibility} in {each.Workbook.Sheets.Count}");
        }

        if (book.Sheet(0).IsSuccess(out Worksheet? first) == true)
        {
            Console.WriteLine(ReferenceEquals(first, sheet));  // True
        }

        await book.RecalculateAsync();
        Result saved = await book.SaveAsync("budget.xlsx");    // saving over the opened file is fine
        #endregion

        Assert.True(saved.IsSuccess(out _, out _));
        Assert.Equal((CellValue)156_250, sheet["C5"].Value);
        Assert.True(sheet["B6"].IsEmpty);
    }

    [Fact]
    public async Task Saving_to_a_stream()
    {
        await CreateBudgetAsync("budget.xlsx");
        await using FileStream input = File.OpenRead("budget.xlsx");
        await using MemoryStream output = new();
        await ToStreamAsync(input, output);

        Assert.True(output.Length > 0);
    }

    [Fact]
    public async Task Formulas()
    {
        await CreateBudgetAsync("budget.xlsx");

        #region example: formulas
        Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

        if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
        {
            return;
        }

        sheet["E1"] = Cell.FromFormula("B14/12");
        sheet["E2"] = Cell.FromFormula("E1+E2");               // a circular reference
        sheet["E3"] = Cell.FromFormula("NOSUCHFUNCTION(1)");

        RecalculationReport changed = await book.RecalculateAsync();                     // what changed, and what depends on it
        RecalculationReport all = await book.RecalculateAsync(RecalculationScope.All);  // every formula

        Console.WriteLine($"{changed.Calculated} then {all.Calculated} calculated, E1 = {sheet["E1"].Value}");
        Console.WriteLine("Circular: " + string.Join(", ", all.CircularReferences));    // Budget!E2
        Console.WriteLine("Unknown: " + string.Join(", ", all.UnknownFunctions));       // NOSUCHFUNCTION
        Console.WriteLine("Broken: " + string.Join(", ", all.InvalidFormulas));

        // A formula that is not in a cell. References without a sheet name are on `sheet`.
        Result<CellValue> average = book.Evaluate("AVERAGE(B2:B13)", sheet);
        Result<CellValue> fromAnywhere = book.Evaluate("SUM(Budget!B2:B13)*1.25");

        // Text read as a number or a date follows Workbook.Culture, which is the machine's unless set.
        book.Culture = CultureInfo.GetCultureInfo("en-GB");
        Result<CellValue> february = book.Evaluate("MONTH(\"3/2/2013\")");             // 2
        Console.WriteLine($"Dates count from 1904: {book.Uses1904Dates}");
        #endregion

        Assert.Contains("Budget!E2", all.CircularReferences);
        Assert.Contains("NOSUCHFUNCTION", all.UnknownFunctions);
        Assert.True(average.IsSuccess(out CellValue value));
        Assert.Equal((CellValue)6500, value);
        Assert.True(fromAnywhere.IsSuccess(out _));
        Assert.True(february.IsSuccess(out CellValue month));
        Assert.Equal((CellValue)2, month);
    }

    [Fact]
    public async Task Charts()
    {
        await CreateBudgetAsync("budget.xlsx");

        #region example: charts
        Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

        if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
        {
            return;
        }

        // A series is one row or one column of numbers; the categories are its labels (the x values of a scatter chart).
        ChartSeries amount = new("Amount", "B2:B13", "A2:A13", "#4472C4");
        ChartSeries target = new("Target", "C2:C13", "A2:A13");

        Chart columns = new(ChartType.Column)
        {
            Title = "Budget 2026",
            Series = [amount, target],
            CategoryAxisTitle = "Month",
            ValueAxisTitle = "kr",
            Legend = LegendPosition.Bottom,
            Description = "Amount and target per month",    // read by screen readers
        };

        Result added = sheet.AddChart(columns, "F2:M18");     // covers F2:M18
        sheet.AddChart(new Chart(ChartType.Bar) { Stacked = true, Series = [amount, target] }, new CellRange(19, 5, 35, 12));
        sheet.AddChart(new Chart(ChartType.Line) { Series = [amount], Legend = LegendPosition.None }, "F38:M54");
        sheet.AddChart(new Chart(ChartType.Area) { Stacked = true, Series = [amount, target] }, "F56:M72");
        sheet.AddChart(new Chart(ChartType.Pie) { Title = "Share per month", Series = [amount] }, "O2:U18");
        sheet.AddChart(new Chart(ChartType.Doughnut) { Series = [amount] }, "O20:U36");
        sheet.AddChart(new Chart(ChartType.Scatter) { Series = [new ChartSeries("Amount by target", "B2:B13", "C2:C13")] }, "O38:U54");

        // Another sheet's cells: name the sheet, in quotes when the name has spaces.
        if (book.AddSheet("Summary").IsFailure(out string? taken, out _, out Worksheet? summary) == true)
        {
            Console.WriteLine(taken);

            return;
        }

        Result other = summary.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries("Amount", "'Budget'!B2:B13", "Budget!A2:A13")] }, "A1:H16");
        Result missing = summary.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries(null, "Nowhere!B2:B13")] }, "A20:H36");

        if (missing.IsFailure(out string? reason, out ResultStatus status) == true)
        {
            Console.WriteLine($"{status}: {reason}");          // NotFound: there is no sheet Nowhere
        }

        await book.SaveAsync("budget-charts.xlsx");
        #endregion

        Assert.True(added.IsSuccess(out _, out _));
        Assert.True(other.IsSuccess(out _, out _));
        Assert.Equal(ResultStatus.NotFound, status);
        Assert.Empty(TestFiles.Validate(new MemoryStream(await File.ReadAllBytesAsync("budget-charts.xlsx"))));
    }

    [Fact]
    public async Task Pictures()
    {
        await File.WriteAllBytesAsync("logo.png", DrawingTests.Png);

        #region example: pictures
        Result<Picture> read = await Picture.FromFileAsync("logo.png");   // or Picture.FromBytes(bytes)

        if (read.IsFailure(out string? error, out _, out Picture? logo) == true)
        {
            Console.WriteLine(error);   // missing, or not a PNG, JPEG or GIF

            return;
        }

        Console.WriteLine($"{logo.Format}, {logo.Width}x{logo.Height} pixels");
        Picture described = logo with { Description = "Company logo" };   // alternative text

        if ((await XlsxWriter.CreateAsync("pictures.xlsx")).IsSuccess(out XlsxWriter? writer) == false)
        {
            return;
        }

        await using (writer)
        {
            if ((await writer.AddSheetAsync("Sheet1")).IsSuccess(out SheetWriter? sheet) == true)
            {
                sheet.AddPicture(described, "A1");                         // at its own size, 96 pixels to the inch
                sheet.AddPicture(described, new CellAddress(0, 3), 0.5);   // D1, at half size
                sheet.AddPicture(described, "F1:H6");                      // stretched over F1:H6
            }
        }
        #endregion

        Assert.Empty(TestFiles.Validate(new MemoryStream(await File.ReadAllBytesAsync("pictures.xlsx"))));
    }

    [Fact]
    public async Task Conditional_formatting()
    {
        await CreateBudgetAsync("budget.xlsx");

        #region example: conditional-formatting
        Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

        if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
        {
            return;
        }

        // A rule's style may set bold, italic, underline, font colour, fill and number format.
        Style bad = new() { FontColor = "#9C0006", Background = "#FFC7CE" };

        Style good = new() { Bold = true, FontColor = "#006100", Background = "#C6EFCE" };


        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThan, 10_000, good));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenCell(ComparisonOperator.LessThan, "$C$2", bad));   // compared with a formula
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenBetween(4000, 6000, new Style { Italic = true }));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenBetween("$C$2/2", "$C$2", new Style { Underline = true }));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenNotBetween(0, 100_000, bad));
        sheet.AddConditionalFormat("A2:C13", ConditionalFormat.WhenFormula("$B2>$C2", new Style { Background = "#FFEB9C" }));   // whole rows
        sheet.AddConditionalFormat("D2:D13", ConditionalFormat.WhenTextContains("late", bad));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.Top(3, good));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.Bottom(10, bad, percent: true));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.AboveAverage(new Style { Bold = true }));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.BelowAverage(new Style { NumberFormat = "0.0" }));
        sheet.AddConditionalFormat("D2:D13", ConditionalFormat.Duplicates(bad));
        sheet.AddConditionalFormat("D2:D13", ConditionalFormat.Unique(good));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.ColorScale("#F8696B", "#63BE7B"));
        sheet.AddConditionalFormat("C2:C13", ConditionalFormat.ColorScale("#F8696B", "#FFEB84", "#63BE7B"));
        sheet.AddConditionalFormat(new CellRange(1, 1, 12, 1), ConditionalFormat.DataBar("#638EC6"));
        sheet.AddConditionalFormat("C2:C13", ConditionalFormat.IconSet(IconSetKind.ThreeTrafficLights));
        sheet.AddConditionalFormat("B2:B13", ConditionalFormat.IconSet(IconSetKind.FiveRating, reverse: true, showValue: false));

        await book.SaveAsync("budget-rules.xlsx");
        #endregion

        Assert.Empty(TestFiles.Validate(new MemoryStream(await File.ReadAllBytesAsync("budget-rules.xlsx"))));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.WhenCell(ComparisonOperator.Equal, 1, new Style { FontSize = 20 }));
    }

    [Fact]
    public async Task Macros()
    {
        await CreateBudgetAsync("upload.xlsm");

        #region example: macros
        if ((await Workbook.OpenAsync("upload.xlsm")).IsSuccess(out Workbook? book) == false)
        {
            return;
        }

        if (book.HasMacros == true)
        {
            Result removed = book.RemoveMacros();   // Invalid for macro sheets and ActiveX controls, which cannot be removed

            if (removed.IsFailure(out string? error, out _) == true)
            {
                Console.WriteLine(error);

                return;
            }
        }

        await book.SaveAsync("clean.xlsx");         // without the VBA project, as an ordinary .xlsx
        #endregion

        Assert.True(File.Exists("clean.xlsx"));
    }

    [Fact]
    public async Task ActiveContentAndLinks()
    {
        await using (MemoryStream upload = new ActiveWorkbook().AddMacroSheet().AddExternalWorkbookLink().AddDdeLink().Build())
        {
            await File.WriteAllBytesAsync("upload.xlsx", upload.ToArray());
        }

        #region example: active-content
        // A zip file that says two things (KalkylLimits.StrictZip) is refused as damaged.
        if ((await Workbook.OpenAsync("upload.xlsx")).IsSuccess(out Workbook? book) == false)
        {
            return;
        }

        if (book.ActiveContent != ActiveContent.None)
        {
            Console.WriteLine(book.ActiveContent);   // Macros, ExternalReferences

            foreach (OpcExternalRelationship link in book.ExternalReferences)
            {
                Console.WriteLine($"{link.Kind}: {link.Target}");   // externalLinkPath: file:///\fileserver\share\other.xlsx, ddeLink: cmd|/c calc
            }

            Result removed = book.RemoveActiveContent();   // Invalid for a pivot table or data model fed by a connection

            if (removed.IsFailure(out string? error, out _) == true)
            {
                Console.WriteLine(error);

                return;
            }
        }

        await book.SaveAsync("clean.xlsx");   // no macro sheet, no links; the linked cells keep their last values
        #endregion

        Assert.Equal(ActiveContent.None, book.ActiveContent);
        Assert.True((await Workbook.OpenAsync("clean.xlsx")).IsSuccess(out Workbook? clean));
        Assert.Equal(ActiveContent.None, clean.ActiveContent);
        Assert.Equal((CellValue)84, clean.Sheets[0]["B1"].Value);
    }

    [Fact]
    public void Every_README_example_is_code_that_runs_here()
    {
        string source = File.ReadAllText(SourcePath());
        string readme = File.ReadAllText(Path.Combine(Path.GetDirectoryName(SourcePath())!, "..", "..", "README.md"));
        Dictionary<string, string> regions = [];

        foreach (Match match in Regex.Matches(source, @"^[ \t]*#region example: (?<name>[\w-]+)\r?\n(?<code>.*?)^[ \t]*#endregion", RegexOptions.Multiline | RegexOptions.Singleline))
        {
            regions[match.Groups["name"].Value] = Dedent(match.Groups["code"].Value);
        }

        Dictionary<string, string> blocks = [];

        foreach (Match match in Regex.Matches(readme, @"<!-- example: (?<name>[\w-]+) -->\r?\n```csharp\r?\n(?<code>.*?)```", RegexOptions.Singleline))
        {
            blocks[match.Groups["name"].Value] = Normalize(match.Groups["code"].Value);
        }

        Assert.Equal(regions.Keys.Order(StringComparer.Ordinal), blocks.Keys.Order(StringComparer.Ordinal));

        foreach ((string name, string code) in regions)
        {
            Assert.True(code == blocks[name], $"The README example '{name}' is not the code in ReadmeExamples.cs.\n--- README ---\n{blocks[name]}\n--- test ---\n{code}");
        }

        // Every csharp block in the README is a tested example.
        Assert.Equal(blocks.Count, Regex.Count(readme, "```csharp"));
    }

    private static async Task WriteToResponseAsync(Stream response)
    {
        #region example: stream-write-to-stream
        // Any writable stream, also one that cannot seek, such as an HTTP response body.
        // leaveOpen: true, so that the response is not closed with the writer.
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(response, leaveOpen: true);

        if ((await writer.AddSheetAsync("Export")).IsSuccess(out SheetWriter? sheet) == true)
        {
            await sheet.WriteRowAsync(["Exported", DateTime.UtcNow]);
        }

        await writer.CompleteAsync();   // or let `await using` complete it
        #endregion
    }

    private static async Task ToStreamAsync(Stream input, Stream output)
    {
        #region example: document-streams
        // A workbook from a stream (an upload, a blob) and back to one (a response, a blob).
        Result<Workbook> opened = await Workbook.OpenAsync(input, KalkylLimits.Default);

        if (opened.IsSuccess(out Workbook? book) == true)
        {
            book.Sheets[0]["A1"] = "Checked";
            await book.SaveAsync(output);   // the stream is left open
        }
        #endregion
    }

    private static async Task<string> HandleUploadAsync(Stream upload)
    {
        #region example: upload
        // The default limits fit files people make by hand; lower them for uploads, and limit the time too.
        KalkylLimits limits = KalkylLimits.Default with { MaxCells = 1_000_000, MaxUncompressedBytes = 100_000_000 };

        using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(10));

        Result<XlsxReader> opened = await XlsxReader.OpenAsync(upload, limits, leaveOpen: true, timeout.Token);

        if (opened.IsFailure(out string? error, out _, out XlsxReader? reader) == true)
        {
            // The error names paths and cell contents: log it, and show the user a message of your own.
            Console.WriteLine($"Upload rejected: {error}");

            return opened.ErrorKind() switch
            {
                KalkylErrorKind.TooLarge => "The file is too large.",
                KalkylErrorKind.NotXlsx => "Upload an .xlsx file.",
                KalkylErrorKind.Damaged => "The file is damaged.",
                _ => "The file could not be read.",
            };
        }

        await using (reader)
        {
            if (reader.Sheet(0).IsSuccess(out SheetReader? sheet) == false)
            {
                return "The file has no sheets.";
            }

            try
            {
                await foreach (Row row in sheet.ReadRowsAsync(timeout.Token))
                {
                    Console.WriteLine(row.Cells.Count);
                }
            }
            catch (InvalidDataException damaged)
            {
                // A row damaged half-way through the file is found when it is reached.
                Console.WriteLine($"{damaged.ErrorKind()}: {damaged.Message}");

                return "The file is damaged.";
            }
        }

        // The document model takes the same limits and keeps them for its formulas: Workbook.Limits.
        // A file you made yourself can be opened without them.
        if ((await Workbook.OpenAsync("upload.xlsx", KalkylLimits.Unlimited)).IsSuccess(out Workbook? book) == true)
        {
            Console.WriteLine(book.Limits.MaxArrayCells);
        }

        return "ok";
        #endregion
    }

    private static IEnumerable<Customer> LoadCustomers()
    {
        yield return new Customer { Name = "Ada", Balance = 1m };

        yield return new Customer { Name = "Alan", Balance = 2m };
    }

    private static async IAsyncEnumerable<Customer> LoadCustomersAsync()
    {
        foreach (Customer customer in LoadCustomers())
        {
            await Task.Yield();

            yield return customer;
        }
    }

    private static async Task WriteSampleAsync(string path)
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);
        Assert.True(created.IsSuccess(out XlsxWriter? writer));

        await using (writer)
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Sales");
            await sheet.WriteRowAsync(["Product", "Units", "Sold on", "Paid"]);
            await sheet.WriteRowAsync(["Tea", 3, new DateTime(2026, 1, 2), true]);
            await sheet.WriteRowAsync([]);
        }
    }

    private static async Task CreateBudgetAsync(string path)
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Budget").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = "Month";
        sheet["B1"] = "Amount";
        sheet["C1"] = "Target";
        sheet["D1"] = "Note";

        for (int month = 1; month <= 12; month++)
        {
            sheet[month, 0] = new DateOnly(2026, month, 1);
            sheet[month, 1] = month * 1000;
            sheet[month, 2] = 6000;
            sheet[month, 3] = "on time";
        }

        sheet.SetFormula(13, 1, "SUM(B2:B13)");
        sheet.SetColumnWidth(0, 14);
        sheet.FreezePanes(1);
        await book.RecalculateAsync();
        Assert.True((await book.SaveAsync(path)).IsSuccess(out _, out _));
    }

    private static async Task<int> CountRowsAsync(string path, string sheetName)
    {
        await using FileStream file = File.OpenRead(path);
        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Sheet(sheetName).IsSuccess(out SheetReader? sheet));
        int count = 0;

        await foreach (Row row in sheet.ReadRowsAsync())
        {
            count++;
        }

        return count;
    }

    private static string SourcePath([CallerFilePath] string path = "")
    {
        return path;
    }

    private static string Dedent(string code)
    {
        string[] lines = Normalize(code).Split('\n');
        int indent = lines.Where(line => line.Trim().Length > 0).Select(line => line.Length - line.TrimStart().Length).DefaultIfEmpty(0).Min();
        List<string> dedented = [];

        foreach (string line in lines)
        {
            if (line.Length >= indent)
            {
                dedented.Add(line[indent..]);
            }
            else
            {
                dedented.Add(line.TrimStart());
            }
        }

        return Normalize(string.Join('\n', dedented));
    }

    private static string Normalize(string code)
    {
        return code.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd() + "\n";
    }
}

/// <summary>The README examples change the current directory, so they run alone.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class ReadmeGroup
{
    public const string Name = "README examples";
}
