namespace Avig.Kalkyl.Tests;

public class RoundTripTests
{
    public static TheoryData<string, CellValue> Values => new()
    {
        { "text", "Ada Lovelace" },
        { "spaces kept", "  padded  " },
        { "new line and tab", "line one\nline two\tend" },
        { "control character", "a\u0001b" },
        { "looks like an escape", "_x0041_" },
        { "emoji", "kalkyl 😀" },
        { "longest text", new string('k', CellValue.MaxTextLength) },
        { "zero", 0 },
        { "decimal fraction", 0.1 + 0.2 },
        { "large", 1e300 },
        { "tiny negative", -1.5e-300 },
        { "int max", int.MaxValue },
        { "decimal", 1234.56m },
        { "true", true },
        { "false", false },
        { "date", new DateTime(2026, 11, 23) },
        { "date and time", new DateTime(2026, 11, 23, 13, 45, 30, 123) },
        { "first day", new DateTime(1900, 1, 1) },
        { "before the made up day", new DateTime(1900, 2, 28) },
        { "after the made up day", new DateTime(1900, 3, 1) },
        { "time only", new DateTime(1899, 12, 31, 6, 30, 0) },
        { "last day", new DateTime(9999, 12, 31) },
        { "date only", new DateOnly(2026, 6, 6) },
        { "error", CellValue.Error("#N/A") },
        { "division error", CellValue.Error("#DIV/0!") },
    };

    [Theory]
    [MemberData(nameof(Values))]
    public async Task A_value_written_reads_back_the_same(string description, CellValue value)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([description, value]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal(value, rows[0][1].Value);
        Assert.Equal(description, rows[0][0].Value.ToString());
    }

    [Fact]
    public async Task A_file_with_every_feature_passes_Microsofts_schema_validation()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter customers = await TestFiles.AddSheetAsync(writer, "Kunder & 'år' 2026");
            customers.SetColumnWidth(0, 30);
            customers.SetColumnWidth(2, 12.5);
            customers.FreezePanes(1, 1);
            customers.EnableAutoFilter();
            await customers.WriteRowAsync([Cell.Of("Namn").Bold().Background("#DDEBF7"), Cell.Of("Datum").Bold(), Cell.Of("Belopp").Bold().Align(HorizontalAlignment.Right)]);
            await customers.WriteRowAsync(["Ada", new DateTime(2026, 1, 2), Cell.Of(1234.5).NumberFormat("# ##0.00 \"kr\"")]);
            await customers.WriteRowAsync([Cell.Of("Lång text som bryts").Wrap().Align(VerticalAlignment.Top), new DateTime(2026, 1, 2, 8, 30, 0), Cell.Of(-5).FontColor("#C00000").Italic().Underline()]);
            await customers.WriteRowAsync([Cell.Of("Summa").Bold(), Cell.Empty, Cell.FromFormula("=SUM(C2:C3)", 1229.5).Bold()]);
            await customers.WriteRowAsync([]);
            await customers.WriteRowAsync([Cell.Of("Arial 14").Font("Arial").FontSize(14), true, CellValue.Error("#N/A"), Cell.FromFormula("A1&\"!\"")]);
            customers.AddConditionalFormat("C2:C4", ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, new Style { FontColor = "#9C0006", Background = "#FFC7CE", NumberFormat = "0.0" }));
            customers.AddConditionalFormat("C2:C4", ConditionalFormat.IconSet(IconSetKind.ThreeArrows));
            Assert.True(customers.AddChart(new Chart(ChartType.Column) { Title = "Belopp", Series = [new ChartSeries("Belopp", "C2:C3", "A2:A3")] }, "F2:L16").IsSuccess(out _, out _));
            customers.AddPicture(DrawingTests.PngPicture(), "N2");

            SheetWriter second = await TestFiles.AddSheetAsync(writer, "Andra");
            second.FreezePanes(2);
            await second.WriteRowAsync(["bara rader frysta"]);

            SheetWriter third = await TestFiles.AddSheetAsync(writer, "Tredje");
            third.FreezePanes(0, 3);
        });

        List<string> problems = TestFiles.Validate(file);

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    [Fact]
    public async Task Formulas_keep_their_text_and_cached_result()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([1, 2, Cell.FromFormula("=SUM(A1:B1)", 3), Cell.FromFormula("A1&\"x\"", "1x"), Cell.FromFormula("NOW()")]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal("SUM(A1:B1)", rows[0][2].Formula);
        Assert.Equal((CellValue)3, rows[0][2].Value);
        Assert.Equal("A1&\"x\"", rows[0][3].Formula);
        Assert.Equal((CellValue)"1x", rows[0][3].Value);
        Assert.Equal("NOW()", rows[0][4].Formula);
        Assert.True(rows[0][4].Value.IsEmpty);
        Assert.Contains("fullCalcOnLoad=\"1\"", TestFiles.ReadPart(file, "xl/workbook.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Number_formats_come_back_and_dates_get_one_when_none_is_given()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([Cell.Of(1.5).NumberFormat("0.00"), new DateTime(2026, 1, 2), new DateTime(2026, 1, 2, 3, 4, 5), Cell.Of(new DateTime(2026, 1, 2)).NumberFormat("d mmmm yyyy")]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal("0.00", rows[0][0].Style.NumberFormat);
        Assert.Equal("yyyy-mm-dd", rows[0][1].Style.NumberFormat);
        Assert.Equal("yyyy-mm-dd hh:mm:ss", rows[0][2].Style.NumberFormat);
        Assert.Equal("d mmmm yyyy", rows[0][3].Style.NumberFormat);
        Assert.Equal(CellKind.Date, rows[0][3].Kind);
    }

    [Fact]
    public async Task Gaps_are_kept_in_rows_and_columns()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["A1", Cell.Empty, Cell.Empty, "D1"]);
            await sheet.WriteRowAsync([]);
            await sheet.WriteRowAsync([Cell.Empty, "B3"]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        Assert.Equal(3, rows.Count);
        Assert.Equal(4, rows[0].ColumnCount);
        Assert.True(rows[0][1].Value.IsEmpty);
        Assert.Equal((CellValue)"D1", rows[0][3].Value);
        Assert.True(rows[1].IsEmpty);
        Assert.Equal(2, rows[2].Index);
        Assert.True(rows[2].TryGetText(1, out string? text));
        Assert.Equal("B3", text);
        Assert.True(rows[2][99].Value.IsEmpty);
    }

    [Fact]
    public async Task Many_rows_stream_through_both_ways()
    {
        const int rowCount = 20_000;

        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");

            for (int index = 0; index < rowCount; index++)
            {
                await sheet.WriteRowAsync([index, "rad " + index, Cell.Of(index * 1.5).Bold(), new DateTime(2026, 1, 1).AddMinutes(index), index % 2 == 0]);
            }
        });

        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Sheet("Data").IsSuccess(out SheetReader? sheet));
        int read = 0;

        await foreach (Row row in sheet.ReadRowsAsync())
        {
            Assert.True(row.TryGetNumber(0, out double number));
            Assert.Equal(row.Index, (int)number);
            read++;
        }

        Assert.Equal(rowCount, read);
    }

    [Fact]
    public async Task A_file_on_disk_can_be_written_and_read_by_path()
    {
        string path = Path.Combine(Path.GetTempPath(), "avig-kalkyl-" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);
            Assert.True(created.IsSuccess(out XlsxWriter? writer), created.Error);

            await using (writer)
            {
                SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
                await sheet.WriteRowAsync(["på disk"]);
            }

            Result<XlsxReader> opened = await XlsxReader.OpenAsync(path);
            Assert.True(opened.IsSuccess(out XlsxReader? reader), opened.Error);

            await using (reader)
            {
                Assert.Equal("Data", reader.Sheets[0].Name);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }
}
