using Avig.Kalkyl.Samples;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// The README gallery (samples/Avig.Kalkyl.Samples): every workbook builds, passes Microsoft's
/// schema validator, opens again in the document model, and calculates without errors. Excel
/// itself opened them without repair (tools/Avig.Kalkyl.ExcelCheck/open-in-excel.ps1), and the
/// pictures in docs/images/gallery are Excel's rendering of them.
/// </summary>
public class GalleryTests
{
    public static TheoryData<string> Samples()
    {
        TheoryData<string> names = [];

        foreach (GallerySample sample in Gallery.All)
        {
            if (sample.Name != "large-export")
            {
                names.Add(sample.Name);
            }
        }

        return names;
    }

    [Theory]
    [MemberData(nameof(Samples))]
    public async Task A_gallery_workbook_is_valid_opens_again_and_calculates_without_errors(string name)
    {
        GallerySample sample = Gallery.All.Single(item => item.Name == name);
        string path = Path.Combine(Path.GetTempPath(), "avig-gallery-" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            await sample.WriteAsync(path);

            await using (FileStream file = File.OpenRead(path))
            {
                List<string> problems = TestFiles.Validate(file);
                Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
            }

            Result<Workbook> opened = await Workbook.OpenAsync(path);
            Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
            Assert.True(book.Sheet(sample.Sheet).IsSuccess(out _));
            Assert.True(CellRange.TryParse(sample.Range, out _));

            RecalculationReport report = await book.RecalculateAsync(RecalculationScope.All);
            Assert.Empty(report.UnknownFunctions);
            Assert.Empty(report.InvalidFormulas);
            Assert.Empty(report.CircularReferences);
            Assert.Equal([], Errors(book));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task The_sales_dashboard_reads_its_pivot_tables_with_GETPIVOTDATA()
    {
        Workbook book = SalesDashboard.Build();
        await book.RecalculateAsync(RecalculationScope.All);

        Assert.True(book.Sheet("Chart data").IsSuccess(out Worksheet? chartData));
        Assert.True(chartData["C2"].TryGetNumber(out double january) && january > 0);
        Assert.True(book.Sheet(SalesDashboard.Dashboard).IsSuccess(out Worksheet? dashboard));
        Assert.True(dashboard["B11"].TryGetText(out string? sales) && sales.EndsWith(" M SEK", StringComparison.Ordinal), dashboard["B11"].ToString());
        Assert.Equal(3, book.PivotTables.Count);
    }

    [Fact]
    public async Task The_large_export_streams_rows_that_validate_and_read_back()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(writer => LargeExport.WriteAsync(writer, 2_000));

        Assert.Empty(TestFiles.Validate(file));
        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        Assert.Equal(2_001, rows.Count);
        Assert.True(rows[1].TryGetNumber(7, out double amount) && amount > 0);
    }

    /// <summary>Every cell that holds an error value, as "Sheet!A1 #VALUE!".</summary>
    private static List<string> Errors(Workbook book)
    {
        List<string> errors = [];

        foreach (Worksheet sheet in book.Sheets)
        {
            foreach (Row row in sheet.Rows)
            {
                for (int column = 0; column < row.ColumnCount; column++)
                {
                    if (row[column].TryGetError(out string? code) == true)
                    {
                        errors.Add($"{sheet.Name}!{new CellAddress(row.Index, column)} {code}");
                    }
                }
            }
        }

        return errors;
    }
}
