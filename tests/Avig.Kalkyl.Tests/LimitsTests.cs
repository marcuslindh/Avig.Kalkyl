using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Tests;

/// <summary>KalkylLimits: what a file from outside may make the package do.</summary>
public class LimitsTests
{
    private static async Task<MemoryStream> FileWithCellsAsync(int rows, int columns)
    {
        return await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");

            for (int row = 0; row < rows; row++)
            {
                Cell[] cells = new Cell[columns];

                for (int column = 0; column < columns; column++)
                {
                    cells[column] = Cell.Of(new string('x', 100));
                }

                await sheet.WriteRowAsync(cells);
            }
        });
    }

    [Fact]
    public async Task A_file_that_unpacks_past_the_limit_is_Invalid_before_anything_is_read()
    {
        await using MemoryStream file = await FileWithCellsAsync(200, 10);
        KalkylLimits limits = KalkylLimits.Default with { MaxUncompressedBytes = 50_000 };

        Result<XlsxReader> reader = await XlsxReader.OpenAsync(file, limits, true);
        file.Position = 0;
        Result<Workbook> book = await Workbook.OpenAsync(file, limits);

        Assert.True(reader.IsFailure(out string? readerError, out _));
        Assert.Contains("MaxUncompressedBytes", readerError, StringComparison.Ordinal);
        Assert.True(book.IsFailure(out string? bookError, out _));
        Assert.Contains("MaxUncompressedBytes", bookError, StringComparison.Ordinal);
    }

    [Fact]
    public async Task The_default_limits_open_an_ordinary_file()
    {
        await using MemoryStream file = await FileWithCellsAsync(200, 10);

        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        Assert.Equal(200, book.Sheets[0].RowCount);
    }

    [Fact]
    public async Task A_stream_longer_than_the_limit_is_not_read_into_memory()
    {
        await using MemoryStream junk = new(new byte[200_000]);

        Result<Workbook> opened = await Workbook.OpenAsync(junk, KalkylLimits.Default with { MaxUncompressedBytes = 10_000 });

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("longer than", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_part_that_unpacks_to_more_than_its_header_says_is_damaged()
    {
        await using Stream part = Avig.OpenXml.OpcZip.LimitStream(new MemoryStream(new byte[100]), 50, "xl/worksheets/sheet1.xml");
        byte[] buffer = new byte[30];

        await part.ReadExactlyAsync(buffer);
        InvalidDataException damaged = await Assert.ThrowsAsync<InvalidDataException>(async () => await part.ReadExactlyAsync(buffer));
        Assert.Contains("sheet1.xml", damaged.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_part_exactly_as_large_as_its_header_is_read()
    {
        await using Stream part = Avig.OpenXml.OpcZip.LimitStream(new MemoryStream(new byte[50]), 50, "part");
        await using MemoryStream copy = new();

        await part.CopyToAsync(copy);

        Assert.Equal(50, copy.Length);
    }

    [Fact]
    public async Task A_workbook_with_more_cells_than_the_limit_is_Invalid()
    {
        await using MemoryStream file = await FileWithCellsAsync(10, 3);

        Result<Workbook> tooMany = await Workbook.OpenAsync(file, KalkylLimits.Default with { MaxCells = 29 });
        file.Position = 0;
        Result<Workbook> enough = await Workbook.OpenAsync(file, KalkylLimits.Default with { MaxCells = 30 });

        Assert.True(tooMany.IsFailure(out string? error, out _));
        Assert.Contains("MaxCells", error, StringComparison.Ordinal);
        Assert.True(enough.IsSuccess(out Workbook? book));
        Assert.Equal(30, book.Limits.MaxCells);
    }

    private static Workbook NumbersIn(int rows, int columns)
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        for (int row = 0; row < rows; row++)
        {
            sheet[row, 0] = row + 1;
        }

        for (int column = 1; column < columns; column++)
        {
            sheet[0, column] = column;
        }

        return book;
    }

    [Fact]
    public void An_array_larger_than_the_limit_is_NUM()
    {
        Workbook book = NumbersIn(20, 11);
        book.Limits = KalkylLimits.Default with { MaxArrayCells = 100 };

        Assert.True(book.Evaluate("SUM(A1:A20*B1:K1)").IsSuccess(out CellValue tooLarge));
        Assert.Equal(CellValue.Error("#NUM!"), tooLarge);

        book.Limits = KalkylLimits.Default with { MaxArrayCells = 200 };

        Assert.True(book.Evaluate("SUM(A1:A20*B1:K1)").IsSuccess(out CellValue fits));
        Assert.Equal((CellValue)(210 * 55), fits);
    }

    /// <summary>60 000 x 60 000 used to overflow int in ValueArray and throw OverflowException.</summary>
    [Fact]
    public void The_default_limit_stops_an_array_of_billions_of_values()
    {
        Workbook book = NumbersIn(60_000, 1);
        Worksheet sheet = book.Sheets[0];

        for (int row = 0; row < 60_000; row++)
        {
            sheet[0, row % 16_000 + 1] = 1;
        }

        Assert.True(book.Evaluate("SUMPRODUCT(A1:A60000*B1:WQJ1)").IsSuccess(out CellValue value));
        Assert.Equal(CellValue.Error("#NUM!"), value);
    }

    [Fact]
    public async Task A_formula_cell_past_the_array_limit_is_recalculated_to_NUM()
    {
        Workbook book = NumbersIn(20, 11);
        book.Limits = KalkylLimits.Default with { MaxArrayCells = 100 };

        book.Sheets[0]["M1"] = Cell.FromFormula("SUM(A1:A20*B1:K1)");

        await book.RecalculateAsync();

        Assert.Equal(CellValue.Error("#NUM!"), book.Sheets[0]["M1"].Value);
    }

    [Fact]
    public async Task A_cancelled_evaluation_or_recalculation_stops()
    {
        Workbook book = NumbersIn(20, 11);
        book.Sheets[0]["M1"] = Cell.FromFormula("SUM(A1:A20)");
        using CancellationTokenSource cancelled = new();
        await cancelled.CancelAsync();

        Assert.Throws<OperationCanceledException>(() => book.Evaluate("SUM(A1:A20)", null, cancelled.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => book.RecalculateAsync(RecalculationScope.All, cancelled.Token));
    }

    [Fact]
    public void A_formula_longer_than_Excel_allows_is_Invalid()
    {
        Workbook book = NumbersIn(1, 1);
        string longest = "\"" + new string('a', 8190) + "\"";

        Assert.True(book.Evaluate(longest).IsSuccess(out CellValue value));
        Assert.True(value.TryGetText(out string? text));
        Assert.Equal(8190, text.Length);
        Assert.True(book.Evaluate("=" + longest).IsSuccess(out _));
        Assert.True(book.Evaluate(longest + "&\"\"").IsFailure(out string? error, out _));
        Assert.Contains("8,192", error, StringComparison.Ordinal);
    }
}
