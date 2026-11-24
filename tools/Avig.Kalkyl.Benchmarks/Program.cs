using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;

namespace Avig.Kalkyl.Benchmarks;

/// <summary>
/// Baselines for docs/testplan.md (P-01 to P-03, P-11): reading, writing, opening and
/// recalculating. Run on a quiet machine:
/// <c>dotnet run -c Release --project tools/Avig.Kalkyl.Benchmarks -- --filter *</c>.
/// A change that makes one more than 20 % slower or larger than docs/benchmarks/baseline.md should be looked into.
/// </summary>
public static class Program
{
    public static void Main(string[] args)
    {
        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
    }
}

/// <summary>Reading and writing a sheet of rows x 10 columns, half numbers and half text.</summary>
[MemoryDiagnoser]
public class RowBenchmarks
{
    private byte[] file = [];

    [Params(100_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public async Task SetupAsync()
    {
        file = await WriteAsync();
    }

    [Benchmark]
    public async Task<byte[]> WriteAsync()
    {
        await using MemoryStream stream = new();

        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(stream, true))
        {
            Result<SheetWriter> added = await writer.AddSheetAsync("Data");
            added.IsSuccess(out SheetWriter? sheet);
            Cell[] cells = new Cell[10];

            for (int row = 0; row < Rows; row++)
            {
                for (int column = 0; column < 10; column++)
                {
                    cells[column] = Cell.Of(Value(row, column));
                }

                await sheet!.WriteRowAsync(cells);
            }
        }

        return stream.ToArray();
    }

    [Benchmark]
    public async Task<int> StreamAsync()
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(new MemoryStream(file));
        opened.IsSuccess(out XlsxReader? reader);
        int cells = 0;

        await using (reader)
        {
            reader!.Sheet(0).IsSuccess(out SheetReader? sheet);

            await foreach (Row row in sheet!.ReadRowsAsync())
            {
                cells += row.ColumnCount;
            }
        }

        return cells;
    }

    [Benchmark]
    public async Task<int> OpenDocumentAsync()
    {
        Result<Workbook> opened = await Workbook.OpenAsync(new MemoryStream(file));
        opened.IsSuccess(out Workbook? book);

        return book!.Sheets[0].RowCount;
    }

    private static CellValue Value(int row, int column)
    {
        if (column % 2 == 0)
        {
            return (row * 10) + column;
        }

        return "Text " + row;
    }
}

/// <summary>Recalculating a chain of formulas, each depending on the one above, and a block of aggregates.</summary>
[MemoryDiagnoser]
public class FormulaBenchmarks
{
    private Workbook chain = Workbook.Create();
    private Workbook aggregates = Workbook.Create();

    [Params(100_000)]
    public int Formulas { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        chain = Workbook.Create();
        chain.AddSheet("Data").IsSuccess(out Worksheet? sheet);
        sheet!["A1"] = 1;

        for (int row = 1; row < Formulas; row++)
        {
            sheet[row, 0] = Cell.FromFormula($"A{row}+1");
        }

        aggregates = Workbook.Create();
        aggregates.AddSheet("Data").IsSuccess(out Worksheet? data);

        for (int row = 0; row < 10_000; row++)
        {
            data![row, 0] = row;
            data[row, 1] = row % 7;
            data[row, 2] = Cell.FromFormula($"SUMIF(B:B,B{row + 1},A:A)");
        }
    }

    [Benchmark]
    public Task<RecalculationReport> RecalculateChainAsync()
    {
        return chain.RecalculateAsync(RecalculationScope.All);
    }

    [Benchmark]
    public Task<RecalculationReport> RecalculateSumIfAsync()
    {
        return aggregates.RecalculateAsync(RecalculationScope.All);
    }
}
