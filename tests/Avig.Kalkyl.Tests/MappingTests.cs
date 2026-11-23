using System.Collections.Immutable;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace Avig.Kalkyl.Tests;

public enum CustomerLevel
{
    Bronze,
    Silver,
    Gold,
}

[KalkylRow]
public partial class Customer
{
    [KalkylColumn("Namn", Width = 30)]
    public required string Name { get; init; }

    [KalkylColumn(NumberFormat = "# ##0.00")]
    public decimal Balance { get; init; }

    public DateOnly? Since { get; init; }

    public int Visits { get; set; }

    public bool Active { get; set; }

    public CustomerLevel Level { get; set; }

    public CustomerLevel? Previous { get; set; }

    public string? Note { get; set; }

    public TimeOnly? Opens { get; set; }

    public Guid? Id { get; set; }

    public double? Score { get; set; }

    public long Big { get; set; }

    [KalkylIgnore]
    public string Secret { get; set; } = string.Empty;

    [KalkylColumn("Nr", Order = 0)]
    public short Number { get; set; }
}

[KalkylRow]
public partial record Point
{
    public int X { get; init; }

    public int Y { get; init; }
}

public class MappingTests
{
    private static readonly string[] CustomerHeaders = ["Nr", "Namn", "Balance", "Since", "Visits", "Active", "Level", "Previous", "Note", "Opens", "Id", "Score", "Big"];

    private static readonly Customer Ada = new()
    {
        Number = 1,
        Name = "Ada",
        Balance = 1234.56m,
        Since = new DateOnly(2020, 5, 17),
        Visits = 12,
        Active = true,
        Level = CustomerLevel.Gold,
        Previous = CustomerLevel.Silver,
        Note = "Första kunden",
        Opens = new TimeOnly(8, 30),
        Id = Guid.Parse("6f9619ff-8b86-d011-b42d-00cf4fc964ff"),
        Score = 4.5,
        Big = 9_007_199_254_740_991,
        Secret = "hemligt",
    };

    private static readonly Customer Bertil = new()
    {
        Number = 2,
        Name = "Bertil",
        Balance = -0.1m,
        Level = CustomerLevel.Bronze,
    };

    [Fact]
    public async Task A_list_written_reads_back_the_same()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Kunder");
            await sheet.WriteAsync([Ada, Bertil]);
        });

        List<Customer> read = await ReadAllAsync<Customer>(file);

        Assert.Equal(2, read.Count);
        AssertSame(Ada, read[0]);
        AssertSame(Bertil, read[1]);
        Assert.Equal(string.Empty, read[0].Secret);
    }

    [Fact]
    public async Task The_header_follows_order_and_names_and_is_bold()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Kunder");
            await sheet.WriteAsync([Ada]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        string sheetXml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");

        Assert.Equal(CustomerHeaders, rows[0].Cells.Select(cell => cell.Value.ToString()));
        Assert.Equal("# ##0.00", rows[1][2].Style.NumberFormat);
        Assert.Contains("<col min=\"2\" max=\"2\" width=\"30\" customWidth=\"1\" />", sheetXml, StringComparison.Ordinal);
        Assert.Contains("state=\"frozen\"", sheetXml, StringComparison.Ordinal);
        Assert.Empty(TestFiles.Validate(file));
    }

    [Fact]
    public async Task Columns_are_matched_by_header_in_any_order_and_case()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([]);
            await sheet.WriteRowAsync(["  y ", "Extra", "x"]);
            await sheet.WriteRowAsync([2, "ignoreras", 1]);
            await sheet.WriteRowAsync([]);
            await sheet.WriteRowAsync(["4", Cell.Empty, "3"]);
        });

        List<Point> points = await ReadAllAsync<Point>(file);

        Assert.Equal(new[] { new Point { X = 1, Y = 2 }, new Point { X = 3, Y = 4 } }, points);
    }

    [Fact]
    public async Task Text_from_other_systems_is_understood()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["Nr", "Namn", "Balance", "Since", "Visits", "Active", "Level", "Big"]);
            await sheet.WriteRowAsync(["7", "Cecilia", "1,5", "2021-03-04", "1 200", "ja", "gold", 45000]);
        });

        Customer customer = Assert.Single(await ReadAllAsync<Customer>(file));

        Assert.Equal(1.5m, customer.Balance);
        Assert.Equal(new DateOnly(2021, 3, 4), customer.Since);
        Assert.Equal(1200, customer.Visits);
        Assert.True(customer.Active);
        Assert.Equal(CustomerLevel.Gold, customer.Level);
        Assert.Null(customer.Note);
        Assert.Null(customer.Previous);
    }

    [Fact]
    public async Task A_bad_cell_names_its_address_and_reading_goes_on()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["X", "Y"]);
            await sheet.WriteRowAsync([1, "abc"]);
            await sheet.WriteRowAsync([1.5, 2]);
            await sheet.WriteRowAsync([5, 6]);
        });

        List<Result<Point>> results = await ReadResultsAsync<Point>(file);

        Assert.Equal(3, results.Count);
        Assert.True(results[0].IsFailure(out string? first, out _));
        Assert.Equal("Cell B2 (row 1, column 1) in 'Data', column 'Y', holds 'abc', which is not a whole number.", first);
        Assert.True(results[1].IsFailure(out string? second, out _));
        Assert.Contains("'1.5', which is not a whole number", second, StringComparison.Ordinal);
        Assert.True(results[2].IsSuccess(out Point? point));
        Assert.Equal(new Point { X = 5, Y = 6 }, point);
    }

    [Fact]
    public async Task An_empty_required_value_is_Invalid()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["Nr", "Namn", "Balance", "Visits", "Active", "Level", "Big"]);
            await sheet.WriteRowAsync([1, Cell.Empty, 0, 0, false, "Gold", 0]);
        });

        Result<Customer> result = Assert.Single(await ReadResultsAsync<Customer>(file));

        Assert.True(result.IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Contains("column 'Namn', is empty, but a value is required", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_missing_required_column_stops_with_one_failure()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["X", "Z"]);
            await sheet.WriteRowAsync([1, 2]);
        });

        Result<Point> result = Assert.Single(await ReadResultsAsync<Point>(file));

        Assert.True(result.IsFailure(out string? error, out _));
        Assert.Equal("The header in row 1 of 'Data' lacks the required column(s) 'Y'. It has 'X', 'Z'.", error);
    }

    [Fact]
    public async Task Whole_files_can_be_written_and_read_by_path()
    {
        string path = Path.Combine(Path.GetTempPath(), "avig-kalkyl-" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            Result written = await Xlsx.WriteAsync(path, [new Point { X = 1, Y = 2 }], "Punkter");
            Assert.True(written.IsSuccess(out _, out _));

            List<Point> points = [];

            await foreach (Result<Point> result in Xlsx.ReadAsync<Point>(path, "punkter"))
            {
                Assert.True(result.IsSuccess(out Point? point), result.Error);
                points.Add(point);
            }

            Assert.Equal(new[] { new Point { X = 1, Y = 2 } }, points);
        }
        finally
        {
            File.Delete(path);
        }

        Result<Point> missing = Assert.Single(await ToListAsync(Xlsx.ReadAsync<Point>(path)));
        Assert.True(missing.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
    }

    [Theory]
    [InlineData("public class Plain { public int X { get; set; } }", "KALK001")]
    [InlineData("public partial class Odd { public System.Uri Link { get; set; } }", "KALK002")]
    [InlineData("public partial class Positional(int X) { public int Y { get; set; } }", "KALK003")]
    [InlineData("public partial class Outer { [Avig.Kalkyl.KalkylRow] public partial class Inner { public int X { get; set; } } }", "KALK004")]
    public void The_generator_explains_what_it_cannot_map(string declaration, string diagnosticId)
    {
        string source = declaration;

        if (declaration.Contains("Outer", StringComparison.Ordinal) == false)
        {
            source = "[Avig.Kalkyl.KalkylRow] " + declaration;
        }

        ImmutableArray<Diagnostic> diagnostics = RunGenerator(source);

        Assert.Contains(diagnostics, diagnostic => diagnostic.Id == diagnosticId);
    }

    private static ImmutableArray<Diagnostic> RunGenerator(string source)
    {
        string[] platform = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES")!).Split(Path.PathSeparator);
        List<MetadataReference> references = platform.Select(path => (MetadataReference)MetadataReference.CreateFromFile(path)).ToList();
        references.Add(MetadataReference.CreateFromFile(typeof(KalkylRowAttribute).Assembly.Location));
        references.Add(MetadataReference.CreateFromFile(typeof(Result).Assembly.Location));
        CSharpCompilation compilation = CSharpCompilation.Create(
            "Probe",
            [CSharpSyntaxTree.ParseText(source)],
            references,
            new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        GeneratorDriver driver = CSharpGeneratorDriver.Create(new Avig.Kalkyl.Generators.KalkylRowGenerator());
        driver.RunGeneratorsAndUpdateCompilation(compilation, out _, out ImmutableArray<Diagnostic> diagnostics);

        return diagnostics;
    }

    private static void AssertSame(Customer expected, Customer actual)
    {
        Assert.Equal(expected.Number, actual.Number);
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.Balance, actual.Balance);
        Assert.Equal(expected.Since, actual.Since);
        Assert.Equal(expected.Visits, actual.Visits);
        Assert.Equal(expected.Active, actual.Active);
        Assert.Equal(expected.Level, actual.Level);
        Assert.Equal(expected.Previous, actual.Previous);
        Assert.Equal(expected.Note, actual.Note);
        Assert.Equal(expected.Opens, actual.Opens);
        Assert.Equal(expected.Id, actual.Id);
        Assert.Equal(expected.Score, actual.Score);
        Assert.Equal(expected.Big, actual.Big);
    }

    private static async Task<List<T>> ReadAllAsync<T>(MemoryStream file)
        where T : IKalkylRow<T>
    {
        List<T> items = [];

        foreach (Result<T> result in await ReadResultsAsync<T>(file))
        {
            Assert.True(result.IsSuccess(out T? item), result.Error);
            items.Add(item);
        }

        return items;
    }

    private static async Task<List<Result<T>>> ReadResultsAsync<T>(MemoryStream file)
        where T : IKalkylRow<T>
    {
        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Sheet(0).IsSuccess(out SheetReader? sheet));

        return await ToListAsync(sheet.ReadAsync<T>());
    }

    private static async Task<List<TItem>> ToListAsync<TItem>(IAsyncEnumerable<TItem> items)
    {
        List<TItem> list = [];

        await foreach (TItem item in items)
        {
            list.Add(item);
        }

        return list;
    }
}
