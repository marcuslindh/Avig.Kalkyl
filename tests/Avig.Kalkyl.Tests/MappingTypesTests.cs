using System.Globalization;

namespace Avig.Kalkyl.Tests;

/// <summary>The enum of <see cref="EveryType"/>.</summary>
public enum Shade
{
    Light,
    Dark,
}

/// <summary>A row with one property of every type the generator maps, plain and nullable.</summary>
[KalkylRow]
public partial class EveryType
{
    public int Whole { get; set; }

    public int? WholeOrNull { get; set; }

    public long Big { get; set; }

    public long? BigOrNull { get; set; }

    public short Small { get; set; }

    public short? SmallOrNull { get; set; }

    public byte Tiny { get; set; }

    public byte? TinyOrNull { get; set; }

    public double Ratio { get; set; }

    public double? RatioOrNull { get; set; }

    public float Fraction { get; set; }

    public float? FractionOrNull { get; set; }

    public decimal Money { get; set; }

    public decimal? MoneyOrNull { get; set; }

    public bool Bool { get; set; }

    public bool? BoolOrNull { get; set; }

    public DateTime DateTime { get; set; }

    public DateTime? DateTimeOrNull { get; set; }

    public DateOnly DateOnly { get; set; }

    public DateOnly? DateOnlyOrNull { get; set; }

    public TimeOnly TimeOnly { get; set; }

    public TimeOnly? TimeOnlyOrNull { get; set; }

    public Guid Id { get; set; }

    public Guid? IdOrNull { get; set; }

    public Shade Shade { get; set; }

    public Shade? ShadeOrNull { get; set; }

    public CellValue Value { get; set; }

    public string? Text { get; set; }
}

/// <summary>The typed mapping for every property type: written, read back, empty, missing and wrong.</summary>
public class MappingTypesTests
{
    private static readonly string[] Headers =
    [
        "Whole", "WholeOrNull", "Big", "BigOrNull", "Small", "SmallOrNull", "Tiny", "TinyOrNull",
        "Ratio", "RatioOrNull", "Fraction", "FractionOrNull", "Money", "MoneyOrNull", "Bool", "BoolOrNull",
        "DateTime", "DateTimeOrNull", "DateOnly", "DateOnlyOrNull", "TimeOnly", "TimeOnlyOrNull",
        "Id", "IdOrNull", "Shade", "ShadeOrNull", "Value", "Text",
    ];

    private static readonly EveryType Full = new()
    {
        Whole = -7,
        WholeOrNull = 2_000_000_000,
        Big = 9_007_199_254_740_991,
        BigOrNull = -9_007_199_254_740_991,
        Small = -32768,
        SmallOrNull = 32767,
        Tiny = 0,
        TinyOrNull = 255,
        Ratio = 1.5,
        RatioOrNull = -0.25,
        Fraction = 0.1f,
        FractionOrNull = 3.75f,
        Money = 1234.56m,
        MoneyOrNull = 0.1m,
        Bool = true,
        BoolOrNull = false,
        DateTime = new DateTime(2026, 11, 24, 13, 45, 30, DateTimeKind.Unspecified),
        DateTimeOrNull = new DateTime(1999, 12, 31, 0, 0, 0, DateTimeKind.Unspecified),
        DateOnly = new DateOnly(2000, 2, 29),
        DateOnlyOrNull = new DateOnly(1900, 3, 1),
        TimeOnly = new TimeOnly(23, 59, 59),
        TimeOnlyOrNull = new TimeOnly(6, 0),
        Id = Guid.Parse("6f9619ff-8b86-d011-b42d-00cf4fc964ff"),
        IdOrNull = Guid.Parse("00000000-0000-0000-0000-000000000001"),
        Shade = Shade.Dark,
        ShadeOrNull = Shade.Light,
        Value = CellValue.Error("#N/A"),
        Text = "Åsa",
    };

    [Fact]
    public async Task Every_type_written_reads_back_the_same()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Typer");
            await sheet.WriteAsync([Full]);
        });

        Assert.Empty(TestFiles.Validate(file));

        EveryType read = Assert.Single(await ReadAllAsync(file));

        AssertSame(Full, read);
    }

    [Fact]
    public async Task Nulls_are_written_as_empty_cells_and_read_back_as_nulls()
    {
        // DateTime and DateOnly default to the year 1, which Excel cannot store.
        EveryType empty = new() { DateTime = Full.DateTime, DateOnly = Full.DateOnly };

        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Typer");
            await sheet.WriteAsync([empty]);
        });

        List<Row> rows = await TestFiles.ReadRowsAsync(file);

        for (int column = 0; column < Headers.Length; column++)
        {
            if (Nullable(Headers[column]) == true)
            {
                Assert.True(column >= rows[1].Cells.Count || rows[1][column].Value.IsEmpty == true, Headers[column]);
            }
        }

        EveryType read = Assert.Single(await ReadAllAsync(file));

        AssertSame(empty, read);
    }

    [Fact]
    public async Task Blank_text_in_a_nullable_column_reads_as_null()
    {
        List<(string Header, Cell Cell)> blanks = [.. Headers.Where(header => header.EndsWith("OrNull", StringComparison.Ordinal) == true).Select(header => (header, (Cell)"  "))];

        await using MemoryStream file = await WriteRowAsync([.. blanks]);

        EveryType read = Assert.Single(await ReadAllAsync(file));

        EveryType expected = Copy(Full);
        expected.WholeOrNull = null;
        expected.BigOrNull = null;
        expected.SmallOrNull = null;
        expected.TinyOrNull = null;
        expected.RatioOrNull = null;
        expected.FractionOrNull = null;
        expected.MoneyOrNull = null;
        expected.BoolOrNull = null;
        expected.DateTimeOrNull = null;
        expected.DateOnlyOrNull = null;
        expected.TimeOnlyOrNull = null;
        expected.IdOrNull = null;
        expected.ShadeOrNull = null;
        AssertSame(expected, read);
    }

    [Fact]
    public async Task Missing_nullable_columns_read_as_null()
    {
        string[] required = [.. Headers.Where(header => Nullable(header) == false)];

        await using MemoryStream file = await WriteRowAsync([], required);

        EveryType read = Assert.Single(await ReadAllAsync(file));

        Assert.Null(read.WholeOrNull);
        Assert.Null(read.DateTimeOrNull);
        Assert.Null(read.IdOrNull);
        Assert.Null(read.ShadeOrNull);
        Assert.Null(read.Text);
        Assert.True(read.Value.IsEmpty);
        Assert.Equal(Full.Whole, read.Whole);
    }

    [Theory]
    [InlineData("Whole")]
    [InlineData("Big")]
    [InlineData("Small")]
    [InlineData("Tiny")]
    [InlineData("Ratio")]
    [InlineData("Fraction")]
    [InlineData("Money")]
    [InlineData("Bool")]
    [InlineData("DateTime")]
    [InlineData("DateOnly")]
    [InlineData("TimeOnly")]
    [InlineData("Id")]
    [InlineData("Shade")]
    public async Task An_empty_cell_in_a_column_that_is_not_nullable_is_Invalid(string header)
    {
        await using MemoryStream file = await WriteRowAsync([(header, Cell.Empty)]);

        Result<EveryType> result = Assert.Single(await ReadResultsAsync(file));

        Assert.True(result.IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Equal($"{Address(header)} in 'Typer', column '{header}', is empty, but a value is required.", error);
    }

    [Fact]
    public async Task Text_from_other_systems_is_understood_for_every_type()
    {
        (string Header, Cell Cell)[] cells =
        [
            ("Whole", "1 200"),
            ("WholeOrNull", " 42 "),
            ("Big", "-5"),
            ("BigOrNull", "1e3"),
            ("Small", "12"),
            ("SmallOrNull", "-12"),
            ("Tiny", "7"),
            ("TinyOrNull", "255"),
            ("Ratio", "1,5"),
            ("RatioOrNull", "-2.25"),
            ("Fraction", "0.5"),
            ("FractionOrNull", "2"),
            ("Money", "10,25"),
            ("MoneyOrNull", "0.1"),
            ("Bool", "Ja"),
            ("BoolOrNull", 0),
            ("DateTime", "2026-11-24 08:15"),
            ("DateTimeOrNull", 45000),
            ("DateOnly", "2021-03-04"),
            ("DateOnlyOrNull", 45000.75),
            ("TimeOnly", "08:30"),
            ("TimeOnlyOrNull", 0.5),
            ("Id", " 6f9619ff-8b86-d011-b42d-00cf4fc964ff "),
            ("IdOrNull", "{00000000-0000-0000-0000-000000000001}"),
            ("Shade", " dark "),
            ("ShadeOrNull", 1),
            ("Value", 3),
            ("Text", 12.5),
        ];

        await using MemoryStream file = await WriteRowAsync(cells);

        EveryType read = Assert.Single(await ReadAllAsync(file));

        Assert.Equal(1200, read.Whole);
        Assert.Equal(42, read.WholeOrNull);
        Assert.Equal(-5, read.Big);
        Assert.Equal(1000, read.BigOrNull);
        Assert.Equal((short)12, read.Small);
        Assert.Equal((short)-12, read.SmallOrNull);
        Assert.Equal((byte)7, read.Tiny);
        Assert.Equal((byte)255, read.TinyOrNull);
        Assert.Equal(1.5, read.Ratio);
        Assert.Equal(-2.25, read.RatioOrNull);
        Assert.Equal(0.5f, read.Fraction);
        Assert.Equal(2f, read.FractionOrNull);
        Assert.Equal(10.25m, read.Money);
        Assert.Equal(0.1m, read.MoneyOrNull);
        Assert.True(read.Bool);
        Assert.False(read.BoolOrNull);
        Assert.Equal(new DateTime(2026, 11, 24, 8, 15, 0, DateTimeKind.Unspecified), read.DateTime);
        Assert.Equal(new DateTime(2023, 3, 15, 0, 0, 0, DateTimeKind.Unspecified), read.DateTimeOrNull);
        Assert.Equal(new DateOnly(2021, 3, 4), read.DateOnly);
        Assert.Equal(new DateOnly(2023, 3, 15), read.DateOnlyOrNull);
        Assert.Equal(new TimeOnly(8, 30), read.TimeOnly);
        Assert.Equal(new TimeOnly(12, 0), read.TimeOnlyOrNull);
        Assert.Equal(Full.Id, read.Id);
        Assert.Equal(Full.IdOrNull, read.IdOrNull);
        Assert.Equal(Shade.Dark, read.Shade);
        Assert.Equal(Shade.Dark, read.ShadeOrNull);
        Assert.Equal(CellValue.Number(3), read.Value);
        Assert.Equal("12.5", read.Text);
    }

    [Theory]
    [InlineData("Whole", "1.5", "a whole number")]
    [InlineData("WholeOrNull", "3000000000", "a whole number")]
    [InlineData("Big", "abc", "a whole number")]
    [InlineData("BigOrNull", "0.5", "a whole number")]
    [InlineData("Small", "40000", "a whole number from -32768 to 32767")]
    [InlineData("SmallOrNull", "-40000", "a whole number from -32768 to 32767")]
    [InlineData("Tiny", "256", "a whole number from 0 to 255")]
    [InlineData("TinyOrNull", "-1", "a whole number from 0 to 255")]
    [InlineData("Ratio", "tio", "a number")]
    [InlineData("RatioOrNull", "1.2.3", "a number")]
    [InlineData("Fraction", "x", "a number")]
    [InlineData("FractionOrNull", "y", "a number")]
    [InlineData("Money", "1e300", "a number")]
    [InlineData("MoneyOrNull", "z", "a number")]
    [InlineData("Bool", "kanske", "TRUE or FALSE")]
    [InlineData("BoolOrNull", "2", "TRUE or FALSE")]
    [InlineData("DateTime", "igår", "a date")]
    [InlineData("DateTimeOrNull", "-1", "a date")]
    [InlineData("DateOnly", "2026-02-30", "a date")]
    [InlineData("DateOnlyOrNull", "aldrig", "a date")]
    [InlineData("TimeOnly", "25:00", "a time")]
    [InlineData("TimeOnlyOrNull", "sen", "a time")]
    [InlineData("Id", "not-a-guid", "a GUID")]
    [InlineData("IdOrNull", "123", "a GUID")]
    [InlineData("Shade", "Purple", "one of Light, Dark")]
    [InlineData("ShadeOrNull", "5", "one of Light, Dark")]
    public async Task A_wrong_value_names_its_cell_and_what_was_expected(string header, string text, string expected)
    {
        Cell cell = text;

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == true)
        {
            cell = number;
        }

        await using MemoryStream file = await WriteRowAsync([(header, cell)]);

        Result<EveryType> result = Assert.Single(await ReadResultsAsync(file));

        Assert.True(result.IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.StartsWith($"{Address(header)} in 'Typer', column '{header}', holds '", error, StringComparison.Ordinal);
        Assert.EndsWith(", which is not " + expected + ".", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Id")]
    [InlineData("Whole")]
    [InlineData("TimeOnly")]
    public async Task A_boolean_where_a_number_date_or_GUID_belongs_is_wrong(string header)
    {
        await using MemoryStream file = await WriteRowAsync([(header, true)]);

        Result<EveryType> result = Assert.Single(await ReadResultsAsync(file));

        Assert.True(result.IsFailure(out string? error, out _));
        Assert.Contains("holds 'TRUE', which is not", error, StringComparison.Ordinal);
    }

    private static bool Nullable(string header)
    {
        return header.EndsWith("OrNull", StringComparison.Ordinal) == true || header == "Text" || header == "Value";
    }

    private static string Address(string header)
    {
        int column = Array.IndexOf(Headers, header);

        return $"Cell {new CellAddress(1, column)} (row 1, column {column})";
    }

    /// <summary>A sheet "Typer" with the headers and one row of <see cref="Full"/>'s values, except for the cells given.</summary>
    private static async Task<MemoryStream> WriteRowAsync((string Header, Cell Cell)[] changes, string[]? headers = null)
    {
        Dictionary<string, Cell> cells = ValidCells();

        foreach ((string header, Cell cell) in changes)
        {
            cells[header] = cell;
        }

        string[] columns = headers ?? Headers;

        return await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Typer");
            await sheet.WriteRowAsync([.. columns.Select(header => (Cell)header)]);
            await sheet.WriteRowAsync([.. columns.Select(header => cells[header])]);
        });
    }

    private static Dictionary<string, Cell> ValidCells()
    {
        return new Dictionary<string, Cell>
        {
            ["Whole"] = KalkylCells.From((int?)Full.Whole),
            ["WholeOrNull"] = KalkylCells.From(Full.WholeOrNull),
            ["Big"] = KalkylCells.From((long?)Full.Big),
            ["BigOrNull"] = KalkylCells.From(Full.BigOrNull),
            ["Small"] = KalkylCells.From((short?)Full.Small),
            ["SmallOrNull"] = KalkylCells.From(Full.SmallOrNull),
            ["Tiny"] = KalkylCells.From((byte?)Full.Tiny),
            ["TinyOrNull"] = KalkylCells.From(Full.TinyOrNull),
            ["Ratio"] = KalkylCells.From((double?)Full.Ratio),
            ["RatioOrNull"] = KalkylCells.From(Full.RatioOrNull),
            ["Fraction"] = KalkylCells.From((float?)Full.Fraction),
            ["FractionOrNull"] = KalkylCells.From(Full.FractionOrNull),
            ["Money"] = KalkylCells.From((decimal?)Full.Money),
            ["MoneyOrNull"] = KalkylCells.From(Full.MoneyOrNull),
            ["Bool"] = KalkylCells.From((bool?)Full.Bool),
            ["BoolOrNull"] = KalkylCells.From(Full.BoolOrNull),
            ["DateTime"] = KalkylCells.From((DateTime?)Full.DateTime),
            ["DateTimeOrNull"] = KalkylCells.From(Full.DateTimeOrNull),
            ["DateOnly"] = KalkylCells.From((DateOnly?)Full.DateOnly),
            ["DateOnlyOrNull"] = KalkylCells.From(Full.DateOnlyOrNull),
            ["TimeOnly"] = KalkylCells.From((TimeOnly?)Full.TimeOnly),
            ["TimeOnlyOrNull"] = KalkylCells.From(Full.TimeOnlyOrNull),
            ["Id"] = KalkylCells.From((Guid?)Full.Id),
            ["IdOrNull"] = KalkylCells.From(Full.IdOrNull),
            ["Shade"] = KalkylCells.FromEnum<Shade>(Full.Shade),
            ["ShadeOrNull"] = KalkylCells.FromEnum(Full.ShadeOrNull),
            ["Value"] = KalkylCells.From(Full.Value),
            ["Text"] = KalkylCells.From(Full.Text),
        };
    }

    private static EveryType Copy(EveryType source)
    {
        return new EveryType
        {
            Whole = source.Whole,
            WholeOrNull = source.WholeOrNull,
            Big = source.Big,
            BigOrNull = source.BigOrNull,
            Small = source.Small,
            SmallOrNull = source.SmallOrNull,
            Tiny = source.Tiny,
            TinyOrNull = source.TinyOrNull,
            Ratio = source.Ratio,
            RatioOrNull = source.RatioOrNull,
            Fraction = source.Fraction,
            FractionOrNull = source.FractionOrNull,
            Money = source.Money,
            MoneyOrNull = source.MoneyOrNull,
            Bool = source.Bool,
            BoolOrNull = source.BoolOrNull,
            DateTime = source.DateTime,
            DateTimeOrNull = source.DateTimeOrNull,
            DateOnly = source.DateOnly,
            DateOnlyOrNull = source.DateOnlyOrNull,
            TimeOnly = source.TimeOnly,
            TimeOnlyOrNull = source.TimeOnlyOrNull,
            Id = source.Id,
            IdOrNull = source.IdOrNull,
            Shade = source.Shade,
            ShadeOrNull = source.ShadeOrNull,
            Value = source.Value,
            Text = source.Text,
        };
    }

    private static void AssertSame(EveryType expected, EveryType actual)
    {
        Assert.Equal(expected.Whole, actual.Whole);
        Assert.Equal(expected.WholeOrNull, actual.WholeOrNull);
        Assert.Equal(expected.Big, actual.Big);
        Assert.Equal(expected.BigOrNull, actual.BigOrNull);
        Assert.Equal(expected.Small, actual.Small);
        Assert.Equal(expected.SmallOrNull, actual.SmallOrNull);
        Assert.Equal(expected.Tiny, actual.Tiny);
        Assert.Equal(expected.TinyOrNull, actual.TinyOrNull);
        Assert.Equal(expected.Ratio, actual.Ratio);
        Assert.Equal(expected.RatioOrNull, actual.RatioOrNull);
        Assert.Equal(expected.Fraction, actual.Fraction);
        Assert.Equal(expected.FractionOrNull, actual.FractionOrNull);
        Assert.Equal(expected.Money, actual.Money);
        Assert.Equal(expected.MoneyOrNull, actual.MoneyOrNull);
        Assert.Equal(expected.Bool, actual.Bool);
        Assert.Equal(expected.BoolOrNull, actual.BoolOrNull);
        Assert.Equal(expected.DateTime, actual.DateTime);
        Assert.Equal(expected.DateTimeOrNull, actual.DateTimeOrNull);
        Assert.Equal(expected.DateOnly, actual.DateOnly);
        Assert.Equal(expected.DateOnlyOrNull, actual.DateOnlyOrNull);
        Assert.Equal(expected.TimeOnly, actual.TimeOnly);
        Assert.Equal(expected.TimeOnlyOrNull, actual.TimeOnlyOrNull);
        Assert.Equal(expected.Id, actual.Id);
        Assert.Equal(expected.IdOrNull, actual.IdOrNull);
        Assert.Equal(expected.Shade, actual.Shade);
        Assert.Equal(expected.ShadeOrNull, actual.ShadeOrNull);
        Assert.Equal(expected.Value, actual.Value);
        Assert.Equal(expected.Text, actual.Text);
    }

    private static async Task<List<EveryType>> ReadAllAsync(MemoryStream file)
    {
        List<EveryType> items = [];

        foreach (Result<EveryType> result in await ReadResultsAsync(file))
        {
            Assert.True(result.IsSuccess(out EveryType? item), result.Error);
            items.Add(item);
        }

        return items;
    }

    private static async Task<List<Result<EveryType>>> ReadResultsAsync(MemoryStream file)
    {
        await using XlsxReader reader = await TestFiles.OpenAsync(file);
        Assert.True(reader.Sheet(0).IsSuccess(out SheetReader? sheet));
        List<Result<EveryType>> results = [];

        await foreach (Result<EveryType> result in sheet.ReadAsync<EveryType>())
        {
            results.Add(result);
        }

        return results;
    }
}
