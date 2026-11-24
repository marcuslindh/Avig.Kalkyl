using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Fuzz;

/// <summary>
/// <c>roundtrip</c> — <strong>a value the package agrees to write reads back as the same value.</strong>
///
/// This is the promise the whole package rests on, and it fails quietly: a lost leading space,
/// a lone \r folded away by the XML reader, a number written with too few digits. Nothing
/// throws, the file opens, and the value is different. The input is a list of cells decoded
/// from bytes; each is written with <see cref="XlsxWriter"/> and with the document model, and
/// read back with <see cref="XlsxReader"/> and <see cref="Workbook"/>.
/// </summary>
internal static class RoundTripTarget
{
    private const int Columns = 6;

    private static readonly string[] ErrorCodes = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA"];

    private static readonly DateTime LatestDate = new(9999, 12, 31, 23, 59, 59, 999, DateTimeKind.Unspecified);

    public static FuzzTarget Create()
    {
        return new FuzzTarget("roundtrip", "values written with the writer and the document model read back unchanged", RunAsync, Generate);
    }

    private static byte[] Generate(Random random)
    {
        List<byte> bytes = [];
        int cells = random.Next(1, 30);

        for (int cell = 0; cell < cells; cell++)
        {
            int kind = random.Next(6);
            bytes.Add((byte)kind);

            switch (kind)
            {
                case 1:
                    byte[] text = Inputs.Utf16(Inputs.Text(random, 12));
                    bytes.Add((byte)Math.Min(text.Length / 2, 255));
                    bytes.AddRange(text.Take(510));
                    break;
                case 2:
                    double[] numbers = [0, -0.0, 1, -1, 0.1, 1.0 / 3, Math.PI, 1e-310, double.Epsilon, double.MaxValue, -double.MaxValue, 1e15 + 0.3, 123456789.123456789, 5e-324];
                    double number = Inputs.Pick(random, numbers);

                    if (random.Next(3) == 0)
                    {
                        number = BitConverter.Int64BitsToDouble(random.NextInt64());
                    }

                    AddLong(bytes, BitConverter.DoubleToInt64Bits(number));
                    break;
                case 4:
                    DateTime[] dates = [new(1899, 12, 31), new(1900, 1, 1), new(1900, 2, 28, 23, 59, 59), new(1900, 3, 1), new(1904, 1, 1), new(2026, 1, 2, 12, 0, 0, 500), LatestDate, new(2026, 3, 29, 2, 30, 0)];
                    DateTime date = Inputs.Pick(random, dates);

                    if (random.Next(2) == 0)
                    {
                        date = new DateTime(random.NextInt64(DateTime.MaxValue.Ticks), DateTimeKind.Unspecified);
                    }

                    AddLong(bytes, date.Ticks);
                    break;
                default:
                    bytes.Add((byte)random.Next(256));
                    break;
            }
        }

        return [.. bytes];
    }

    private static void AddLong(List<byte> bytes, long value)
    {
        for (int shift = 56; shift >= 0; shift -= 8)
        {
            bytes.Add((byte)(value >> shift));
        }
    }

    /// <summary>
    /// The cells in the input. A value the package refuses (NaN, a date before 1900, text over
    /// 32 767 characters) is left out: refusing it is the documented answer, with an ArgumentException.
    /// </summary>
    private static List<CellValue> Decode(byte[] input)
    {
        ByteSource source = new(input);
        List<CellValue> values = [];

        while (source.HasMore == true && values.Count < 200)
        {
            int kind = source.NextByte() % 6;

            try
            {
                values.Add(kind switch
                {
                    1 => CellValue.Text(ReadText(source)),
                    2 => CellValue.Number(source.NextDouble()),
                    3 => CellValue.Bool((source.NextByte() & 1) == 1),
                    4 => CellValue.Date(ReadDate(source)),
                    5 => CellValue.Error(ErrorCodes[source.NextByte() % ErrorCodes.Length]),
                    _ => Skip(source),
                });
            }
            catch (ArgumentException)
            {
                values.Add(CellValue.Empty);
            }
        }

        return values;
    }

    private static CellValue Skip(ByteSource source)
    {
        source.NextByte();

        return CellValue.Empty;
    }

    private static string ReadText(ByteSource source)
    {
        int length = source.NextByte();
        StringBuilder builder = new(length);

        for (int index = 0; index < length; index++)
        {
            builder.Append(source.NextChar());
        }

        return builder.ToString();
    }

    /// <summary>A date rounded to whole milliseconds, which is what Excel's serial numbers hold.</summary>
    private static DateTime ReadDate(ByteSource source)
    {
        long ticks = Math.Abs(source.NextLong() % (LatestDate.Ticks + 1));

        return new DateTime(ticks - (ticks % TimeSpan.TicksPerMillisecond), DateTimeKind.Unspecified);
    }

    private static async Task RunAsync(byte[] input)
    {
        List<CellValue> values = Decode(input);

        if (values.Count == 0)
        {
            return;
        }

        await StreamingAsync(values);
        await DocumentAsync(values);
    }

    private static async Task StreamingAsync(List<CellValue> values)
    {
        await using MemoryStream file = new();

        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(file, true))
        {
            Result<SheetWriter> added = await writer.AddSheetAsync("Values");
            added.IsSuccess(out SheetWriter? sheet);

            foreach (CellValue[] row in values.Chunk(Columns))
            {
                await sheet!.WriteRowAsync([.. row.Select(Cell.Of)]);
            }
        }

        file.Position = 0;
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(file, true);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            throw new InvariantViolationException($"The writer's file cannot be opened: {opened.Error}");
        }

        Dictionary<int, Row> rows = [];

        await using (reader)
        {
            reader.Sheet(0).IsSuccess(out SheetReader? sheet);

            await foreach (Row row in sheet!.ReadRowsAsync())
            {
                rows[row.Index] = row;
            }
        }

        for (int index = 0; index < values.Count; index++)
        {
            CellValue read = CellValue.Empty;

            if (rows.TryGetValue(index / Columns, out Row? row) == true)
            {
                read = row[index % Columns].Value;
            }

            Compare("writer -> reader", index, values[index], read);
        }
    }

    private static async Task DocumentAsync(List<CellValue> values)
    {
        Workbook book = Workbook.Create();
        book.AddSheet("Values").IsSuccess(out Worksheet? sheet);

        for (int index = 0; index < values.Count; index++)
        {
            sheet!.SetValue(index / Columns, index % Columns, values[index]);
        }

        await using MemoryStream file = new();
        await book.SaveAsync(file);
        file.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(file);

        if (opened.IsSuccess(out Workbook? again) == false)
        {
            throw new InvariantViolationException($"A saved workbook cannot be opened: {opened.Error}");
        }

        for (int index = 0; index < values.Count; index++)
        {
            Compare("document -> document", index, values[index], again.Sheets[0][index / Columns, index % Columns].Value);
        }
    }

    private static void Compare(string path, int index, CellValue written, CellValue read)
    {
        if (written == read)
        {
            return;
        }

        string address = new CellAddress(index / Columns, index % Columns).ToString();

        throw new InvariantViolationException(string.Create(
            CultureInfo.InvariantCulture,
            $"{path}: {address} was written as {Describe(written)} and read back as {Describe(read)}."));
    }

    private static string Describe(CellValue value)
    {
        string text = value.ToString();

        if (value.TryGetNumber(out double number) == true)
        {
            text = number.ToString("R", CultureInfo.InvariantCulture);
        }
        else if (value.TryGetDate(out DateTime date) == true)
        {
            text = date.ToString("O", CultureInfo.InvariantCulture);
        }

        StringBuilder escaped = new();

        foreach (char current in text)
        {
            if (current < ' ' || current > '~')
            {
                escaped.Append(CultureInfo.InvariantCulture, $"\\u{(int)current:X4}");
            }
            else
            {
                escaped.Append(current);
            }
        }

        return $"{value.Kind} \"{escaped}\"";
    }
}
