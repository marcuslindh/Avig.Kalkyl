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
/// read back with <see cref="XlsxReader"/> and <see cref="Workbook"/>. The same cells are then
/// the source of a pivot table, which must save, open again and count them.
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
        await PivotAsync(values);
    }

    private static async Task StreamingAsync(List<CellValue> values)
    {
        await using MemoryStream file = new();

        List<string> shapeTexts = [];

        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(file, true))
        {
            Result<SheetWriter> added = await writer.AddSheetAsync("Values");
            added.IsSuccess(out SheetWriter? sheet);

            // The texts go into text boxes too, as their text and their alternative text.
            foreach (CellValue value in values)
            {
                if (value.TryGetText(out string? text) == true && TryTextBox(text, out Shape? box) == true)
                {
                    sheet!.AddShape(box, DrawingBounds.FromPoints(0, shapeTexts.Count * 20, 100, 20));
                    shapeTexts.Add(text);
                }
            }

            foreach (CellValue[] row in values.Chunk(Columns))
            {
                await sheet!.WriteRowAsync([.. row.Select(Cell.Of)]);
            }
        }

        CompareShapeTexts(file, shapeTexts);

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

    /// <summary>
    /// A text box with the text, or false when the text holds a character XML cannot store
    /// (a control character, half a surrogate pair): refusing it with an ArgumentException is the
    /// documented answer, where a cell escapes it instead.
    /// </summary>
    private static bool TryTextBox(string text, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out Shape? box)
    {
        box = null;

        try
        {
            box = new Shape(ShapeKind.TextBox) { Text = new ShapeText(text), Description = text[..Math.Min(text.Length, 1000)] };
        }
        catch (ArgumentException)
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Each text box's text, read back from the drawing as Excel reads it (the a:t of each a:p,
    /// one line per paragraph), and its alternative text, must be the text written, with
    /// every line break as \n.
    /// </summary>
    private static void CompareShapeTexts(MemoryStream file, List<string> texts)
    {
        if (texts.Count == 0)
        {
            return;
        }

        file.Position = 0;
        using System.IO.Compression.ZipArchive archive = new(file, System.IO.Compression.ZipArchiveMode.Read, true);
        System.IO.Compression.ZipArchiveEntry entry = archive.GetEntry("xl/drawings/drawing1.xml")
            ?? throw new InvariantViolationException("The sheet has text boxes but no drawing.");
        System.Xml.Linq.XDocument drawing;

        using (Stream stream = entry.Open())
        {
            drawing = System.Xml.Linq.XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        }

        List<System.Xml.Linq.XElement> shapes = drawing.Descendants().Where(element => element.Name.LocalName == "sp").ToList();

        if (shapes.Count != texts.Count)
        {
            throw new InvariantViolationException($"{texts.Count} text boxes were written and {shapes.Count} read back.");
        }

        for (int index = 0; index < texts.Count; index++)
        {
            string expected = texts[index].Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
            IEnumerable<string> paragraphs = shapes[index].Descendants().Where(element => element.Name.LocalName == "p")
                .Select(paragraph => string.Concat(paragraph.Descendants().Where(element => element.Name.LocalName == "t").Select(element => element.Value)));
            string read = string.Join('\n', paragraphs);

            if (read != expected)
            {
                throw new InvariantViolationException($"Text box {index + 1} was written as {Describe(CellValue.Text(texts[index]))} and read back as {Describe(CellValue.Text(read))}.");
            }

            string description = (string?)shapes[index].Descendants().First(element => element.Name.LocalName == "cNvPr").Attribute("descr") ?? string.Empty;
            string expectedDescription = texts[index][..Math.Min(texts[index].Length, 1000)];

            if (description != expectedDescription)
            {
                throw new InvariantViolationException($"Text box {index + 1}'s alternative text was written as {Describe(CellValue.Text(expectedDescription))} and read back as {Describe(CellValue.Text(description))}.");
            }
        }

        file.Position = 0;
    }

    /// <summary>
    /// The values as the source of a pivot table, the first row as headings: adding it gives a
    /// result, not an exception; the saved file opens with the pivot table and a cache of every
    /// row; and GETPIVOTDATA's grand total is the number of values in the counted column.
    /// </summary>
    private static async Task PivotAsync(List<CellValue> values)
    {
        if (values.Count < Columns + 1)
        {
            return;
        }

        int rows = (values.Count + Columns - 1) / Columns;
        string rowField = values[0].ToString().Trim();
        string countField = values[1].ToString().Trim();

        if (rows < 2 || rowField.Length == 0 || countField.Length == 0 || string.Equals(rowField, countField, StringComparison.OrdinalIgnoreCase) == true)
        {
            return;
        }

        Workbook book = Workbook.Create();
        book.AddSheet("Values").IsSuccess(out Worksheet? sheet);
        book.AddSheet("Pivot").IsSuccess(out Worksheet? pivots);

        for (int index = 0; index < values.Count; index++)
        {
            sheet!.SetValue(index / Columns, index % Columns, values[index]);
        }

        PivotTable table = new("Fuzz", "Values!A1:F" + rows.ToString(CultureInfo.InvariantCulture))
        {
            Rows = [rowField],
            Values = [new PivotValue(countField, PivotAggregation.Count) { Name = "Counted" }],
        };

        Result added = pivots!.AddPivotTable(table, "A1");

        if (added.IsSuccess(out _, out _) == false)
        {
            if (added.Status is not (ResultStatus.Invalid or ResultStatus.NotFound or ResultStatus.Conflict))
            {
                throw new InvariantViolationException($"AddPivotTable gave {added.Status}: {added.Error}");
            }

            return;
        }

        // Rows of empty cells at the bottom are not in the source: the range ends at the last used row.
        int used = Math.Min(rows, sheet!.RowCount);
        int counted = 0;

        for (int row = 1; row < used; row++)
        {
            if (sheet[row, 1].Value.IsEmpty == false)
            {
                counted++;
            }
        }

        await using MemoryStream file = new();
        await book.SaveAsync(file);
        file.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(file);

        if (opened.IsSuccess(out Workbook? again) == false)
        {
            throw new InvariantViolationException($"A saved workbook with a pivot table cannot be opened: {opened.Error}");
        }

        if (again.PivotTables.Count != 1 || again.PivotCaches.Count != 1 || again.PivotCaches[0].RecordCount != used - 1)
        {
            throw new InvariantViolationException($"The saved pivot table came back as {again.PivotTables.Count} tables and {again.PivotCaches.Count} caches, for {used - 1} rows.");
        }

        again.Evaluate("GETPIVOTDATA(\"Counted\",Pivot!A1)").IsSuccess(out CellValue total);
        CellValue expected = CellValue.Number(counted);

        if (counted == 0)
        {
            expected = CellValue.Error("#REF!");
        }

        if (total != expected)
        {
            throw new InvariantViolationException($"GETPIVOTDATA gave {Describe(total)} for {counted} values in column B.");
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
