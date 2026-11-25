using System.Globalization;
using System.Text;
using Xunit.Abstractions;

namespace Avig.Kalkyl.Compatibility;

/// <summary>
/// Opens real workbooks from Excel tutorials with Avig.Kalkyl and with ExcelDataReader,
/// ClosedXML, Sylvan.Data.Excel and MiniExcel, and checks that Avig.Kalkyl reads what the other
/// libraries agree on. The workbooks are downloaded with download.ps1 (see files.json); without
/// them the tests are skipped.
/// </summary>
public class CompatibilityTests
{
    private const string NoFiles = "(no files: run download.ps1)";

    /// <summary>Functions whose value changes every time they are calculated, so the value Excel saved cannot be checked.</summary>
    private static readonly string[] Volatile = ["NOW(", "TODAY(", "RAND(", "RANDBETWEEN(", "RANDARRAY("];

    private readonly ITestOutputHelper output;

    public CompatibilityTests(ITestOutputHelper output)
    {
        this.output = output;
    }

    public static TheoryData<string> Files()
    {
        TheoryData<string> data = [];
        string directory = Path.Combine(ProjectDirectory(), "files");

        if (Directory.Exists(directory) == true)
        {
            foreach (string path in Directory.GetFiles(directory, "*.xlsx").Order(StringComparer.Ordinal))
            {
                data.Add(Path.GetFileName(path));
            }
        }

        if (data.Count == 0)
        {
            data.Add(NoFiles);
        }

        return data;
    }

    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task Reading_agrees_with_the_other_libraries(string file)
    {
        string path = Find(file);
        WorkbookValues kalkyl = await Readers.KalkylStreamingAsync(path);
        Assert.True(kalkyl.Failure.Length == 0, kalkyl.Failure);
        WorkbookValues[] references = [Readers.WithExcelDataReader(path), Readers.WithClosedXml(path), Readers.WithSylvan(path), Readers.WithMiniExcel(path)];

        ComparisonResult result = Comparison.WithConsensus(kalkyl, references);

        foreach (WorkbookValues values in references.Prepend(kalkyl))
        {
            output.WriteLine($"{values.Library}: {values.Sheets.Count} sheets, {values.Sheets.Values.Sum(cells => cells.Count)} cells {values.Failure}");
        }

        output.WriteLine(Comparison.Summary("Avig.Kalkyl against the other libraries", result, 0));
        Report(file, "reading", result);
        Assert.True(result.Differences.Count == 0, Comparison.Summary("Avig.Kalkyl against the other libraries", result));
    }

    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task The_document_model_reads_what_the_streaming_reader_reads(string file)
    {
        string path = Find(file);
        WorkbookValues streaming = await Readers.KalkylStreamingAsync(path);
        Workbook book = await OpenAsync(path);

        ComparisonResult result = Comparison.Exactly(Readers.KalkylDocument(book), streaming);

        Report(file, "document", result);
        Assert.True(result.Differences.Count == 0, Comparison.Summary("Workbook against XlsxReader", result));
    }

    /// <summary>
    /// Every workbook, saved by Excel or another program, passes the strict zip check
    /// (KalkylLimits.StrictZip, on by default) in both readers. RemoveActiveContent on it gives a
    /// file without active content that opens again; what each file had goes to the test output.
    /// </summary>
    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task Every_file_passes_the_strict_zip_check_and_survives_RemoveActiveContent(string file)
    {
        string path = Find(file);
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(path);
        Assert.True(opened.IsSuccess(out XlsxReader? reader), opened.Error);

        await using (reader)
        {
            output.WriteLine($"Active content: {await reader.GetActiveContentAsync()}");
        }

        Workbook book = await OpenAsync(path);
        Result removed = book.RemoveActiveContent();

        if (removed.IsFailure(out string? error, out _) == true)
        {
            output.WriteLine($"Not removed: {error}");

            return;
        }

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;
        Result<Workbook> clean = await Workbook.OpenAsync(saved);
        Assert.True(clean.IsSuccess(out Workbook? cleaned), clean.Error);
        Assert.Equal(ActiveContent.None, cleaned.ActiveContent);
        Assert.Equal(book.Sheets.Count, cleaned.Sheets.Count);
    }

    /// <summary>
    /// Opens the workbook, changes nothing, saves it, and reads both files with each other
    /// library. Saving must not change what any of them sees: the document model copies what it
    /// does not own and rewrites only the cells.
    /// </summary>
    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task A_saved_copy_reads_the_same_in_the_other_libraries(string file)
    {
        string path = Find(file);
        Workbook book = await OpenAsync(path);
        string copy = ReportPath(file, "saved.xlsx");
        Result saved = await book.SaveAsync(copy);
        Assert.True(saved.IsSuccess(out string? error, out _), error);
        StringBuilder failures = new();

        foreach (Func<string, WorkbookValues> reader in ReferenceReaders())
        {
            WorkbookValues original = reader(path);

            if (original.Failure.Length > 0)
            {
                continue;
            }

            ComparisonResult result = Comparison.Exactly(reader(copy), original);
            Report(file, "saved-" + original.Library, result);

            if (result.Differences.Count > 0)
            {
                failures.Append(Comparison.Summary($"{original.Library}: saved copy against the original", result, 10));
            }
        }

        // The copy stays in report/ when the test fails, to open and look at.
        Assert.True(failures.Length == 0, failures.ToString());
        File.Delete(copy);
    }

    /// <summary>Writes what Avig.Kalkyl read into a new file with XlsxWriter, and checks that every other library reads the same values from it.</summary>
    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task A_written_copy_reads_the_same_in_the_other_libraries(string file)
    {
        string path = Find(file);
        WorkbookValues kalkyl = await Readers.KalkylStreamingAsync(path);
        string copy = ReportPath(file, "written.xlsx");
        await WriteAsync(copy, kalkyl);
        StringBuilder failures = new();

        foreach (Func<string, WorkbookValues> reader in ReferenceReaders())
        {
            WorkbookValues written = reader(copy);
            ComparisonResult result = Comparison.Exactly(written, kalkyl);
            Report(file, "written-" + written.Library, result);

            if (result.Differences.Count > 0)
            {
                failures.Append(Comparison.Summary($"{written.Library}: file written by XlsxWriter", result, 10));
            }
        }

        // The copy stays in report/ when the test fails, to open and look at.
        Assert.True(failures.Length == 0, failures.ToString());
        File.Delete(copy);
    }

    /// <summary>
    /// Recalculates every formula and compares the result with the value Excel saved in the
    /// file. This is the formula engine checked against Excel itself.
    /// </summary>
    [SkippableTheory]
    [MemberData(nameof(Files))]
    public async Task Recalculation_gives_the_values_Excel_saved(string file)
    {
        string path = Find(file);
        Workbook book = await OpenAsync(path);

        // Text dates in the formulas ("<=31/12/2018") mean what they meant to the author.
        book.Culture = CultureOf(file);
        List<(Worksheet Sheet, int Row, int Column, string Formula, CellValue Saved)> formulas = [];

        foreach (Worksheet sheet in book.Sheets)
        {
            foreach (Row row in sheet.Rows)
            {
                for (int column = 0; column < row.ColumnCount; column++)
                {
                    Cell cell = row[column];

                    if (cell.Formula is not null && cell.Value.IsEmpty == false)
                    {
                        formulas.Add((sheet, row.Index, column, cell.Formula, cell.Value));
                    }
                }
            }
        }

        Skip.If(formulas.Count == 0, "The workbook has no formulas with a saved value.");

        // A volatile cell (RANDBETWEEN, NOW) gets a new value on every calculation, and so does
        // everything that depends on it. Fixing it at the value Excel saved lets the cells that
        // depend on it be checked against the same input Excel used.
        foreach ((Worksheet sheet, int row, int column, string formula, CellValue saved) in formulas.Where(cell => IsVolatile(cell.Formula) == true))
        {
            sheet.SetValue(row, column, saved);
        }

        RecalculationReport report = await book.RecalculateAsync(RecalculationScope.All);
        ComparisonResult result = new();

        foreach ((Worksheet sheet, int row, int column, string formula, CellValue saved) in formulas)
        {
            if (IsVolatile(formula) == true)
            {
                continue;
            }

            // Structured references (Table1[Column]) are not supported; the formula gives #REF!.
            // They are counted apart so that the test shows what the engine does wrong elsewhere.
            if (formula.Contains('[', StringComparison.Ordinal) == true)
            {
                result.Unsupported++;

                continue;
            }

            result.Compared++;
            CellValue calculated = sheet[row, column].Value;

            if (SameResult(calculated, saved) == false)
            {
                result.Differences.Add(new Difference(sheet.Name, row, column, Describe(calculated), Describe(saved), $" in ={formula}"));
            }
        }

        if (report.UnknownFunctions.Count > 0)
        {
            result.Notes.Add("Unknown functions: " + string.Join(", ", report.UnknownFunctions));
        }

        output.WriteLine($"{formulas.Count} formulas, {result.Compared} compared, {result.Differences.Count} differ.");
        Report(file, "recalculation", result);
        Assert.True(result.Differences.Count == 0, Comparison.Summary("Recalculated against Excel's saved values", result));
    }

    /// <summary>
    /// A recalculated value against the one Excel saved. Text is compared without whitespace at
    /// the ends: Excel trims a formula's text result when it saves it without xml:space, so
    /// VLOOKUP of "Sheetal Besoya " is saved as "Sheetal Besoya", and " " as nothing.
    /// </summary>
    private static bool SameResult(CellValue calculated, CellValue saved)
    {
        if (calculated.TryGetText(out string? text) == true && saved.TryGetText(out string? savedText) == true)
        {
            return string.Equals(text.Trim(), savedText.Trim(), StringComparison.Ordinal);
        }

        WorkbookValues pair = new("pair");
        Readers.AddKalkyl(pair, "saved", 0, 0, saved);
        Readers.AddKalkyl(pair, "calculated", 0, 0, calculated);
        bool hasSaved = pair.TryGet("saved", 0, 0, out Value expected);
        bool hasCalculated = pair.TryGet("calculated", 0, 0, out Value actual);

        if (hasSaved == false || hasCalculated == false)
        {
            return hasSaved == hasCalculated;
        }

        return actual.Matches(expected);
    }

    private static string Describe(CellValue value)
    {
        WorkbookValues single = new("single");
        Readers.AddKalkyl(single, "value", 0, 0, value);

        if (single.TryGet("value", 0, 0, out Value converted) == false)
        {
            return "(empty)";
        }

        return $"{converted.Kind} {converted}";
    }

    /// <summary>The culture files.json gives for the workbook, or en-US.</summary>
    private static CultureInfo CultureOf(string file)
    {
        using System.Text.Json.JsonDocument manifest = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(ProjectDirectory(), "files.json")));

        foreach (System.Text.Json.JsonElement entry in manifest.RootElement.GetProperty("files").EnumerateArray())
        {
            if (entry.GetProperty("name").GetString() == file && entry.TryGetProperty("culture", out System.Text.Json.JsonElement culture) == true)
            {
                return CultureInfo.GetCultureInfo(culture.GetString() ?? "en-US");
            }
        }

        return CultureInfo.GetCultureInfo("en-US");
    }

    private static bool IsVolatile(string formula)
    {
        return Volatile.Any(name => formula.Contains(name, StringComparison.OrdinalIgnoreCase)) == true;
    }

    private static IEnumerable<Func<string, WorkbookValues>> ReferenceReaders()
    {
        yield return Readers.WithExcelDataReader;
        yield return Readers.WithClosedXml;
        yield return Readers.WithSylvan;
        yield return Readers.WithMiniExcel;
    }

    private static async Task WriteAsync(string path, WorkbookValues values)
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);
        Assert.True(created.IsSuccess(out XlsxWriter? writer), created.Error);

        await using (writer)
        {
            foreach ((string name, Dictionary<(int Row, int Column), Value> cells) in values.Sheets)
            {
                Result<SheetWriter> added = await writer.AddSheetAsync(name);
                Assert.True(added.IsSuccess(out SheetWriter? sheet), added.Error);

                foreach (IGrouping<int, KeyValuePair<(int Row, int Column), Value>> row in cells.GroupBy(cell => cell.Key.Row).OrderBy(group => group.Key))
                {
                    int width = row.Max(cell => cell.Key.Column) + 1;
                    Cell[] written = new Cell[width];

                    foreach (KeyValuePair<(int Row, int Column), Value> cell in row)
                    {
                        written[cell.Key.Column] = Cell.Of(ToCellValue(cell.Value));
                    }

                    while (sheet.RowCount < row.Key)
                    {
                        await sheet.WriteRowAsync([]);
                    }

                    await sheet.WriteRowAsync(written);
                }
            }
        }
    }

    private static CellValue ToCellValue(Value value)
    {
        return value.Kind switch
        {
            ValueKind.Number => CellValue.Number(value.Number),
            ValueKind.Bool => CellValue.Bool(value.Number == 1),
            ValueKind.Date => CellValue.Date(value.Date),
            ValueKind.Error => CellValue.Error(value.Text),
            _ => CellValue.Text(value.Text),
        };
    }

    private static async Task<Workbook> OpenAsync(string path)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(path);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        return book;
    }

    private static string Find(string file)
    {
        Skip.If(file == NoFiles, "No workbooks downloaded. Run tests/Avig.Kalkyl.Compatibility/download.ps1 first.");

        return Path.Combine(ProjectDirectory(), "files", file);
    }

    /// <summary>Writes every difference to report/&lt;file&gt;.&lt;test&gt;.md, since an assertion message only has room for the first ones.</summary>
    private static void Report(string file, string test, ComparisonResult result)
    {
        string directory = Path.Combine(ProjectDirectory(), "report");
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(file)}.{test}.md");

        if (result.Differences.Count == 0 && result.Notes.Count == 0)
        {
            File.Delete(path);

            return;
        }

        File.WriteAllText(path, Comparison.Summary($"{file}, {test}", result, int.MaxValue), Encoding.UTF8);
    }

    private static string ReportPath(string file, string suffix)
    {
        string directory = Path.Combine(ProjectDirectory(), "report");
        Directory.CreateDirectory(directory);

        return Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(file)}.{suffix}");
    }

    private static string ProjectDirectory()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Avig.Kalkyl.Compatibility.csproj")) == true)
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        return Environment.CurrentDirectory;
    }
}
