using System.IO.Compression;
using System.Text;

namespace Avig.Kalkyl.ExcelCheck;

/// <summary>
/// Writes the files to open in real Excel (FUN-04 in docs/testplan.md): what the writer makes,
/// what the document model saves, the files the package repairs or changes, and every
/// compatibility workbook after a round trip. docs/excel-checklista.md says what to look for in
/// each. Excel is not on the build machine, so this is the one test a person does.
/// </summary>
/// <example>
/// dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check
/// dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check --xlsm my-macros.xlsm
/// </example>
internal static class Program
{
    private const string SpreadsheetMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static readonly List<string> Expected = [];

    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.Error.WriteLine("Usage: Avig.Kalkyl.ExcelCheck <output folder> [--xlsm <file with macros>]");

            return 2;
        }

        string output = Path.GetFullPath(args[0]);
        Directory.CreateDirectory(output);
        int xlsm = Array.IndexOf(args, "--xlsm");

        await WriteWithWriterAsync(Path.Combine(output, "01-writer.xlsx"));
        await WriteDocumentAsync(Path.Combine(output, "02-document.xlsx"));
        await EditAsync(Path.Combine(output, "01-writer.xlsx"), Path.Combine(output, "03-edited.xlsx"));
        await WriteWithoutContentTypesAsync(Path.Combine(output, "04-no-content-types.xlsx"));
        await WriteFormulaLikeTextAsync(Path.Combine(output, "05-formula-like-text.xlsx"));
        await WriteEdgeValuesAsync(Path.Combine(output, "06-edge-values.xlsx"));

        if (xlsm >= 0 && xlsm + 1 < args.Length)
        {
            await WriteMacroFilesAsync(Path.GetFullPath(args[xlsm + 1]), output);
        }
        else
        {
            Expected.Add("07, 08: skipped, no --xlsm file given.");
        }

        await RoundTripCompatibilityFilesAsync(output);
        await File.WriteAllLinesAsync(Path.Combine(output, "expected.txt"), Expected);
        Console.WriteLine($"Files in {output}. Open each in Excel and follow docs/excel-checklista.md; expected.txt has the values to compare.");

        return 0;
    }

    private static async Task WriteWithWriterAsync(string path)
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);
        Check(created.IsSuccess(out XlsxWriter? writer), created.Error);

        await using (writer)
        {
            Result<SheetWriter> added = await writer.AddSheetAsync("Rapport");
            Check(added.IsSuccess(out SheetWriter? sheet), added.Error);
            sheet.FreezePanes(1);
            sheet.EnableAutoFilter();
            sheet.SetColumnWidth(0, 24);
            sheet.SetColumnWidth(2, 14);

            await sheet.WriteRowAsync([Cell.Of("Namn").Bold().Background("#FFF2CC"), Cell.Of("Belopp").Bold(), Cell.Of("Datum").Bold(), Cell.Of("Klar").Bold()]);
            await sheet.WriteRowAsync(["Åsa Öberg", Cell.Of(1234.5).NumberFormat("#,##0.00"), new DateTime(2026, 1, 2), true]);
            await sheet.WriteRowAsync(["Ada", Cell.Of(-0.25).NumberFormat("0.0%"), new DateTime(2026, 2, 28, 13, 45, 0), false]);
            await sheet.WriteRowAsync(["Röd text", Cell.Of(42).FontColor("#C00000"), new DateOnly(1900, 3, 1), Cell.Of(CellValue.Error("#N/A"))]);
            await sheet.WriteRowAsync([Cell.Of("Summa").Bold(), Cell.FromFormula("SUM(B2:B4)", CellValue.Number(1276.25)), Cell.FromFormula("MAX(C2:C4)"), Cell.FromFormula("COUNTIF(D2:D4,TRUE)")]);

            Result<SheetWriter> second = await writer.AddSheetAsync("Blad 2 & 'citat' x");
            Check(second.IsSuccess(out SheetWriter? other), second.Error);
            await other.WriteRowAsync([Cell.FromFormula("Rapport!B5*2"), Cell.FromFormula("'Blad 2 & ''citat'' x'!A1+1")]);
        }

        Expected.Add("01-writer.xlsx: Rapport!B5 = 1 276,25 (or 1,276.25), C5 = 46081.57 as a plain number (MAX of dates in a cell without a date format; Excel shows the serial number too), D5 = 1. 'Blad 2 & 'citat' x'!A1 = 2552.5, B1 = 2553.5. Row 1 frozen, filter buttons on row 1, A1 yellow and bold, B4 red.");
    }

    private static async Task WriteDocumentAsync(string path)
    {
        Workbook book = Workbook.Create();
        Check(book.AddSheet("Data").IsSuccess(out Worksheet? data), "Data");
        Check(book.AddSheet("Beräkning").IsSuccess(out Worksheet? calc), "Beräkning");
        data[0, 0] = Cell.Of("Region").Bold();
        data[0, 1] = Cell.Of("Kategori").Bold();
        data[0, 2] = Cell.Of("Belopp").Bold();
        string[] regions = ["Norr", "Syd", "Väst"];

        for (int row = 1; row <= 300; row++)
        {
            data[row, 0] = regions[row % 3];
            data[row, 1] = $"K{row % 7}";
            data[row, 2] = row * 10.5;
        }

        data.FreezePanes(1);
        data.AutoFilter = true;
        calc["A1"] = "Summa Norr";
        calc["B1"] = Cell.FromFormula("SUMIF(Data!A:A,\"Norr\",Data!C:C)");
        calc["A2"] = "Antal K3";
        calc["B2"] = Cell.FromFormula("COUNTIF(Data!B:B,\"K3\")");
        calc["A3"] = "Medel Syd";
        calc["B3"] = Cell.FromFormula("AVERAGEIFS(Data!C:C,Data!A:A,\"Syd\")");
        calc["A4"] = "Uppslag rad 10";
        calc["B4"] = Cell.FromFormula("VLOOKUP(\"Norr\",Data!A2:C301,3,FALSE)");
        calc["A5"] = "Summa allt";
        calc["B5"] = Cell.FromFormula("SUM(Data!C:C)");
        calc["A6"] = "Text";
        calc["B6"] = Cell.FromFormula("TEXT(B5,\"0.00\")&\" kr\"");
        calc["A7"] = "Fel";
        calc["B7"] = Cell.FromFormula("IFERROR(1/0,\"delat med noll\")");
        calc["A8"] = "Datum";
        calc["B8"] = Cell.FromFormula("DATE(2026,2,29)").NumberFormat("yyyy-mm-dd");
        calc.SetColumnWidth(0, 18);
        await book.RecalculateAsync(RecalculationScope.All);
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);

        StringBuilder values = new("02-document.xlsx: the values below are calculated by the package and cached in the file; Ctrl+Alt+F9 in Excel must not change any of them.");

        for (int row = 0; row < 8; row++)
        {
            values.Append(System.Globalization.CultureInfo.InvariantCulture, $" Beräkning!B{row + 1} = {calc[row, 1].Value};");
        }

        Expected.Add(values.ToString());
    }

    private static async Task EditAsync(string source, string path)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(source);
        Check(opened.IsSuccess(out Workbook? book), opened.Error);
        Check(book.Sheet("Rapport").IsSuccess(out Worksheet? sheet), "Rapport");
        sheet.SetValue(1, 1, 2000);
        sheet["E1"] = Cell.Of("Ny kolumn").Bold();
        sheet["E2"] = Cell.FromFormula("B2*2");
        sheet.SetColumnWidth(4, 16);
        await book.RecalculateAsync();
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("03-edited.xlsx: as 01, but Rapport!B2 = 2 000,00 with the same number format, B5 = 2 041,75, E2 = 4000, column E wider. 'Blad 2 & 'citat' x'!A1 = 4083.5.");
    }

    private static async Task WriteWithoutContentTypesAsync(string path)
    {
        using MemoryStream package = new();

        using (ZipArchive archive = new(package, ZipArchiveMode.Create, true))
        {
            Add(archive, "_rels/.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
            Add(archive, "xl/workbook.xml", $"""<workbook xmlns="{SpreadsheetMain}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""");
            Add(archive, "xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
            Add(archive, "xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{SpreadsheetMain}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Utan content types</t></is></c><c r="B1"><v>7</v></c></row></sheetData></worksheet>""");
        }

        package.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(package);
        Check(opened.IsSuccess(out Workbook? book), opened.Error);
        book.Sheets[0]["C1"] = Cell.FromFormula("B1*6");
        await book.RecalculateAsync();
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("04-no-content-types.xlsx: made from a package without [Content_Types].xml, which Excel cannot open, and saved by the package. Must open without repair: A1 'Utan content types', B1 = 7, C1 = 42.");
    }

    private static async Task WriteFormulaLikeTextAsync(string path)
    {
        string[] texts = ["=1+1", "=HYPERLINK(\"https://example.invalid\",\"klicka\")", "+cmd|' /C calc'!A0", "-2+3", "@SUM(1)", "=WEBSERVICE(\"https://example.invalid\")", "\t=1+1", "'=quoted"];
        Workbook book = Workbook.Create();
        Check(book.AddSheet("Text").IsSuccess(out Worksheet? sheet), "Text");

        for (int row = 0; row < texts.Length; row++)
        {
            sheet[row, 0] = texts[row];
        }

        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add($"05-formula-like-text.xlsx: column A shows exactly this text, and nothing is calculated, linked or fetched; Excel must not warn about links or external content: {string.Join(" | ", texts)}");
    }

    private static async Task WriteEdgeValuesAsync(string path)
    {
        Workbook book = Workbook.Create();
        Check(book.AddSheet("Gränser").IsSuccess(out Worksheet? sheet), "Gränser");
        sheet["A1"] = new string('x', CellValue.MaxTextLength - 3) + "END";
        sheet["A2"] = 1.7976931348623157e308;
        sheet["A3"] = 5e-324;
        sheet["A4"] = 0.1 + 0.2;
        sheet["A5"] = Cell.Of(new DateTime(9999, 12, 31, 23, 59, 59)).NumberFormat("yyyy-mm-dd hh:mm:ss");
        sheet["A6"] = Cell.Of(new DateTime(1900, 1, 1)).NumberFormat("yyyy-mm-dd");
        sheet["A7"] = Cell.Of(TimeSpan.FromHours(36).TotalDays).NumberFormat("[h]:mm:ss");
        sheet["A8"] = "Rad\nbrytning och \u00e5\u00e4\u00f6 \u2603 \U0001F600";
        sheet["A9"] = "  mellanslag före och efter  ";
        sheet[CellAddress.MaxRows - 1, CellAddress.MaxColumns - 1] = "XFD1048576";
        sheet["B1"] = Cell.FromFormula("LEN(A1)");
        await book.RecalculateAsync(RecalculationScope.All);
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("06-edge-values.xlsx: A1 is 32 767 characters ending in END and B1 = 32767; A2 = 1.79769313486232E+308; A3 = 4.94065645841247E-324 (Excel may show 0); A4 = 0.3 (0.30000000000000004 inside); A5 = 9999-12-31 23:59:59; A6 = 1900-01-01; A7 = 36:00:00; A8 has a line break, åäö, a snowman and a smiley; A9 keeps its spaces; XFD1048576 holds 'XFD1048576' (Ctrl+End goes there).");
    }

    private static async Task WriteMacroFilesAsync(string source, string output)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(source);
        Check(opened.IsSuccess(out Workbook? book), opened.Error);
        Check(book.HasMacros, $"{source} has no macros.");
        book.Sheets[0]["A1"] = book.Sheets[0]["A1"];
        Check((await book.SaveAsync(Path.Combine(output, "07-macros-kept.xlsm"))).IsSuccess(out _, out _), "07");

        Result<Workbook> again = await Workbook.OpenAsync(source);
        Check(again.IsSuccess(out Workbook? stripped), again.Error);
        Result removed = stripped.RemoveMacros();

        if (removed.IsFailure(out string? error, out _) == true)
        {
            Expected.Add($"07-macros-kept.xlsm: the macros run as in the original. 08: not written, because RemoveMacros said: {error}");

            return;
        }

        Check((await stripped.SaveAsync(Path.Combine(output, "08-macros-removed.xlsx"))).IsSuccess(out _, out _), "08");
        Expected.Add($"07-macros-kept.xlsm: saved from {Path.GetFileName(source)}; Excel asks to enable macros, and they run as in the original. 08-macros-removed.xlsx: opens as an ordinary workbook without a macro warning, and Developer > Visual Basic shows no project.");
    }

    private static async Task RoundTripCompatibilityFilesAsync(string output)
    {
        if (TryFindCompatibilityFiles(out string? folder) == false)
        {
            Expected.Add("09: skipped, tests/Avig.Kalkyl.Compatibility/files is empty; run the compatibility tests once to download the workbooks.");

            return;
        }

        string target = Path.Combine(output, "09-roundtrip");
        Directory.CreateDirectory(target);
        int written = 0;

        foreach (string file in Directory.GetFiles(folder, "*.xlsx").Order(StringComparer.Ordinal))
        {
            Result<Workbook> opened = await Workbook.OpenAsync(file);

            if (opened.IsSuccess(out Workbook? book) == false || book.Sheets.Count == 0)
            {
                continue;
            }

            Worksheet first = book.Sheets[0];
            first[first.RowCount, 0] = Cell.Of("Avig.Kalkyl").Bold();
            await book.RecalculateAsync(RecalculationScope.All);

            if ((await book.SaveAsync(Path.Combine(target, Path.GetFileName(file)))).IsSuccess(out _, out _) == true)
            {
                written++;
            }
        }

        Expected.Add($"09-roundtrip: {written} tutorial workbooks, opened, recalculated, one cell added under the last row of the first sheet ('Avig.Kalkyl', bold) and saved. Each must open without repair and look like the original in tests/Avig.Kalkyl.Compatibility/files: formatting, charts, tables, pivot tables, comments, names and data validation.");
    }

    private static bool TryFindCompatibilityFiles([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? folder)
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null)
        {
            string candidate = Path.Combine(directory.FullName, "tests", "Avig.Kalkyl.Compatibility", "files");

            if (Directory.Exists(candidate) == true && Directory.GetFiles(candidate, "*.xlsx").Length > 0)
            {
                folder = candidate;

                return true;
            }

            directory = directory.Parent;
        }

        folder = null;

        return false;
    }

    private static void Add(ZipArchive archive, string path, string content)
    {
        using Stream entry = archive.CreateEntry(path).Open();
        entry.Write(Encoding.UTF8.GetBytes(content));
    }

    private static void Check([System.Diagnostics.CodeAnalysis.DoesNotReturnIf(false)] bool condition, string? message)
    {
        if (condition == false)
        {
            throw new InvalidOperationException(message);
        }
    }
}
