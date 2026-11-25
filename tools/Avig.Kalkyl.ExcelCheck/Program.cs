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
        await WriteChartsWithWriterAsync(Path.Combine(output, "10-charts-writer.xlsx"));
        await WriteChartsWithDocumentAsync(Path.Combine(output, "11-charts-document.xlsx"));
        await AddToCompatibilityFilesAsync(output);
        await WriteActiveContentRemovedAsync(Path.Combine(output, "13-active-content-removed.xlsx"));
        await WritePivotDashboardAsync(Path.Combine(output, "17-pivot-dashboard.xlsx"));
        await WritePivotTemplateAsync(Path.Combine(output, "17-pivot-dashboard.xlsx"), Path.Combine(output, "18-pivot-template-refreshed.xlsx"));
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

    private static readonly (string Month, double Income, double Cost)[] Months =
    [
        ("Jan", 120, 80), ("Feb", 180, 95), ("Mar", 240, 130), ("Apr", 90, 70), ("Maj", 310, 150), ("Jun", 205, 120),
    ];

    private static async Task WriteChartsWithWriterAsync(string path)
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);
        Check(created.IsSuccess(out XlsxWriter? writer), created.Error);
        Picture picture = Gradient();

        await using (writer)
        {
            Result<SheetWriter> added = await writer.AddSheetAsync("Data");
            Check(added.IsSuccess(out SheetWriter? sheet), added.Error);
            await sheet.WriteRowAsync([Cell.Of("Månad").Bold(), Cell.Of("Intäkt").Bold(), Cell.Of("Kostnad").Bold(), Cell.Of("Resultat").Bold()]);

            for (int index = 0; index < Months.Length; index++)
            {
                (string month, double income, double cost) = Months[index];
                string row = (index + 2).ToString(System.Globalization.CultureInfo.InvariantCulture);
                await sheet.WriteRowAsync([month, income, cost, Cell.FromFormula("B" + row + "-C" + row, CellValue.Number(income - cost))]);
            }

            AddRules(sheet.AddConditionalFormat);

            foreach ((Chart chart, CellRange area) in Charts())
            {
                Result chartAdded = sheet.AddChart(chart, area);
                Check(chartAdded.IsSuccess(out _, out _), "chart");
            }

            sheet.AddPicture(picture, "A10");
            sheet.AddPicture(picture, "A16:C22");

            Result<SheetWriter> report = await writer.AddSheetAsync("Rapport");
            Check(report.IsSuccess(out SheetWriter? reportSheet), report.Error);
            await reportSheet.WriteRowAsync(["Diagrammet nedan läser bladet Data."]);
            Result crossAdded = reportSheet.AddChart(new Chart(ChartType.Column) { Title = "Resultat", Series = [new ChartSeries("Resultat", "Data!D2:D7", "Data!A2:A7", "#70AD47")] }, "A3:H18");
            Check(crossAdded.IsSuccess(out _, out _), crossAdded.ToString());
        }

        Expected.Add("10-charts-writer.xlsx: sheet Data has seven charts in column F onwards, top to bottom: clustered column (title 'Intäkt och kostnad', axis titles 'Månad' and 'kr', blue Intäkt), stacked bar (legend below), line (orange, no legend), stacked area (legend on top), pie (title 'Andel', six coloured slices, legend left), doughnut, scatter (green points, Intäkt along x, Kostnad up). A 64×48 gradient picture at A10 at its own size and one stretched over A16:C22. Rules: B2:B7 green bold at 200 or more and a blue data bar; C2:C7 a red–yellow–green colour scale; D2:D7 arrows (3 arrows); rows where D is below 50 have a yellow fill (D2 and D5). Sheet Rapport: a green column chart of Data!D2:D7. Change B2 to 500: the charts follow and B2 turns green.");
    }

    private static async Task WriteChartsWithDocumentAsync(string path)
    {
        Workbook book = Workbook.Create();
        Result<Worksheet> added = book.AddSheet("Data");
        Check(added.IsSuccess(out Worksheet? sheet), added.Error);
        sheet["A1"] = Cell.Of("Månad").Bold();
        sheet["B1"] = Cell.Of("Intäkt").Bold();
        sheet["C1"] = Cell.Of("Kostnad").Bold();
        sheet["D1"] = Cell.Of("Resultat").Bold();

        for (int index = 0; index < Months.Length; index++)
        {
            (string month, double income, double cost) = Months[index];
            string row = (index + 2).ToString(System.Globalization.CultureInfo.InvariantCulture);
            sheet[index + 1, 0] = month;
            sheet[index + 1, 1] = income;
            sheet[index + 1, 2] = cost;
            sheet.SetFormula(index + 1, 3, "B" + row + "-C" + row);
        }

        AddRules(sheet.AddConditionalFormat);

        foreach ((Chart chart, CellRange area) in Charts())
        {
            Check(sheet.AddChart(chart, area).IsSuccess(out _, out _), "chart");
        }

        Picture picture = Gradient();
        Check(sheet.AddPicture(picture, "A10").IsSuccess(out _, out _), "picture");
        Check(sheet.AddPicture(picture, "A16:C22").IsSuccess(out _, out _), "picture");
        await book.RecalculateAsync(RecalculationScope.All);
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), "11");
        Expected.Add("11-charts-document.xlsx: the same sheet Data as 10, written by the document model (no Rapport sheet).");
    }

    /// <summary>
    /// A workbook with every kind of active content (a VBA project, an Excel 4.0 macro sheet, a
    /// dialog sheet, an ActiveX control next to a form button, an OLE object, a workbook link, a
    /// DDE link, a linked picture, a query table on a connection and a Power Query query), saved
    /// after RemoveActiveContent. Only the cleaned file is written.
    /// </summary>
    private static async Task WriteActiveContentRemovedAsync(string path)
    {
        Avig.Kalkyl.Tests.ActiveWorkbook source = new Avig.Kalkyl.Tests.ActiveWorkbook()
            .AddVbaProject().AddMacroSheet().AddDialogSheet().AddActiveXControl().AddFormControl().AddOleObject()
            .AddExternalWorkbookLink().AddDdeLink().AddLinkedPicture().AddHyperlink().AddDataMashup().AddQueryTable();
        Result<Workbook> opened = await Workbook.OpenAsync(source.Build());
        Check(opened.IsSuccess(out Workbook? book), opened.Error);
        Result removed = book.RemoveActiveContent();
        Check(removed.IsSuccess(out _, out _), removed.Error);
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("13-active-content-removed.xlsx: one sheet, Data. Must open without repair and without a security warning, a links prompt or a refresh prompt. A1:A3 'N', 1, 2 as a table with a filter button; B1 = 84 and C1 = #N/A as values, not formulas; the button 'Button 2' and a picture frame without a picture (the link was removed); a hyperlink on A1 to https://example.com/. Formulas > Name Manager: Rate = #REF!, no Auto_Open. Data > Queries & Connections: empty.");
    }

    /// <summary>
    /// A dashboard like the ones Excel users build by hand: a data sheet, two pivot tables on a
    /// sheet of their own, slicers for Year and Month connected to both, and charts on cells
    /// that read the pivot tables with GETPIVOTDATA.
    /// </summary>
    private static async Task WritePivotDashboardAsync(string path)
    {
        Workbook book = Workbook.Create();
        Check(book.AddSheet("Dashboard").IsSuccess(out Worksheet? dashboard), "Dashboard");
        Check(book.AddSheet("Control").IsSuccess(out Worksheet? control), "Control");
        Check(book.AddSheet("Data").IsSuccess(out Worksheet? data), "Data");
        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        string[] regions = ["North", "South", "East", "West"];
        string[] headings = ["Year", "Month", "Region", "Visitors", "Sales"];

        for (int column = 0; column < headings.Length; column++)
        {
            data[0, column] = Cell.Of(headings[column]).Bold();
        }

        int row = 1;

        foreach (int year in new[] { 2025, 2026 })
        {
            for (int month = 0; month < months.Length; month++)
            {
                for (int region = 0; region < regions.Length; region++)
                {
                    data[row, 0] = year;
                    data[row, 1] = months[month];
                    data[row, 2] = regions[region];
                    data[row, 3] = 1000 + ((year - 2025) * 500) + (month * 37) + (region * 11);
                    data[row, 4] = Cell.Of(20000 + ((year - 2025) * 4000) + (month * 750) + (region * 1250) + ((row % 5) * 100.5)).NumberFormat("#,##0.00");
                    row++;
                }
            }
        }

        data.FreezePanes(1);
        PivotTable byMonth = new("ByMonth", "Data!A1:E97")
        {
            Rows = ["Month"],
            Values = [new PivotValue("Sales") { NumberFormat = "#,##0" }, new PivotValue("Visitors")],
        };

        PivotTable byRegion = new("ByRegion", "Data!A1:E97")
        {
            Rows = ["Region"],
            Columns = ["Year"],
            Values = [new PivotValue("Sales", PivotAggregation.Average) { Name = "Average sale", NumberFormat = "#,##0.00" }],
            Layout = PivotLayout.Tabular,
            GrandTotalColumn = false,
            Style = "PivotStyleMedium9",
        };

        Check(control.AddPivotTable(byMonth, "A3").IsSuccess(out _, out _), "ByMonth");
        Check(control.AddPivotTable(byRegion, "F3").IsSuccess(out _, out _), "ByRegion");
        Result year2026 = dashboard.AddSlicer(new Slicer("Year", ["ByMonth", "ByRegion"]) { Columns = 2, Selected = ["2026"], Style = "SlicerStyleLight2" }, "A1:D4");
        Check(year2026.IsSuccess(out _, out _), year2026.Error);
        Result allMonths = dashboard.AddSlicer(new Slicer("Month", ["ByMonth", "ByRegion"]) { Columns = 6, ShowCaption = false, ItemHeight = 16 }, "E1:M4");
        Check(allMonths.IsSuccess(out _, out _), allMonths.Error);

        dashboard["A6"] = Cell.Of("Month").Bold();
        dashboard["B6"] = Cell.Of("Sales").Bold();
        dashboard["C6"] = Cell.Of("Visitors").Bold();

        for (int index = 0; index < months.Length; index++)
        {
            string label = "A" + (7 + index).ToString(System.Globalization.CultureInfo.InvariantCulture);
            dashboard[6 + index, 0] = months[index];
            dashboard[6 + index, 1] = Cell.FromFormula("IFERROR(GETPIVOTDATA(\"Sum of Sales\",Control!$A$3,\"Month\"," + label + "),0)").NumberFormat("#,##0");
            dashboard[6 + index, 2] = Cell.FromFormula("IFERROR(GETPIVOTDATA(\"Visitors\",Control!$A$3,\"Month\"," + label + "),0)");
        }

        dashboard["A20"] = Cell.Of("Region").Bold();
        dashboard["B20"] = Cell.Of("Average sale 2026").Bold();

        for (int index = 0; index < regions.Length; index++)
        {
            string label = "A" + (21 + index).ToString(System.Globalization.CultureInfo.InvariantCulture);
            dashboard[20 + index, 0] = regions[index];
            dashboard[20 + index, 1] = Cell.FromFormula("IFERROR(GETPIVOTDATA(\"Average sale\",Control!$F$3,\"Region\"," + label + ",\"Year\",2026),NA())").NumberFormat("#,##0");
        }

        dashboard["E20"] = Cell.Of("Total sales").Bold();
        dashboard["F20"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",Control!$A$3)").NumberFormat("#,##0");
        Check(dashboard.AddChart(new Chart(ChartType.Column) { Title = "Sales per month", Series = [new ChartSeries("Sales", "B7:B18", "A7:A18")], Legend = LegendPosition.None }, "E6:M18").IsSuccess(out _, out _), "chart");
        Check(dashboard.AddChart(new Chart(ChartType.Bar) { Title = "Average sale per region", Series = [new ChartSeries("2026", "B21:B24", "A21:A24")], Legend = LegendPosition.None }, "E22:M34").IsSuccess(out _, out _), "chart");
        dashboard.SetColumnWidth(1, 14);
        await book.RecalculateAsync(RecalculationScope.All);
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("17-pivot-dashboard.xlsx: three sheets. Control has the pivot tables ByMonth (A3, Month down, Sum of Sales and Sum of Visitors) and ByRegion (F3, tabular, Region down and Year across, Average sale, no grand total column, style Medium 9). Dashboard has a Year slicer (2026 selected, two columns) and a Month slicer (six columns, no caption) connected to both; clicking Jan in the Month slicer filters both pivot tables and the charts. Must open without repair and show the same numbers before and after Data > Refresh All. Dashboard!F20 = "
            + dashboard["F20"].Value.ToString() + " (sales 2026), B7 = " + dashboard["B7"].Value.ToString() + ", B21 = " + dashboard["B21"].Value.ToString() + ".");
    }

    /// <summary>
    /// File 17 opened again as a template: new sales figures in Data and RefreshPivotTablesOnOpen,
    /// so the pivot tables, which the package now only keeps, show the new figures once Excel has
    /// refreshed them.
    /// </summary>
    private static async Task WritePivotTemplateAsync(string template, string path)
    {
        Result<Workbook> opened = await Workbook.OpenAsync(template);
        Check(opened.IsSuccess(out Workbook? book), opened.Error);
        Check(book.Sheet("Data").IsSuccess(out Worksheet? data), "Data");

        for (int row = 1; row < data.RowCount; row++)
        {
            data.SetValue(row, 4, 100);
        }

        book.RefreshPivotTablesOnOpen = true;
        Check((await book.SaveAsync(path)).IsSuccess(out _, out _), path);
        Expected.Add("18-pivot-template-refreshed.xlsx: file 17 with every sale in Data set to 100, saved with RefreshPivotTablesOnOpen. When Excel has opened it, ByMonth shows Sum of Sales 400 for each month (four regions, 2026 selected) and ByRegion 100 everywhere; Dashboard!F20 = 4 800 once Excel has recalculated.");
    }

    private static async Task AddToCompatibilityFilesAsync(string output)
    {
        if (TryFindCompatibilityFiles(out string? folder) == false)
        {
            Expected.Add("12: skipped, tests/Avig.Kalkyl.Compatibility/files is empty.");

            return;
        }

        string target = Path.Combine(output, "12-added");
        Directory.CreateDirectory(target);
        Picture picture = Gradient();
        int written = 0;

        foreach (string file in Directory.GetFiles(folder, "*.xlsx").Order(StringComparer.Ordinal))
        {
            Result<Workbook> opened = await Workbook.OpenAsync(file);

            if (opened.IsSuccess(out Workbook? book) == false || book.Sheets.Count == 0)
            {
                continue;
            }

            // Three numbers two rows under the sheet's last row, with a chart, a picture and a
            // data bar of their own: added next to whatever drawings and rules the file has.
            Worksheet first = book.Sheets[0];
            int top = first.RowCount + 1;
            first[top, 0] = Cell.Of("Avig.Kalkyl").Bold();

            for (int index = 0; index < 3; index++)
            {
                first[top + 1 + index, 0] = "Värde " + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                first[top + 1 + index, 1] = 10 * (index + 1);
            }

            CellRange values = new(top + 1, 1, top + 3, 1);
            CellRange labels = new(top + 1, 0, top + 3, 0);
            Chart chart = new(ChartType.Column) { Title = "Avig.Kalkyl", Series = [new ChartSeries("Värde", values.ToString(), labels.ToString())] };

            Result chartAdded = first.AddChart(chart, new CellRange(top, 3, top + 14, 9));
            Result pictureAdded = first.AddPicture(picture, new CellAddress(top, 10));

            if (chartAdded.IsSuccess(out _, out _) == false || pictureAdded.IsSuccess(out _, out _) == false)
            {
                // Strict OOXML: nothing is added, and that is what the result says.
                continue;
            }

            first.AddConditionalFormat(values, ConditionalFormat.DataBar("#638EC6"));

            if ((await book.SaveAsync(Path.Combine(target, Path.GetFileName(file)))).IsSuccess(out _, out _) == true)
            {
                written++;
            }
        }

        Expected.Add($"12-added: {written} tutorial workbooks with, on the first sheet two rows under the last one: 'Avig.Kalkyl' in bold, three values 10, 20, 30 with blue data bars, a column chart of them titled 'Avig.Kalkyl' to their right and the gradient picture beside it. The file's own charts, pictures and conditional formatting must all still be there and look as in the original.");
    }

    private static void AddRules(Action<string, ConditionalFormat> add)
    {
        add("B2:B7", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThanOrEqual, 200, new Style { Bold = true, FontColor = "#006100", Background = "#C6EFCE" }));
        add("B2:B7", ConditionalFormat.DataBar("#638EC6"));
        add("C2:C7", ConditionalFormat.ColorScale("#F8696B", "#FFEB84", "#63BE7B"));
        add("D2:D7", ConditionalFormat.IconSet(IconSetKind.ThreeArrows));
        add("A2:D7", ConditionalFormat.WhenFormula("$D2<50", new Style { Background = "#FFEB9C" }));
    }

    private static IEnumerable<(Chart Chart, CellRange Area)> Charts()
    {
        Chart[] charts =
        [
            new(ChartType.Column) { Title = "Intäkt och kostnad", CategoryAxisTitle = "Månad", ValueAxisTitle = "kr", Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7", "#4472C4"), new ChartSeries("Kostnad", "C2:C7", "A2:A7")] },
            new(ChartType.Bar) { Stacked = true, Legend = LegendPosition.Bottom, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7"), new ChartSeries("Kostnad", "C2:C7", "A2:A7")] },
            new(ChartType.Line) { Legend = LegendPosition.None, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7", "#ED7D31")] },
            new(ChartType.Area) { Stacked = true, Legend = LegendPosition.Top, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7"), new ChartSeries("Kostnad", "C2:C7", "A2:A7")] },
            new(ChartType.Pie) { Title = "Andel", Legend = LegendPosition.Left, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7")] },
            new(ChartType.Doughnut) { Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7")] },
            new(ChartType.Scatter) { CategoryAxisTitle = "Intäkt", ValueAxisTitle = "Kostnad", Series = [new ChartSeries("Månader", "C2:C7", "B2:B7", "#70AD47")] },
        ];

        for (int index = 0; index < charts.Length; index++)
        {
            yield return (charts[index], new CellRange(index * 17, 5, index * 17 + 15, 12));
        }
    }

    /// <summary>
    /// A 64×48 PNG with a colour gradient, made here so that the tool needs no image file: an
    /// IHDR, one IDAT of zlib data and an IEND, each with its CRC.
    /// </summary>
    private static Picture Gradient()
    {
        const int width = 64;
        const int height = 48;
        using MemoryStream pixels = new();

        using (ZLibStream zlib = new(pixels, CompressionLevel.Optimal, true))
        {
            for (int row = 0; row < height; row++)
            {
                zlib.WriteByte(0);

                for (int column = 0; column < width; column++)
                {
                    zlib.WriteByte((byte)(column * 4));
                    zlib.WriteByte((byte)(row * 5));
                    zlib.WriteByte(200);
                }
            }
        }

        using MemoryStream png = new();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        byte[] header = new byte[13];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header, width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        header[9] = 2;
        WriteChunk(png, "IHDR", header);
        WriteChunk(png, "IDAT", pixels.ToArray());
        WriteChunk(png, "IEND", []);
        Result<Picture> picture = Picture.FromBytes(png.ToArray());
        Check(picture.IsSuccess(out Picture? made), picture.Error);

        return made with { Description = "Färgskala" };
    }

    private static void WriteChunk(Stream png, string type, byte[] data)
    {
        byte[] length = new byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(length, data.Length);
        png.Write(length);
        byte[] typeAndData = [.. Encoding.ASCII.GetBytes(type), .. data];
        png.Write(typeAndData);
        byte[] crc = new byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(typeAndData));
        png.Write(crc);
    }

    private static uint Crc32(byte[] bytes)
    {
        uint crc = 0xFFFFFFFF;

        foreach (byte value in bytes)
        {
            crc ^= value;

            for (int bit = 0; bit < 8; bit++)
            {
                uint mask = (uint)-(int)(crc & 1);
                crc = (crc >> 1) ^ (0xEDB88320 & mask);
            }
        }

        return ~crc;
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
