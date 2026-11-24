using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

namespace Avig.Kalkyl.Tests;

/// <summary>Charts, pictures and conditional formatting: written by both writers, and added to opened files.</summary>
public class DrawingTests
{
    /// <summary>A 1×1 PNG.</summary>
    internal static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");

    private static readonly Style Red = new() { FontColor = "#9C0006", Background = "#FFC7CE" };

    private static readonly Style Green = new() { Bold = true, FontColor = "#006100", Background = "#C6EFCE" };

    [Fact]
    public async Task Every_chart_type_picture_and_rule_from_the_streaming_writer_passes_schema_validation()
    {
        Picture picture = PngPicture();

        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter data = await TestFiles.AddSheetAsync(writer, "Försäljning 2026");
            await WriteDataAsync(data);
            AddEveryRule(data);
            AddEveryChart(data, null);
            data.AddPicture(picture, "N2");
            data.AddPicture(picture, new CellRange(30, 13, 35, 16));

            SheetWriter report = await TestFiles.AddSheetAsync(writer, "Rapport");
            await report.WriteRowAsync(["Diagram över ett annat blad"]);
            Result added = report.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries("Intäkt", "'Försäljning 2026'!B2:B7", "'Försäljning 2026'!A2:A7")] }, "B3:I18");
            Assert.True(added.IsSuccess(out _, out _));
        });

        AssertValid(file);

        using SpreadsheetDocument document = SpreadsheetDocument.Open(file, false);
        WorksheetPart first = SheetPart(document, "Försäljning 2026");
        Assert.NotNull(first.DrawingsPart);
        Assert.Equal(7, first.DrawingsPart!.ChartParts.Count());
        Assert.Equal(2, first.DrawingsPart.ImageParts.Count());
        Assert.Equal(12, first.Worksheet!.Elements<Excel.ConditionalFormatting>().Count());
        Assert.Single(SheetPart(document, "Rapport").DrawingsPart!.ChartParts);
    }

    [Theory]
    [InlineData(IconSetKind.ThreeArrows, "3Arrows", 3)]
    [InlineData(IconSetKind.ThreeArrowsGray, "3ArrowsGray", 3)]
    [InlineData(IconSetKind.ThreeFlags, "3Flags", 3)]
    [InlineData(IconSetKind.ThreeTrafficLights, "3TrafficLights1", 3)]
    [InlineData(IconSetKind.ThreeTrafficLightsRimmed, "3TrafficLights2", 3)]
    [InlineData(IconSetKind.ThreeSigns, "3Signs", 3)]
    [InlineData(IconSetKind.ThreeSymbols, "3Symbols", 3)]
    [InlineData(IconSetKind.ThreeSymbolsUncircled, "3Symbols2", 3)]
    [InlineData(IconSetKind.FourArrows, "4Arrows", 4)]
    [InlineData(IconSetKind.FourArrowsGray, "4ArrowsGray", 4)]
    [InlineData(IconSetKind.FourRedToBlack, "4RedToBlack", 4)]
    [InlineData(IconSetKind.FourRating, "4Rating", 4)]
    [InlineData(IconSetKind.FourTrafficLights, "4TrafficLights", 4)]
    [InlineData(IconSetKind.FiveArrows, "5Arrows", 5)]
    [InlineData(IconSetKind.FiveArrowsGray, "5ArrowsGray", 5)]
    [InlineData(IconSetKind.FiveRating, "5Rating", 5)]
    [InlineData(IconSetKind.FiveQuarters, "5Quarters", 5)]
    public async Task Every_icon_set_gets_Excels_name_and_one_threshold_per_icon(IconSetKind icons, string name, int count)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter data = await TestFiles.AddSheetAsync(writer, "Data");
            await WriteDataAsync(data);
            data.AddConditionalFormat("B2:B7", ConditionalFormat.IconSet(icons));
        });

        AssertValid(file);

        using SpreadsheetDocument document = SpreadsheetDocument.Open(file, false);
        Excel.IconSet iconSet = SheetPart(document, "Data").Worksheet!.Descendants<Excel.IconSet>().Single();
        Assert.Equal(name, iconSet.IconSetValue!.InnerText);
        Assert.Equal(count, iconSet.Elements<Excel.ConditionalFormatValueObject>().Count());
    }

    [Fact]
    public async Task The_document_model_writes_the_same_in_a_new_file()
    {
        Workbook book = Workbook.Create();
        Worksheet data = Added(book, "Data");
        WriteData(data);
        AddEveryRule(data);
        AddEveryChart(null, data);
        Assert.True(data.AddPicture(PngPicture(), new CellAddress(1, 13), 4).IsSuccess(out _, out _));

        await using MemoryStream file = new();
        await book.SaveAsync(file);

        AssertValid(file);
        string chart = TestFiles.ReadPart(file, "xl/charts/chart1.xml");
        Assert.Contains("<c:f>'Data'!$B$2:$B$7</c:f>", chart, StringComparison.Ordinal);
        Assert.Contains("<c:f>'Data'!$A$2:$A$7</c:f>", chart, StringComparison.Ordinal);

        // At scale 4, the 1-pixel picture is 4 × 9525 EMU.
        Assert.Contains("cx=\"38100\"", TestFiles.ReadPart(file, "xl/drawings/drawing1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Rules_get_their_own_differential_formats_and_priorities()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await WriteDataAsync(sheet);
            sheet.AddConditionalFormat("B2:B7", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThan, 200, Green));
            sheet.AddConditionalFormat("C2:C7", ConditionalFormat.WhenCell(ComparisonOperator.LessThan, 0, Red));
            sheet.AddConditionalFormat("D2:D7", ConditionalFormat.WhenCell(ComparisonOperator.Equal, "\"Ja\"", Green));
            sheet.AddConditionalFormat("A2:A7", ConditionalFormat.WhenFormula("=$B2>$C2", new Style { NumberFormat = "0.0%" }));
        });

        string styles = TestFiles.ReadPart(file, "xl/styles.xml");
        string sheetXml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");

        // Equal styles share a dxf: three distinct styles for four rules.
        Assert.Contains("<dxfs count=\"3\">", styles, StringComparison.Ordinal);
        Assert.Contains("formatCode=\"0.0%\"", styles, StringComparison.Ordinal);
        Assert.Equal(["0", "1", "0", "2"], Attribute(sheetXml, "dxfId"));
        Assert.Equal(["1", "2", "3", "4"], Attribute(sheetXml, "priority"));
        Assert.Contains("<formula>$B2&gt;$C2</formula>", sheetXml, StringComparison.Ordinal);
        Assert.Contains("<formula>\"Ja\"</formula>", sheetXml, StringComparison.Ordinal);
        AssertValid(file);
    }

    [Fact]
    public async Task Text_rules_compare_from_the_top_left_cell_and_quote_the_text()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            sheet.AddConditionalFormat("C3:C9", ConditionalFormat.WhenTextContains("säger \"hej\"", Red));
        });

        string sheetXml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");

        Assert.Contains("<formula>NOT(ISERROR(SEARCH(\"säger \"\"hej\"\"\",C3)))</formula>", sheetXml, StringComparison.Ordinal);
        AssertValid(file);
    }

    [Fact]
    public async Task An_opened_file_gets_a_drawing_and_rules_and_keeps_its_cells()
    {
        await using MemoryStream original = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await WriteDataAsync(sheet);
        });

        Workbook book = await OpenAsync(original);
        Worksheet data = book.Sheets[0];
        Assert.True(data.AddChart(new Chart(ChartType.Line) { Title = "Intäkt", Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7")] }, "F2:M18").IsSuccess(out _, out _));
        Assert.True(data.AddPicture(PngPicture(), "O2").IsSuccess(out _, out _));
        data.AddConditionalFormat("B2:B7", ConditionalFormat.DataBar("#638EC6"));

        await using MemoryStream saved = await SaveAsync(book);

        AssertValid(saved);
        List<Row> rows = await TestFiles.ReadRowsAsync(saved);
        Assert.Equal((CellValue)"Jan", rows[1][0].Value);
        Assert.Equal((CellValue)120, rows[1][1].Value);

        using SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false);
        WorksheetPart part = SheetPart(document, "Data");
        Assert.Single(part.DrawingsPart!.ChartParts);
        Assert.Single(part.DrawingsPart.ImageParts);
        Assert.Single(part.Worksheet!.Elements<Excel.ConditionalFormatting>());
    }

    [Fact]
    public async Task A_sheet_that_has_a_drawing_and_rules_gets_more_added_to_them()
    {
        await using MemoryStream original = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await WriteDataAsync(sheet);
            sheet.AddConditionalFormat("B2:B7", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThan, 200, Green));
            sheet.AddConditionalFormat("C2:C7", ConditionalFormat.ColorScale("#F8696B", "#63BE7B"));
            Assert.True(sheet.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries("Intäkt", "B2:B7")] }, "F2:M18").IsSuccess(out _, out _));
            sheet.AddPicture(PngPicture(), "O2");
        });

        // Twice: the second time, the drawing and rules being added to are the package's own from the first.
        await using MemoryStream once = await AddMoreAsync(original);
        await using MemoryStream twice = await AddMoreAsync(once);

        AssertValid(twice);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(twice, false);
        WorksheetPart part = SheetPart(document, "Data");
        Assert.Single(part.Worksheet!.Elements<Excel.Drawing>());
        Assert.Equal(3, part.DrawingsPart!.ChartParts.Count());
        Assert.Equal(3, part.DrawingsPart.ImageParts.Count());

        string drawing = TestFiles.ReadPart(twice, "xl/drawings/drawing1.xml");
        List<string> shapeIds = Regex.Matches(drawing, "<xdr:cNvPr id=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();
        Assert.Equal(6, shapeIds.Count);
        Assert.Equal(shapeIds.Count, shapeIds.Distinct().Count());

        string sheetXml = TestFiles.ReadPart(twice, "xl/worksheets/sheet1.xml");
        Assert.Equal(["1", "2", "3", "4"], Attribute(sheetXml, "priority"));

        // The first file had one dxf; each round adds one style that is new to the registry.
        Assert.Equal(["0", "1", "2"], Attribute(sheetXml, "dxfId"));
        Assert.Contains("<dxfs count=\"3\">", TestFiles.ReadPart(twice, "xl/styles.xml"), StringComparison.Ordinal);

        // png is declared once, however many pictures there are.
        Assert.Single(Regex.Matches(TestFiles.ReadPart(twice, "[Content_Types].xml"), "Extension=\"png\""));
    }

    [Fact]
    public async Task A_new_sheet_in_an_opened_file_can_chart_the_old_ones()
    {
        await using MemoryStream original = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await WriteDataAsync(sheet);
        });

        Workbook book = await OpenAsync(original);
        Worksheet report = Added(book, "Rapport");
        Assert.True(report.AddChart(new Chart(ChartType.Pie) { Series = [new ChartSeries(null, "Data!B2:B7", "Data!A2:A7")] }, "A1:H16").IsSuccess(out _, out _));
        report.AddConditionalFormat("A1:A10", ConditionalFormat.Top(3, Green));

        await using MemoryStream saved = await SaveAsync(book);

        AssertValid(saved);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false);
        Assert.Single(SheetPart(document, "Rapport").DrawingsPart!.ChartParts);
    }

    [Fact]
    public async Task A_series_on_a_sheet_that_does_not_exist_is_NotFound()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Added(book, "Data");
        Chart chart = new(ChartType.Column) { Series = [new ChartSeries(null, "Saknas!B2:B7")] };

        Assert.True(sheet.AddChart(chart, "A1:H16").IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
        Assert.Contains("Saknas", error, StringComparison.Ordinal);

        // The streaming writer sees only the sheets added before: data comes first.
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter report = await TestFiles.AddSheetAsync(writer, "Rapport");
            Assert.True(report.AddChart(chart, "A1:H16").IsFailure(out _, out ResultStatus streamed));
            Assert.Equal(ResultStatus.NotFound, streamed);
        });
    }

    [Fact]
    public async Task A_Strict_OOXML_file_gets_no_drawings()
    {
        const string strictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        const string packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        await using MemoryStream file = TestFiles.Package(
            ("_rels/.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", $"""<workbook xmlns="{strictMain}" xmlns:r="{strictRelationships}"><sheets><sheet name="Strikt" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{strictMain}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"""));

        Workbook book = await OpenAsync(file);

        Assert.True(book.Sheets[0].AddPicture(PngPicture(), "B2").IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Contains("Strict", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Mistakes_in_a_rule_or_a_chart_are_programming_errors()
    {
        Assert.Throws<ArgumentException>(() => ConditionalFormat.WhenCell(ComparisonOperator.Equal, 1, new Style { FontSize = 14 }));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.WhenCell(ComparisonOperator.Equal, 1, new Style { Horizontal = HorizontalAlignment.Center }));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.WhenCell(ComparisonOperator.Equal, 1, Style.Default));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.WhenFormula("SUM(", Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => ConditionalFormat.WhenCell(ComparisonOperator.Equal, double.NaN, Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => ConditionalFormat.Top(0, Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => ConditionalFormat.Top(101, Red, true));
        Assert.Throws<ArgumentException>(() => ConditionalFormat.ColorScale("red", "#00FF00"));
        Assert.Throws<ArgumentException>(() => new ChartSeries(null, "B2:C7"));
        Assert.Throws<ArgumentException>(() => new ChartSeries(null, "B2:B7", "A2:A6"));
        Assert.Throws<ArgumentException>(() => new ChartSeries(null, "B2-B7"));
        Assert.Throws<ArgumentException>(() => new ChartSeries(null, "Bad[name]!B2:B7"));
        Worksheet empty = Added(Workbook.Create(), "Data");
        Assert.Throws<ArgumentException>(() => empty.AddChart(new Chart(ChartType.Column), "A1:B2"));
    }

    [Fact]
    public void Ranges_parse_in_either_corner_order_and_print_in_A1_notation()
    {
        Assert.True(CellRange.Parse("$D$10:B2").IsSuccess(out CellRange range));
        Assert.Equal(new CellRange(1, 1, 9, 3), range);
        Assert.Equal("B2:D10", range.ToString());
        Assert.Equal(9, range.RowCount);
        Assert.Equal(3, range.ColumnCount);
        Assert.True(CellRange.Parse("C5").IsSuccess(out CellRange single));
        Assert.Equal("C5", single.ToString());
        Assert.True(range.Contains(new CellAddress(5, 2)));
        Assert.False(range.Contains(new CellAddress(0, 2)));
        Assert.False(CellRange.TryParse("B2:", out _));
        Assert.False(CellRange.TryParse("A0", out _));
        Assert.Throws<ArgumentException>(() => new CellRange(new CellAddress(5, 5), new CellAddress(1, 1)));
    }

    [Fact]
    public void Picture_headers_give_the_format_and_size()
    {
        Assert.True(Picture.FromBytes(Png).IsSuccess(out Picture? png));
        Assert.Equal((PictureFormat.Png, 1, 1), (png!.Format, png.Width, png.Height));

        byte[] gif = [.. "GIF89a"u8, 0x40, 0x01, 0xC8, 0x00, 0, 0, 0];
        Assert.True(Picture.FromBytes(gif).IsSuccess(out Picture? gifPicture));
        Assert.Equal((PictureFormat.Gif, 320, 200), (gifPicture!.Format, gifPicture.Width, gifPicture.Height));

        Assert.True(Picture.FromBytes(Jpeg(640, 480)).IsSuccess(out Picture? jpeg));
        Assert.Equal((PictureFormat.Jpeg, 640, 480), (jpeg!.Format, jpeg.Width, jpeg.Height));
    }

    [Theory]
    [InlineData(new byte[0])]
    [InlineData(new byte[] { 0x42, 0x4D, 0, 0, 0, 0 })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x00 })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0xFF, 0xFF, 0x00 })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08 })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xDA, 0x00, 0x02 })]
    [InlineData(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0, 0, 0, 0 })]
    public void Damaged_or_unknown_images_are_Invalid(byte[] data)
    {
        Assert.True(Picture.FromBytes(data).IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
    }

    [Fact]
    public void Random_bytes_after_an_image_signature_never_throw()
    {
        Random random = new(4711);
        byte[][] signatures = [[0xFF, 0xD8], [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], [.. "GIF89a"u8]];

        for (int round = 0; round < 20_000; round++)
        {
            byte[] signature = signatures[round % signatures.Length];
            byte[] data = new byte[signature.Length + random.Next(0, 64)];
            random.NextBytes(data);
            signature.CopyTo(data, 0);

            // Either answer is fine; an exception or a hang is not.
            _ = Picture.FromBytes(data);
        }
    }

    [Fact]
    public async Task A_missing_picture_file_is_NotFound()
    {
        Result<Picture> read = await Picture.FromFileAsync(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png"));

        Assert.True(read.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
    }

    internal static Picture PngPicture()
    {
        Assert.True(Picture.FromBytes(Png).IsSuccess(out Picture? picture));

        return picture!;
    }

    /// <summary>A JPEG header: SOI, an APP0 segment, then a baseline frame header with the size.</summary>
    private static byte[] Jpeg(int width, int height)
    {
        return
        [
            0xFF, 0xD8,
            0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00,
            0xFF, 0xC0, 0x00, 0x11, 0x08, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 0x03,
            0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
            0xFF, 0xD9,
        ];
    }

    private static readonly (string Month, double Income, double Cost, string Done)[] Months =
    [
        ("Jan", 120, -30, "Ja"), ("Feb", 180, 20, "Nej"), ("Mar", 240, 15, "Ja"), ("Apr", 90, -5, "Ja"), ("Maj", 310, 40, "Nej"), ("Jun", 205, 10, "Ja"),
    ];

    private static async Task WriteDataAsync(SheetWriter sheet)
    {
        await sheet.WriteRowAsync(["Månad", "Intäkt", "Kostnad", "Klar"]);

        foreach ((string month, double income, double cost, string done) in Months)
        {
            await sheet.WriteRowAsync([month, income, cost, done]);
        }
    }

    private static void WriteData(Worksheet sheet)
    {
        sheet["A1"] = "Månad";
        sheet["B1"] = "Intäkt";
        sheet["C1"] = "Kostnad";
        sheet["D1"] = "Klar";

        for (int index = 0; index < Months.Length; index++)
        {
            sheet[index + 1, 0] = Months[index].Month;
            sheet[index + 1, 1] = Months[index].Income;
            sheet[index + 1, 2] = Months[index].Cost;
            sheet[index + 1, 3] = Months[index].Done;
        }
    }

    /// <summary>Twelve rules, one of every kind, on whichever writer is given.</summary>
    private static void AddEveryRule(object sheet)
    {
        (string Range, ConditionalFormat Format)[] rules =
        [
            ("B2:B7", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThanOrEqual, 200, Green)),
            ("B2:B7", ConditionalFormat.WhenBetween(100, 150, Red)),
            ("C2:C7", ConditionalFormat.WhenNotBetween(-10, 10, new Style { Italic = true, Underline = true })),
            ("A2:D7", ConditionalFormat.WhenFormula("$D2=\"Nej\"", new Style { Background = "#FFEB9C" })),
            ("D2:D7", ConditionalFormat.WhenTextContains("ej", Red)),
            ("B2:B7", ConditionalFormat.Top(2, Green)),
            ("B2:B7", ConditionalFormat.Bottom(20, Red, true)),
            ("C2:C7", ConditionalFormat.BelowAverage(Red)),
            ("D2:D7", ConditionalFormat.Duplicates(new Style { Bold = true })),
            ("C2:C7", ConditionalFormat.ColorScale("#F8696B", "#FFEB84", "#63BE7B")),
            ("B2:B7", ConditionalFormat.DataBar("#638EC6")),
            ("C2:C7", ConditionalFormat.IconSet(IconSetKind.FiveQuarters, reverse: true, showValue: false)),
        ];

        foreach ((string range, ConditionalFormat format) in rules)
        {
            if (sheet is SheetWriter writer)
            {
                writer.AddConditionalFormat(range, format);
            }
            else
            {
                ((Worksheet)sheet).AddConditionalFormat(range, format);
            }
        }
    }

    /// <summary>One chart of each type, stacked and with titles where they apply.</summary>
    private static void AddEveryChart(SheetWriter? writer, Worksheet? sheet)
    {
        Chart[] charts =
        [
            new(ChartType.Column) { Title = "Intäkt och kostnad", Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7", "#4472C4"), new ChartSeries("Kostnad", "C2:C7", "A2:A7")], CategoryAxisTitle = "Månad", ValueAxisTitle = "kr", Description = "Staplar per månad" },
            new(ChartType.Bar) { Stacked = true, Legend = LegendPosition.Bottom, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7"), new ChartSeries("Kostnad", "C2:C7", "A2:A7")] },
            new(ChartType.Line) { Legend = LegendPosition.None, Series = [new ChartSeries("Intäkt", "$B$2:$B$7", "A2:A7", "#ED7D31")] },
            new(ChartType.Area) { Stacked = true, Legend = LegendPosition.Top, Series = [new ChartSeries("Intäkt", "B2:B7"), new ChartSeries("Kostnad", "C2:C7")] },
            new(ChartType.Pie) { Title = "Andel", Legend = LegendPosition.Left, Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7", "#FF0000")] },
            new(ChartType.Doughnut) { Series = [new ChartSeries("Intäkt", "B2:B7", "A2:A7")] },
            new(ChartType.Scatter) { CategoryAxisTitle = "Intäkt", ValueAxisTitle = "Kostnad", Series = [new ChartSeries("Punkter", "C2:C7", "B2:B7", "#70AD47")] },
        ];

        for (int index = 0; index < charts.Length; index++)
        {
            CellRange area = new(index * 18, 5, index * 18 + 15, 12);
            Result added;

            if (writer is not null)
            {
                added = writer.AddChart(charts[index], area);
            }
            else
            {
                added = sheet!.AddChart(charts[index], area);
            }

            Assert.True(added.IsSuccess(out _, out _));
        }
    }

    private static async Task<MemoryStream> AddMoreAsync(Stream file)
    {
        Workbook book = await OpenAsync(file);
        Worksheet data = book.Sheets[0];
        Assert.True(data.AddChart(new Chart(ChartType.Line) { Series = [new ChartSeries("Kostnad", "C2:C7")] }, "F20:M36").IsSuccess(out _, out _));
        Assert.True(data.AddPicture(PngPicture(), "O20").IsSuccess(out _, out _));
        data.AddConditionalFormat("D2:D7", ConditionalFormat.Unique(new Style { Italic = true }));

        return await SaveAsync(book);
    }

    private static async Task<Workbook> OpenAsync(Stream file)
    {
        file.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        return book;
    }

    private static async Task<MemoryStream> SaveAsync(Workbook book)
    {
        MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        return saved;
    }

    private static Worksheet Added(Workbook book, string name)
    {
        Result<Worksheet> added = book.AddSheet(name);
        Assert.True(added.IsSuccess(out Worksheet? sheet), added.Error);

        return sheet;
    }

    private static WorksheetPart SheetPart(SpreadsheetDocument document, string name)
    {
        WorkbookPart workbook = document.WorkbookPart!;
        Excel.Sheet sheet = workbook.Workbook!.Descendants<Excel.Sheet>().Single(candidate => candidate.Name == name);

        return (WorksheetPart)workbook.GetPartById(sheet.Id!);
    }

    private static List<string> Attribute(string xml, string name)
    {
        return Regex.Matches(xml, " " + name + "=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();
    }

    private static void AssertValid(MemoryStream file)
    {
        List<string> problems = TestFiles.Validate(file);
        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));

        // Every part the package names exists, and the zip has no name twice.
        file.Position = 0;
        using ZipArchive archive = new(file, ZipArchiveMode.Read, true);
        Assert.Equal(archive.Entries.Count, archive.Entries.Select(entry => entry.FullName).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        file.Position = 0;
    }
}
