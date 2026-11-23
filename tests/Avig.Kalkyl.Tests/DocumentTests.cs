using System.IO.Compression;
using System.Text;

namespace Avig.Kalkyl.Tests;

public class DocumentTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

    private const string DrawingXml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>""";

    /// <summary>
    /// A file the way Excel writes it, with parts the package does not understand: merged cells,
    /// conditional formatting, data validation, a drawing, a calculation chain, document
    /// properties, rich shared text, a border style and a shared formula.
    /// </summary>
    private static MemoryStream ExcelFile()
    {
        return TestFiles.Package(
            ("[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
                  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
                  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
                </Types>
                """),
            ("_rels/.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationships}">
                  <Relationship Id="rId1" Type="{Relationships}/officeDocument" Target="xl/workbook.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
                </Relationships>
                """),
            ("docProps/app.xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>"""),
            ("xl/workbook.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="{Main}" xmlns:r="{Relationships}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                  <workbookPr defaultThemeVersion="166925"/>
                  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews>
                  <sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets>
                  <definedNames><definedName name="Moms">0.25</definedName></definedNames>
                  <calcPr calcId="191029"/>
                </workbook>
                """),
            ("xl/_rels/workbook.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationships}">
                  <Relationship Id="rId1" Type="{Relationships}/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="{Relationships}/styles" Target="styles.xml"/>
                  <Relationship Id="rId3" Type="{Relationships}/sharedStrings" Target="sharedStrings.xml"/>
                  <Relationship Id="rId4" Type="{Relationships}/calcChain" Target="calcChain.xml"/>
                </Relationships>
                """),
            ("xl/sharedStrings.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <sst xmlns="{Main}" count="2" uniqueCount="2"><si><t>Post</t></si><si><r><rPr><b/></rPr><t>Fet</t></r><r><t xml:space="preserve"> och normal</t></r></si></sst>
                """),
            ("xl/styles.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <styleSheet xmlns="{Main}">
                  <fonts count="1"><font><sz val="11"/><name val="Aptos Narrow"/></font></fonts>
                  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
                  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs>
                  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
                </styleSheet>
                """),
            ("xl/calcChain.xml", $"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="{Main}"><c r="C2" i="1"/></calcChain>"""),
            ("xl/worksheets/_rels/sheet1.xml.rels", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/drawing" Target="../drawings/drawing1.xml"/></Relationships>
                """),
            ("xl/drawings/drawing1.xml", DrawingXml),
            ("xl/worksheets/sheet1.xml", $"""
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <worksheet xmlns="{Main}" xmlns:r="{Relationships}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
                  <dimension ref="A1:C4"/>
                  <sheetViews><sheetView tabSelected="1" zoomScale="120" workbookViewId="0"/></sheetViews>
                  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
                  <cols><col min="1" max="3" width="14.5" customWidth="1"/></cols>
                  <sheetData>
                    <row r="1" spans="1:3" ht="20" customHeight="1" x14ac:dyDescent="0.25"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
                    <row r="2" spans="1:3"><c r="A2"><v>100</v></c><c r="B2" s="1"><v>200</v></c><c r="C2"><f t="shared" ref="C2:C3" si="0">A2+B2</f><v>300</v></c></row>
                    <row r="3" spans="1:3"><c r="A3"><v>1</v></c><c r="B3" s="1"><v>2</v></c><c r="C3"><f t="shared" si="0"/><v>3</v></c></row>
                  </sheetData>
                  <mergeCells count="1"><mergeCell ref="A4:C4"/></mergeCells>
                  <conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>
                  <dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="A2:A3"><formula1>0</formula1><formula2>1000</formula2></dataValidation></dataValidations>
                  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
                  <drawing r:id="rId1"/>
                </worksheet>
                """));
    }

    [Fact]
    public async Task A_new_workbook_is_saved_as_a_valid_file_and_reads_back()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Kunder").IsSuccess(out Worksheet? customers));
        Assert.True(book.AddSheet("Tom").IsSuccess(out _));
        customers["A1"] = Cell.Of("Namn").Bold().Background("#DDEBF7");
        customers[0, 1] = Cell.Of("Datum").Bold();
        customers[1, 0] = "Ada";
        customers[1, 1] = new DateTime(2026, 1, 2);
        customers[5, 3] = Cell.FromFormula("LEN(A2)", 3);
        customers.SetColumnWidth(0, 25);
        customers.FreezePanes(1);
        customers.AutoFilter = true;

        await using MemoryStream file = new();
        await book.SaveAsync(file);

        Assert.Empty(TestFiles.Validate(file));
        List<Row> rows = await TestFiles.ReadRowsAsync(file);
        Assert.Equal((CellValue)"Ada", rows[1][0].Value);
        Assert.Equal((CellValue)new DateTime(2026, 1, 2), rows[1][1].Value);
        Assert.True(rows[0][0].Style.Bold);
        Assert.Equal("#DDEBF7", rows[0][0].Style.Background);
        Assert.Equal("LEN(A2)", rows[2][3].Formula);
        Assert.Equal(5, rows[2].Index);
    }

    [Fact]
    public async Task An_unchanged_file_keeps_every_part_and_cell_as_it_was()
    {
        await using MemoryStream original = ExcelFile();
        Result<Workbook> opened = await Workbook.OpenAsync(original);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);

        Assert.Equal(DrawingXml, TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml"));
        Assert.Equal(TestFiles.ReadPart(original, "xl/sharedStrings.xml"), TestFiles.ReadPart(saved, "xl/sharedStrings.xml"));
        Assert.Equal(TestFiles.ReadPart(original, "docProps/app.xml"), TestFiles.ReadPart(saved, "docProps/app.xml"));

        string sheet = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("<c r=\"A1\" t=\"s\"><v>0</v></c>", sheet, StringComparison.Ordinal);
        Assert.Contains("<c r=\"B2\" s=\"1\"><v>200</v></c>", sheet, StringComparison.Ordinal);
        Assert.Contains("ht=\"20\" customHeight=\"1\" x14ac:dyDescent=\"0.25\"", sheet, StringComparison.Ordinal);
        Assert.Contains("zoomScale=\"120\"", sheet, StringComparison.Ordinal);
        Assert.Contains("<mergeCell ref=\"A4:C4\" />", sheet, StringComparison.Ordinal);
        Assert.Contains("<conditionalFormatting sqref=\"A2:A3\">", sheet, StringComparison.Ordinal);
        Assert.Contains("<dataValidation type=\"whole\"", sheet, StringComparison.Ordinal);
        Assert.Contains("<drawing r:id=\"rId1\" />", sheet, StringComparison.Ordinal);
        Assert.DoesNotContain("<dimension", sheet, StringComparison.Ordinal);
        Assert.Contains("<definedName name=\"Moms\">0.25</definedName>", TestFiles.ReadPart(saved, "xl/workbook.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Shared_formulas_become_explicit_and_the_calculation_chain_is_dropped()
    {
        await using MemoryStream original = ExcelFile();
        Assert.True((await Workbook.OpenAsync(original)).IsSuccess(out Workbook? book));
        Assert.True(book.Sheet("Budget").IsSuccess(out Worksheet? sheet));

        Assert.Equal("A2+B2", sheet["C2"].Formula);
        Assert.Equal("A3+B3", sheet["C3"].Formula);

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        string sheetXml = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");

        Assert.DoesNotContain("t=\"shared\"", sheetXml, StringComparison.Ordinal);
        Assert.Contains("<c r=\"C3\"><f>A3+B3</f><v>3</v></c>", sheetXml, StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => TestFiles.ReadPart(saved, "xl/calcChain.xml"));
        Assert.DoesNotContain("calcChain", TestFiles.ReadPart(saved, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
        Assert.DoesNotContain("calcChain", TestFiles.ReadPart(saved, "[Content_Types].xml"), StringComparison.Ordinal);
        Assert.Contains("fullCalcOnLoad=\"1\"", TestFiles.ReadPart(saved, "xl/workbook.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Changes_are_saved_and_everything_else_survives()
    {
        await using MemoryStream original = ExcelFile();
        Assert.True((await Workbook.OpenAsync(original)).IsSuccess(out Workbook? book));
        Assert.True(book.Sheet(0).IsSuccess(out Worksheet? sheet));

        Assert.Equal("#,##0.00", sheet["B2"].Style.NumberFormat);
        Assert.Equal("Fet och normal", sheet["B1"].Value.ToString());

        sheet.SetValue(1, 1, 250);
        sheet["A3"] = Cell.Of(7).Italic().FontColor("#FF0000");
        sheet["D1"] = Cell.Of("Ny").Bold();
        sheet.SetColumnWidth(1, 30);
        Assert.True(book.AddSheet("Anteckningar").IsSuccess(out Worksheet? notes));
        notes["A1"] = "Hej";

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        Assert.True((await Workbook.OpenAsync(saved)).IsSuccess(out Workbook? reopened));
        Assert.True(reopened.Sheet("Budget").IsSuccess(out Worksheet? budget));
        Assert.Equal((CellValue)250, budget["B2"].Value);
        Assert.Equal("#,##0.00", budget["B2"].Style.NumberFormat);
        Assert.True(budget["A3"].Style.Italic);
        Assert.Equal("#FF0000", budget["A3"].Style.FontColor);
        Assert.True(budget["D1"].Style.Bold);
        Assert.Equal(30, budget.GetColumnWidth(1));
        Assert.Equal(14.5, budget.GetColumnWidth(0));
        Assert.Equal(14.5, budget.GetColumnWidth(2));
        Assert.True(reopened.Sheet("Anteckningar").IsSuccess(out Worksheet? reopenedNotes));
        Assert.Equal((CellValue)"Hej", reopenedNotes["A1"].Value);

        string sheetXml = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.Contains("<c r=\"B2\" s=\"1\"><v>250</v></c>", sheetXml, StringComparison.Ordinal);
        Assert.Contains("<drawing r:id=\"rId1\" />", sheetXml, StringComparison.Ordinal);

        string styles = TestFiles.ReadPart(saved, "xl/styles.xml");
        Assert.Contains("<cellXfs count=\"4\">", styles, StringComparison.Ordinal);
        Assert.Contains("borderId=\"1\"", styles, StringComparison.Ordinal);
        Assert.Contains("<name val=\"Aptos Narrow\" />", styles, StringComparison.Ordinal);
        Assert.Equal(DrawingXml, TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml"));
    }

    [Fact]
    public async Task A_file_from_the_writer_survives_open_and_save_with_valid_schema()
    {
        await using MemoryStream written = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            sheet.FreezePanes(1);
            sheet.SetColumnWidth(0, 20);
            sheet.EnableAutoFilter();
            await sheet.WriteRowAsync([Cell.Of("Namn").Bold(), Cell.Of("Belopp").Bold()]);
            await sheet.WriteRowAsync(["Ada", Cell.Of(12.5).NumberFormat("0.00")]);
            await sheet.WriteRowAsync(["Summa", Cell.FromFormula("SUM(B2:B2)", 12.5)]);
        });

        Assert.True((await Workbook.OpenAsync(written)).IsSuccess(out Workbook? book));
        Assert.True(book.Sheet(0).IsSuccess(out Worksheet? data));
        data["C2"] = Cell.Of(new DateOnly(2026, 11, 23));
        data.AutoFilter = false;
        data.FreezePanes(0, 1);

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);

        Assert.Empty(TestFiles.Validate(saved));
        List<Row> rows = await TestFiles.ReadRowsAsync(saved);
        Assert.Equal((CellValue)12.5, rows[1][1].Value);
        Assert.Equal("SUM(B2:B2)", rows[2][1].Formula);
        Assert.Equal((CellValue)new DateTime(2026, 11, 23), rows[1][2].Value);
        string sheetXml = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("<autoFilter", sheetXml, StringComparison.Ordinal);
        Assert.Contains("xSplit=\"1\"", sheetXml, StringComparison.Ordinal);
        Assert.DoesNotContain("_FilterDatabase", TestFiles.ReadPart(saved, "xl/workbook.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public void SetValue_keeps_the_style_and_the_indexer_replaces_it()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet[0, 0] = Cell.Of(1).Bold().NumberFormat("0.0");

        sheet.SetValue(0, 0, 2);
        Assert.True(sheet[0, 0].Style.Bold);
        Assert.Equal("0.0", sheet[0, 0].Style.NumberFormat);

        sheet.SetValue(0, 0, new DateTime(2026, 1, 1));
        Assert.Equal("yyyy-mm-dd", sheet[0, 0].Style.NumberFormat);

        sheet.SetFormula(0, 0, "=1+1");
        Assert.Equal("1+1", sheet[0, 0].Formula);
        Assert.True(sheet[0, 0].Style.Bold);

        sheet[0, 0] = 3;
        Assert.True(sheet[0, 0].Style.IsDefault);

        sheet.Clear(0, 0);
        Assert.True(sheet[0, 0].IsEmpty);
        Assert.Throws<ArgumentException>(() => sheet["nope"]);
    }

    [Fact]
    public async Task Hidden_sheets_stay_hidden_and_visibility_can_change()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Synlig").IsSuccess(out _));
        Assert.True(book.AddSheet("Dold").IsSuccess(out Worksheet? hidden));
        hidden.Visibility = SheetVisibility.Hidden;

        await using MemoryStream file = new();
        await book.SaveAsync(file);
        file.Position = 0;
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? reopened));

        Assert.Equal(SheetVisibility.Hidden, reopened.Sheets[1].Visibility);
    }

    [Fact]
    public async Task Problems_come_back_as_results()
    {
        Result<Workbook> missing = await Workbook.OpenAsync(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx"));
        Assert.True(missing.IsFailure(out _, out ResultStatus missingStatus));
        Assert.Equal(ResultStatus.NotFound, missingStatus);

        await using MemoryStream text = new(Encoding.UTF8.GetBytes("inte en arbetsbok"));
        Assert.True((await Workbook.OpenAsync(text)).IsFailure(out _, out ResultStatus textStatus));
        Assert.Equal(ResultStatus.Invalid, textStatus);

        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out _));
        Assert.True(book.AddSheet("DATA").IsFailure(out _, out ResultStatus duplicate));
        Assert.Equal(ResultStatus.Conflict, duplicate);
        Assert.True(book.Sheet("Saknas").IsFailure(out _, out ResultStatus noSheet));
        Assert.Equal(ResultStatus.NotFound, noSheet);

        Result saved = await book.SaveAsync(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "fil.xlsx"));
        Assert.True(saved.IsFailure(out _, out ResultStatus folder));
        Assert.Equal(ResultStatus.NotFound, folder);
    }

    [Fact]
    public async Task A_file_can_be_saved_over_itself()
    {
        string path = Path.Combine(Path.GetTempPath(), "avig-kalkyl-" + Guid.NewGuid().ToString("N") + ".xlsx");

        try
        {
            await using (MemoryStream source = ExcelFile())
            {
                await File.WriteAllBytesAsync(path, source.ToArray());
            }

            Assert.True((await Workbook.OpenAsync(path)).IsSuccess(out Workbook? book));
            book.Sheets[0]["A2"] = 999;
            Assert.True((await book.SaveAsync(path)).IsSuccess(out _, out _));

            Assert.True((await Workbook.OpenAsync(path)).IsSuccess(out Workbook? reopened));
            Assert.Equal((CellValue)999, reopened.Sheets[0]["A2"].Value);

            using ZipArchive archive = ZipFile.OpenRead(path);
            Assert.NotNull(archive.GetEntry("xl/drawings/drawing1.xml"));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
