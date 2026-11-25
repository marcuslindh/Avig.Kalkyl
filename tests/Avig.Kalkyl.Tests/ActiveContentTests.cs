using System.IO.Compression;
using Avig.OpenXml;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Active content (docs/testplan.md, S-16 and S-17): what Workbook.ActiveContent and
/// XlsxReader.GetActiveContentAsync find, and what a save after RemoveActiveContent leaves out.
/// The files are built by <see cref="ActiveWorkbook"/> from XML Excel 365 wrote.
/// </summary>
public class ActiveContentTests
{
    public static TheoryData<string, ActiveContent> Kinds()
    {
        return new TheoryData<string, ActiveContent>
        {
            { "vba", ActiveContent.Macros },
            { "macrosheet", ActiveContent.Macros },
            { "dialogsheet", ActiveContent.Macros },
            { "activex", ActiveContent.ActiveX },
            { "ole", ActiveContent.EmbeddedObjects },
            { "externallink", ActiveContent.ExternalReferences },
            { "dde", ActiveContent.ExternalReferences },
            { "linkedpicture", ActiveContent.ExternalReferences },
            { "connection", ActiveContent.DataConnections },
            { "querytable", ActiveContent.DataConnections },
            { "datamashup", ActiveContent.DataConnections },
        };
    }

    private static ActiveWorkbook With(string kind)
    {
        ActiveWorkbook book = new();

        return kind switch
        {
            "vba" => book.AddVbaProject(),
            "macrosheet" => book.AddMacroSheet(),
            "dialogsheet" => book.AddDialogSheet(),
            "activex" => book.AddActiveXControl(),
            "ole" => book.AddOleObject(),
            "externallink" => book.AddExternalWorkbookLink(),
            "dde" => book.AddDdeLink(),
            "linkedpicture" => book.AddLinkedPicture(),
            "connection" => book.AddConnection(),
            "querytable" => book.AddQueryTable(),
            "datamashup" => book.AddDataMashup(),
            _ => throw new ArgumentException(kind, nameof(kind)),
        };
    }

    [Theory]
    [MemberData(nameof(Kinds))]
    public async Task Each_kind_of_active_content_is_found_by_the_document_model_and_the_streaming_reader(string kind, ActiveContent expected)
    {
        await using MemoryStream file = With(kind).Build();

        Workbook book = await OpenAsync(file);
        ActiveContent streaming = await StreamingAsync(file);

        Assert.Equal(expected, book.ActiveContent);
        Assert.Equal(expected, streaming);
    }

    [Fact]
    public async Task A_plain_workbook_with_a_hyperlink_and_a_form_button_has_no_active_content()
    {
        await using MemoryStream file = new ActiveWorkbook().AddHyperlink().AddFormControl().Build();

        Workbook book = await OpenAsync(file);

        Assert.Equal(ActiveContent.None, book.ActiveContent);
        Assert.Empty(book.ExternalReferences);
        Assert.Equal(ActiveContent.None, await StreamingAsync(file));
        Assert.Equal(ActiveContent.None, Workbook.Create().ActiveContent);
    }

    [Fact]
    public async Task External_references_list_the_linked_workbook_and_the_picture_but_not_hyperlinks()
    {
        await using MemoryStream file = new ActiveWorkbook().AddExternalWorkbookLink().AddLinkedPicture().AddHyperlink().Build();

        Workbook book = await OpenAsync(file);

        Assert.Equal(2, book.ExternalReferences.Count);
        OpcExternalRelationship link = Assert.Single(book.ExternalReferences, reference => reference.Kind == "externalLinkPath");
        Assert.Equal("xl/externalLinks/externalLink1.xml", link.SourcePart);
        Assert.Contains("other.xlsx", link.Target, StringComparison.Ordinal);
        OpcExternalRelationship picture = Assert.Single(book.ExternalReferences, reference => reference.Kind == "image");
        Assert.Equal("https://tracker.example/pixel.png", picture.Target);
        Assert.DoesNotContain(book.ExternalReferences, reference => reference.IsHyperlink == true);

        file.Position = 0;
        Assert.True((await XlsxReader.OpenAsync(file, true)).IsSuccess(out XlsxReader? reader));

        await using (reader)
        {
            Assert.Equal(2, (await reader.GetExternalReferencesAsync()).Count);
        }
    }

    [Fact]
    public async Task A_dde_link_is_listed_with_the_program_it_starts()
    {
        await using MemoryStream file = new ActiveWorkbook().AddDdeLink().Build();

        Workbook book = await OpenAsync(file);

        OpcExternalRelationship dde = Assert.Single(book.ExternalReferences);
        Assert.Equal("ddeLink", dde.Kind);
        Assert.Equal("cmd|/c calc", dde.Target);
        Assert.Equal("xl/externalLinks/externalLink1.xml", dde.SourcePart);
    }

    [Fact]
    public async Task Removing_the_vba_project_makes_a_plain_xlsx()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddVbaProject());

        Assert.DoesNotContain(Entries(saved), entry => entry.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) == true);
        string types = TestFiles.ReadPart(saved, "[Content_Types].xml");
        Assert.DoesNotContain("macroEnabled", types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject", TestFiles.ReadPart(saved, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_macro_sheet_takes_its_tab_and_Auto_Open_and_moves_the_other_positions()
    {
        ActiveWorkbook source = new ActiveWorkbook().AddMacroSheet().AddQueryTable();
        source.BookViews = "<bookViews><workbookView activeTab=\"1\"/></bookViews>";
        MemoryStream saved = await RemoveAsync(source);

        Assert.DoesNotContain(Entries(saved), entry => entry.Contains("macrosheets", StringComparison.OrdinalIgnoreCase) == true);
        string workbook = TestFiles.ReadPart(saved, "xl/workbook.xml");
        Assert.DoesNotContain("Makro1", workbook, StringComparison.Ordinal);
        Assert.DoesNotContain("Auto_Open", workbook, StringComparison.Ordinal);

        // Data was the second tab and is now the first: the active tab and the sheet's own name move with it.
        Assert.Contains("activeTab=\"0\"", workbook, StringComparison.Ordinal);
        Assert.Contains("name=\"ExternalData_1\" localSheetId=\"0\"", workbook, StringComparison.Ordinal);
        Assert.DoesNotContain("xlMacrosheet", TestFiles.ReadPart(saved, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_dialog_sheet_takes_its_drawing_too()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddDialogSheet());

        string[] entries = Entries(saved);
        Assert.DoesNotContain(entries, entry => entry.Contains("dialogsheets", StringComparison.OrdinalIgnoreCase) == true);
        Assert.DoesNotContain("xl/drawings/vmlDrawing9.vml", entries);
        Assert.DoesNotContain("Dialog1", TestFiles.ReadPart(saved, "xl/workbook.xml"), StringComparison.Ordinal);
        Assert.DoesNotContain("dialogsheet", TestFiles.ReadPart(saved, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_an_ActiveX_control_keeps_the_form_button_next_to_it()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddActiveXControl().AddFormControl());

        string[] entries = Entries(saved);
        Assert.DoesNotContain(entries, entry => entry.Contains("activeX", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains("xl/ctrlProps/ctrlProp1.xml", entries);
        Assert.DoesNotContain("xl/media/image1.emf", entries);

        string sheet = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("CommandButton1", sheet, StringComparison.Ordinal);
        Assert.Contains("name=\"Button 2\"", sheet, StringComparison.Ordinal);
        Assert.DoesNotContain("/control\"", TestFiles.ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"), StringComparison.Ordinal);

        // The shape Excel drew for the control goes from both drawings; the button's stays.
        string vml = TestFiles.ReadPart(saved, "xl/drawings/vmlDrawing1.vml");
        Assert.DoesNotContain("_x0000_s1025", vml, StringComparison.Ordinal);
        Assert.Contains("_x0000_s1026", vml, StringComparison.Ordinal);
        string drawing = TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml");
        Assert.DoesNotContain("_x0000_s1025", drawing, StringComparison.Ordinal);
        Assert.Contains("_x0000_s1026", drawing, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_an_OLE_object_takes_the_object_its_icon_and_its_shapes()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddOleObject());

        string[] entries = Entries(saved);
        Assert.DoesNotContain(entries, entry => entry.Contains("embeddings", StringComparison.OrdinalIgnoreCase) == true);
        Assert.DoesNotContain("xl/media/image1.emf", entries);
        string sheet = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("oleObject", sheet, StringComparison.Ordinal);
        Assert.Contains("legacyDrawing", sheet, StringComparison.Ordinal);
        Assert.DoesNotContain("_x0000_s1025", TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml"), StringComparison.Ordinal);
        Assert.DoesNotContain("oleObject", TestFiles.ReadPart(saved, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_workbook_link_keeps_the_values_its_formulas_last_had()
    {
        ActiveWorkbook source = new ActiveWorkbook().AddExternalWorkbookLink();
        await using MemoryStream file = source.Build();
        Workbook book = await OpenAsync(file);
        Assert.Equal("[1]Blad1!A1*2", book.Sheets[0]["B1"].Formula);

        Assert.True(book.RemoveActiveContent().IsSuccess(out string? error, out _), error);

        // The model changes at once, so what is read now is what is saved.
        Assert.Null(book.Sheets[0]["B1"].Formula);
        Assert.Equal((CellValue)84, book.Sheets[0]["B1"].Value);
        MemoryStream saved = await SaveAndCheckAsync(book);
        Assert.DoesNotContain(Entries(saved), entry => entry.Contains("externalLink", StringComparison.OrdinalIgnoreCase) == true);
        string workbook = TestFiles.ReadPart(saved, "xl/workbook.xml");
        Assert.DoesNotContain("externalReference", workbook, StringComparison.Ordinal);
        Assert.Contains("<definedName name=\"Rate\">#REF!</definedName>", workbook, StringComparison.Ordinal);
        Assert.DoesNotContain("[1]", TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_dde_link_leaves_the_error_the_cell_showed()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddDdeLink());

        string sheet = TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml");
        Assert.DoesNotContain("<f>", sheet, StringComparison.Ordinal);
        Assert.Contains("#N/A", sheet, StringComparison.Ordinal);
        Assert.DoesNotContain(Entries(saved), entry => entry.Contains("externalLink", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public async Task Removing_a_linked_picture_leaves_the_picture_without_its_link()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddLinkedPicture().AddHyperlink());

        Assert.DoesNotContain("tracker.example", TestFiles.ReadPart(saved, "xl/drawings/_rels/drawing1.xml.rels"), StringComparison.Ordinal);
        string drawing = TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml");
        Assert.DoesNotContain("r:link", drawing, StringComparison.Ordinal);
        Assert.Contains("Picture 1", drawing, StringComparison.Ordinal);

        // A hyperlink is only followed when it is clicked, and stays.
        Assert.Contains("https://example.com/", TestFiles.ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_connection_takes_connections_xml()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddConnection());

        Assert.DoesNotContain("xl/connections.xml", Entries(saved));
        Assert.DoesNotContain("connections", TestFiles.ReadPart(saved, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
        Assert.DoesNotContain("connections", TestFiles.ReadPart(saved, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_a_query_table_leaves_an_ordinary_table_with_the_same_data()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddQueryTable());

        string[] entries = Entries(saved);
        Assert.DoesNotContain(entries, entry => entry.Contains("queryTable", StringComparison.OrdinalIgnoreCase) == true);
        string table = TestFiles.ReadPart(saved, "xl/tables/table1.xml");
        Assert.DoesNotContain("queryTable", table, StringComparison.Ordinal);
        Assert.DoesNotContain("connectionId", table, StringComparison.Ordinal);
        Assert.Contains("ref=\"A1:A3\"", table, StringComparison.Ordinal);
        Assert.DoesNotContain("queryTable", TestFiles.ReadPart(saved, "xl/tables/_rels/table1.xml.rels"), StringComparison.Ordinal);
        saved.Position = 0;
        Workbook clean = await OpenAsync(saved);
        Assert.Equal((CellValue)2, clean.Sheets[0]["A3"].Value);
    }

    [Fact]
    public async Task Removing_a_Power_Query_query_takes_the_DataMashup_and_its_properties()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddDataMashup());

        Assert.DoesNotContain(Entries(saved), entry => entry.StartsWith("customXml", StringComparison.OrdinalIgnoreCase) == true);
        Assert.DoesNotContain("customXml", TestFiles.ReadPart(saved, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
        Assert.DoesNotContain("itemProps", TestFiles.ReadPart(saved, "[Content_Types].xml"), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Removing_everything_at_once_leaves_a_workbook_the_schema_validator_accepts()
    {
        ActiveWorkbook source = new ActiveWorkbook()
            .AddVbaProject().AddMacroSheet().AddDialogSheet().AddActiveXControl().AddFormControl().AddOleObject()
            .AddExternalWorkbookLink().AddDdeLink().AddLinkedPicture().AddHyperlink().AddDataMashup().AddQueryTable();

        await using MemoryStream file = source.Build();
        Workbook book = await OpenAsync(file);
        Assert.Equal(ActiveContent.Macros | ActiveContent.ActiveX | ActiveContent.EmbeddedObjects | ActiveContent.ExternalReferences | ActiveContent.DataConnections, book.ActiveContent);

        Assert.True(book.RemoveActiveContent().IsSuccess(out string? error, out _), error);
        MemoryStream saved = await SaveAndCheckAsync(book);

        Assert.Equal(["Data"], (await OpenAsync(saved)).Sheets.Select(sheet => sheet.Name));
    }

    [Fact]
    public async Task A_part_nothing_points_to_is_left_out()
    {
        MemoryStream saved = await RemoveAsync(new ActiveWorkbook().AddOleObject().AddStrayPart("xl/payload.exe"));

        Assert.DoesNotContain("xl/payload.exe", Entries(saved));
    }

    [Fact]
    public async Task A_pivot_table_on_a_connection_cannot_be_cleaned()
    {
        await using MemoryStream file = new ActiveWorkbook().AddPivotCacheOnConnection().Build();
        Workbook book = await OpenAsync(file);

        Assert.True(book.RemoveActiveContent().IsFailure(out string? error, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Contains("pivot table", error, StringComparison.Ordinal);
        Assert.Equal(ActiveContent.DataConnections, book.ActiveContent);
    }

    [Fact]
    public async Task A_chart_drawn_from_another_workbook_cannot_be_cleaned()
    {
        await using MemoryStream file = new ActiveWorkbook().AddChartOnExternalLink().Build();
        Workbook book = await OpenAsync(file);

        Assert.True(book.RemoveActiveContent().IsFailure(out string? error, out _));
        Assert.Contains("chart", error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Without_RemoveActiveContent_a_save_keeps_the_macro_sheet_in_its_place()
    {
        ActiveWorkbook source = new ActiveWorkbook().AddMacroSheet().AddQueryTable();
        await using MemoryStream file = source.Build();
        Workbook book = await OpenAsync(file);
        book.Sheets[0]["B1"] = "changed";
        await using MemoryStream saved = new();

        await book.SaveAsync(saved);

        // The model holds only Data, but the sheet list still has Makro1 first, so localSheetId 1 is still Data.
        string workbook = TestFiles.ReadPart(saved, "xl/workbook.xml");
        Assert.True(workbook.IndexOf("Makro1", StringComparison.Ordinal) < workbook.IndexOf("name=\"Data\"", StringComparison.Ordinal));
        Assert.Contains("localSheetId=\"1\"", workbook, StringComparison.Ordinal);
        Assert.Contains("Auto_Open", workbook, StringComparison.Ordinal);
        saved.Position = 0;
        Assert.Equal(ActiveContent.Macros | ActiveContent.DataConnections, (await OpenAsync(saved)).ActiveContent);
    }

    [Fact]
    public async Task A_name_on_a_sheet_after_a_macro_sheet_belongs_to_that_sheet_in_formulas()
    {
        ActiveWorkbook source = new ActiveWorkbook().AddMacroSheet().AddQueryTable();
        await using MemoryStream file = source.Build();
        Workbook book = await OpenAsync(file);

        Assert.True(book.Evaluate("SUM(ExternalData_1)").IsSuccess(out CellValue value));
        Assert.Equal((CellValue)3, value);
    }

    private static async Task<Workbook> OpenAsync(MemoryStream file)
    {
        file.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        return book;
    }

    private static async Task<ActiveContent> StreamingAsync(MemoryStream file)
    {
        file.Position = 0;
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(file, true);
        Assert.True(opened.IsSuccess(out XlsxReader? reader), opened.Error);

        await using (reader)
        {
            return await reader.GetActiveContentAsync();
        }
    }

    /// <summary>Opens the file, removes the active content, saves, and checks the result (see <see cref="SaveAndCheckAsync"/>).</summary>
    private static async Task<MemoryStream> RemoveAsync(ActiveWorkbook source)
    {
        await using MemoryStream file = source.Build();
        Workbook book = await OpenAsync(file);
        Assert.NotEqual(ActiveContent.None, book.ActiveContent);

        Assert.True(book.RemoveActiveContent().IsSuccess(out string? error, out _), error);

        Assert.Equal(ActiveContent.None, book.ActiveContent);

        return await SaveAndCheckAsync(book);
    }

    /// <summary>Saves, and checks that the saved file has no active content, opens in both readers, and passes Microsoft's schema validator.</summary>
    private static async Task<MemoryStream> SaveAndCheckAsync(Workbook book)
    {
        MemoryStream saved = new();
        await book.SaveAsync(saved);
        Workbook clean = await OpenAsync(saved);

        Assert.Equal(ActiveContent.None, clean.ActiveContent);
        Assert.Empty(clean.ExternalReferences);
        Assert.Equal(ActiveContent.None, await StreamingAsync(saved));
        Assert.Empty(TestFiles.Validate(saved));
        saved.Position = 0;

        return saved;
    }

    private static string[] Entries(MemoryStream zip)
    {
        zip.Position = 0;
        using ZipArchive archive = new(zip, ZipArchiveMode.Read, true);

        return [.. archive.Entries.Select(entry => entry.FullName)];
    }
}
