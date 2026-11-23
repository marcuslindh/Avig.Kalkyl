using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Tests;

public class WriterTests
{
    [Theory]
    [InlineData("")]
    [InlineData("abcdefghijklmnopqrstuvwxyz123456")]
    [InlineData("a/b")]
    [InlineData("a:b")]
    [InlineData("[x]")]
    [InlineData("'citat")]
    [InlineData("citat'")]
    [InlineData("history")]
    public async Task Sheet_names_Excel_would_refuse_are_Invalid(string name)
    {
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(new MemoryStream());
        Result<SheetWriter> added = await writer.AddSheetAsync(name);

        Assert.True(added.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
    }

    [Fact]
    public async Task A_sheet_name_already_taken_regardless_of_case_is_a_Conflict()
    {
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(new MemoryStream());
        await TestFiles.AddSheetAsync(writer, "Data");
        Result<SheetWriter> again = await writer.AddSheetAsync("DATA");

        Assert.True(again.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Conflict, status);
    }

    [Fact]
    public async Task Column_widths_and_panes_must_come_before_the_first_row()
    {
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(new MemoryStream());
        SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
        await sheet.WriteRowAsync(["rad"]);

        Assert.Throws<InvalidOperationException>(() => sheet.SetColumnWidth(0, 10));
        Assert.Throws<InvalidOperationException>(() => sheet.FreezePanes(1));
    }

    [Fact]
    public async Task A_sheet_is_finished_when_the_next_one_is_added()
    {
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(new MemoryStream());
        SheetWriter first = await TestFiles.AddSheetAsync(writer, "Ett");
        await TestFiles.AddSheetAsync(writer, "Två");

        Assert.True(first.IsFinished);
        await Assert.ThrowsAsync<InvalidOperationException>(() => first.WriteRowAsync(["för sent"]));
    }

    [Fact]
    public async Task A_workbook_without_sheets_gets_an_empty_one_so_Excel_can_open_it()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(_ => Task.CompletedTask);
        await using XlsxReader reader = await TestFiles.OpenAsync(file);

        Assert.Equal("Sheet1", Assert.Single(reader.Sheets).Name);
        Assert.Empty(TestFiles.Validate(file));
    }

    [Fact]
    public async Task Equal_styles_share_one_entry_however_many_cells_use_them()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");

            for (int index = 0; index < 1000; index++)
            {
                await sheet.WriteRowAsync([Cell.Of(index).Bold(), Cell.Of("x").Bold().Background("#FFFF00"), Cell.Of(index).Bold()]);
            }
        });

        string styles = TestFiles.ReadPart(file, "xl/styles.xml");

        Assert.Matches(new Regex("<cellXfs count=\"3\">"), styles);
        Assert.Matches(new Regex("<fonts count=\"2\">"), styles);
        Assert.Matches(new Regex("<fills count=\"3\">"), styles);
    }

    [Fact]
    public async Task Frozen_panes_widths_and_autofilter_end_up_in_the_sheet()
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            sheet.SetColumnWidth(1, 25);
            sheet.FreezePanes(1);
            sheet.EnableAutoFilter();
            await sheet.WriteRowAsync(["Namn", "Ort", "Belopp"]);
            await sheet.WriteRowAsync(["Ada", "London", 1]);
        });

        string sheetXml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");
        string workbookXml = TestFiles.ReadPart(file, "xl/workbook.xml");

        Assert.Contains("<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\" />", sheetXml, StringComparison.Ordinal);
        Assert.Contains("<col min=\"2\" max=\"2\" width=\"25\" customWidth=\"1\" />", sheetXml, StringComparison.Ordinal);
        Assert.Contains("<autoFilter ref=\"A1:C2\" />", sheetXml, StringComparison.Ordinal);
        Assert.Contains("'Data'!$A$1:$C$2", workbookXml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Creating_a_file_in_a_folder_that_does_not_exist_is_NotFound()
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "fil.xlsx");
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);

        Assert.True(created.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.NotFound, status);
    }

    [Fact]
    public async Task A_stream_is_left_open_when_asked_and_closed_otherwise()
    {
        MemoryStream kept = new();
        XlsxWriter keeping = await XlsxWriter.CreateAsync(kept, true);
        await keeping.DisposeAsync();
        Assert.True(kept.CanRead);

        MemoryStream closed = new();
        XlsxWriter closing = await XlsxWriter.CreateAsync(closed);
        await closing.DisposeAsync();
        Assert.False(closed.CanRead);
    }
}
