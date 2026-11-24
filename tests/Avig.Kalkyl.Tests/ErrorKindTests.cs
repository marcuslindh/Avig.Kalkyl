namespace Avig.Kalkyl.Tests;

/// <summary>
/// KalkylError.ErrorKind: every way a file can be refused maps to a kind an application can show
/// its users without the message, which names parts, cells and text of the file (SEC-10 in docs/testplan.md).
/// </summary>
public class ErrorKindTests
{
    private static async Task<MemoryStream> OrdinaryFileAsync(int rows)
    {
        return await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");

            for (int row = 0; row < rows; row++)
            {
                await sheet.WriteRowAsync([Cell.Of(row), Cell.Of("text")]);
            }
        });
    }

    [Fact]
    public async Task A_success_has_no_kind()
    {
        await using MemoryStream file = await OrdinaryFileAsync(3);

        Result<Workbook> opened = await Workbook.OpenAsync(file);

        Assert.Equal(KalkylErrorKind.Unknown, opened.ErrorKind());
    }

    [Fact]
    public async Task A_missing_file_is_NotFound()
    {
        string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        Assert.Equal(KalkylErrorKind.NotFound, (await Workbook.OpenAsync(path)).ErrorKind());
        Assert.Equal(KalkylErrorKind.NotFound, (await XlsxReader.OpenAsync(path)).ErrorKind());
    }

    [Fact]
    public async Task A_missing_sheet_is_NotFound()
    {
        await using MemoryStream file = await OrdinaryFileAsync(1);
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));

        Assert.Equal(KalkylErrorKind.NotFound, book.Sheet("Nope").ErrorKind());
    }

    [Theory]
    [InlineData(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0 })]
    [InlineData(new byte[] { (byte)'a', (byte)';', (byte)'b', (byte)'\n', (byte)'1', (byte)';', (byte)'2' })]
    [InlineData(new byte[] { })]
    public async Task An_xls_a_protected_file_or_a_csv_is_NotXlsx(byte[] bytes)
    {
        await using MemoryStream file = new(bytes);

        Assert.Equal(KalkylErrorKind.NotXlsx, (await Workbook.OpenAsync(file)).ErrorKind());
        file.Position = 0;
        Assert.Equal(KalkylErrorKind.NotXlsx, (await XlsxReader.OpenAsync(file, true)).ErrorKind());
    }

    [Fact]
    public async Task A_zip_without_a_workbook_is_NotXlsx()
    {
        await using MemoryStream file = TestFiles.Package(("word/document.xml", "<document/>"));

        Assert.Equal(KalkylErrorKind.NotXlsx, (await Workbook.OpenAsync(file)).ErrorKind());
    }

    [Fact]
    public async Task A_truncated_file_is_Damaged()
    {
        await using MemoryStream whole = await OrdinaryFileAsync(50);
        await using MemoryStream file = new(whole.ToArray()[..(int)(whole.Length / 2)]);

        Assert.Equal(KalkylErrorKind.Damaged, (await Workbook.OpenAsync(file)).ErrorKind());
    }

    /// <summary>MaxSharedStrings is checked in HostileFilesTests, on a file that has shared strings.</summary>
    [Fact]
    public async Task Each_limit_is_TooLarge()
    {
        await using MemoryStream file = await OrdinaryFileAsync(200);
        KalkylLimits[] limits =
        [
            KalkylLimits.Default with { MaxUncompressedBytes = 2_000 },
            KalkylLimits.Default with { MaxCells = 100 },
        ];

        foreach (KalkylLimits limit in limits)
        {
            file.Position = 0;
            Result<Workbook> opened = await Workbook.OpenAsync(file, limit);

            Assert.True(opened.ErrorKind() == KalkylErrorKind.TooLarge, $"{limit}: {opened.Error}");
        }
    }

    [Fact]
    public async Task A_file_another_program_holds_is_Inaccessible()
    {
        string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        await using (MemoryStream file = await OrdinaryFileAsync(1))
        {
            await System.IO.File.WriteAllBytesAsync(path, file.ToArray());
        }

        try
        {
            await using FileStream held = new(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

            Assert.Equal(KalkylErrorKind.Inaccessible, (await Workbook.OpenAsync(path)).ErrorKind());
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void Exceptions_from_reading_have_kinds()
    {
        Assert.Equal(KalkylErrorKind.Damaged, new InvalidDataException("Block length does not match with its complement.").ErrorKind());
        Assert.Equal(KalkylErrorKind.TooLarge, new InvalidDataException($"The workbook has more than 5 cells, {KalkylError.LimitMarker}MaxCells.").ErrorKind());
        Assert.Equal(KalkylErrorKind.NotFound, new FileNotFoundException().ErrorKind());
        Assert.Equal(KalkylErrorKind.Unknown, new OperationCanceledException().ErrorKind());
    }

    [Fact]
    public void A_failure_from_elsewhere_is_Unknown()
    {
        Assert.Equal(KalkylErrorKind.Unknown, Result.Invalid("Something else went wrong.").ErrorKind());
    }
}
