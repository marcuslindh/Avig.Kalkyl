using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Avig.Kalkyl.Tests;

/// <summary>Hjälp för att skriva, läsa och bygga xlsx-filer i minnet.</summary>
internal static class TestFiles
{
    public static async Task<MemoryStream> WriteAsync(Func<XlsxWriter, Task> write)
    {
        MemoryStream stream = new();
        XlsxWriter writer = await XlsxWriter.CreateAsync(stream, true);
        await write(writer);
        await writer.DisposeAsync();
        stream.Position = 0;

        return stream;
    }

    public static async Task<SheetWriter> AddSheetAsync(XlsxWriter writer, string name)
    {
        Result<SheetWriter> added = await writer.AddSheetAsync(name);
        Assert.True(added.IsSuccess(out SheetWriter? sheet), added.Error);

        return sheet;
    }

    public static async Task<XlsxReader> OpenAsync(Stream stream)
    {
        stream.Position = 0;
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(stream, true);
        Assert.True(opened.IsSuccess(out XlsxReader? reader), opened.Error);

        return reader;
    }

    public static async Task<List<Row>> ReadRowsAsync(Stream stream, int sheetIndex = 0)
    {
        await using XlsxReader reader = await OpenAsync(stream);
        Result<SheetReader> found = reader.Sheet(sheetIndex);
        Assert.True(found.IsSuccess(out SheetReader? sheet), found.Error);
        List<Row> rows = [];

        await foreach (Row row in sheet.ReadRowsAsync())
        {
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>Bygger ett paket av delar, som en fil från ett annat program.</summary>
    public static MemoryStream Package(params (string Path, string Content)[] parts)
    {
        MemoryStream stream = new();

        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, true))
        {
            foreach ((string path, string content) in parts)
            {
                ZipArchiveEntry entry = archive.CreateEntry(path);
                using Stream entryStream = entry.Open();
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                entryStream.Write(bytes, 0, bytes.Length);
            }
        }

        stream.Position = 0;

        return stream;
    }

    public static string ReadPart(Stream stream, string path)
    {
        stream.Position = 0;
        using ZipArchive archive = new(stream, ZipArchiveMode.Read, true);
        ZipArchiveEntry entry = archive.GetEntry(path) ?? throw new InvalidOperationException($"{path} saknas i paketet.");
        using StreamReader reader = new(entry.Open(), Encoding.UTF8);
        string content = reader.ReadToEnd();
        stream.Position = 0;

        return content;
    }

    /// <summary>Microsofts schemavalidator. En tom lista betyder att Excel kan öppna filen utan att reparera den.</summary>
    public static List<string> Validate(Stream stream)
    {
        stream.Position = 0;
        List<string> problems = [];

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
        {
            OpenXmlValidator validator = new();

            foreach (ValidationErrorInfo error in validator.Validate(document))
            {
                problems.Add($"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}");
            }
        }

        stream.Position = 0;

        return problems;
    }
}
