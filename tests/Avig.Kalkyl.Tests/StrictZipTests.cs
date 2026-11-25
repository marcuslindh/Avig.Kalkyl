using System.Buffers.Binary;
using System.Text;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Zip files that say two things (docs/testplan.md, S-10): an entry that only the local
/// headers have, which a scanner that walks the headers sees and Excel does not, and bytes after
/// the end record. Every way the package opens a file refuses them as damaged, unless
/// KalkylLimits.StrictZip is turned off.
/// </summary>
public class StrictZipTests
{
    public static TheoryData<string> Hostile()
    {
        return ["hidden-entry", "trailing-bytes"];
    }

    public static TheoryData<string, string> HostileByPath()
    {
        TheoryData<string, string> data = [];

        foreach (string file in new[] { "hidden-entry", "trailing-bytes" })
        {
            foreach (string path in new[] { "reader-stream", "reader-unseekable", "reader-file", "workbook-stream", "workbook-file", "typed-file" })
            {
                data.Add(file, path);
            }
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(HostileByPath))]
    public async Task Every_way_in_refuses_a_zip_that_says_two_things(string file, string path)
    {
        byte[] bytes = await HostileAsync(file);

        (string error, KalkylErrorKind kind) = await OpenAsync(bytes, path);

        Assert.Equal(KalkylErrorKind.Damaged, kind);
        Assert.StartsWith("The file is damaged: the zip file does not say one thing only", error, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Hostile))]
    public async Task With_StrictZip_off_the_same_file_opens(string file)
    {
        byte[] bytes = await HostileAsync(file);
        KalkylLimits lax = KalkylLimits.Default with { StrictZip = false };

        Result<XlsxReader> reader = await XlsxReader.OpenAsync(new MemoryStream(bytes), lax);
        Result<Workbook> book = await Workbook.OpenAsync(new MemoryStream(bytes), lax);

        Assert.True(reader.IsSuccess(out XlsxReader? opened), reader.Error);
        await opened.DisposeAsync();
        Assert.True(book.IsSuccess(out Workbook? workbook), book.Error);
        Assert.Equal("Data", workbook.Sheets[0].Name);
    }

    [Fact]
    public async Task A_file_the_package_writes_and_one_it_saves_pass_the_check()
    {
        await using MemoryStream written = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["x", 1]);
        });

        Assert.True((await Workbook.OpenAsync(written)).IsSuccess(out Workbook? book));
        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        Assert.True((await Workbook.OpenAsync(saved)).IsSuccess(out _));
        saved.Position = 0;
        Assert.True((await Avig.OpenXml.OpcZip.CheckStructureAsync(saved, int.MaxValue, CancellationToken.None)).IsSuccess(out string? error, out _), error);
    }

    [Fact]
    public void StrictZip_is_on_by_default_and_in_Unlimited()
    {
        Assert.True(KalkylLimits.Default.StrictZip);
        Assert.True(KalkylLimits.Unlimited.StrictZip);
    }

    /// <summary>A workbook made into a zip that says two things.</summary>
    private static async Task<byte[]> HostileAsync(string kind)
    {
        await using MemoryStream file = new ActiveWorkbook().Build();
        byte[] bytes = file.ToArray();

        return kind switch
        {
            "hidden-entry" => WithHiddenEntry(bytes, "xl/hidden.bin", Encoding.ASCII.GetBytes("payload only a header walker sees")),
            "trailing-bytes" => [.. bytes, .. Encoding.ASCII.GetBytes("appended after the end record")],
            _ => throw new ArgumentException(kind, nameof(kind)),
        };
    }

    /// <summary>
    /// Puts a stored entry with a local header just before the central directory and moves the
    /// directory's offset past it. The directory does not list it, so .NET and Excel do not see it;
    /// a program that walks the local headers does.
    /// </summary>
    private static byte[] WithHiddenEntry(byte[] zip, string name, byte[] content)
    {
        int end = zip.Length - 22;

        while (BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(end)) != 0x06054b50)
        {
            end--;
        }

        int directory = (int)BinaryPrimitives.ReadUInt32LittleEndian(zip.AsSpan(end + 16));
        byte[] nameBytes = Encoding.UTF8.GetBytes(name);
        byte[] header = new byte[30 + nameBytes.Length];
        BinaryPrimitives.WriteUInt32LittleEndian(header, 0x04034b50);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(4), 20);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(18), (uint)content.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(header.AsSpan(22), (uint)content.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(header.AsSpan(26), (ushort)nameBytes.Length);
        nameBytes.CopyTo(header, 30);
        byte[] result = [.. zip.AsSpan(0, directory), .. header, .. content, .. zip.AsSpan(directory)];
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(end + header.Length + content.Length + 16), (uint)(directory + header.Length + content.Length));

        return result;
    }

    private static async Task<(string Error, KalkylErrorKind Kind)> OpenAsync(byte[] bytes, string path)
    {
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        await File.WriteAllBytesAsync(file, bytes);

        try
        {
            switch (path)
            {
                case "reader-stream":
                    return Failure(await XlsxReader.OpenAsync(new MemoryStream(bytes)));

                case "reader-unseekable":
                    return Failure(await XlsxReader.OpenAsync(new UnseekableStream(bytes)));

                case "reader-file":
                    return Failure(await XlsxReader.OpenAsync(file));

                case "workbook-stream":
                    return Failure(await Workbook.OpenAsync(new MemoryStream(bytes)));

                case "workbook-file":
                    return Failure(await Workbook.OpenAsync(file));

                default:
                    await foreach (Result<Customer> row in Xlsx.ReadAsync<Customer>(file))
                    {
                        return Failure(row);
                    }

                    throw new InvalidOperationException("The typed reader gave no result.");
            }
        }
        finally
        {
            File.Delete(file);
        }
    }

    private static (string Error, KalkylErrorKind Kind) Failure<T>(Result<T> result)
    {
        Assert.True(result.IsFailure(out string? error, out _), $"The file was opened.");

        return (error, result.ErrorKind());
    }

    /// <summary>A stream that cannot seek, like a network upload, which the reader copies to memory first.</summary>
    private sealed class UnseekableStream(byte[] bytes) : MemoryStream(bytes)
    {
        public override bool CanSeek => false;
    }
}
