using System.IO.Compression;
using System.Xml;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl;

/// <summary>
/// Reads an .xlsx file. Opening reads the list of sheets, the shared text and the number
/// formats; the rows are read one at a time with <see cref="SheetReader.ReadRowsAsync"/>, so
/// memory use stays small however many rows a sheet has.
/// </summary>
/// <example>
/// <code>
/// Result&lt;XlsxReader&gt; opened = await XlsxReader.OpenAsync("customers.xlsx");
///
/// if (opened.IsFailure(out string? error, out _))
/// {
///     Console.WriteLine(error);   // not an xlsx, password-protected, not found ...
///     return;
/// }
///
/// opened.IsSuccess(out XlsxReader? reader);
/// await using (reader)
/// {
///     if (reader.Sheet("Customers").IsSuccess(out SheetReader? sheet))
///     {
///         await foreach (Row row in sheet.ReadRowsAsync())
///         {
///             if (row.TryGetNumber(2, out double amount)) { ... }
///         }
///     }
/// }
/// </code>
/// </example>
public sealed class XlsxReader : IAsyncDisposable
{
    private static readonly byte[] ZipSignature = [0x50, 0x4B, 0x03, 0x04];
    private static readonly byte[] OleSignature = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    private readonly Stream stream;
    private readonly bool leaveOpen;
    private readonly ZipArchive archive;
    private readonly List<SheetPart> sheets;
    private readonly IReadOnlyList<string> sharedStrings;
    private readonly StyleTable styles;

    private XlsxReader(Stream stream, bool leaveOpen, ZipArchive archive, List<SheetPart> sheets, IReadOnlyList<string> sharedStrings, StyleTable styles, bool uses1904Dates, string workbookPath)
    {
        this.stream = stream;
        this.leaveOpen = leaveOpen;
        this.archive = archive;
        this.sheets = sheets;
        this.sharedStrings = sharedStrings;
        this.styles = styles;
        Uses1904Dates = uses1904Dates;
        Sheets = sheets.ConvertAll(sheet => sheet.Info);
        WorkbookPath = workbookPath;
    }

    /// <summary>The worksheets, in tab order. Chart sheets are not listed.</summary>
    public IReadOnlyList<SheetInfo> Sheets { get; }

    /// <summary>True when the workbook counts dates from 1904, as old Mac files do. Dates are converted either way.</summary>
    public bool Uses1904Dates { get; }

    internal ZipArchive Archive => archive;

    internal IReadOnlyList<string> SharedStrings => sharedStrings;

    internal StyleTable Styles => styles;

    internal string WorkbookPath { get; }

    internal IReadOnlyList<SheetPart> SheetParts => sheets;


    /// <summary>Opens a file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>
    /// The reader; NotFound when the file does not exist, Forbidden when it may not be read,
    /// Conflict when another program holds it locked, and Invalid when it is not an .xlsx
    /// workbook (including old .xls files and password-protected workbooks).
    /// </returns>
    public static Task<Result<XlsxReader>> OpenAsync(string path, CancellationToken cancellationToken = default)
    {
        return OpenAsync(path, KalkylLimits.Default, cancellationToken);
    }

    /// <summary>Opens a file for reading, with limits on how much it may unpack to.</summary>
    /// <param name="path">The file.</param>
    /// <param name="limits">The limits; <see cref="KalkylLimits.Default"/> unless the file is trusted.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>As <see cref="OpenAsync(string, CancellationToken)"/>, and Invalid when the file goes past a limit.</returns>
    public static async Task<Result<XlsxReader>> OpenAsync(string path, KalkylLimits limits, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(limits);
        FileStream file;

        try
        {
            file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 65_536, FileOptions.Asynchronous);
        }
        catch (FileNotFoundException)
        {
            return Result<XlsxReader>.NotFound($"'{path}' does not exist.");
        }
        catch (DirectoryNotFoundException)
        {
            return Result<XlsxReader>.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<XlsxReader>.Forbidden($"'{path}' may not be read.");
        }
        catch (IOException exception)
        {
            return Result<XlsxReader>.Conflict($"'{path}' could not be opened: {exception.Message}");
        }

        Result<XlsxReader> opened = await OpenAsync(file, limits, false, cancellationToken);

        if (opened.IsFailure(out _, out _) == true)
        {
            await file.DisposeAsync();
        }

        return opened;
    }

    /// <summary>
    /// Reads from a stream. A stream that cannot seek is copied to memory first, because the
    /// zip directory sits at the end of the file.
    /// </summary>
    /// <param name="input">The stream.</param>
    /// <param name="leaveOpen">Leave the stream open when the reader is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The reader, or Invalid when the content is not an .xlsx workbook.</returns>
    public static Task<Result<XlsxReader>> OpenAsync(Stream input, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        return OpenAsync(input, KalkylLimits.Default, leaveOpen, cancellationToken);
    }

    /// <summary>Reads from a stream, with limits on how much the content may unpack to.</summary>
    /// <param name="input">The stream.</param>
    /// <param name="limits">The limits; <see cref="KalkylLimits.Default"/> unless the content is trusted.</param>
    /// <param name="leaveOpen">Leave the stream open when the reader is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The reader, or Invalid when the content is not an .xlsx workbook or goes past a limit.</returns>
    public static async Task<Result<XlsxReader>> OpenAsync(Stream input, KalkylLimits limits, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(limits);
        Stream source = input;
        bool ownsSource = false;

        if (input.CanSeek == false)
        {
            MemoryStream copy = new();

            // An endless stream must not fill memory either; the zip file is never larger than what it unpacks to.
            if (await CopyLimitedAsync(input, copy, limits.MaxUncompressedBytes, cancellationToken) == false)
            {
                await copy.DisposeAsync();

                return Result<XlsxReader>.Invalid($"The stream is longer than {limits.MaxUncompressedBytes:N0} bytes, the limit set in KalkylLimits.MaxUncompressedBytes.");
            }

            copy.Position = 0;
            source = copy;
            ownsSource = true;
        }

        Result<XlsxReader> opened = await OpenSeekableAsync(source, ownsSource || leaveOpen == false, limits, cancellationToken);

        if (ownsSource == true && leaveOpen == false)
        {
            await input.DisposeAsync();
        }

        if (opened.IsFailure(out _, out _) == true && ownsSource == true)
        {
            await source.DisposeAsync();
        }

        return opened;
    }

    /// <summary>Gets a worksheet by name. Excel compares sheet names without regard to case, and so does this.</summary>
    /// <param name="name">The name on the tab.</param>
    /// <returns>The sheet, or NotFound.</returns>
    public Result<SheetReader> Sheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        int index = sheets.FindIndex(sheet => string.Equals(sheet.Info.Name, name, StringComparison.OrdinalIgnoreCase) == true);

        if (index < 0)
        {
            string names = string.Join(", ", sheets.Select(sheet => "'" + sheet.Info.Name + "'"));

            return Result<SheetReader>.NotFound($"There is no sheet named '{name}'. The workbook has {names}.");
        }

        return Sheet(index);
    }

    /// <summary>Gets a worksheet by position.</summary>
    /// <param name="index">The position, 0-based, among the worksheets.</param>
    /// <returns>The sheet, NotFound when there are fewer sheets, or Invalid when the file lacks the sheet's part.</returns>
    public Result<SheetReader> Sheet(int index)
    {
        if (index < 0 || index >= sheets.Count)
        {
            return Result<SheetReader>.NotFound($"There is no sheet {index}. The workbook has {sheets.Count} worksheet(s), numbered from 0.");
        }

        (SheetInfo info, string path, _, _) = sheets[index];

        if (XmlParts.TryFindEntry(archive, path, out ZipArchiveEntry? entry) == false)
        {
            return Result<SheetReader>.Invalid($"The file is damaged: the sheet '{info.Name}' points to '{path}', which is not in the file.");
        }

        return new SheetReader(info, entry, sharedStrings, styles, Uses1904Dates);
    }

    /// <summary>Closes the file.</summary>
    /// <returns>A task that completes when the file is closed.</returns>
    public async ValueTask DisposeAsync()
    {
        await archive.DisposeAsync();

        if (leaveOpen == false)
        {
            await stream.DisposeAsync();
        }
    }

    /// <summary>Copies at most <paramref name="limit"/> bytes; false when the source has more.</summary>
    internal static async Task<bool> CopyLimitedAsync(Stream source, Stream target, long limit, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81_920];
        long total = 0;
        int read = await source.ReadAsync(buffer, cancellationToken);

        while (read > 0)
        {
            total += read;

            if (total > limit)
            {
                return false;
            }

            await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            read = await source.ReadAsync(buffer, cancellationToken);
        }

        return true;
    }

    private static async Task<Result<XlsxReader>> OpenSeekableAsync(Stream source, bool disposeSource, KalkylLimits limits, CancellationToken cancellationToken)
    {
        Result signature = await CheckSignatureAsync(source, cancellationToken);

        if (signature.IsFailure(out string? signatureError, out ResultStatus signatureStatus) == true)
        {
            return Result<XlsxReader>.Failure(signatureError, signatureStatus);
        }

        ZipArchive archive;

        try
        {
            archive = await ZipArchive.CreateAsync(source, ZipArchiveMode.Read, true, null, cancellationToken);
        }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException)
        {
            return Result<XlsxReader>.Invalid($"The file is not a valid .xlsx workbook: the zip container is damaged ({exception.Message}).");
        }

        Result size = XmlParts.CheckSize(archive, limits);

        if (size.IsFailure(out string? sizeError, out ResultStatus sizeStatus) == true)
        {
            await archive.DisposeAsync();

            return Result<XlsxReader>.Failure(sizeError, sizeStatus);
        }

        Result entries = await XmlParts.CheckEntriesAsync(archive, cancellationToken);

        if (entries.IsFailure(out string? entriesError, out ResultStatus entriesStatus) == true)
        {
            await archive.DisposeAsync();

            return Result<XlsxReader>.Failure(entriesError, entriesStatus);
        }

        try
        {
            Result<XlsxReader> loaded = await LoadAsync(source, disposeSource == false, archive, limits, cancellationToken);

            if (loaded.IsFailure(out _, out _) == true)
            {
                await archive.DisposeAsync();
            }

            return loaded;
        }
        catch (XmlException exception)
        {
            await archive.DisposeAsync();

            return Result<XlsxReader>.Invalid($"The file is not a valid .xlsx workbook: a part has broken XML ({exception.Message}).");
        }
        catch (InvalidDataException exception)
        {
            await archive.DisposeAsync();

            return Result<XlsxReader>.Invalid(KalkylError.Describe(exception));
        }
    }

    private static async Task<Result> CheckSignatureAsync(Stream source, CancellationToken cancellationToken)
    {
        byte[] header = new byte[8];
        long start = source.Position;
        int read = await source.ReadAtLeastAsync(header, header.Length, false, cancellationToken);
        source.Position = start;

        if (read >= 8 && header.AsSpan().SequenceEqual(OleSignature) == true)
        {
            return Result.Invalid("The file is not an .xlsx workbook. It is either an old .xls workbook or a password-protected .xlsx; neither can be read. Save it as an unprotected .xlsx in Excel.");
        }

        if (read < 4 || header.AsSpan(0, 4).SequenceEqual(ZipSignature) == false)
        {
            return Result.Invalid("The file is not an .xlsx workbook.");
        }

        return Result.Success();
    }

    private static async Task<Result<XlsxReader>> LoadAsync(Stream source, bool leaveOpen, ZipArchive archive, KalkylLimits limits, CancellationToken cancellationToken)
    {
        Dictionary<string, (string Type, string Target)> rootRelationships = await XmlParts.ReadRelationshipsAsync(archive, "_rels/.rels", cancellationToken);
        string workbookPath = "xl/workbook.xml";

        foreach ((string type, string target) in rootRelationships.Values)
        {
            if (Ooxml.IsRelationship(type, "officeDocument") == true)
            {
                workbookPath = Ooxml.ResolveTarget(string.Empty, target);
            }
        }

        if (XmlParts.TryFindEntry(archive, workbookPath, out ZipArchiveEntry? workbookEntry) == false)
        {
            return Result<XlsxReader>.Invalid("The file is not an .xlsx workbook: it has no workbook part. It may be a Word or PowerPoint file.");
        }

        Dictionary<string, (string Type, string Target)> workbookRelationships = await XmlParts.ReadRelationshipsAsync(archive, Ooxml.RelationshipsPath(workbookPath), cancellationToken);
        (List<(string Name, string RelationshipId, int SheetId, SheetVisibility Visibility)> declared, bool uses1904) = await ReadWorkbookAsync(workbookEntry, cancellationToken);
        List<SheetPart> sheets = [];

        foreach ((string name, string relationshipId, int sheetId, SheetVisibility visibility) in declared)
        {
            if (workbookRelationships.TryGetValue(relationshipId, out (string Type, string Target) relationship) == false)
            {
                continue;
            }

            if (Ooxml.IsRelationship(relationship.Type, "worksheet") == false)
            {
                continue;
            }

            string path = Ooxml.ResolveTarget(workbookPath, relationship.Target);
            SheetPart? sharing = sheets.Find(sheet => string.Equals(sheet.Path, path, StringComparison.OrdinalIgnoreCase) == true);

            // Two sheets on one part: an edit to one would change the other, and saving wrote
            // the part twice, which made a zip file no program can read.
            if (sharing is not null)
            {
                return Result<XlsxReader>.Invalid($"The file is damaged: the sheets '{sharing.Info.Name}' and '{name}' point to the same part, '{path}'.");
            }

            sheets.Add(new SheetPart(new SheetInfo(name, sheets.Count, visibility), path, relationshipId, sheetId));
        }

        IReadOnlyList<string> sharedStrings = [];
        StyleTable styles = StyleTable.Empty;

        foreach ((string type, string target) in workbookRelationships.Values)
        {
            string path = Ooxml.ResolveTarget(workbookPath, target);

            if (Ooxml.IsRelationship(type, "sharedStrings") == true && XmlParts.TryFindEntry(archive, path, out ZipArchiveEntry? stringsEntry) == true)
            {
                sharedStrings = await ReadSharedStringsAsync(stringsEntry, limits.MaxSharedStrings, cancellationToken);
            }
            else if (Ooxml.IsRelationship(type, "styles") == true && XmlParts.TryFindEntry(archive, path, out ZipArchiveEntry? stylesEntry) == true)
            {
                styles = await StyleTable.ReadAsync(stylesEntry, cancellationToken);
            }
        }

        return new XlsxReader(source, leaveOpen, archive, sheets, sharedStrings, styles, uses1904, workbookPath);
    }

    private static async Task<(List<(string Name, string RelationshipId, int SheetId, SheetVisibility Visibility)> Sheets, bool Uses1904)> ReadWorkbookAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        List<(string, string, int, SheetVisibility)> sheets = [];
        bool uses1904 = false;
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        // Only the sheets in <sheets>, the workbook's own child: the saver rewrites that element and
        // copies the others, so a <sheet> anywhere else (found by libFuzzer in a <sheetz>) came back
        // as a second copy of every sheet.
        string section = string.Empty;

        while (await reader.ReadAsync() == true)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
            {
                section = reader.LocalName;
            }

            if (reader.Depth == 1 && XmlParts.IsElement(reader, "workbookPr") == true)
            {
                string? date1904 = reader.GetAttribute("date1904");
                uses1904 = date1904 is "1" or "true";

                continue;
            }

            if (section == "sheets" && reader.Depth == 2 && XmlParts.IsElement(reader, "sheet") == true)
            {
                string? name = reader.GetAttribute("name");

                if (name is null || XmlParts.TryGetRelationshipId(reader, out string? relationshipId) == false)
                {
                    continue;
                }

                SheetVisibility visibility = reader.GetAttribute("state") switch
                {
                    "hidden" => SheetVisibility.Hidden,
                    "veryHidden" => SheetVisibility.VeryHidden,
                    _ => SheetVisibility.Visible,
                };

                int.TryParse(reader.GetAttribute("sheetId"), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int sheetId);
                sheets.Add((name, relationshipId, sheetId, visibility));
            }
        }

        return (sheets, uses1904);
    }

    private static async Task<IReadOnlyList<string>> ReadSharedStringsAsync(ZipArchiveEntry entry, long limit, CancellationToken cancellationToken)
    {
        List<string> strings = [];
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "si") == true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (strings.Count >= limit)
                {
                    throw new InvalidDataException($"The file has more than {limit:N0} shared strings, the limit set in KalkylLimits.MaxSharedStrings.");
                }

                strings.Add(await XmlParts.ReadRichTextAsync(reader));
            }
        }

        return strings;
    }
}

/// <summary>A worksheet's entry in the workbook: its info, part path, relationship id and sheet id.</summary>
internal sealed record SheetPart(SheetInfo Info, string Path, string RelationshipId, int SheetId);
