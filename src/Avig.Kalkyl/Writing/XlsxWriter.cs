using System.Buffers;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Writing;

namespace Avig.Kalkyl;

/// <summary>
/// Writes an .xlsx file from the top down: add a sheet, write its rows, add the next sheet.
/// Memory use stays flat however many rows are written. Complete the workbook with
/// <see cref="CompleteAsync"/>, or dispose it, which completes it.
/// </summary>
/// <example>
/// <code>
/// await using XlsxWriter writer = await XlsxWriter.CreateAsync(stream);
/// Result&lt;SheetWriter&gt; added = await writer.AddSheetAsync("Customers");
///
/// if (added.IsSuccess(out SheetWriter? sheet))
/// {
///     sheet.FreezePanes(1);
///     await sheet.WriteRowAsync([Cell.Of("Name").Bold(), Cell.Of("Amount").Bold()]);
///     await sheet.WriteRowAsync(["Ada", 42.5]);
/// }
/// </code>
/// </example>
public sealed class XlsxWriter : IAsyncDisposable
{
    internal static readonly XmlWriterSettings XmlSettings = new()
    {
        Async = true,
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = false,

        // The default, Replace, writes every \n in a cell as Environment.NewLine: "a\nb" became
        // "a\r\nb" on Windows, and ClosedXML read it back that way. Entitize leaves text as it
        // is and writes line breaks in attributes as references, which a reader cannot fold away.
        NewLineHandling = NewLineHandling.Entitize,
    };

    private static readonly SearchValues<char> ForbiddenNameCharacters = SearchValues.Create(":\\/?*[]");

    private readonly Stream stream;
    private readonly bool leaveOpen;
    private readonly ZipArchive archive;
    private readonly StyleRegistry styles = new();
    private readonly List<SheetWriter> sheets = [];
    private readonly DrawingPackage drawings;
    private bool completed;

    private XlsxWriter(Stream stream, bool leaveOpen, ZipArchive archive)
    {
        this.stream = stream;
        this.leaveOpen = leaveOpen;
        this.archive = archive;
        drawings = new DrawingPackage(archive, []);
    }

    /// <summary>The sheets added so far, in order.</summary>
    public IReadOnlyList<SheetWriter> Sheets => sheets;

    /// <summary>Creates a file, replacing one that already exists.</summary>
    /// <param name="path">Where to write the file.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>
    /// The writer, or NotFound when the folder does not exist, Forbidden when the file may not be
    /// written, and Conflict when another program has it open.
    /// </returns>
    public static async Task<Result<XlsxWriter>> CreateAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        FileStream file;

        try
        {
            file = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 65_536, FileOptions.Asynchronous);
        }
        catch (DirectoryNotFoundException)
        {
            return Result<XlsxWriter>.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<XlsxWriter>.Forbidden($"'{path}' may not be written.");
        }
        catch (IOException exception)
        {
            return Result<XlsxWriter>.Conflict($"'{path}' could not be created: {exception.Message}");
        }

        return await CreateAsync(file, false, cancellationToken);
    }

    /// <summary>Writes to a stream. The stream must be writable; it does not need to be seekable.</summary>
    /// <param name="output">The stream.</param>
    /// <param name="leaveOpen">Leave the stream open when the writer is disposed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The writer.</returns>
    public static async Task<XlsxWriter> CreateAsync(Stream output, bool leaveOpen = false, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);

        if (output.CanWrite == false)
        {
            throw new ArgumentException("The stream is not writable.", nameof(output));
        }

        ZipArchive archive = await ZipArchive.CreateAsync(output, ZipArchiveMode.Create, true, null, cancellationToken);

        return new XlsxWriter(output, leaveOpen, archive);
    }

    /// <summary>Adds a sheet and finishes the one before it.</summary>
    /// <param name="name">
    /// The name shown on the tab: 1 to 31 characters, none of <c>: \ / ? * [ ]</c>, not starting
    /// or ending with an apostrophe, and different from the other sheets regardless of case.
    /// </param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The sheet, or Invalid for a name Excel would refuse.</returns>
    public async Task<Result<SheetWriter>> AddSheetAsync(string name, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(name);
        ThrowIfCompleted();
        Result validation = ValidateName(name);

        if (validation.IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result<SheetWriter>.Failure(error, status);
        }

        if (sheets.Exists(existing => string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase) == true))
        {
            return Result<SheetWriter>.Conflict($"A sheet named '{name}' already exists. Excel compares sheet names without regard to case.");
        }

        await FinishCurrentAsync(cancellationToken);

        return await AddSheetCoreAsync(name, cancellationToken);
    }

    /// <summary>Finishes the last sheet and writes the workbook parts. The writer cannot be used afterwards.</summary>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when the file is written.</returns>
    public async Task CompleteAsync(CancellationToken cancellationToken = default)
    {
        if (completed == true)
        {
            return;
        }

        await FinishCurrentAsync(cancellationToken);

        if (sheets.Count == 0)
        {
            // Excel does not open a workbook without sheets.
            SheetWriter empty = await AddSheetCoreAsync("Sheet1", cancellationToken);
            await empty.FinishAsync(cancellationToken);
        }

        await WritePartAsync("xl/workbook.xml", WriteWorkbookAsync, cancellationToken);
        await WritePartAsync("xl/_rels/workbook.xml.rels", WriteWorkbookRelationshipsAsync, cancellationToken);
        await WritePartAsync("xl/styles.xml", styles.WriteAsync, cancellationToken);
        await WritePartAsync("[Content_Types].xml", WriteContentTypesAsync, cancellationToken);
        await WritePartAsync("_rels/.rels", WriteRootRelationshipsAsync, cancellationToken);
        completed = true;
        await archive.DisposeAsync();

        if (leaveOpen == false)
        {
            await stream.DisposeAsync();
        }
    }

    /// <summary>Completes the workbook if that has not been done.</summary>
    /// <returns>A task that completes when the file is written and closed.</returns>
    public async ValueTask DisposeAsync()
    {
        await CompleteAsync(CancellationToken.None);
    }

    /// <summary>Whether Excel accepts a sheet name. Shared with the document model.</summary>
    internal static Result ValidateName(string name)
    {
        if (name.Length == 0 || name.Length > 31)
        {
            return Result.Invalid($"The sheet name '{name}' must be 1 to 31 characters long.");
        }

        if (name.AsSpan().IndexOfAny(ForbiddenNameCharacters) >= 0)
        {
            return Result.Invalid($"The sheet name '{name}' contains one of : \\ / ? * [ ], which Excel does not allow.");
        }

        if (name.StartsWith('\'') == true || name.EndsWith('\'') == true)
        {
            return Result.Invalid($"The sheet name '{name}' starts or ends with an apostrophe, which Excel does not allow.");
        }

        if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase) == true)
        {
            return Result.Invalid("'History' is reserved by Excel and cannot be a sheet name.");
        }

        return Result.Success();
    }

    private async Task<SheetWriter> AddSheetCoreAsync(string name, CancellationToken cancellationToken)
    {
        int index = sheets.Count;
        ZipArchiveEntry entry = archive.CreateEntry(SheetPath(index), CompressionLevel.Optimal);
        Stream entryStream = await entry.OpenAsync(cancellationToken);
        SheetWriter sheet = new(name, index, entryStream, styles, SheetExists);
        sheets.Add(sheet);

        return sheet;
    }

    private async Task FinishCurrentAsync(CancellationToken cancellationToken)
    {
        if (sheets.Count == 0 || sheets[^1].IsFinished == true)
        {
            return;
        }

        SheetWriter sheet = sheets[^1];
        await sheet.FinishAsync(cancellationToken);

        // The sheet's entry is closed now, so the parts its drawing points to can follow it in the zip.
        if (sheet.Extras.Drawings.Count > 0 || sheet.Extras.ExternalHyperlinks.Any() == true)
        {
            await WriteSheetPartsAsync(sheet, cancellationToken);
        }
    }

    /// <summary>A sheet's drawing with its charts and pictures, and the relationships from the sheet to the drawing and its web links.</summary>
    private async Task WriteSheetPartsAsync(SheetWriter sheet, CancellationToken cancellationToken)
    {
        string drawingPath = string.Empty;

        if (sheet.Extras.Drawings.Count > 0)
        {
            drawingPath = await drawings.WriteDrawingAsync(sheet.Extras.Drawings, cancellationToken);
        }

        List<PartRelationship> relationships = sheet.Relationships("/" + drawingPath);
        await WritePartAsync("xl/worksheets/_rels/sheet" + Text(sheet.Index + 1) + ".xml.rels", xml => DrawingXml.WriteRelationshipsAsync(xml, relationships), cancellationToken);
    }

    private bool SheetExists(string name)
    {
        return sheets.Exists(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase) == true);
    }

    private async Task WritePartAsync(string path, Func<XmlWriter, Task> write, CancellationToken cancellationToken)
    {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        await using Stream entryStream = await entry.OpenAsync(cancellationToken);
        await using XmlWriter xml = XmlWriter.Create(entryStream, XmlSettings);
        await write(xml);
        await xml.FlushAsync();
    }

    private async Task WriteWorkbookAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "workbook", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await xml.WriteStartElementAsync(null, "bookViews", null);
        await xml.WriteElementStringAsync(null, "workbookView", null, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(null, "sheets", null);

        foreach (SheetWriter sheet in sheets)
        {
            await xml.WriteStartElementAsync(null, "sheet", null);
            await xml.WriteAttributeStringAsync(null, "name", null, sheet.Name);
            await xml.WriteAttributeStringAsync(null, "sheetId", null, Text(sheet.Index + 1));
            await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, "rId" + Text(sheet.Index + 1));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();

        List<SheetWriter> filtered = sheets.FindAll(sheet => sheet.AutoFilterEnd is not null);

        if (filtered.Count > 0)
        {
            await xml.WriteStartElementAsync(null, "definedNames", null);

            foreach (SheetWriter sheet in filtered)
            {
                await xml.WriteStartElementAsync(null, "definedName", null);
                await xml.WriteAttributeStringAsync(null, "name", null, "_xlnm._FilterDatabase");
                await xml.WriteAttributeStringAsync(null, "localSheetId", null, Text(sheet.Index));
                await xml.WriteAttributeStringAsync(null, "hidden", null, "1");
                await xml.WriteStringAsync(AbsoluteRange(sheet.Name, sheet.AutoFilterEnd!.Value));
                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        if (sheets.Exists(sheet => sheet.HasFormulas == true) == true)
        {
            // Formulas without a stored result are to be calculated when the file is opened.
            await xml.WriteStartElementAsync(null, "calcPr", null);
            await xml.WriteAttributeStringAsync(null, "calcId", null, "191029");
            await xml.WriteAttributeStringAsync(null, "fullCalcOnLoad", null, "1");
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private async Task WriteWorkbookRelationshipsAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);

        foreach (SheetWriter sheet in sheets)
        {
            await WriteRelationshipAsync(xml, "rId" + Text(sheet.Index + 1), Ooxml.WorksheetRelationship, "worksheets/sheet" + Text(sheet.Index + 1) + ".xml");
        }

        await WriteRelationshipAsync(xml, "rId" + Text(sheets.Count + 1), Ooxml.StylesRelationship, "styles.xml");
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private async Task WriteContentTypesAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Types", Ooxml.ContentTypesNamespace);
        await WriteDefaultTypeAsync(xml, "rels", Ooxml.RelationshipsContentType);
        await WriteDefaultTypeAsync(xml, "xml", "application/xml");

        foreach ((string extension, string contentType) in drawings.ImageTypes)
        {
            await WriteDefaultTypeAsync(xml, extension, contentType);
        }

        await WriteOverrideTypeAsync(xml, "/xl/workbook.xml", Ooxml.WorkbookContentType);

        foreach (SheetWriter sheet in sheets)
        {
            await WriteOverrideTypeAsync(xml, "/" + SheetPath(sheet.Index), Ooxml.WorksheetContentType);
        }

        await WriteOverrideTypeAsync(xml, "/xl/styles.xml", Ooxml.StylesContentType);

        foreach ((string partName, string contentType) in drawings.Overrides)
        {
            await WriteOverrideTypeAsync(xml, partName, contentType);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteRootRelationshipsAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
        await WriteRelationshipAsync(xml, "rId1", Ooxml.OfficeDocumentRelationship, "xl/workbook.xml");
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteRelationshipAsync(XmlWriter xml, string id, string type, string target)
    {
        await xml.WriteStartElementAsync(null, "Relationship", null);
        await xml.WriteAttributeStringAsync(null, "Id", null, id);
        await xml.WriteAttributeStringAsync(null, "Type", null, type);
        await xml.WriteAttributeStringAsync(null, "Target", null, target);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteDefaultTypeAsync(XmlWriter xml, string extension, string contentType)
    {
        await xml.WriteStartElementAsync(null, "Default", null);
        await xml.WriteAttributeStringAsync(null, "Extension", null, extension);
        await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteOverrideTypeAsync(XmlWriter xml, string partName, string contentType)
    {
        await xml.WriteStartElementAsync(null, "Override", null);
        await xml.WriteAttributeStringAsync(null, "PartName", null, partName);
        await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await xml.WriteEndElementAsync();
    }

    private static string AbsoluteRange(string sheetName, CellAddress end)
    {
        StringBuilder builder = new();
        builder.Append('\'').Append(sheetName.Replace("'", "''", StringComparison.Ordinal)).Append("'!$A$1:");
        builder.Append('$').Append(CellAddress.ColumnName(end.Column)).Append('$').Append(Text(end.Row + 1));

        return builder.ToString();
    }

    private static string SheetPath(int index)
    {
        return "xl/worksheets/sheet" + Text(index + 1) + ".xml";
    }

    private static string Text(int number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }

    private void ThrowIfCompleted()
    {
        if (completed == true)
        {
            throw new InvalidOperationException("The workbook is complete. Create a new writer to write another file.");
        }
    }
}
