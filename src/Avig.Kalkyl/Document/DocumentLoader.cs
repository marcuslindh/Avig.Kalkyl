using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Document;

/// <summary>
/// Reads every worksheet of an opened file into the document model. Each cell keeps its
/// original XML, so a cell that is not changed is written back exactly as it was. Cells in a
/// shared formula are the exception: they get their own formula text and are written out
/// explicitly, so that editing the first cell of the group cannot break the others.
/// </summary>
/// <summary>
/// Counts the cells loaded into the document model across all sheets and stops at
/// KalkylLimits.MaxCells, before the cells fill memory rather than after.
/// </summary>
internal sealed class CellBudget(long limit)
{
    private long used;

    public void Take()
    {
        used++;

        if (used > limit)
        {
            throw new InvalidDataException($"The workbook has more than {limit:N0} cells, the limit set in KalkylLimits.MaxCells. If it is a file you trust, open it with larger limits, or read it with XlsxReader, which holds one row at a time.");
        }
    }
}

internal static class DocumentLoader
{
    public static async Task<Result> LoadAsync(Workbook book, XlsxReader reader, CancellationToken cancellationToken)
    {
        CellBudget budget = new(book.Limits.MaxCells);

        foreach (SheetPart part in reader.SheetParts)
        {
            if (XmlParts.TryFindEntry(reader.Archive, part.Path, out ZipArchiveEntry? entry) == false)
            {
                return Result.Invalid($"The file is damaged: the sheet '{part.Info.Name}' points to '{part.Path}', which is not in the file.");
            }

            try
            {
                Worksheet sheet = await LoadSheetAsync(book, part, entry, reader, budget, cancellationToken);
                book.AddLoadedSheet(sheet);
            }
            catch (XmlException exception)
            {
                return Result.Invalid($"The file is damaged: the sheet '{part.Info.Name}' has broken XML ({exception.Message}).");
            }
            catch (InvalidDataException exception)
            {
                return Result.Invalid(KalkylError.Describe(exception));
            }
        }

        try
        {
            await LoadNamesAsync(book, reader, cancellationToken);
        }
        catch (XmlException exception)
        {
            return Result.Invalid($"The file is damaged: the workbook part has broken XML ({exception.Message}).");
        }

        try
        {
            book.Pivots.Loaded = await PivotReader.LoadAsync(book, reader, cancellationToken);
        }
        catch (XmlException)
        {
            // A broken .rels or workbook part: the file opens without the list of its pivot
            // tables, as it did before the package read them, and the parts are kept as they are.
            book.Pivots.Loaded = new LoadedPivots();
        }
        catch (InvalidDataException exception)
        {
            return Result.Invalid(KalkylError.Describe(exception));
        }

        try
        {
            await CheckWellFormedAsync(reader.Archive, "[Content_Types].xml", cancellationToken);
        }
        catch (XmlException exception)
        {
            return Result.Invalid($"The file is damaged: [Content_Types].xml has broken XML ({exception.Message}).");
        }

        // Loading is not a change; the formula engine starts from a clean slate.
        book.ChangedCells.Clear();

        return Result.Success();
    }

    /// <summary>
    /// Reads a part through once without keeping anything. The model does not need the content
    /// types, but saving rewrites them, and a broken part has to be reported when the file is
    /// opened rather than thrown when it is saved.
    /// </summary>
    private static async Task CheckWellFormedAsync(ZipArchive archive, string path, CancellationToken cancellationToken)
    {
        if (XmlParts.TryFindEntry(archive, path, out ZipArchiveEntry? entry) == false)
        {
            return;
        }

        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader xml = XmlParts.CreateReader(stream);
        bool more = await xml.ReadAsync();

        while (more == true)
        {
            more = await xml.ReadAsync();
        }
    }

    /// <summary>The defined names, which the formula engine resolves. The file keeps its own copy, which is saved as it was.</summary>
    private static async Task LoadNamesAsync(Workbook book, XlsxReader reader, CancellationToken cancellationToken)
    {
        if (XmlParts.TryFindEntry(reader.Archive, reader.WorkbookPath, out ZipArchiveEntry? entry) == false)
        {
            return;
        }

        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader xml = XmlParts.CreateReader(stream);

        // Only the names in <definedNames>, the element the saver copies them from (see ReadWorkbookAsync).
        string section = string.Empty;

        while (xml.EOF == false)
        {
            if (xml.NodeType == XmlNodeType.Element && xml.Depth == 1)
            {
                section = xml.LocalName;
            }

            if (section != "definedNames" || xml.Depth != 2 || XmlParts.IsElement(xml, "definedName") == false)
            {
                await xml.ReadAsync();

                continue;
            }

            string? name = xml.GetAttribute("name");
            int scope = -1;

            if (int.TryParse(xml.GetAttribute("localSheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int sheetIndex) == true)
            {
                scope = ScopeOf(reader, sheetIndex);
            }

            string formula = await xml.ReadElementContentAsStringAsync();

            if (name is not null && formula.Length > 0)
            {
                book.DefinedNames.Add(new Formulas.DefinedName(name, scope, formula));
            }
        }
    }

    /// <summary>
    /// The worksheet a name belongs to. localSheetId counts every sheet in the sheet list, macro
    /// and chart sheets too, and the model only holds the worksheets. A name of another kind of
    /// sheet gets -2, which no formula looks up.
    /// </summary>
    private static int ScopeOf(XlsxReader reader, int localSheetId)
    {
        if (localSheetId >= reader.DeclaredSheets.Count)
        {
            return localSheetId;
        }

        int worksheet = reader.DeclaredSheets[localSheetId].WorksheetIndex;

        if (worksheet < 0)
        {
            return -2;
        }

        return worksheet;
    }

    private static async Task<Worksheet> LoadSheetAsync(Workbook book, SheetPart part, ZipArchiveEntry entry, XlsxReader reader, CellBudget budget, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader xml = XmlParts.CreateReader(stream);
        await xml.MoveToContentAsync();
        string sheetNamespace = xml.NamespaceURI;
        XmlNamespaceManager namespaces = new(new NameTable());

        if (xml.MoveToFirstAttribute() == true)
        {
            do
            {
                if (string.Equals(xml.Prefix, "xmlns", StringComparison.Ordinal) == true)
                {
                    namespaces.AddNamespace(xml.LocalName, xml.Value);
                }
                else if (string.Equals(xml.LocalName, "xmlns", StringComparison.Ordinal) == true && xml.Prefix.Length == 0)
                {
                    namespaces.AddNamespace(string.Empty, xml.Value);
                }
            }
            while (xml.MoveToNextAttribute() == true);

            xml.MoveToElement();
        }

        List<ColumnRange> columns = [];
        int frozenRows = 0;
        int frozenColumns = 0;
        bool autoFilter = false;
        bool tabSelected = false;
        SheetView view = SheetView.Default;
        string? tabColor = null;
        List<(int Row, int Column, CellEntry Entry)> cells = [];
        List<(int Row, IReadOnlyList<(string, string, string, string)> Attributes)> rowAttributes = [];
        CellContext context = new(reader.SharedStrings, reader.Styles, reader.Uses1904Dates, part.Info.Name);
        XmlParserContext fragmentContext = new(null, namespaces, null, XmlSpace.None);
        await xml.ReadAsync();

        while (xml.EOF == false && XmlParts.IsEnd(xml, 0) == false)
        {
            if (xml.NodeType != XmlNodeType.Element || xml.Depth != 1)
            {
                await xml.ReadAsync();

                continue;
            }

            switch (xml.LocalName)
            {
                case "cols":
                    columns = ReadColumns((XElement)await XNode.ReadFromAsync(xml, cancellationToken));

                    break;

                case "sheetPr":
                    if (TryReadTabColor((XElement)await XNode.ReadFromAsync(xml, cancellationToken), out string? rgb) == true)
                    {
                        tabColor = rgb;
                    }

                    break;

                case "sheetViews":
                    XElement views = (XElement)await XNode.ReadFromAsync(xml, cancellationToken);
                    (frozenRows, frozenColumns, tabSelected) = ReadPanes(views);
                    view = ReadView(views);

                    break;

                case "autoFilter":
                    autoFilter = true;
                    await xml.SkipAsync();

                    break;

                case "sheetData":
                    await ReadSheetDataAsync(xml, context, fragmentContext, cells, rowAttributes, budget, cancellationToken);

                    break;

                default:
                    await xml.SkipAsync();

                    break;
            }
        }

        SheetSource source = new(part.Path, part.RelationshipId, part.SheetId, sheetNamespace, columns, frozenRows, frozenColumns, autoFilter, part.Info.Visibility, tabSelected)
        {
            View = view with { TabColor = tabColor },
        };

        Worksheet sheet = new(book, part.Info.Name, part.Info.Index, source);

        foreach ((int row, IReadOnlyList<(string, string, string, string)> attributes) in rowAttributes)
        {
            sheet.GetRow(row).Attributes = attributes;
        }

        foreach ((int row, int column, CellEntry cellEntry) in cells)
        {
            sheet.Load(row, column, cellEntry);
        }

        return sheet;
    }

    private static async Task ReadSheetDataAsync(
        XmlReader xml,
        CellContext context,
        XmlParserContext fragmentContext,
        List<(int Row, int Column, CellEntry Entry)> cells,
        List<(int Row, IReadOnlyList<(string, string, string, string)> Attributes)> rowAttributes,
        CellBudget budget,
        CancellationToken cancellationToken)
    {
        if (xml.IsEmptyElement == true)
        {
            await xml.ReadAsync();

            return;
        }

        int depth = xml.Depth;
        int previousRow = -1;
        XmlReaderSettings fragmentSettings = XmlParts.Settings.Clone();
        fragmentSettings.ConformanceLevel = ConformanceLevel.Fragment;
        fragmentSettings.Async = false;

        // One writer for the sheet copies each cell's XML into one buffer. ReadOuterXml made a
        // new writer and string writer for every cell, which with the reader for the fragment
        // came to 67 kB of allocations per cell: 6.7 GB to open 100 000 formulas.
        StringBuilder rawText = new();
        await using StringWriter rawWriter = new(rawText, CultureInfo.InvariantCulture);
        XmlWriterSettings rawSettings = new() { ConformanceLevel = ConformanceLevel.Fragment, Async = true, NewLineHandling = NewLineHandling.None };

        await using XmlWriter copier = XmlWriter.Create(rawWriter, rawSettings);
        await xml.ReadAsync();

        while (XmlParts.IsEnd(xml, depth) == false)
        {
            if (XmlParts.IsElement(xml, "row") == false)
            {
                await xml.ReadAsync();

                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            int rowIndex = CellXml.RowIndex(xml, previousRow, context);

            previousRow = rowIndex;
            rowAttributes.Add((rowIndex, ReadRowAttributes(xml)));

            if (xml.IsEmptyElement == true)
            {
                await xml.ReadAsync();

                continue;
            }

            int rowDepth = xml.Depth;
            int previousColumn = -1;
            await xml.ReadAsync();

            while (XmlParts.IsEnd(xml, rowDepth) == false)
            {
                if (XmlParts.IsElement(xml, "c") == false)
                {
                    await xml.ReadAsync();

                    continue;
                }

                rawText.Clear();
                await copier.WriteNodeAsync(xml, true);
                await copier.FlushAsync();
                string raw = rawText.ToString();
                using XmlReader fragment = XmlParts.CreateReader(new StringReader(raw), fragmentSettings, fragmentContext);
                await fragment.MoveToContentAsync();
                ParsedCell parsed = await CellXml.ReadCellAsync(fragment, rowIndex, previousColumn, context);
                previousColumn = parsed.Column;
                string? kept = raw;

                if (parsed.SharedFormula == true)
                {
                    kept = null;
                }

                budget.Take();
                cells.Add((rowIndex, parsed.Column, new CellEntry(parsed.Cell, kept, parsed.StyleIndex, true)));
            }

            await xml.ReadAsync();
        }

        await xml.ReadAsync();
    }

    /// <summary>The row's attributes except r (rewritten) and spans (only a hint, and it may no longer be right).</summary>
    private static List<(string, string, string, string)> ReadRowAttributes(XmlReader xml)
    {
        List<(string, string, string, string)> attributes = [];

        if (xml.MoveToFirstAttribute() == true)
        {
            do
            {
                bool isDeclaration = string.Equals(xml.Prefix, "xmlns", StringComparison.Ordinal) == true || (xml.Prefix.Length == 0 && string.Equals(xml.LocalName, "xmlns", StringComparison.Ordinal) == true);

                if (isDeclaration == false && (xml.Prefix.Length > 0 || (xml.LocalName != "r" && xml.LocalName != "spans")))
                {
                    attributes.Add((xml.Prefix, xml.LocalName, xml.NamespaceURI, xml.Value));
                }
            }
            while (xml.MoveToNextAttribute() == true);

            xml.MoveToElement();
        }

        return attributes;
    }

    private static List<ColumnRange> ReadColumns(XElement cols)
    {
        List<ColumnRange> columns = [];

        foreach (XElement col in cols.Elements().Where(element => element.Name.LocalName == "col"))
        {
            if (int.TryParse((string?)col.Attribute("min"), NumberStyles.None, CultureInfo.InvariantCulture, out int min) == false
                || int.TryParse((string?)col.Attribute("max"), NumberStyles.None, CultureInfo.InvariantCulture, out int max) == false
                || min < 1 || max < min)
            {
                continue;
            }

            double? width = null;

            if (double.TryParse((string?)col.Attribute("width"), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) == true)
            {
                width = parsed;
            }

            List<KeyValuePair<string, string>> others = col.Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration == false && attribute.Name.LocalName is not ("min" or "max" or "width" or "customWidth"))
                .Select(attribute => new KeyValuePair<string, string>(attribute.Name.LocalName, attribute.Value))
                .ToList();

            columns.Add(new ColumnRange(min - 1, Math.Min(max, CellAddress.MaxColumns) - 1, width, others));
        }

        return columns;
    }

    /// <summary>
    /// The first view's settings. Values the model cannot hold (a zoom outside 10 to 400, an
    /// active cell that is not an address) are read as the default.
    /// </summary>
    private static SheetView ReadView(XElement views)
    {
        XElement? element = views.Elements().FirstOrDefault(candidate => candidate.Name.LocalName == "sheetView");

        if (element is null)
        {
            return SheetView.Default;
        }

        int zoom = 100;

        if (int.TryParse((string?)element.Attribute("zoomScale"), NumberStyles.None, CultureInfo.InvariantCulture, out int scale) == true && scale >= 10 && scale <= 400)
        {
            zoom = scale;
        }

        CellAddress? active = null;
        XElement? selection = element.Elements().LastOrDefault(candidate => candidate.Name.LocalName == "selection");

        if (selection is not null && CellAddress.TryParse((string?)selection.Attribute("activeCell") ?? string.Empty, out CellAddress cell) == true)
        {
            active = cell;
        }

        return new SheetView
        {
            ShowGridLines = IsOff((string?)element.Attribute("showGridLines")) == false,
            ShowHeaders = IsOff((string?)element.Attribute("showRowColHeaders")) == false,
            RightToLeft = (string?)element.Attribute("rightToLeft") is "1" or "true",
            Zoom = zoom,
            ActiveCell = active,
        };
    }

    private static bool IsOff(string? value)
    {
        return value is "0" or "false";
    }

    /// <summary>The tab's colour when sheetPr gives it as RGB; a theme or indexed colour reads as none and is kept in the file.</summary>
    private static bool TryReadTabColor(XElement properties, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? tabColor)
    {
        tabColor = null;
        XElement? color = properties.Elements().FirstOrDefault(candidate => candidate.Name.LocalName == "tabColor");
        string? rgb = (string?)color?.Attribute("rgb");

        if (rgb is null || (rgb.Length != 8 && rgb.Length != 6) || int.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _) == false)
        {
            return false;
        }

        tabColor = "#" + rgb[^6..].ToUpperInvariant();

        return true;
    }

    private static (int Rows, int Columns, bool Selected) ReadPanes(XElement views)
    {
        XElement? view = views.Elements().FirstOrDefault(element => element.Name.LocalName == "sheetView");

        if (view is null)
        {
            return (0, 0, false);
        }

        bool selected = (string?)view.Attribute("tabSelected") is "1" or "true";
        XElement? pane = view.Elements().FirstOrDefault(element => element.Name.LocalName == "pane");

        if (pane is null || ((string?)pane.Attribute("state") is "frozen" or "frozenSplit") == false)
        {
            return (0, 0, selected);
        }

        double.TryParse((string?)pane.Attribute("xSplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out double columns);
        double.TryParse((string?)pane.Attribute("ySplit"), NumberStyles.Float, CultureInfo.InvariantCulture, out double rows);

        return ((int)rows, (int)columns, selected);
    }
}
