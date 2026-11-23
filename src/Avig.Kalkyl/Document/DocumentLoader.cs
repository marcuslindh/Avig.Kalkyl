using System.Globalization;
using System.IO.Compression;
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
internal static class DocumentLoader
{
    public static async Task<Result> LoadAsync(Workbook book, XlsxReader reader, CancellationToken cancellationToken)
    {
        foreach (SheetPart part in reader.SheetParts)
        {
            if (XmlParts.TryFindEntry(reader.Archive, part.Path, out ZipArchiveEntry? entry) == false)
            {
                return Result.Invalid($"The file is damaged: the sheet '{part.Info.Name}' points to '{part.Path}', which is not in the file.");
            }

            try
            {
                Worksheet sheet = await LoadSheetAsync(book, part, entry, reader, cancellationToken);
                book.AddLoadedSheet(sheet);
            }
            catch (XmlException exception)
            {
                return Result.Invalid($"The file is damaged: the sheet '{part.Info.Name}' has broken XML ({exception.Message}).");
            }
            catch (InvalidDataException exception)
            {
                return Result.Invalid(exception.Message);
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

        // Loading is not a change; the formula engine starts from a clean slate.
        book.ChangedCells.Clear();

        return Result.Success();
    }

    /// <summary>The defined names, which the formula engine resolves. The file keeps its own copy, which is saved as it was.</summary>
    private static async Task LoadNamesAsync(Workbook book, XlsxReader reader, CancellationToken cancellationToken)
    {
        if (XmlParts.TryFindEntry(reader.Archive, reader.WorkbookPath, out ZipArchiveEntry? entry) == false)
        {
            return;
        }

        await using Stream stream = await entry.OpenAsync(cancellationToken);
        using XmlReader xml = XmlReader.Create(stream, XmlParts.Settings);

        while (xml.EOF == false)
        {
            if (XmlParts.IsElement(xml, "definedName") == false)
            {
                await xml.ReadAsync();

                continue;
            }

            string? name = xml.GetAttribute("name");
            int scope = -1;

            if (int.TryParse(xml.GetAttribute("localSheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int sheetIndex) == true)
            {
                scope = sheetIndex;
            }

            string formula = await xml.ReadElementContentAsStringAsync();

            if (name is not null && formula.Length > 0)
            {
                book.DefinedNames.Add(new Formulas.DefinedName(name, scope, formula));
            }
        }
    }

    private static async Task<Worksheet> LoadSheetAsync(Workbook book, SheetPart part, ZipArchiveEntry entry, XlsxReader reader, CancellationToken cancellationToken)
    {
        await using Stream stream = await entry.OpenAsync(cancellationToken);
        using XmlReader xml = XmlReader.Create(stream, XmlParts.Settings);
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

                case "sheetViews":
                    (frozenRows, frozenColumns, tabSelected) = ReadPanes((XElement)await XNode.ReadFromAsync(xml, cancellationToken));

                    break;

                case "autoFilter":
                    autoFilter = true;
                    await xml.SkipAsync();

                    break;

                case "sheetData":
                    await ReadSheetDataAsync(xml, context, fragmentContext, cells, rowAttributes, cancellationToken);

                    break;

                default:
                    await xml.SkipAsync();

                    break;
            }
        }

        SheetSource source = new(part.Path, part.RelationshipId, part.SheetId, sheetNamespace, columns, frozenRows, frozenColumns, autoFilter, part.Info.Visibility, tabSelected);
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
        await xml.ReadAsync();

        while (XmlParts.IsEnd(xml, depth) == false)
        {
            if (XmlParts.IsElement(xml, "row") == false)
            {
                await xml.ReadAsync();

                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            int rowIndex = previousRow + 1;

            if (int.TryParse(xml.GetAttribute("r"), NumberStyles.None, CultureInfo.InvariantCulture, out int rowNumber) == true && rowNumber > 0)
            {
                rowIndex = rowNumber - 1;
            }

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

                string raw = await xml.ReadOuterXmlAsync();
                using XmlReader fragment = XmlReader.Create(new StringReader(raw), fragmentSettings, fragmentContext);
                await fragment.MoveToContentAsync();
                ParsedCell parsed = await CellXml.ReadCellAsync(fragment, rowIndex, previousColumn, context);
                previousColumn = parsed.Column;
                string? kept = raw;

                if (parsed.SharedFormula == true)
                {
                    kept = null;
                }

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
