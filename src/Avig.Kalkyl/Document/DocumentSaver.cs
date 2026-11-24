using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;
using Avig.Kalkyl.Writing;

namespace Avig.Kalkyl.Document;

/// <summary>
/// Writes the document model to a new .xlsx package.
/// <para>
/// For an opened file the original package is the starting point. Every part the model does
/// not own is copied byte for byte. The worksheets are copied element by element with only
/// &lt;sheetData&gt; rewritten (and the panes, columns and filter when they changed), so
/// conditional formatting, merged cells, drawings and the like stay where they were. Unchanged
/// cells are written from their original XML. styles.xml gets new entries appended; the
/// workbook part gets a fresh sheet list. calcChain.xml is dropped, since it may name cells that
/// no longer hold formulas; Excel rebuilds it, and fullCalcOnLoad makes Excel recalculate.
/// </para>
/// </summary>
internal static class DocumentSaver
{
    private static readonly string[] WorkbookOrder =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
        "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
        "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
    ];

    public static async Task SaveAsync(Workbook book, Stream output, CancellationToken cancellationToken)
    {
        await using ZipArchive archive = await ZipArchive.CreateAsync(output, ZipArchiveMode.Create, true, null, cancellationToken);

        if (book.Package is null || book.WorkbookPath is null)
        {
            await SaveNewAsync(book, archive, cancellationToken);

            return;
        }

        await using ZipArchive original = await ZipArchive.CreateAsync(new MemoryStream(book.Package, false), ZipArchiveMode.Read, false, null, cancellationToken);
        await SaveOpenedAsync(book, original, book.WorkbookPath, archive, cancellationToken);
    }

    private static async Task SaveNewAsync(Workbook book, ZipArchive archive, CancellationToken cancellationToken)
    {
        StyleRegistry styles = new();
        List<Worksheet> sheets = book.Sheets.ToList();

        if (sheets.Count == 0)
        {
            // Excel does not open a workbook without sheets, so the workbook gets one, as Excel would show it.
            book.AddSheet("Sheet1");
            sheets = book.Sheets.ToList();
        }

        List<NewSheet> parts = [];

        foreach (Worksheet sheet in sheets)
        {
            string number = SheetXml.Text(sheet.Index + 1);
            NewSheet part = new(sheet, "xl/worksheets/sheet" + number + ".xml", "worksheets/sheet" + number + ".xml", "rId" + number, sheet.Index + 1);
            parts.Add(part);
            await WritePartAsync(archive, part.Path, xml => WriteNewSheetAsync(xml, sheet, styles, sheet.Index == 0, false), cancellationToken);
        }

        await WritePartAsync(archive, "xl/styles.xml", styles.WriteAsync, cancellationToken);
        await WritePartAsync(archive, "xl/workbook.xml", xml => WriteNewWorkbookAsync(xml, book, parts), cancellationToken);
        string stylesId = "rId" + SheetXml.Text(parts.Count + 1);
        await WritePartAsync(archive, "xl/_rels/workbook.xml.rels", xml => WriteNewWorkbookRelationshipsAsync(xml, parts, stylesId), cancellationToken);
        await WritePartAsync(archive, "[Content_Types].xml", xml => WriteNewContentTypesAsync(xml, parts), cancellationToken);
        await WritePartAsync(archive, "_rels/.rels", WriteRootRelationshipsAsync, cancellationToken);
    }

    private static async Task SaveOpenedAsync(Workbook book, ZipArchive original, string workbookPath, ZipArchive archive, CancellationToken cancellationToken)
    {
        string workbookRelsPath = Ooxml.RelationshipsPath(workbookPath);
        (string? stylesPath, string? calcChainPath, HashSet<string> relationshipIds, HashSet<string> partPaths) = await ReadWorkbookRelationshipsAsync(original, workbookPath, workbookRelsPath, cancellationToken);
        bool mergeStyles = stylesPath is not null && XmlParts.TryFindEntry(original, stylesPath, out _);
        StyleRegistry styles = new();

        if (mergeStyles == true)
        {
            styles = new StyleRegistry(book.Styles);
        }

        HashSet<string> written = new(StringComparer.OrdinalIgnoreCase) { workbookPath, workbookRelsPath, "[Content_Types].xml" };

        if (calcChainPath is not null)
        {
            written.Add(calcChainPath);
        }

        List<NewSheet> added = [];
        int nextSheetId = book.Sheets.Where(sheet => sheet.Source is not null).Select(sheet => sheet.Source!.SheetId).DefaultIfEmpty(0).Max() + 1;
        string workbookFolder = Folder(workbookPath);
        bool hasFormulas = book.Sheets.Any(sheet => sheet.RowEntries.Values.Any(row => row.Cells.Values.Any(cell => cell.Cell.Formula is not null)));

        foreach (Worksheet sheet in book.Sheets)
        {
            if (sheet.Source is SheetSource source && XmlParts.TryFindEntry(original, source.Path, out ZipArchiveEntry? sheetEntry) == true)
            {
                written.Add(sheetEntry.FullName);
                await WritePartAsync(archive, sheetEntry.FullName, xml => WritePatchedSheetAsync(xml, sheetEntry, sheet, styles, book.Uses1904Dates, cancellationToken), cancellationToken);

                continue;
            }

            int number = 1;

            while (partPaths.Contains(workbookFolder + "worksheets/sheet" + SheetXml.Text(number) + ".xml") == true)
            {
                number++;
            }

            string target = "worksheets/sheet" + SheetXml.Text(number) + ".xml";
            partPaths.Add(workbookFolder + target);
            int idNumber = 1;

            while (relationshipIds.Contains("rId" + SheetXml.Text(idNumber)) == true)
            {
                idNumber++;
            }

            string relationshipId = "rId" + SheetXml.Text(idNumber);
            relationshipIds.Add(relationshipId);
            NewSheet part = new(sheet, workbookFolder + target, target, relationshipId, nextSheetId);
            nextSheetId++;
            added.Add(part);
            written.Add(part.Path);
            await WritePartAsync(archive, part.Path, xml => WriteNewSheetAsync(xml, sheet, styles, false, book.Uses1904Dates), cancellationToken);
        }

        string? addedStylesId = null;

        if (mergeStyles == true)
        {
            written.Add(stylesPath!);
            XmlParts.TryFindEntry(original, stylesPath!, out ZipArchiveEntry? stylesEntry);
            await WritePartAsync(archive, stylesEntry!.FullName, async xml =>
            {
                await using Stream stream = await XmlParts.OpenEntryAsync(stylesEntry, cancellationToken);
                using XmlReader reader = XmlParts.CreateReader(stream);
                await styles.WriteMergedAsync(reader, xml);
            }, cancellationToken);
        }
        else
        {
            stylesPath = workbookFolder + "styles.xml";
            written.Add(stylesPath);
            int idNumber = 1;

            while (relationshipIds.Contains("rId" + SheetXml.Text(idNumber)) == true)
            {
                idNumber++;
            }

            addedStylesId = "rId" + SheetXml.Text(idNumber);
            await WritePartAsync(archive, stylesPath, styles.WriteAsync, cancellationToken);
        }

        XmlParts.TryFindEntry(original, workbookPath, out ZipArchiveEntry? workbookEntry);
        await WritePartAsync(archive, workbookPath, xml => WritePatchedWorkbookAsync(xml, workbookEntry!, book, added, hasFormulas, cancellationToken), cancellationToken);
        await WritePartAsync(archive, workbookRelsPath, xml => WritePatchedWorkbookRelationshipsAsync(xml, original, workbookRelsPath, added, addedStylesId, book.RemovingMacros, cancellationToken), cancellationToken);
        await WritePartAsync(archive, "[Content_Types].xml", xml => WritePatchedContentTypesAsync(xml, original, book, workbookPath, added, calcChainPath, addedStylesId is not null, stylesPath, cancellationToken), cancellationToken);

        foreach (ZipArchiveEntry entry in original.Entries)
        {
            string name = entry.FullName.Replace('\\', '/');

            if (written.Contains(name) == true || name.EndsWith('/') == true)
            {
                continue;
            }

            if (book.RemovingMacros == true && Workbook.IsVbaPart(name) == true)
            {
                continue;
            }

            ZipArchiveEntry copy = archive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            await using Stream source = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            await using Stream target = await copy.OpenAsync(cancellationToken);
            await source.CopyToAsync(target, cancellationToken);
        }
    }

    private static async Task<(string? Styles, string? CalcChain, HashSet<string> Ids, HashSet<string> Paths)> ReadWorkbookRelationshipsAsync(ZipArchive original, string workbookPath, string relsPath, CancellationToken cancellationToken)
    {
        Dictionary<string, (string Type, string Target)> relationships = await XmlParts.ReadRelationshipsAsync(original, relsPath, cancellationToken);
        string? stylesPath = null;
        string? calcChainPath = null;
        HashSet<string> ids = new(relationships.Keys, StringComparer.Ordinal);
        HashSet<string> paths = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string type, string target) in relationships.Values)
        {
            string path = Ooxml.ResolveTarget(workbookPath, target);
            paths.Add(path);

            if (Ooxml.IsRelationship(type, "styles") == true)
            {
                stylesPath = path;
            }
            else if (Ooxml.IsRelationship(type, "calcChain") == true)
            {
                calcChainPath = path;
            }
        }

        foreach (ZipArchiveEntry entry in original.Entries)
        {
            paths.Add(entry.FullName.Replace('\\', '/'));
        }

        // Ids of external relationships are not in the dictionary but must not be reused either.
        if (XmlParts.TryFindEntry(original, relsPath, out ZipArchiveEntry? relsEntry) == true)
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(relsEntry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);

            while (await reader.ReadAsync() == true)
            {
                if (XmlParts.IsElement(reader, "Relationship") == true && reader.GetAttribute("Id") is string id)
                {
                    ids.Add(id);
                }
            }
        }

        return (stylesPath, calcChainPath, ids, paths);
    }

    private static async Task WriteNewSheetAsync(XmlWriter xml, Worksheet sheet, StyleRegistry styles, bool selected, bool uses1904)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "worksheet", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await SheetXml.WriteSheetViewsAsync(xml, selected, sheet.FrozenRows, sheet.FrozenColumns);
        await SheetXml.WriteColumnsAsync(xml, sheet.Columns);
        await WriteSheetDataAsync(xml, sheet, styles, Ooxml.SpreadsheetNamespace, uses1904);

        if (TryGetFilterEnd(sheet, out CellAddress end) == true)
        {
            await SheetXml.WriteAutoFilterAsync(xml, end);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WritePatchedSheetAsync(XmlWriter xml, ZipArchiveEntry entry, Worksheet sheet, StyleRegistry styles, bool uses1904, CancellationToken cancellationToken)
    {
        SheetSource source = sheet.Source!;
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        await reader.MoveToContentAsync();
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        await xml.WriteAttributesAsync(reader, true);

        // Elements to write in place of the original ones (or insert where they belong when the
        // original has none), keyed by element name. dimension is dropped: it is optional and
        // may no longer match.
        Dictionary<string, Func<Task>> replacements = new(StringComparer.Ordinal)
        {
            ["dimension"] = () => Task.CompletedTask,
            ["sheetData"] = () => WriteSheetDataAsync(xml, sheet, styles, source.Namespace, uses1904),
        };

        if (sheet.ViewChanged == true)
        {
            replacements["sheetViews"] = () => SheetXml.WriteSheetViewsAsync(xml, source.TabSelected, sheet.FrozenRows, sheet.FrozenColumns);
        }

        if (sheet.ColumnsChanged == true)
        {
            replacements["cols"] = () => SheetXml.WriteColumnsAsync(xml, sheet.Columns);
        }

        if (sheet.FilterChanged == true)
        {
            replacements["autoFilter"] = () => WriteFilterAsync(xml, sheet);
        }

        HashSet<string> owned = [.. replacements.Keys];

        if (reader.IsEmptyElement == false)
        {
            await reader.ReadAsync();

            while (reader.EOF == false && XmlParts.IsEnd(reader, 0) == false)
            {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1)
                {
                    await reader.ReadAsync();

                    continue;
                }

                string name = reader.LocalName;

                // An element the schema does not name (mc:AlternateContent, another program's
                // extension, damage) says nothing about where the replacements belong. Treating
                // it as "after everything" wrote sheetData before it and then copied the
                // original sheetData as well, so the file held two and the edit was lost.
                if (SheetXml.IsKnown(name) == true)
                {
                    await WritePendingBeforeAsync(replacements, SheetXml.OrderOf(name), SheetXml.OrderOf);
                }

                if (replacements.Remove(name, out Func<Task>? replacement) == true)
                {
                    await replacement();
                    await reader.SkipAsync();

                    continue;
                }

                // A second copy of an element already written from the model is dropped.
                if (owned.Contains(name) == true)
                {
                    await reader.SkipAsync();

                    continue;
                }

                await xml.WriteNodeAsync(reader, true);
            }
        }

        await WritePendingBeforeAsync(replacements, int.MaxValue, SheetXml.OrderOf);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>Writes the pending elements that belong before an element of the given order, in schema order.</summary>
    private static async Task WritePendingBeforeAsync(Dictionary<string, Func<Task>> pending, int order, Func<string, int> orderOf)
    {
        foreach (string name in pending.Keys.Where(key => orderOf(key) < order).OrderBy(orderOf).ToList())
        {
            await pending[name]();
            pending.Remove(name);
        }
    }

    private static async Task WriteFilterAsync(XmlWriter xml, Worksheet sheet)
    {
        if (TryGetFilterEnd(sheet, out CellAddress end) == true)
        {
            await SheetXml.WriteAutoFilterAsync(xml, end);
        }
    }

    private static bool TryGetFilterEnd(Worksheet sheet, out CellAddress end)
    {
        end = default;
        int rows = sheet.RowCount;
        int columns = sheet.ColumnCount;

        if (sheet.AutoFilter == false || rows == 0 || columns == 0)
        {
            return false;
        }

        end = new CellAddress(rows - 1, columns - 1);

        return true;
    }

    private static async Task WriteSheetDataAsync(XmlWriter xml, Worksheet sheet, StyleRegistry styles, string sheetNamespace, bool uses1904)
    {
        string declaration = " xmlns=\"" + sheetNamespace + "\"";
        await xml.WriteStartElementAsync(null, "sheetData", null);

        foreach ((int rowIndex, RowEntry row) in sheet.RowEntries)
        {
            if (row.Cells.Count == 0 && row.Attributes.Count == 0)
            {
                continue;
            }

            string rowNumber = SheetXml.Text(rowIndex + 1);
            await xml.WriteStartElementAsync(null, "row", null);
            await xml.WriteAttributeStringAsync(null, "r", null, rowNumber);

            foreach ((string prefix, string localName, string ns, string value) in row.Attributes)
            {
                await xml.WriteAttributeStringAsync(NullIfEmpty(prefix), localName, NullIfEmpty(ns), value);
            }

            foreach ((int column, CellEntry entry) in row.Cells)
            {
                if (entry.RawXml is not null)
                {
                    await xml.WriteRawAsync(RemoveFirst(entry.RawXml, declaration));

                    continue;
                }

                if (entry.Cell.IsEmpty == true)
                {
                    continue;
                }

                int styleIndex = entry.StyleIndex;

                if (entry.StyleFromFile == false)
                {
                    styleIndex = styles.GetIndex(entry.Cell.Style, CellXml.DefaultFormat(entry.Cell.Value));
                }

                await CellXml.WriteCellAsync(xml, entry.Cell, CellAddress.ColumnName(column) + rowNumber, styleIndex, uses1904);
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteNewWorkbookAsync(XmlWriter xml, Workbook book, List<NewSheet> parts)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "workbook", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await xml.WriteStartElementAsync(null, "bookViews", null);
        await xml.WriteElementStringAsync(null, "workbookView", null, string.Empty);
        await xml.WriteEndElementAsync();
        await WriteSheetListAsync(xml, book, parts, Ooxml.RelationshipsNamespace);
        await WriteFilterNamesAsync(xml, book, [], true);

        if (book.Sheets.Any(sheet => sheet.RowEntries.Values.Any(row => row.Cells.Values.Any(cell => cell.Cell.Formula is not null))) == true)
        {
            await xml.WriteStartElementAsync(null, "calcPr", null);
            await xml.WriteAttributeStringAsync(null, "calcId", null, "191029");
            await xml.WriteAttributeStringAsync(null, "fullCalcOnLoad", null, "1");
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WritePatchedWorkbookAsync(XmlWriter xml, ZipArchiveEntry entry, Workbook book, List<NewSheet> added, bool hasFormulas, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        await reader.MoveToContentAsync();
        string relationshipsNamespace = Ooxml.RelationshipsNamespace;

        if (reader.MoveToFirstAttribute() == true)
        {
            do
            {
                if (reader.Value.EndsWith("/relationships", StringComparison.Ordinal) == true && reader.Value.Contains("package", StringComparison.Ordinal) == false)
                {
                    relationshipsNamespace = reader.Value;
                }
            }
            while (reader.MoveToNextAttribute() == true);

            reader.MoveToElement();
        }

        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        await xml.WriteAttributesAsync(reader, true);
        Dictionary<string, Func<Task>> pending = new(StringComparer.Ordinal)
        {
            ["sheets"] = () => WriteSheetListAsync(xml, book, added, relationshipsNamespace),
            ["definedNames"] = () => WriteFilterNamesAsync(xml, book, [], true),
        };

        if (hasFormulas == true)
        {
            pending["calcPr"] = () => WriteCalcPropertiesAsync(xml, null);
        }

        int orderOf(string name) => OrderIn(WorkbookOrder, name);
        HashSet<string> owned = [.. pending.Keys];

        if (reader.IsEmptyElement == false)
        {
            await reader.ReadAsync();

            while (reader.EOF == false && XmlParts.IsEnd(reader, 0) == false)
            {
                if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1)
                {
                    await reader.ReadAsync();

                    continue;
                }

                string name = reader.LocalName;

                // As for worksheets: an element the schema does not name must not pull the sheet
                // list in front of it, or the original list is copied too and every sheet appears twice.
                if (Array.IndexOf(WorkbookOrder, name) >= 0)
                {
                    await WritePendingBeforeAsync(pending, orderOf(name), orderOf);
                }

                if (owned.Contains(name) == true && pending.ContainsKey(name) == false)
                {
                    await reader.SkipAsync();

                    continue;
                }

                switch (name)
                {
                    case "sheets":
                        pending.Remove(name);
                        await WriteSheetListAsync(xml, book, added, relationshipsNamespace);
                        await reader.SkipAsync();

                        break;

                    case "definedNames":
                        pending.Remove(name);
                        await CopyDefinedNamesAsync(xml, reader, book);

                        break;

                    case "calcPr" when hasFormulas == true:
                        pending.Remove(name);
                        await WriteCalcPropertiesAsync(xml, reader);
                        await reader.SkipAsync();

                        break;

                    default:
                        await xml.WriteNodeAsync(reader, true);

                        break;
                }
            }
        }

        await WritePendingBeforeAsync(pending, int.MaxValue, orderOf);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteSheetListAsync(XmlWriter xml, Workbook book, List<NewSheet> added, string relationshipsNamespace)
    {
        await xml.WriteStartElementAsync(null, "sheets", null);

        foreach (Worksheet sheet in book.Sheets)
        {
            string relationshipId;
            int sheetId;
            NewSheet? newPart = added.Find(part => ReferenceEquals(part.Sheet, sheet) == true);

            if (newPart is not null)
            {
                relationshipId = newPart.RelationshipId;
                sheetId = newPart.SheetId;
            }
            else if (sheet.Source is not null)
            {
                relationshipId = sheet.Source.RelationshipId;
                sheetId = sheet.Source.SheetId;
            }
            else
            {
                continue;
            }

            await xml.WriteStartElementAsync(null, "sheet", null);
            await xml.WriteAttributeStringAsync(null, "name", null, sheet.Name);
            await xml.WriteAttributeStringAsync(null, "sheetId", null, SheetXml.Text(sheetId));

            if (sheet.Visibility == SheetVisibility.Hidden)
            {
                await xml.WriteAttributeStringAsync(null, "state", null, "hidden");
            }
            else if (sheet.Visibility == SheetVisibility.VeryHidden)
            {
                await xml.WriteAttributeStringAsync(null, "state", null, "veryHidden");
            }

            await xml.WriteAttributeStringAsync("r", "id", relationshipsNamespace, relationshipId);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// Copies the defined names, except the filter ranges of sheets whose filter changed, and
    /// adds the filter ranges the model needs.
    /// </summary>
    private static async Task CopyDefinedNamesAsync(XmlWriter xml, XmlReader reader, Workbook book)
    {
        List<string> kept = [];
        int depth = reader.Depth;
        HashSet<int> filterNamesPresent = [];

        if (reader.IsEmptyElement == false)
        {
            await reader.ReadAsync();

            while (XmlParts.IsEnd(reader, depth) == false)
            {
                if (XmlParts.IsElement(reader, "definedName") == false)
                {
                    await reader.ReadAsync();

                    continue;
                }

                bool isFilter = string.Equals(reader.GetAttribute("name"), "_xlnm._FilterDatabase", StringComparison.Ordinal);
                int.TryParse(reader.GetAttribute("localSheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int localSheet);
                bool replace = isFilter == true && localSheet < book.Sheets.Count && book.Sheets[localSheet].FilterChanged == true;

                if (isFilter == true)
                {
                    filterNamesPresent.Add(localSheet);
                }

                string declaration = " xmlns=\"" + reader.NamespaceURI + "\"";
                string outer = RemoveFirst(await reader.ReadOuterXmlAsync(), declaration);

                if (replace == false)
                {
                    kept.Add(outer);
                }
            }
        }

        await reader.ReadAsync();
        await WriteFilterNamesAsync(xml, book, kept, false, filterNamesPresent);
    }

    private static async Task WriteFilterNamesAsync(XmlWriter xml, Workbook book, List<string> kept, bool all, HashSet<int>? present = null)
    {
        List<(int Index, string Range)> filters = [];

        foreach (Worksheet sheet in book.Sheets)
        {
            bool needsName = all == true || sheet.FilterChanged == true || sheet.Source is null || present?.Contains(sheet.Index) == false;

            if (needsName == true && sheet.AutoFilter == true && TryGetFilterEnd(sheet, out CellAddress end) == true)
            {
                filters.Add((sheet.Index, "'" + sheet.Name.Replace("'", "''", StringComparison.Ordinal) + "'!$A$1:$" + CellAddress.ColumnName(end.Column) + "$" + SheetXml.Text(end.Row + 1)));
            }
        }

        if (kept.Count == 0 && filters.Count == 0)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, "definedNames", null);

        foreach (string raw in kept)
        {
            await xml.WriteRawAsync(raw);
        }

        foreach ((int index, string range) in filters)
        {
            await xml.WriteStartElementAsync(null, "definedName", null);
            await xml.WriteAttributeStringAsync(null, "name", null, "_xlnm._FilterDatabase");
            await xml.WriteAttributeStringAsync(null, "localSheetId", null, SheetXml.Text(index));
            await xml.WriteAttributeStringAsync(null, "hidden", null, "1");
            await xml.WriteStringAsync(range);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>calcPr with fullCalcOnLoad, keeping the original's other attributes.</summary>
    private static async Task WriteCalcPropertiesAsync(XmlWriter xml, XmlReader? original)
    {
        await xml.WriteStartElementAsync(null, "calcPr", null);
        bool hasCalcId = false;

        if (original is not null && original.MoveToFirstAttribute() == true)
        {
            do
            {
                if (original.LocalName == "fullCalcOnLoad" || original.Prefix == "xmlns" || original.LocalName == "xmlns")
                {
                    continue;
                }

                hasCalcId |= original.LocalName == "calcId";
                await xml.WriteAttributeStringAsync(NullIfEmpty(original.Prefix), original.LocalName, NullIfEmpty(original.NamespaceURI), original.Value);
            }
            while (original.MoveToNextAttribute() == true);

            original.MoveToElement();
        }

        if (hasCalcId == false)
        {
            await xml.WriteAttributeStringAsync(null, "calcId", null, "191029");
        }

        await xml.WriteAttributeStringAsync(null, "fullCalcOnLoad", null, "1");
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePatchedWorkbookRelationshipsAsync(XmlWriter xml, ZipArchive original, string relsPath, List<NewSheet> added, string? addedStylesId, bool removeMacros, CancellationToken cancellationToken)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);

        if (XmlParts.TryFindEntry(original, relsPath, out ZipArchiveEntry? entry) == true)
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);

            while (reader.EOF == false)
            {
                if (XmlParts.IsElement(reader, "Relationship") == false)
                {
                    await reader.ReadAsync();

                    continue;
                }

                string relationshipType = reader.GetAttribute("Type") ?? string.Empty;

                if (Ooxml.IsRelationship(relationshipType, "calcChain") == true || (removeMacros == true && relationshipType.EndsWith("/vbaProject", StringComparison.Ordinal) == true))
                {
                    await reader.SkipAsync();

                    continue;
                }

                // Copied exactly. Rebuilding it filled a missing Target with "", which the reader
                // takes as a real target, so a relationship it had skipped now replaced another
                // with the same Id and a sheet disappeared on the next open.
                await xml.WriteNodeAsync(reader, true);
            }
        }

        foreach (NewSheet part in added)
        {
            await WriteRelationshipAsync(xml, part.RelationshipId, Ooxml.WorksheetRelationship, part.Target, null);
        }

        if (addedStylesId is not null)
        {
            await WriteRelationshipAsync(xml, addedStylesId, Ooxml.StylesRelationship, "styles.xml", null);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";

    private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

    /// <summary>The macro-enabled workbook types and the plain types they become when the VBA project is removed.</summary>
    private static readonly Dictionary<string, string> MacroEnabledTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["application/vnd.ms-excel.sheet.macroEnabled.main+xml"] = Ooxml.WorkbookContentType,
        ["application/vnd.ms-excel.template.macroEnabled.main+xml"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    };

    private static async Task WriteDefaultAsync(XmlWriter xml, string extension, string contentType)
    {
        await xml.WriteStartElementAsync(null, "Default", null);
        await xml.WriteAttributeStringAsync(null, "Extension", null, extension);
        await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePatchedContentTypesAsync(XmlWriter xml, ZipArchive original, Workbook book, string workbookPath, List<NewSheet> added, string? calcChainPath, bool addStyles, string? stylesPath, CancellationToken cancellationToken)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Types", Ooxml.ContentTypesNamespace);
        string? calcChainPart = null;

        if (calcChainPath is not null)
        {
            calcChainPart = "/" + calcChainPath;
        }

        if (XmlParts.TryFindEntry(original, "[Content_Types].xml", out ZipArchiveEntry? entry) == true)
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);

            while (reader.EOF == false)
            {
                if (reader.NodeType != XmlNodeType.Element || (reader.LocalName != "Default" && reader.LocalName != "Override"))
                {
                    await reader.ReadAsync();

                    continue;
                }

                string? partName = reader.GetAttribute("PartName");
                string contentType = reader.GetAttribute("ContentType") ?? string.Empty;
                bool calcChain = reader.LocalName == "Override" && calcChainPart is not null && string.Equals(partName, calcChainPart, StringComparison.OrdinalIgnoreCase) == true;
                bool macro = book.RemovingMacros == true && (string.Equals(contentType, VbaProjectContentType, StringComparison.OrdinalIgnoreCase) == true || (partName is not null && Workbook.IsVbaPart(partName) == true));

                if (calcChain == true || macro == true)
                {
                    await reader.SkipAsync();

                    continue;
                }

                // Without its VBA project the workbook is a plain one, and says so.
                if (book.RemovingMacros == true && MacroEnabledTypes.TryGetValue(contentType, out string? plain) == true)
                {
                    contentType = plain;
                }

                await xml.WriteStartElementAsync(null, reader.LocalName, null);

                if (reader.LocalName == "Default")
                {
                    await xml.WriteAttributeStringAsync(null, "Extension", null, reader.GetAttribute("Extension") ?? string.Empty);
                }
                else
                {
                    await xml.WriteAttributeStringAsync(null, "PartName", null, partName ?? string.Empty);
                }

                await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
                await xml.WriteEndElementAsync();
                await reader.SkipAsync();
            }
        }
        else
        {
            // A file without content types (some tools write none) still opens here; written
            // back as it was, Excel could not open it, so the package's own parts are declared.
            await WriteDefaultAsync(xml, "rels", Ooxml.RelationshipsContentType);
            await WriteDefaultAsync(xml, "xml", "application/xml");
            await WriteOverrideAsync(xml, "/" + workbookPath, Ooxml.WorkbookContentType);

            foreach (Worksheet sheet in book.Sheets.Where(sheet => sheet.Source is not null))
            {
                await WriteOverrideAsync(xml, "/" + sheet.Source!.Path, Ooxml.WorksheetContentType);
            }

            if (addStyles == false && stylesPath is not null)
            {
                await WriteOverrideAsync(xml, "/" + stylesPath, Ooxml.StylesContentType);
            }

            foreach (ZipArchiveEntry part in original.Entries.Where(part => part.Name.Equals("sharedStrings.xml", StringComparison.OrdinalIgnoreCase) == true))
            {
                await WriteOverrideAsync(xml, "/" + part.FullName.Replace('\\', '/'), SharedStringsContentType);
            }
        }

        foreach (NewSheet part in added)
        {
            await WriteOverrideAsync(xml, "/" + part.Path, Ooxml.WorksheetContentType);
        }

        if (addStyles == true && stylesPath is not null)
        {
            await WriteOverrideAsync(xml, "/" + stylesPath, Ooxml.StylesContentType);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteNewWorkbookRelationshipsAsync(XmlWriter xml, List<NewSheet> parts, string stylesId)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);

        foreach (NewSheet part in parts)
        {
            await WriteRelationshipAsync(xml, part.RelationshipId, Ooxml.WorksheetRelationship, part.Target, null);
        }

        await WriteRelationshipAsync(xml, stylesId, Ooxml.StylesRelationship, "styles.xml", null);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteNewContentTypesAsync(XmlWriter xml, List<NewSheet> parts)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Types", Ooxml.ContentTypesNamespace);

        foreach ((string extension, string contentType) in new[] { ("rels", Ooxml.RelationshipsContentType), ("xml", "application/xml") })
        {
            await xml.WriteStartElementAsync(null, "Default", null);
            await xml.WriteAttributeStringAsync(null, "Extension", null, extension);
            await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
            await xml.WriteEndElementAsync();
        }

        await WriteOverrideAsync(xml, "/xl/workbook.xml", Ooxml.WorkbookContentType);

        foreach (NewSheet part in parts)
        {
            await WriteOverrideAsync(xml, "/" + part.Path, Ooxml.WorksheetContentType);
        }

        await WriteOverrideAsync(xml, "/xl/styles.xml", Ooxml.StylesContentType);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteRootRelationshipsAsync(XmlWriter xml)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
        await WriteRelationshipAsync(xml, "rId1", Ooxml.OfficeDocumentRelationship, "xl/workbook.xml", null);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteRelationshipAsync(XmlWriter xml, string id, string type, string target, string? targetMode)
    {
        await xml.WriteStartElementAsync(null, "Relationship", null);
        await xml.WriteAttributeStringAsync(null, "Id", null, id);
        await xml.WriteAttributeStringAsync(null, "Type", null, type);
        await xml.WriteAttributeStringAsync(null, "Target", null, target);

        if (targetMode is not null)
        {
            await xml.WriteAttributeStringAsync(null, "TargetMode", null, targetMode);
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteOverrideAsync(XmlWriter xml, string partName, string contentType)
    {
        await xml.WriteStartElementAsync(null, "Override", null);
        await xml.WriteAttributeStringAsync(null, "PartName", null, partName);
        await xml.WriteAttributeStringAsync(null, "ContentType", null, contentType);
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePartAsync(ZipArchive archive, string path, Func<XmlWriter, Task> write, CancellationToken cancellationToken)
    {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        await using Stream entryStream = await entry.OpenAsync(cancellationToken);
        await using XmlWriter xml = XmlWriter.Create(entryStream, XlsxWriter.XmlSettings);
        await write(xml);
        await xml.FlushAsync();
    }

    private static int OrderIn(string[] order, string name)
    {
        int index = Array.IndexOf(order, name);

        if (index < 0)
        {
            return order.Length;
        }

        return index;
    }

    private static string Folder(string path)
    {
        int slash = path.LastIndexOf('/');

        if (slash < 0)
        {
            return string.Empty;
        }

        return path[..(slash + 1)];
    }

    private static string RemoveFirst(string text, string part)
    {
        int index = text.IndexOf(part, StringComparison.Ordinal);

        if (index < 0)
        {
            return text;
        }

        StringBuilder builder = new(text.Length);
        builder.Append(text, 0, index).Append(text, index + part.Length, text.Length - index - part.Length);

        return builder.ToString();
    }

    private static string? NullIfEmpty(string value)
    {
        string? result = value;

        if (value.Length == 0)
        {
            result = null;
        }

        return result;
    }

    private sealed record NewSheet(Worksheet Sheet, string Path, string Target, string RelationshipId, int SheetId);
}
