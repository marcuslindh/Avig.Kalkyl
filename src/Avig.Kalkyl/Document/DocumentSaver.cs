using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
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
    private static readonly HashSet<string> NoShapes = new(StringComparer.Ordinal);

    /// <summary>The defined names that start Excel 4.0 macros when the workbook opens or closes; Excel runs any name that begins so.</summary>
    private static readonly string[] AutoMacroNames = ["Auto_Open", "Auto_Close", "Auto_Activate", "Auto_Deactivate"];

    private static readonly string[] WorkbookOrder =
    [
        "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
        "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
        "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
    ];

    public static async Task SaveAsync(Workbook book, Stream output, CancellationToken cancellationToken)
    {
        // The pivot tables the package added are written as their sources are now.
        book.RenderPivotTables();
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
        DrawingPackage drawings = new(archive, []);
        PivotPackage pivots = new(archive, []);
        await pivots.WriteAsync(book, styles, sheet => sheet.Index + 1, cancellationToken);

        foreach (Worksheet sheet in sheets)
        {
            string number = SheetXml.Text(sheet.Index + 1);
            NewSheet part = new(sheet, "xl/worksheets/sheet" + number + ".xml", "worksheets/sheet" + number + ".xml", "rId" + number, sheet.Index + 1);
            parts.Add(part);
            (string? drawingId, List<string> linkIds, string? slicersId) = await WriteNewSheetRelationshipsAsync(archive, drawings, pivots, sheet, part.Path, cancellationToken);
            await WritePartAsync(archive, part.Path, xml => WriteNewSheetAsync(xml, sheet, styles, sheet.Index == 0, false, drawingId, linkIds, slicersId), cancellationToken);
        }

        await WritePartAsync(archive, "xl/styles.xml", styles.WriteAsync, cancellationToken);
        SheetList sheetList = SheetList.Create(book, new HashSet<string>(StringComparer.Ordinal));
        List<string> pivotIds = [.. pivots.WorkbookRelationships.Select((_, index) => "rId" + SheetXml.Text(parts.Count + 2 + index))];
        await WritePartAsync(archive, "xl/workbook.xml", xml => WriteNewWorkbookAsync(xml, book, parts, sheetList, pivots, pivotIds), cancellationToken);
        string stylesId = "rId" + SheetXml.Text(parts.Count + 1);
        await WritePartAsync(archive, "xl/_rels/workbook.xml.rels", xml => WriteNewWorkbookRelationshipsAsync(xml, parts, stylesId, pivots, pivotIds), cancellationToken);
        await WritePartAsync(archive, "[Content_Types].xml", xml => WriteNewContentTypesAsync(xml, parts, drawings, pivots), cancellationToken);
        await WritePartAsync(archive, "_rels/.rels", WriteRootRelationshipsAsync, cancellationToken);
    }

    private static async Task SaveOpenedAsync(Workbook book, ZipArchive original, string workbookPath, ZipArchive archive, CancellationToken cancellationToken)
    {
        string workbookRelsPath = Ooxml.RelationshipsPath(workbookPath);
        (string? stylesPath, string? calcChainPath, HashSet<string> relationshipIds, HashSet<string> partPaths) = await ReadWorkbookRelationshipsAsync(original, workbookPath, workbookRelsPath, cancellationToken);
        ActiveContentRemoval removal = ActiveContentRemoval.None;

        if (book.RemovingActiveContent == true)
        {
            removal = await ActiveContentRemoval.CreateAsync(original, book.Scan, workbookPath, cancellationToken);
        }

        SheetList sheetList = SheetList.Create(book, removal.RemovedIds(workbookPath));
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
        DrawingPackage drawings = new(archive, partPaths) { Removal = removal };

        Dictionary<Worksheet, int> sheetIds = SheetIds(book, original, nextSheetId);
        PivotPackage pivots = new(archive, partPaths);
        await pivots.WriteAsync(book, styles, sheet => sheetIds[sheet], cancellationToken);

        foreach (Worksheet sheet in book.Sheets)
        {
            if (sheet.Source is SheetSource source && XmlParts.TryFindEntry(original, source.Path, out ZipArchiveEntry? sheetEntry) == true)
            {
                written.Add(sheetEntry.FullName);
                string? drawingId = null;
                List<string> linkIds = [];
                string? slicersId = null;

                if (sheet.Extras.Drawings.Count > 0 || sheet.Extras.ExternalHyperlinks.Any() == true || pivots.Sheets.ContainsKey(sheet) == true)
                {
                    string sheetPath = sheetEntry.FullName.Replace('\\', '/');

                    // The sheet's .rels and its drawing are read for the first time here.
                    try
                    {
                        (string? openedId, List<string> openedLinks, string? openedSlicersId, List<string> drawingParts) = await WriteOpenedSheetRelationshipsAsync(original, archive, drawings, pivots, sheetPath, sheet, cancellationToken);
                        drawingId = openedId;
                        linkIds = openedLinks;
                        slicersId = openedSlicersId;
                        written.UnionWith(drawingParts);
                    }
                    catch (XmlException exception)
                    {
                        throw new InvalidDataException($"{KalkylError.DamagedPrefix}: the relationships or the drawing of {sheetPath} are not valid XML ({exception.Message}), so nothing can be added to them.", exception);
                    }
                }

                int firstPriority = 1;

                if (sheet.Extras.ConditionalFormats.Count > 0)
                {
                    firstPriority = await MaxPriorityAsync(sheetEntry, cancellationToken) + 1;
                }

                await WritePartAsync(archive, sheetEntry.FullName, xml => WritePatchedSheetAsync(xml, sheetEntry, sheet, styles, book.Uses1904Dates, new SheetTargets(drawingId, linkIds, slicersId), firstPriority, removal, cancellationToken), cancellationToken);

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
            (string? newDrawingId, List<string> newLinkIds, string? newSlicersId) = await WriteNewSheetRelationshipsAsync(archive, drawings, pivots, sheet, part.Path, cancellationToken);

            if (newDrawingId is not null || newLinkIds.Count > 0 || newSlicersId is not null || pivots.Sheets.ContainsKey(sheet) == true)
            {
                written.Add(Ooxml.RelationshipsPath(part.Path));
            }

            await WritePartAsync(archive, part.Path, xml => WriteNewSheetAsync(xml, sheet, styles, false, book.Uses1904Dates, newDrawingId, newLinkIds, newSlicersId), cancellationToken);
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

        if (addedStylesId is not null)
        {
            relationshipIds.Add(addedStylesId);
        }

        List<string> pivotIds = [.. pivots.WorkbookRelationships.Select(_ => DrawingPackage.NextId(relationshipIds))];
        XmlParts.TryFindEntry(original, workbookPath, out ZipArchiveEntry? workbookEntry);
        await WritePartAsync(archive, workbookPath, xml => WritePatchedWorkbookAsync(xml, workbookEntry!, book, added, hasFormulas, sheetList, removal, pivots, pivotIds, cancellationToken), cancellationToken);
        await WritePartAsync(archive, workbookRelsPath, xml => WritePatchedWorkbookRelationshipsAsync(xml, original, workbookRelsPath, added, addedStylesId, book.DropsVba, removal.RemovedIds(workbookPath), pivots, pivotIds, cancellationToken), cancellationToken);
        await WritePartAsync(archive, "[Content_Types].xml", xml => WritePatchedContentTypesAsync(xml, original, book, workbookPath, added, calcChainPath, addedStylesId is not null, stylesPath, drawings, pivots, removal, cancellationToken), cancellationToken);

        foreach (ZipArchiveEntry entry in original.Entries)
        {
            string name = entry.FullName.Replace('\\', '/');

            if (written.Contains(name) == true || name.EndsWith('/') == true)
            {
                continue;
            }

            if ((book.DropsVba == true && Workbook.IsVbaPart(name) == true) || removal.Keeps(name) == false)
            {
                continue;
            }

            if (Avig.OpenXml.OpcPath.TryGetSourcePart(name, out string? sourcePart) == true && removal.RemovedIds(sourcePart).Count > 0)
            {
                IReadOnlySet<string> removedIds = removal.RemovedIds(sourcePart);
                await WritePartAsync(archive, entry.FullName, async xml =>
                {
                    await xml.WriteStartDocumentAsync(true);
                    await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
                    await ActiveContentRemoval.CopyRelationshipsAsync(entry, xml, removedIds, cancellationToken);
                    await xml.WriteEndElementAsync();
                    await xml.WriteEndDocumentAsync();
                }, cancellationToken);

                continue;
            }

            ZipArchiveEntry copy = archive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            await using Stream target = await copy.OpenAsync(cancellationToken);

            if (removal.Rewrites(name) == true)
            {
                await removal.WriteCleanedAsync(entry, target, cancellationToken);

                continue;
            }

            if (book.RefreshPivotTablesOnOpen == true && book.Pivots.Loaded.CacheDefinitionPaths.Contains(name, StringComparer.OrdinalIgnoreCase) == true)
            {
                await CopyWithRefreshOnLoadAsync(entry, target, cancellationToken);

                continue;
            }

            await using Stream source = await XmlParts.OpenEntryAsync(entry, cancellationToken);
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

    /// <summary>
    /// The drawing of a sheet the package writes from scratch, and the sheet's .rels pointing to
    /// it and to the web addresses its cells link to. Returns the relationship id for
    /// &lt;drawing&gt; (null without drawings) and one per web link, in order. The pivot tables
    /// and slicers on the sheet get theirs after them, and the slicers part's id is returned for
    /// the sheet's extLst.
    /// </summary>
    private static async Task<(string? DrawingId, List<string> LinkIds, string? SlicersId)> WriteNewSheetRelationshipsAsync(ZipArchive archive, DrawingPackage drawings, PivotPackage pivots, Worksheet sheet, string sheetPath, CancellationToken cancellationToken)
    {
        List<PartRelationship> relationships = [];
        string? drawingId = null;

        if (sheet.Extras.Drawings.Count > 0)
        {
            string drawingPath = await drawings.WriteDrawingAsync(sheet.Extras.Drawings, cancellationToken);
            drawingId = "rId1";
            relationships.Add(new PartRelationship(drawingId, DrawingXml.DrawingRelationship, "/" + drawingPath));
        }

        List<string> linkIds = [];
        int number = 2;

        foreach (Hyperlink link in sheet.Extras.ExternalHyperlinks)
        {
            string id = "rId" + SheetXml.Text(number);
            number++;
            linkIds.Add(id);
            relationships.Add(new PartRelationship(id, DrawingXml.HyperlinkRelationship, link.Url!, true));
        }

        string? slicersId = null;

        if (pivots.Sheets.TryGetValue(sheet, out SheetPivotParts? parts) == true)
        {
            foreach ((string type, string target) in parts.Relationships)
            {
                string id = "rId" + SheetXml.Text(number);
                number++;
                relationships.Add(new PartRelationship(id, type, target));

                if (string.Equals(target, parts.SlicersTarget, StringComparison.Ordinal) == true)
                {
                    slicersId = id;
                }
            }
        }

        if (relationships.Count > 0)
        {
            await WritePartAsync(archive, Ooxml.RelationshipsPath(sheetPath), xml => DrawingXml.WriteRelationshipsAsync(xml, relationships), cancellationToken);
        }

        return (drawingId, linkIds, slicersId);
    }

    /// <summary>
    /// The new drawings and web links of an opened sheet. A sheet that has a drawing already gets
    /// the new anchors added to it, since a sheet has at most one; the drawing id is then null.
    /// Otherwise a new drawing part is added to the sheet's .rels, and the id is what
    /// &lt;drawing&gt; must refer to. Each web link gets a relationship too. Also returns the
    /// parts written in place of the original ones.
    /// </summary>
    private static async Task<(string? DrawingId, List<string> LinkIds, string? SlicersId, List<string> Written)> WriteOpenedSheetRelationshipsAsync(ZipArchive original, ZipArchive archive, DrawingPackage drawings, PivotPackage pivots, string sheetPath, Worksheet sheet, CancellationToken cancellationToken)
    {
        string relsPath = Ooxml.RelationshipsPath(sheetPath);
        HashSet<string> ids = await DrawingPackage.ReadRelationshipIdsAsync(original, relsPath, cancellationToken);
        List<PartRelationship> added = [];
        List<string> written = [];
        string? drawingId = null;

        if (sheet.Extras.Drawings.Count > 0)
        {
            bool appended = false;
            Dictionary<string, (string Type, string Target)> relationships = await XmlParts.ReadRelationshipsAsync(original, relsPath, cancellationToken);

            foreach ((string type, string target) in relationships.Values)
            {
                if (Ooxml.IsRelationship(type, "drawing") == true && XmlParts.TryFindEntry(original, Ooxml.ResolveTarget(sheetPath, target), out ZipArchiveEntry? drawing) == true)
                {
                    written.AddRange(await drawings.AppendToDrawingAsync(original, drawing, sheet.Extras.Drawings, cancellationToken));
                    appended = true;

                    break;
                }
            }

            if (appended == false)
            {
                string drawingPath = await drawings.WriteDrawingAsync(sheet.Extras.Drawings, cancellationToken);
                drawingId = DrawingPackage.NextId(ids);
                added.Add(new PartRelationship(drawingId, DrawingXml.DrawingRelationship, "/" + drawingPath));
            }
        }

        List<string> linkIds = [];

        foreach (Hyperlink link in sheet.Extras.ExternalHyperlinks)
        {
            string id = DrawingPackage.NextId(ids);
            linkIds.Add(id);
            added.Add(new PartRelationship(id, DrawingXml.HyperlinkRelationship, link.Url!, true));
        }

        // The pivot tables and slicers added to the sheet.
        string? slicersId = null;

        if (pivots.Sheets.TryGetValue(sheet, out SheetPivotParts? parts) == true)
        {
            foreach ((string type, string target) in parts.Relationships)
            {
                string pivotId = DrawingPackage.NextId(ids);
                added.Add(new PartRelationship(pivotId, type, target));

                if (string.Equals(target, parts.SlicersTarget, StringComparison.Ordinal) == true)
                {
                    slicersId = pivotId;
                }
            }
        }

        if (added.Count > 0)
        {
            await WritePartAsync(archive, relsPath, async xml =>
            {
                await xml.WriteStartDocumentAsync(true);
                await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
                await DrawingPackage.CopyRelationshipsAsync(original, relsPath, xml, drawings.Removal.RemovedIds(sheetPath), cancellationToken);
                await DrawingXml.WriteRelationshipElementsAsync(xml, added);
                await xml.WriteEndElementAsync();
                await xml.WriteEndDocumentAsync();
            }, cancellationToken);

            written.Add(relsPath);
        }

        return (drawingId, linkIds, slicersId, written);
    }

    /// <summary>
    /// What a saved sheet refers to through its .rels: the drawing (null when none is new), the
    /// web links, in order, and the slicers part (null when the sheet gets no slicers).
    /// </summary>
    private sealed record SheetTargets(string? DrawingId, IReadOnlyList<string> LinkIds, string? SlicersId);

    /// <summary>
    /// The sheetId each sheet gets in the saved file: its own for a sheet the file had, and the
    /// next free ones, in tab order, for the sheets added, as the save gives them out.
    /// </summary>
    private static Dictionary<Worksheet, int> SheetIds(Workbook book, ZipArchive original, int nextSheetId)
    {
        Dictionary<Worksheet, int> ids = [];

        foreach (Worksheet sheet in book.Sheets)
        {
            if (sheet.Source is SheetSource source && XmlParts.TryFindEntry(original, source.Path, out _) == true)
            {
                ids[sheet] = source.SheetId;

                continue;
            }

            ids[sheet] = nextSheetId;
            nextSheetId++;
        }

        return ids;
    }

    /// <summary>Copies a pivot cache definition with refreshOnLoad="1" on its root, so that Excel refreshes it when it opens the file.</summary>
    private static async Task CopyWithRefreshOnLoadAsync(ZipArchiveEntry entry, Stream target, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        await using XmlWriter xml = XmlWriter.Create(target, XlsxWriter.XmlSettings);
        await reader.MoveToContentAsync();
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI);
        bool empty = reader.IsEmptyElement;

        if (reader.MoveToFirstAttribute() == true)
        {
            do
            {
                if (reader.LocalName == "refreshOnLoad" && reader.NamespaceURI.Length == 0)
                {
                    continue;
                }

                await xml.WriteAttributeStringAsync(NullIfEmpty(reader.Prefix), reader.LocalName, NullIfEmpty(reader.NamespaceURI), reader.Value);
            }
            while (reader.MoveToNextAttribute() == true);

            reader.MoveToElement();
        }

        await xml.WriteAttributeStringAsync(null, "refreshOnLoad", null, "1");

        if (empty == false)
        {
            await reader.ReadAsync();

            while (reader.EOF == false && XmlParts.IsEnd(reader, 0) == false)
            {
                await xml.WriteNodeAsync(reader, true);
            }
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
        await xml.FlushAsync();
    }

    /// <summary>The highest priority of the sheet's conditional formatting rules, extensions (x14) included; 0 when it has none.</summary>
    private static async Task<int> MaxPriorityAsync(ZipArchiveEntry sheetEntry, CancellationToken cancellationToken)
    {
        int max = 0;
        await using Stream stream = await XmlParts.OpenEntryAsync(sheetEntry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "sheetData") == true)
            {
                // The cells are most of a sheet and hold no rules.
                await reader.SkipAsync();

                continue;
            }

            if (XmlParts.IsElement(reader, "cfRule") == true
                && int.TryParse(reader.GetAttribute("priority"), NumberStyles.None, CultureInfo.InvariantCulture, out int priority) == true)
            {
                max = Math.Max(max, priority);
            }
        }

        return max;
    }

    private static async Task WriteNewSheetAsync(XmlWriter xml, Worksheet sheet, StyleRegistry styles, bool selected, bool uses1904, string? drawingId, IReadOnlyList<string> linkIds, string? slicersId)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "worksheet", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await SheetXml.WriteSheetPropertiesAsync(xml, sheet.View.TabColor);
        await SheetXml.WriteSheetViewsAsync(xml, selected, sheet.FrozenRows, sheet.FrozenColumns, sheet.View);
        await SheetXml.WriteColumnsAsync(xml, sheet.Columns);
        await WriteSheetDataAsync(xml, sheet, styles, Ooxml.SpreadsheetNamespace, uses1904);

        if (TryGetFilterEnd(sheet, out CellAddress end) == true)
        {
            await SheetXml.WriteAutoFilterAsync(xml, end);
        }

        if (sheet.Extras.ConditionalFormats.Count > 0)
        {
            await ConditionalFormatXml.WriteAsync(xml, sheet.Extras.ConditionalFormats, styles, 1);
        }

        if (sheet.Extras.Hyperlinks.Count > 0)
        {
            await SheetXml.WriteHyperlinksAsync(xml, sheet.Extras.Hyperlinks, linkIds);
        }

        if (drawingId is not null)
        {
            await DrawingXml.WriteSheetDrawingAsync(xml, drawingId);
        }

        if (slicersId is not null)
        {
            await xml.WriteStartElementAsync(null, "extLst", null);
            await PivotXml.WriteSheetSlicerListAsync(xml, slicersId);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WritePatchedSheetAsync(XmlWriter xml, ZipArchiveEntry entry, Worksheet sheet, StyleRegistry styles, bool uses1904, SheetTargets targets, int firstPriority, ActiveContentRemoval removal, CancellationToken cancellationToken)
    {
        string? drawingId = targets.DrawingId;
        string? slicersId = targets.SlicersId;
        SheetSource source = sheet.Source!;
        IReadOnlySet<string> removedIds = removal.RemovedIds(entry.FullName.Replace('\\', '/'));
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

        // Elements written from the model and the file's own element together; the file's one
        // is passed in, or null when it has none.
        Dictionary<string, Func<XElement?, Task>> patches = new(StringComparer.Ordinal);

        if (sheet.ViewChanged == true)
        {
            patches["sheetViews"] = original => SheetXml.WriteSheetViewsAsync(xml, source.TabSelected, sheet.FrozenRows, sheet.FrozenColumns, sheet.View, original);
        }

        if (sheet.TabColorChanged == true)
        {
            patches["sheetPr"] = original => SheetXml.WriteSheetPropertiesAsync(xml, sheet.View.TabColor, original);
        }

        if (sheet.Extras.Hyperlinks.Count > 0)
        {
            patches["hyperlinks"] = original => SheetXml.WriteHyperlinksAsync(xml, sheet.Extras.Hyperlinks, targets.LinkIds, original);
        }

        foreach ((string name, Func<XElement?, Task> patch) in patches)
        {
            replacements[name] = () => patch(null);
        }

        if (sheet.ColumnsChanged == true)
        {
            replacements["cols"] = () => SheetXml.WriteColumnsAsync(xml, sheet.Columns);
        }

        if (sheet.FilterChanged == true)
        {
            replacements["autoFilter"] = () => WriteFilterAsync(xml, sheet);
        }

        if (drawingId is not null)
        {
            replacements["drawing"] = () => DrawingXml.WriteSheetDrawingAsync(xml, drawingId);
        }

        // New rules go after the sheet's own: they are pending like the replacements, so they are
        // written in schema order with them, but the original elements of that name are kept.
        HashSet<string> insertOnly = new(StringComparer.Ordinal);

        if (sheet.Extras.ConditionalFormats.Count > 0)
        {
            replacements["conditionalFormatting"] = () => ConditionalFormatXml.WriteAsync(xml, sheet.Extras.ConditionalFormats, styles, firstPriority);
            insertOnly.Add("conditionalFormatting");
        }

        HashSet<string> owned = [.. replacements.Keys.Where(key => insertOnly.Contains(key) == false)];

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

                if (patches.Remove(name, out Func<XElement?, Task>? patch) == true)
                {
                    replacements.Remove(name);
                    XElement element = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);
                    await patch(element);

                    continue;
                }

                // The slicers go into the sheet's own extension list, next to what it has.
                if (slicersId is not null && name == "extLst")
                {
                    XElement extensions = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);

                    if (removal.IsActive == true && ActiveContentRemoval.Clean(extensions, removedIds, NoShapes) == false)
                    {
                        extensions = new XElement(extensions.Name);
                    }

                    AddExtension(extensions, PivotXml.SheetSlicerListUri, "slicerList", "slicer", slicersId);
                    await extensions.WriteToAsync(xml, cancellationToken);
                    slicersId = null;

                    continue;
                }

                if (insertOnly.Contains(name) == false && replacements.Remove(name, out Func<Task>? replacement) == true)
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

                // Without active content the OLE objects and ActiveX controls go, and elements that
                // need a relationship that is left out (see ActiveContentRemoval).
                if (removal.IsActive == true)
                {
                    XElement element = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);

                    if (ActiveContentRemoval.Clean(element, removedIds, NoShapes) == true)
                    {
                        await element.WriteToAsync(xml, cancellationToken);
                    }

                    continue;
                }

                await xml.WriteNodeAsync(reader, true);
            }
        }

        await WritePendingBeforeAsync(replacements, int.MaxValue, SheetXml.OrderOf);

        if (slicersId is not null)
        {
            await xml.WriteStartElementAsync(null, "extLst", null);
            await PivotXml.WriteSheetSlicerListAsync(xml, slicersId);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>
    /// Adds an x14 list entry (a slicer, a slicer cache) to an extension list: to the extension
    /// with that uri when there is one, and in a new one otherwise.
    /// </summary>
    private static void AddExtension(XElement extensions, string uri, string listName, string itemName, string relationshipId)
    {
        XNamespace main = extensions.Name.Namespace;
        XNamespace x14 = PivotXml.X14Namespace;
        XNamespace relationships = Ooxml.RelationshipsNamespace;
        XElement item = new(x14 + itemName, new XAttribute(relationships + "id", relationshipId));
        XElement? extension = extensions.Elements().FirstOrDefault(element => element.Name.LocalName == "ext" && string.Equals((string?)element.Attribute("uri"), uri, StringComparison.OrdinalIgnoreCase) == true);
        XElement? list = extension?.Elements().FirstOrDefault(element => element.Name.LocalName == listName);

        if (list is not null)
        {
            list.Add(item);

            return;
        }

        extensions.AddFirst(new XElement(main + "ext", new XAttribute("uri", uri), new XAttribute(XNamespace.Xmlns + "x14", PivotXml.X14Namespace), new XElement(x14 + listName, item)));
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

    private static async Task WriteNewWorkbookAsync(XmlWriter xml, Workbook book, List<NewSheet> parts, SheetList sheetList, PivotPackage pivots, List<string> pivotIds)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "workbook", Ooxml.SpreadsheetNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await xml.WriteStartElementAsync(null, "bookViews", null);
        await xml.WriteElementStringAsync(null, "workbookView", null, string.Empty);
        await xml.WriteEndElementAsync();
        await WriteSheetListAsync(xml, sheetList, parts, Ooxml.RelationshipsNamespace);
        await WriteFilterNamesAsync(xml, book, sheetList, SlicerNames(pivots), true);

        if (book.Sheets.Any(sheet => sheet.RowEntries.Values.Any(row => row.Cells.Values.Any(cell => cell.Cell.Formula is not null))) == true)
        {
            await xml.WriteStartElementAsync(null, "calcPr", null);
            await xml.WriteAttributeStringAsync(null, "calcId", null, "191029");
            await xml.WriteAttributeStringAsync(null, "fullCalcOnLoad", null, "1");
            await xml.WriteEndElementAsync();
        }

        await WritePivotCachesAsync(xml, null, pivots, pivotIds, Ooxml.RelationshipsNamespace, CancellationToken.None);
        await WriteWorkbookExtensionsAsync(xml, null, pivots, pivotIds, CancellationToken.None);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WritePatchedWorkbookAsync(XmlWriter xml, ZipArchiveEntry entry, Workbook book, List<NewSheet> added, bool hasFormulas, SheetList sheetList, ActiveContentRemoval removal, PivotPackage pivots, List<string> pivotIds, CancellationToken cancellationToken)
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
            ["sheets"] = () => WriteSheetListAsync(xml, sheetList, added, relationshipsNamespace),
            ["definedNames"] = () => WriteFilterNamesAsync(xml, book, sheetList, SlicerNames(pivots), true),
        };

        if (hasFormulas == true)
        {
            pending["calcPr"] = () => WriteCalcPropertiesAsync(xml, null);
        }

        if (pivots.PivotCaches.Count > 0)
        {
            pending["pivotCaches"] = () => WritePivotCachesAsync(xml, null, pivots, pivotIds, relationshipsNamespace, cancellationToken);
        }

        if (pivots.SlicerCaches.Count > 0)
        {
            pending["extLst"] = () => WriteWorkbookExtensionsAsync(xml, null, pivots, pivotIds, cancellationToken);
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
                        await WriteSheetListAsync(xml, sheetList, added, relationshipsNamespace);
                        await reader.SkipAsync();

                        break;

                    case "definedNames":
                        pending.Remove(name);
                        await CopyDefinedNamesAsync(xml, reader, book, sheetList, removal, SlicerNames(pivots));

                        break;

                    case "pivotCaches" when pivots.PivotCaches.Count > 0:
                        pending.Remove(name);
                        await WritePivotCachesAsync(xml, (XElement)await XNode.ReadFromAsync(reader, cancellationToken), pivots, pivotIds, relationshipsNamespace, cancellationToken);

                        break;

                    case "extLst" when pivots.SlicerCaches.Count > 0:
                        pending.Remove(name);
                        await WriteWorkbookExtensionsAsync(xml, (XElement)await XNode.ReadFromAsync(reader, cancellationToken), pivots, pivotIds, cancellationToken);

                        break;

                    case "calcPr" when hasFormulas == true:
                        pending.Remove(name);
                        await WriteCalcPropertiesAsync(xml, reader);
                        await reader.SkipAsync();

                        break;

                    // Every external link is left out, so the list of them goes too.
                    case "externalReferences" when removal.IsActive == true:
                        await reader.SkipAsync();

                        break;

                    // The active tab and the first tab shown count positions in the sheet list.
                    case "bookViews" when sheetList.HasDropped == true:
                        XElement views = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);
                        MoveTabs(views, sheetList);
                        await views.WriteToAsync(xml, cancellationToken);

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

    /// <summary>The defined names of the new slicer caches, as raw definedName elements: Excel keeps one per cache, with #N/A.</summary>
    private static List<string> SlicerNames(PivotPackage pivots)
    {
        return [.. pivots.DefinedNames.Select(name => "<definedName name=\"" + System.Security.SecurityElement.Escape(name) + "\">#N/A</definedName>")];
    }

    /// <summary>pivotCaches: the original's entries, when it has the element, and one per new cache.</summary>
    private static async Task WritePivotCachesAsync(XmlWriter xml, XElement? original, PivotPackage pivots, List<string> pivotIds, string relationshipsNamespace, CancellationToken cancellationToken)
    {
        if (pivots.PivotCaches.Count == 0)
        {
            if (original is not null)
            {
                await original.WriteToAsync(xml, cancellationToken);
            }

            return;
        }

        XNamespace relationships = relationshipsNamespace;
        XElement caches = original ?? new XElement(XName.Get("pivotCaches", Ooxml.SpreadsheetNamespace));
        XNamespace main = caches.Name.Namespace;

        foreach ((int cacheId, int relationship) in pivots.PivotCaches)
        {
            caches.Add(new XElement(main + "pivotCache", new XAttribute("cacheId", SheetXml.Text(cacheId)), new XAttribute(relationships + "id", pivotIds[relationship])));
        }

        await caches.WriteToAsync(xml, cancellationToken);
    }

    /// <summary>The workbook's extLst: the original's extensions, with the new slicer caches in the x14 list of them.</summary>
    private static async Task WriteWorkbookExtensionsAsync(XmlWriter xml, XElement? original, PivotPackage pivots, List<string> pivotIds, CancellationToken cancellationToken)
    {
        if (pivots.SlicerCaches.Count == 0)
        {
            if (original is not null)
            {
                await original.WriteToAsync(xml, cancellationToken);
            }

            return;
        }

        XElement extensions = original ?? new XElement(XName.Get("extLst", Ooxml.SpreadsheetNamespace));

        foreach (int relationship in pivots.SlicerCaches)
        {
            AddExtension(extensions, PivotXml.WorkbookSlicerCachesUri, "slicerCaches", "slicerCache", pivotIds[relationship]);
        }

        await extensions.WriteToAsync(xml, cancellationToken);
    }

    /// <summary>Points activeTab and firstSheet at the positions their sheets have in the list written, or at the next sheet kept.</summary>
    private static void MoveTabs(XElement views, SheetList sheetList)
    {
        foreach (XElement view in views.Elements().Where(element => element.Name.LocalName == "workbookView"))
        {
            foreach (string attribute in new[] { "activeTab", "firstSheet" })
            {
                if (int.TryParse((string?)view.Attribute(attribute), NumberStyles.None, CultureInfo.InvariantCulture, out int position) == true)
                {
                    view.SetAttributeValue(attribute, SheetXml.Text(sheetList.MapNearest(position)));
                }
            }
        }
    }

    private static async Task WriteSheetListAsync(XmlWriter xml, SheetList sheetList, List<NewSheet> added, string relationshipsNamespace)
    {
        await xml.WriteStartElementAsync(null, "sheets", null);

        foreach (SheetSlot slot in sheetList.Slots)
        {
            string name;
            string relationshipId;
            int sheetId;
            SheetVisibility visibility;

            if (slot.Model is Worksheet sheet)
            {
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

                name = sheet.Name;
                visibility = sheet.Visibility;
            }
            else if (slot.Original is DeclaredSheet original)
            {
                // A chart, macro or dialog sheet the model does not hold, written as it was.
                name = original.Name;
                relationshipId = original.RelationshipId;
                sheetId = original.SheetId;
                visibility = original.Visibility;
            }
            else
            {
                continue;
            }

            await xml.WriteStartElementAsync(null, "sheet", null);
            await xml.WriteAttributeStringAsync(null, "name", null, name);
            await xml.WriteAttributeStringAsync(null, "sheetId", null, SheetXml.Text(sheetId));

            if (visibility == SheetVisibility.Hidden)
            {
                await xml.WriteAttributeStringAsync(null, "state", null, "hidden");
            }
            else if (visibility == SheetVisibility.VeryHidden)
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
    /// adds the filter ranges the model needs. localSheetId is moved to the sheet's position in
    /// the list written, and a name of a sheet left out goes with it. When active content is
    /// removed, so are the names that start macros, and a name that refers to another workbook
    /// becomes #REF!.
    /// </summary>
    private static async Task CopyDefinedNamesAsync(XmlWriter xml, XmlReader reader, Workbook book, SheetList sheetList, ActiveContentRemoval removal, List<string> added)
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
                bool hasSheet = int.TryParse(reader.GetAttribute("localSheetId"), NumberStyles.None, CultureInfo.InvariantCulture, out int localSheet);
                int written = -1;
                bool keep = true;

                if (hasSheet == true)
                {
                    keep = sheetList.TryMapOriginal(localSheet, out written);
                }

                bool replace = isFilter == true && keep == true && IsFilterChanged(sheetList, written) == true;

                if (isFilter == true && keep == true)
                {
                    filterNamesPresent.Add(written);
                }

                string declaration = " xmlns=\"" + reader.NamespaceURI + "\"";
                string outer = await reader.ReadOuterXmlAsync();

                if (keep == false || replace == true)
                {
                    continue;
                }

                XElement name = XElement.Parse(outer);
                bool changed = false;

                if (removal.IsActive == true)
                {
                    if (StartsMacro(name, sheetList) == true)
                    {
                        continue;
                    }

                    if (ActiveContentRules.HasExternalReference(name.Value) == true)
                    {
                        name.Value = "#REF!";
                        changed = true;
                    }
                }

                if (hasSheet == true && written != localSheet)
                {
                    name.SetAttributeValue("localSheetId", SheetXml.Text(written));
                    changed = true;
                }

                if (changed == true)
                {
                    outer = name.ToString(SaveOptions.DisableFormatting);
                }

                kept.Add(RemoveFirst(outer, declaration));
            }
        }

        await reader.ReadAsync();
        kept.AddRange(added);
        await WriteFilterNamesAsync(xml, book, sheetList, kept, false, filterNamesPresent);
    }

    private static bool IsFilterChanged(SheetList sheetList, int written)
    {
        return written >= 0 && written < sheetList.Slots.Count && sheetList.Slots[written].Model is Worksheet sheet && sheet.FilterChanged == true;
    }

    /// <summary>
    /// True for a defined name that starts or belongs to a macro: Auto_Open and its kin, a name
    /// marked as an Excel 4.0 macro, function or VBA procedure, or one that points into a macro
    /// or dialog sheet that is left out.
    /// </summary>
    private static bool StartsMacro(XElement name, SheetList sheetList)
    {
        string text = (string?)name.Attribute("name") ?? string.Empty;

        if (AutoMacroNames.Any(prefix => text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) == true) == true)
        {
            return true;
        }

        foreach (string flag in new[] { "xlm", "function", "vbProcedure" })
        {
            if ((string?)name.Attribute(flag) is "1" or "true")
            {
                return true;
            }
        }

        return sheetList.DroppedNames.Any(sheet => RefersToSheet(name.Value, sheet) == true);
    }

    /// <summary>True when a formula refers to the sheet by name: Sheet1!A1 or 'Sheet 1'!A1.</summary>
    private static bool RefersToSheet(string formula, string sheet)
    {
        if (formula.Contains("'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!", StringComparison.OrdinalIgnoreCase) == true)
        {
            return true;
        }

        string plain = sheet + "!";
        int index = formula.IndexOf(plain, StringComparison.OrdinalIgnoreCase);

        while (index >= 0)
        {
            if (index == 0 || (char.IsLetterOrDigit(formula[index - 1]) == false && formula[index - 1] is not ('_' or '.')))
            {
                return true;
            }

            index = formula.IndexOf(plain, index + 1, StringComparison.OrdinalIgnoreCase);
        }

        return false;
    }

    private static async Task WriteFilterNamesAsync(XmlWriter xml, Workbook book, SheetList sheetList, List<string> kept, bool all, HashSet<int>? present = null)
    {
        List<(int Index, string Range)> filters = [];

        foreach (Worksheet sheet in book.Sheets)
        {
            int position = sheetList.WrittenIndexOf(sheet);
            bool needsName = all == true || sheet.FilterChanged == true || sheet.Source is null || present?.Contains(position) == false;

            if (needsName == true && sheet.AutoFilter == true && TryGetFilterEnd(sheet, out CellAddress end) == true)
            {
                filters.Add((position, "'" + sheet.Name.Replace("'", "''", StringComparison.Ordinal) + "'!$A$1:$" + CellAddress.ColumnName(end.Column) + "$" + SheetXml.Text(end.Row + 1)));
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

    private static async Task WritePatchedWorkbookRelationshipsAsync(XmlWriter xml, ZipArchive original, string relsPath, List<NewSheet> added, string? addedStylesId, bool removeMacros, IReadOnlySet<string> removedIds, PivotPackage pivots, List<string> pivotIds, CancellationToken cancellationToken)
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

                bool removed = reader.GetAttribute("Id") is string id && removedIds.Contains(id) == true;

                if (removed == true || Ooxml.IsRelationship(relationshipType, "calcChain") == true || (removeMacros == true && relationshipType.EndsWith("/vbaProject", StringComparison.Ordinal) == true))
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

        await WritePivotRelationshipsAsync(xml, pivots, pivotIds);
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

    private static async Task WritePatchedContentTypesAsync(XmlWriter xml, ZipArchive original, Workbook book, string workbookPath, List<NewSheet> added, string? calcChainPath, bool addStyles, string? stylesPath, DrawingPackage drawings, PivotPackage pivots, ActiveContentRemoval removal, CancellationToken cancellationToken)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Types", Ooxml.ContentTypesNamespace);
        HashSet<string> defaults = new(StringComparer.OrdinalIgnoreCase);
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
                bool macro = book.DropsVba == true && (string.Equals(contentType, VbaProjectContentType, StringComparison.OrdinalIgnoreCase) == true || (partName is not null && Workbook.IsVbaPart(partName) == true));

                // An override for a part that is left out would name a part the package does not have.
                bool removed = reader.LocalName == "Override" && partName is not null && removal.Keeps(partName.TrimStart('/')) == false;

                if (calcChain == true || macro == true || removed == true)
                {
                    await reader.SkipAsync();

                    continue;
                }

                // Without its VBA project the workbook is a plain one, and says so.
                if (book.DropsVba == true && MacroEnabledTypes.TryGetValue(contentType, out string? plain) == true)
                {
                    contentType = plain;
                }

                await xml.WriteStartElementAsync(null, reader.LocalName, null);

                if (reader.LocalName == "Default")
                {
                    string extension = reader.GetAttribute("Extension") ?? string.Empty;
                    defaults.Add(extension);
                    await xml.WriteAttributeStringAsync(null, "Extension", null, extension);
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
            defaults.UnionWith(["rels", "xml"]);
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

        // A second Default for an extension makes the package invalid, so only missing ones are added.
        foreach ((string extension, string contentType) in drawings.ImageTypes)
        {
            if (defaults.Contains(extension) == false)
            {
                await WriteDefaultAsync(xml, extension, contentType);
            }
        }

        foreach ((string partName, string contentType) in drawings.Overrides.Concat(pivots.Overrides))
        {
            await WriteOverrideAsync(xml, partName, contentType);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>The workbook's relationships to the new pivot caches and slicer caches.</summary>
    private static async Task WritePivotRelationshipsAsync(XmlWriter xml, PivotPackage pivots, List<string> pivotIds)
    {
        for (int index = 0; index < pivots.WorkbookRelationships.Count; index++)
        {
            (string type, string target) = pivots.WorkbookRelationships[index];
            await WriteRelationshipAsync(xml, pivotIds[index], type, target, null);
        }
    }

    private static async Task WriteNewWorkbookRelationshipsAsync(XmlWriter xml, List<NewSheet> parts, string stylesId, PivotPackage pivots, List<string> pivotIds)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);

        foreach (NewSheet part in parts)
        {
            await WriteRelationshipAsync(xml, part.RelationshipId, Ooxml.WorksheetRelationship, part.Target, null);
        }

        await WriteRelationshipAsync(xml, stylesId, Ooxml.StylesRelationship, "styles.xml", null);
        await WritePivotRelationshipsAsync(xml, pivots, pivotIds);
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    private static async Task WriteNewContentTypesAsync(XmlWriter xml, List<NewSheet> parts, DrawingPackage drawings, PivotPackage pivots)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "Types", Ooxml.ContentTypesNamespace);

        foreach ((string extension, string contentType) in new[] { ("rels", Ooxml.RelationshipsContentType), ("xml", "application/xml") }.Concat(drawings.ImageTypes.Select(pair => (pair.Key, pair.Value))))
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

        foreach ((string partName, string contentType) in drawings.Overrides.Concat(pivots.Overrides))
        {
            await WriteOverrideAsync(xml, partName, contentType);
        }

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
