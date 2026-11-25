using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Reads what an opened file has of pivot tables, pivot caches, slicers and Excel tables: enough
/// to list them, to let GETPIVOTDATA find values in the pivot tables' cells, and to keep names
/// and ids apart from the ones the package adds. The parts themselves are saved as they were.
/// The cache records are not read. A pivot part that is not valid XML is left out of the lists;
/// Excel repairs it or not when it opens the file.
/// </summary>
internal static class PivotReader
{
    public static async Task<LoadedPivots> LoadAsync(Workbook book, XlsxReader reader, CancellationToken cancellationToken)
    {
        ZipArchive archive = reader.Archive;
        LoadedPivots loaded = new();
        string workbookPath = reader.WorkbookPath;
        Dictionary<string, (string Type, string Target)> workbookRelationships = await XmlParts.ReadRelationshipsAsync(archive, Ooxml.RelationshipsPath(workbookPath), cancellationToken);
        Dictionary<string, int> cacheIds = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string relationshipId, int cacheId) in await ReadPivotCacheIdsAsync(archive, workbookPath, cancellationToken))
        {
            loaded.MaxCacheId = Math.Max(loaded.MaxCacheId, cacheId);

            if (workbookRelationships.TryGetValue(relationshipId, out (string Type, string Target) relationship) == true)
            {
                cacheIds[Ooxml.ResolveTarget(workbookPath, relationship.Target)] = cacheId;
            }
        }

        Dictionary<string, CacheDefinition> caches = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string path, int cacheId) in cacheIds)
        {
            if ((await ReadCacheAsync(archive, path, cacheId, cancellationToken)).IsSuccess(out CacheDefinition? cache) == true)
            {
                caches[path] = cache;
                loaded.CacheDefinitionPaths.Add(path);
                loaded.Caches.Add(cache.Info);

                if (cache.PivotCacheId is long id)
                {
                    loaded.PivotCacheIds.Add(id);
                }
            }
        }

        foreach ((string type, string target) in workbookRelationships.Values)
        {
            if (type.EndsWith("/slicerCache", StringComparison.Ordinal) == true && (await ReadRootAsync(archive, Ooxml.ResolveTarget(workbookPath, target), cancellationToken)).IsSuccess(out XElement? slicerCache) == true && (string?)slicerCache.Attribute("name") is string name)
            {
                loaded.SlicerCacheNames.Add(name);
            }
        }

        foreach (Worksheet sheet in book.Sheets)
        {
            if (sheet.Source is not SheetSource source)
            {
                continue;
            }

            Dictionary<string, (string Type, string Target)> relationships = await XmlParts.ReadRelationshipsAsync(archive, Ooxml.RelationshipsPath(source.Path), cancellationToken);

            foreach ((string type, string target) in relationships.Values)
            {
                string path = Ooxml.ResolveTarget(source.Path, target);

                if (Ooxml.IsRelationship(type, "pivotTable") == true)
                {
                    await TryReadPivotTableAsync(archive, path, sheet, caches, loaded, cancellationToken);
                }
                else if (Ooxml.IsRelationship(type, "table") == true && (await ReadRootAsync(archive, path, cancellationToken)).IsSuccess(out XElement? table) == true)
                {
                    AddTable(loaded, table, sheet);
                }
                else if (type.EndsWith("/slicer", StringComparison.Ordinal) == true && (await ReadRootAsync(archive, path, cancellationToken)).IsSuccess(out XElement? slicers) == true)
                {
                    foreach (XElement slicer in slicers.Elements().Where(element => element.Name.LocalName == "slicer"))
                    {
                        if ((string?)slicer.Attribute("name") is string name)
                        {
                            loaded.SlicerNames.Add(name);
                        }
                    }
                }
            }
        }

        return loaded;
    }

    /// <summary>The pivotCache entries of workbook.xml: relationship id and cache id.</summary>
    private static async Task<List<(string RelationshipId, int CacheId)>> ReadPivotCacheIdsAsync(ZipArchive archive, string workbookPath, CancellationToken cancellationToken)
    {
        List<(string, int)> entries = [];

        if (XmlParts.TryFindEntry(archive, workbookPath, out ZipArchiveEntry? entry) == false)
        {
            return entries;
        }

        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        string section = string.Empty;

        while (await reader.ReadAsync() == true)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.Depth == 1)
            {
                section = reader.LocalName;
            }

            if (section == "pivotCaches" && reader.Depth == 2 && reader.LocalName == "pivotCache"
                && int.TryParse(reader.GetAttribute("cacheId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int cacheId) == true
                && XmlParts.TryGetRelationshipId(reader, out string? id) == true)
            {
                entries.Add((id, cacheId));
            }
        }

        return entries;
    }

    private static async Task<Result<CacheDefinition>> ReadCacheAsync(ZipArchive archive, string path, int cacheId, CancellationToken cancellationToken)
    {
        if ((await ReadRootAsync(archive, path, cancellationToken)).IsFailure(out string? error, out ResultStatus status, out XElement? root) == true)
        {
            return Result<CacheDefinition>.Failure(error, status);
        }

        string source = "external";
        XElement? worksheetSource = Child(Child(root, "cacheSource"), "worksheetSource");

        if (worksheetSource is not null)
        {
            if ((string?)worksheetSource.Attribute("name") is string name)
            {
                source = name;
            }
            else
            {
                string sheet = (string?)worksheetSource.Attribute("sheet") ?? string.Empty;
                source = "'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!" + ((string?)worksheetSource.Attribute("ref") ?? string.Empty);
            }
        }

        List<string> names = [];
        List<List<CellValue>> shared = [];

        foreach (XElement field in Children(Child(root, "cacheFields"), "cacheField"))
        {
            names.Add(CellText.Decode((string?)field.Attribute("name") ?? string.Empty));
            List<CellValue> items = [];

            foreach (XElement item in Child(field, "sharedItems")?.Elements() ?? [])
            {
                items.Add(ItemValue(item));
            }

            shared.Add(items);
        }

        int.TryParse((string?)root.Attribute("recordCount"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int records);
        bool refresh = (string?)root.Attribute("refreshOnLoad") is "1" or "true";
        long? pivotCacheId = null;

        foreach (XElement extension in root.Descendants().Where(element => element.Name.LocalName == "pivotCacheDefinition"))
        {
            if (long.TryParse((string?)extension.Attribute("pivotCacheId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) == true)
            {
                pivotCacheId = id;
            }
        }

        return new CacheDefinition(new PivotCacheInfo(cacheId, source, names, records, refresh), names, shared, pivotCacheId);
    }

    private static async Task TryReadPivotTableAsync(ZipArchive archive, string path, Worksheet sheet, Dictionary<string, CacheDefinition> caches, LoadedPivots loaded, CancellationToken cancellationToken)
    {
        if ((await ReadRootAsync(archive, path, cancellationToken)).IsFailure(out _, out _, out XElement? root) == true)
        {
            return;
        }

        // The pivot table's own relationship names its cache; cacheId is the workbook's number for it.
        CacheDefinition? cache = null;
        Dictionary<string, (string Type, string Target)> relationships = await XmlParts.ReadRelationshipsAsync(archive, Ooxml.RelationshipsPath(path), cancellationToken);

        foreach ((string type, string target) in relationships.Values)
        {
            if (Ooxml.IsRelationship(type, "pivotCacheDefinition") == true)
            {
                caches.TryGetValue(Ooxml.ResolveTarget(path, target), out cache);
            }
        }

        XElement? location = Child(root, "location");

        if (cache is null || location is null || CellRange.TryParse((string?)location.Attribute("ref"), out CellRange range) == false)
        {
            return;
        }

        List<IReadOnlyList<PivotItem>> items = [];

        foreach (XElement field in Children(Child(root, "pivotFields"), "pivotField"))
        {
            int index = items.Count;
            List<PivotItem> list = [];

            foreach (XElement item in Children(Child(field, "items"), "item"))
            {
                string kind = (string?)item.Attribute("t") ?? "data";
                bool hidden = (string?)item.Attribute("h") is "1" or "true";

                if (kind != "data" || int.TryParse((string?)item.Attribute("x"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int shared) == false)
                {
                    list.Add(new PivotItem(CellValue.Empty, -1, hidden, false));

                    continue;
                }

                CellValue value = CellValue.Empty;

                if (index < cache.SharedItems.Count && shared >= 0 && shared < cache.SharedItems[index].Count)
                {
                    value = cache.SharedItems[index][shared];
                }

                if ((string?)item.Attribute("n") is string caption)
                {
                    value = CellValue.Text(CellText.Decode(caption));
                }

                list.Add(new PivotItem(value, shared, hidden, true));
            }

            items.Add(list);
        }

        List<int> rowFields = Fields(Child(root, "rowFields"));
        List<int> columnFields = Fields(Child(root, "colFields"));
        List<PivotDataField> dataFields = [];

        foreach (XElement dataField in Children(Child(root, "dataFields"), "dataField"))
        {
            int.TryParse((string?)dataField.Attribute("fld"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int field);
            dataFields.Add(new PivotDataField(CellText.Decode((string?)dataField.Attribute("name") ?? string.Empty), field));
        }

        // Fields without a pivotField entry get no items, so that an index never goes past the list.
        while (items.Count < cache.Names.Count)
        {
            items.Add([]);
        }

        loaded.Grids.Add(new PivotGrid
        {
            Name = CellText.Decode((string?)root.Attribute("name") ?? string.Empty),
            Sheet = sheet,
            Location = range,
            FirstHeaderRow = Number(location, "firstHeaderRow"),
            FirstDataRow = Number(location, "firstDataRow"),
            FirstDataColumn = Number(location, "firstDataCol"),
            CacheId = cache.Info.CacheId,
            FieldNames = cache.Names,
            Items = items,
            RowFields = rowFields,
            ColumnFields = columnFields,
            RowLines = Lines(Child(root, "rowItems")),
            ColumnLines = Lines(Child(root, "colItems")),
            DataFields = dataFields,
        });
    }

    private static void AddTable(LoadedPivots loaded, XElement table, Worksheet sheet)
    {
        if (CellRange.TryParse((string?)table.Attribute("ref"), out CellRange range) == false)
        {
            return;
        }

        foreach (string attribute in new[] { "displayName", "name" })
        {
            if ((string?)table.Attribute(attribute) is string name && name.Length > 0)
            {
                loaded.Tables.TryAdd(name, (sheet.Name, range));
            }
        }
    }

    /// <summary>rowItems or colItems: each line's members are those it repeats from the line before and its own.</summary>
    private static List<PivotLine> Lines(XElement? element)
    {
        List<PivotLine> lines = [];
        int[] previous = [];

        foreach (XElement line in Children(element, "i"))
        {
            string kind = (string?)line.Attribute("t") ?? "data";
            int repeat = Math.Clamp(Number(line, "r"), 0, previous.Length);
            int dataIndex = Number(line, "i");
            List<int> members = [.. previous[..repeat]];

            foreach (XElement member in Children(line, "x"))
            {
                members.Add(Number(member, "v"));
            }

            PivotLineKind lineKind = kind switch
            {
                "data" => PivotLineKind.Data,
                "grand" => PivotLineKind.Grand,
                "blank" => PivotLineKind.Other,
                _ => PivotLineKind.Subtotal,
            };

            if (lineKind == PivotLineKind.Grand)
            {
                members.Clear();
            }

            int[] array = [.. members];
            lines.Add(new PivotLine(lineKind, array, dataIndex, repeat));
            previous = array;
        }

        return lines;
    }

    private static List<int> Fields(XElement? element)
    {
        List<int> fields = [];

        foreach (XElement field in Children(element, "field"))
        {
            fields.Add(Number(field, "x"));
        }

        return fields;
    }

    private static CellValue ItemValue(XElement item)
    {
        string? text = (string?)item.Attribute("v");

        switch (item.Name.LocalName)
        {
            case "s" when text is not null:
                return CellValue.Text(CellText.Decode(text));
            case "n" when double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == true:
                return CellValue.Number(number);
            case "b":
                return CellValue.Bool(text is "1" or "true");
            case "e" when text is not null:
                return CellValue.Error(text);
            case "d" when DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date) == true:
                return CellValue.Date(date);
            default:
                return CellValue.Empty;
        }
    }

    private static int Number(XElement element, string attribute)
    {
        if (int.TryParse((string?)element.Attribute(attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) == true)
        {
            return number;
        }

        return 0;
    }

    private static XElement? Child(XElement? element, string name)
    {
        return element?.Elements().FirstOrDefault(child => child.Name.LocalName == name);
    }

    private static IEnumerable<XElement> Children(XElement? element, string name)
    {
        if (element is null)
        {
            return [];
        }

        return element.Elements().Where(child => child.Name.LocalName == name);
    }

    /// <summary>A part's root element; NotFound when the part is missing, Invalid when it is not valid XML.</summary>
    private static async Task<Result<XElement>> ReadRootAsync(ZipArchive archive, string path, CancellationToken cancellationToken)
    {
        if (XmlParts.TryFindEntry(archive, path, out ZipArchiveEntry? entry) == false)
        {
            return Result<XElement>.NotFound($"The file has no part {path}.");
        }

        try
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);

            return await XElement.LoadAsync(reader, LoadOptions.None, cancellationToken);
        }
        catch (XmlException exception)
        {
            return Result<XElement>.Invalid($"The file is damaged: {path} is not valid XML ({exception.Message}).");
        }
    }

    private sealed record CacheDefinition(PivotCacheInfo Info, IReadOnlyList<string> Names, IReadOnlyList<List<CellValue>> SharedItems, long? PivotCacheId);
}
