using System.IO.Compression;
using System.Xml;
using Avig.Kalkyl.Writing;

namespace Avig.Kalkyl.Internal;

/// <summary>What a sheet needs for the pivot tables and slicers on it: relationships, and the slicers part for its extLst.</summary>
internal sealed class SheetPivotParts
{
    /// <summary>Relationship types and targets, from the package root, to add to the sheet's .rels.</summary>
    public List<(string Type, string Target)> Relationships { get; } = [];

    /// <summary>The slicers part, when the sheet shows slicers; its relationship must be listed in the sheet's extLst.</summary>
    public string? SlicersTarget { get; set; }
}

/// <summary>
/// Writes the parts of the pivot tables and slicers the package adds, picking names no other
/// part has, and collects what the workbook, the sheets and [Content_Types].xml must point to.
/// </summary>
internal sealed class PivotPackage
{
    private readonly ZipArchive archive;
    private readonly HashSet<string> paths;

    public PivotPackage(ZipArchive archive, IEnumerable<string> existingPaths)
    {
        this.archive = archive;
        paths = new HashSet<string>(existingPaths, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Workbook relationships to add: type and target from the package root.</summary>
    public List<(string Type, string Target)> WorkbookRelationships { get; } = [];

    /// <summary>The pivotCache entries for workbook.xml: cache id and the index of its relationship in <see cref="WorkbookRelationships"/>.</summary>
    public List<(int CacheId, int Relationship)> PivotCaches { get; } = [];

    /// <summary>The indexes in <see cref="WorkbookRelationships"/> of the slicer caches, for the workbook's extLst.</summary>
    public List<int> SlicerCaches { get; } = [];

    /// <summary>The defined names Excel keeps for slicer caches (Slicer_Year), each #N/A.</summary>
    public List<string> DefinedNames { get; } = [];

    public Dictionary<Worksheet, SheetPivotParts> Sheets { get; } = [];

    /// <summary>Overrides for [Content_Types].xml.</summary>
    public List<(string PartName, string ContentType)> Overrides { get; } = [];

    public bool IsEmpty => WorkbookRelationships.Count == 0 && Sheets.Count == 0;

    /// <summary>Writes every part. The pivot tables are rendered already, by <see cref="PivotRegistry.RenderAll"/>.</summary>
    public async Task WriteAsync(Workbook book, StyleRegistry styles, Func<Worksheet, int> sheetId, CancellationToken cancellationToken)
    {
        PivotRegistry registry = book.Pivots;
        Dictionary<string, (PivotCacheData Data, long PivotCacheId)> caches = new(StringComparer.Ordinal);
        HashSet<long> usedIds = [.. registry.Loaded.PivotCacheIds];

        foreach ((int cacheId, PivotSource source, List<CreatedPivot> pivots) in registry.CacheGroups)
        {
            Result<PivotCacheData> read = PivotCacheData.Read(source);

            if (read.IsFailure(out string? error, out _, out PivotCacheData? data) == true)
            {
                throw new InvalidOperationException($"The pivot table cache for {source.Display} cannot be written: {error}");
            }

            long pivotCacheId = 1_000_000L + cacheId;

            while (usedIds.Add(pivotCacheId) == false)
            {
                pivotCacheId++;
            }

            caches[source.Key] = (data, pivotCacheId);
            bool[] shared = new bool[data.FieldNames.Count];

            for (int field = 0; field < shared.Length; field++)
            {
                shared[field] = data.SharedItems[field].Exists(value => value.Kind == CellKind.Text) == true
                    || pivots.Exists(pivot => pivot.Rendering is not null && (pivot.Rendering.Axes[field] != PivotAxis.None || pivot.Rendering.Grid.Items[field].Count > 0)) == true;
            }

            string definitionPath = Unused("xl/pivotCache/pivotCacheDefinition", ".xml");
            string recordsPath = Unused("xl/pivotCache/pivotCacheRecords", ".xml");
            await WritePartAsync(definitionPath, xml => PivotXml.WriteCacheDefinitionAsync(xml, data, source, shared, "rId1", pivotCacheId, format => FormatId(styles, format)), cancellationToken);
            await WritePartAsync(Ooxml.RelationshipsPath(definitionPath), xml => DrawingXml.WriteRelationshipsAsync(xml, [new PartRelationship("rId1", PivotXml.PivotCacheRecordsRelationship, "/" + recordsPath)]), cancellationToken);
            await WritePartAsync(recordsPath, xml => PivotXml.WriteCacheRecordsAsync(xml, data, shared), cancellationToken);
            Overrides.Add(("/" + definitionPath, PivotXml.PivotCacheDefinitionContentType));
            Overrides.Add(("/" + recordsPath, PivotXml.PivotCacheRecordsContentType));
            PivotCaches.Add((cacheId, WorkbookRelationships.Count));
            WorkbookRelationships.Add((PivotXml.PivotCacheDefinitionRelationship, "/" + definitionPath));

            foreach (CreatedPivot pivot in pivots)
            {
                PivotRendering rendering = pivot.Rendering!;
                string tablePath = Unused("xl/pivotTables/pivotTable", ".xml");
                await WritePartAsync(tablePath, xml => PivotXml.WritePivotTableAsync(xml, pivot, rendering, format => FormatId(styles, format)), cancellationToken);
                await WritePartAsync(Ooxml.RelationshipsPath(tablePath), xml => DrawingXml.WriteRelationshipsAsync(xml, [new PartRelationship("rId1", PivotXml.PivotCacheDefinitionRelationship, "/" + definitionPath)]), cancellationToken);
                Overrides.Add(("/" + tablePath, PivotXml.PivotTableContentType));
                SheetParts(pivot.Sheet).Relationships.Add((PivotXml.PivotTableRelationship, "/" + tablePath));
            }
        }

        foreach (SlicerCacheModel slicerCache in registry.SlicerCaches)
        {
            (PivotCacheData data, long pivotCacheId) = caches[slicerCache.SourceKey];
            string path = Unused("xl/slicerCaches/slicerCache", ".xml");
            await WritePartAsync(path, xml => PivotXml.WriteSlicerCacheAsync(xml, slicerCache, data, sheetId, pivotCacheId), cancellationToken);
            Overrides.Add(("/" + path, PivotXml.SlicerCacheContentType));
            SlicerCaches.Add(WorkbookRelationships.Count);
            WorkbookRelationships.Add((PivotXml.SlicerCacheRelationship, "/" + path));
            DefinedNames.Add(slicerCache.Name);
        }

        foreach (IGrouping<Worksheet, CreatedSlicer> sheet in registry.Slicers.GroupBy(slicer => slicer.Sheet))
        {
            string path = Unused("xl/slicers/slicer", ".xml");
            await WritePartAsync(path, xml => PivotXml.WriteSlicersAsync(xml, sheet), cancellationToken);
            Overrides.Add(("/" + path, PivotXml.SlicerContentType));
            SheetPivotParts parts = SheetParts(sheet.Key);
            parts.SlicersTarget = "/" + path;
            parts.Relationships.Add((PivotXml.SlicerRelationship, "/" + path));
        }
    }

    /// <summary>The numFmtId of a number format, registered in the styles when it is new; 0 for General.</summary>
    private static int FormatId(StyleRegistry styles, string? format)
    {
        if (format is null)
        {
            return 0;
        }

        return styles.GetFormatId(format);
    }

    private SheetPivotParts SheetParts(Worksheet sheet)
    {
        if (Sheets.TryGetValue(sheet, out SheetPivotParts? parts) == false)
        {
            parts = new SheetPivotParts();
            Sheets[sheet] = parts;
        }

        return parts;
    }

    private string Unused(string prefix, string suffix)
    {
        int number = 1;

        while (paths.Contains(prefix + SheetXml.Text(number) + suffix) == true)
        {
            number++;
        }

        string path = prefix + SheetXml.Text(number) + suffix;
        paths.Add(path);

        return path;
    }

    private async Task WritePartAsync(string path, Func<XmlWriter, Task> write, CancellationToken cancellationToken)
    {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        await using Stream entryStream = await entry.OpenAsync(cancellationToken);
        await using XmlWriter xml = XmlWriter.Create(entryStream, XlsxWriter.XmlSettings);
        await write(xml);
        await xml.FlushAsync();
    }
}
