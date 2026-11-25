using System.Globalization;
using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes the parts behind pivot tables and slicers the way Excel 365 writes them: the pivot
/// cache definition and its records, the pivot table definitions, the slicer caches and the
/// slicers. Shared by the document model's saver through <see cref="PivotPackage"/>.
/// </summary>
internal static class PivotXml
{
    public const string X14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    public const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    public const string SlicerDrawingNamespace = "http://schemas.microsoft.com/office/drawing/2010/slicer";
    public const string A14Namespace = "http://schemas.microsoft.com/office/drawing/2010/main";
    public const string PivotDefaultLayoutNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout";

    public const string PivotCacheDefinitionRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    public const string PivotCacheRecordsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
    public const string PivotTableRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    public const string SlicerCacheRelationship = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
    public const string SlicerRelationship = "http://schemas.microsoft.com/office/2007/relationships/slicer";

    public const string PivotCacheDefinitionContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
    public const string PivotCacheRecordsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
    public const string PivotTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
    public const string SlicerCacheContentType = "application/vnd.ms-excel.slicerCache+xml";
    public const string SlicerContentType = "application/vnd.ms-excel.slicer+xml";

    /// <summary>The extension of a worksheet that lists its slicers.</summary>
    public const string SheetSlicerListUri = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";

    /// <summary>The extension of the workbook that lists the slicer caches.</summary>
    public const string WorkbookSlicerCachesUri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";

    /// <summary>A slicer button's height when none is given: 0.26 inch, Excel's default.</summary>
    public const long DefaultSlicerRowHeight = 241300;

    private const string MainNamespace = Ooxml.SpreadsheetNamespace;

    /// <summary>pivotCacheDefinition, with refreshOnLoad so that Excel fetches the source again when it opens the file.</summary>
    public static async Task WriteCacheDefinitionAsync(XmlWriter xml, PivotCacheData data, PivotSource source, bool[] shared, string recordsId, long pivotCacheId, Func<string?, int> formatId)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "pivotCacheDefinition", MainNamespace);
        await xml.WriteAttributeStringAsync(null, "xmlns", null, MainNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, recordsId);
        await xml.WriteAttributeStringAsync(null, "refreshOnLoad", null, "1");
        await xml.WriteAttributeStringAsync(null, "createdVersion", null, "8");
        await xml.WriteAttributeStringAsync(null, "refreshedVersion", null, "8");
        await xml.WriteAttributeStringAsync(null, "minRefreshableVersion", null, "3");
        await xml.WriteAttributeStringAsync(null, "recordCount", null, Text(data.Records.Count));

        await xml.WriteStartElementAsync(null, "cacheSource", null);
        await xml.WriteAttributeStringAsync(null, "type", null, "worksheet");
        await xml.WriteStartElementAsync(null, "worksheetSource", null);

        if (source.TableName is not null)
        {
            await xml.WriteAttributeStringAsync(null, "name", null, source.TableName);
        }
        else
        {
            await xml.WriteAttributeStringAsync(null, "ref", null, source.Range.ToString());
            await xml.WriteAttributeStringAsync(null, "sheet", null, source.Sheet.Name);
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await xml.WriteStartElementAsync(null, "cacheFields", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(data.FieldNames.Count));

        for (int field = 0; field < data.FieldNames.Count; field++)
        {
            await xml.WriteStartElementAsync(null, "cacheField", null);
            await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(data.FieldNames[field]));
            await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(formatId(data.Formats[field])));
            await WriteSharedItemsAsync(xml, data.SharedItems[field], shared[field]);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(null, "extLst", null);
        await xml.WriteStartElementAsync(null, "ext", null);
        await xml.WriteAttributeStringAsync(null, "uri", null, "{725AE2AE-9491-48be-B2B4-4EB974FC3084}");
        await xml.WriteAttributeStringAsync("xmlns", "x14", null, X14Namespace);
        await xml.WriteStartElementAsync("x14", "pivotCacheDefinition", X14Namespace);
        await xml.WriteAttributeStringAsync(null, "pivotCacheId", null, pivotCacheId.ToString(CultureInfo.InvariantCulture));
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>pivotCacheRecords: one &lt;r&gt; per source row, an index into the shared items or the value itself.</summary>
    public static async Task WriteCacheRecordsAsync(XmlWriter xml, PivotCacheData data, bool[] shared)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "pivotCacheRecords", MainNamespace);
        await xml.WriteAttributeStringAsync(null, "xmlns", null, MainNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "r", null, Ooxml.RelationshipsNamespace);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(data.Records.Count));

        for (int record = 0; record < data.Records.Count; record++)
        {
            await xml.WriteStartElementAsync(null, "r", null);

            for (int field = 0; field < data.FieldNames.Count; field++)
            {
                if (shared[field] == true)
                {
                    await xml.WriteStartElementAsync(null, "x", null);
                    await xml.WriteAttributeStringAsync(null, "v", null, Text(data.Records[record][field]));
                    await xml.WriteEndElementAsync();

                    continue;
                }

                await WriteItemAsync(xml, data.Value(record, field));
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>pivotTableDefinition, with the rows and columns Excel would show and the x14 extension Excel 2010 and later write.</summary>
    public static async Task WritePivotTableAsync(XmlWriter xml, CreatedPivot pivot, PivotRendering rendering, Func<string?, int> formatId)
    {
        PivotTable definition = pivot.Definition;
        PivotGrid grid = rendering.Grid;
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "pivotTableDefinition", MainNamespace);
        await xml.WriteAttributeStringAsync(null, "xmlns", null, MainNamespace);
        await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(definition.Name));
        await xml.WriteAttributeStringAsync(null, "cacheId", null, Text(pivot.CacheId));

        foreach (string attribute in new[] { "applyNumberFormats", "applyBorderFormats", "applyFontFormats", "applyPatternFormats", "applyAlignmentFormats" })
        {
            await xml.WriteAttributeStringAsync(null, attribute, null, "0");
        }

        await xml.WriteAttributeStringAsync(null, "applyWidthHeightFormats", null, "1");
        await xml.WriteAttributeStringAsync(null, "dataCaption", null, "Values");
        await xml.WriteAttributeStringAsync(null, "updatedVersion", null, "8");
        await xml.WriteAttributeStringAsync(null, "minRefreshableVersion", null, "3");
        await xml.WriteAttributeStringAsync(null, "useAutoFormatting", null, "1");

        // rowGrandTotals is the grand total of the row items, the row at the bottom (Excel 365 checked through COM); colGrandTotals the column at the right.
        if (definition.GrandTotalRow == false)
        {
            await xml.WriteAttributeStringAsync(null, "rowGrandTotals", null, "0");
        }

        if (definition.GrandTotalColumn == false)
        {
            await xml.WriteAttributeStringAsync(null, "colGrandTotals", null, "0");
        }

        await xml.WriteAttributeStringAsync(null, "itemPrintTitles", null, "1");
        await xml.WriteAttributeStringAsync(null, "createdVersion", null, "8");
        await xml.WriteAttributeStringAsync(null, "indent", null, "0");

        switch (definition.Layout)
        {
            case PivotLayout.Outline:
                await xml.WriteAttributeStringAsync(null, "compact", null, "0");
                await xml.WriteAttributeStringAsync(null, "outline", null, "1");
                await xml.WriteAttributeStringAsync(null, "outlineData", null, "1");
                await xml.WriteAttributeStringAsync(null, "compactData", null, "0");

                break;
            case PivotLayout.Tabular:
                await xml.WriteAttributeStringAsync(null, "compact", null, "0");
                await xml.WriteAttributeStringAsync(null, "compactData", null, "0");

                break;
            default:
                await xml.WriteAttributeStringAsync(null, "outline", null, "1");
                await xml.WriteAttributeStringAsync(null, "outlineData", null, "1");

                break;
        }

        await xml.WriteAttributeStringAsync(null, "multipleFieldFilters", null, "0");

        await xml.WriteStartElementAsync(null, "location", null);
        await xml.WriteAttributeStringAsync(null, "ref", null, grid.Location.ToString());
        await xml.WriteAttributeStringAsync(null, "firstHeaderRow", null, Text(grid.FirstHeaderRow));
        await xml.WriteAttributeStringAsync(null, "firstDataRow", null, Text(grid.FirstDataRow));
        await xml.WriteAttributeStringAsync(null, "firstDataCol", null, Text(grid.FirstDataColumn));

        if (rendering.PageFields.Count > 0)
        {
            await xml.WriteAttributeStringAsync(null, "rowPageCount", null, Text(rendering.PageFields.Count));
            await xml.WriteAttributeStringAsync(null, "colPageCount", null, "1");
        }

        await xml.WriteEndElementAsync();

        HashSet<int> dataFields = [.. rendering.Values.Select(value => value.Field)];
        await xml.WriteStartElementAsync(null, "pivotFields", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(grid.FieldNames.Count));

        for (int field = 0; field < grid.FieldNames.Count; field++)
        {
            await xml.WriteStartElementAsync(null, "pivotField", null);

            switch (rendering.Axes[field])
            {
                case PivotAxis.Row:
                    await xml.WriteAttributeStringAsync(null, "axis", null, "axisRow");

                    break;
                case PivotAxis.Column:
                    await xml.WriteAttributeStringAsync(null, "axis", null, "axisCol");

                    break;
                case PivotAxis.Page:
                    await xml.WriteAttributeStringAsync(null, "axis", null, "axisPage");

                    break;
            }

            if (dataFields.Contains(field) == true)
            {
                await xml.WriteAttributeStringAsync(null, "dataField", null, "1");
            }

            if (definition.Layout != PivotLayout.Compact)
            {
                await xml.WriteAttributeStringAsync(null, "compact", null, "0");
            }

            if (definition.Layout == PivotLayout.Tabular)
            {
                await xml.WriteAttributeStringAsync(null, "outline", null, "0");
            }

            await xml.WriteAttributeStringAsync(null, "showAll", null, "0");
            IReadOnlyList<PivotItem> items = grid.Items[field];

            if (items.Count > 0)
            {
                await xml.WriteStartElementAsync(null, "items", null);
                await xml.WriteAttributeStringAsync(null, "count", null, Text(items.Count));

                foreach (PivotItem item in items)
                {
                    await xml.WriteStartElementAsync(null, "item", null);

                    if (item.IsMember == false)
                    {
                        await xml.WriteAttributeStringAsync(null, "t", null, "default");
                    }
                    else
                    {
                        if (item.Hidden == true)
                        {
                            await xml.WriteAttributeStringAsync(null, "h", null, "1");
                        }

                        await xml.WriteAttributeStringAsync(null, "x", null, Text(item.SharedIndex));
                    }

                    await xml.WriteEndElementAsync();
                }

                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await WriteFieldsAsync(xml, "rowFields", grid.RowFields);
        await WriteLinesAsync(xml, "rowItems", grid.RowLines, false);
        await WriteFieldsAsync(xml, "colFields", grid.ColumnFields);
        await WriteLinesAsync(xml, "colItems", grid.ColumnLines, grid.ColumnFields.Contains(PivotGrid.ValuesField));

        if (rendering.PageFields.Count > 0)
        {
            await xml.WriteStartElementAsync(null, "pageFields", null);
            await xml.WriteAttributeStringAsync(null, "count", null, Text(rendering.PageFields.Count));

            foreach (int field in rendering.PageFields)
            {
                await xml.WriteStartElementAsync(null, "pageField", null);
                await xml.WriteAttributeStringAsync(null, "fld", null, Text(field));
                await xml.WriteAttributeStringAsync(null, "hier", null, "-1");
                await xml.WriteEndElementAsync();
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteStartElementAsync(null, "dataFields", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(rendering.Values.Count));

        foreach ((PivotValue value, int field) in rendering.Values)
        {
            await xml.WriteStartElementAsync(null, "dataField", null);
            await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(value.DisplayName));
            await xml.WriteAttributeStringAsync(null, "fld", null, Text(field));

            if (value.Aggregation != PivotAggregation.Sum)
            {
                await xml.WriteAttributeStringAsync(null, "subtotal", null, value.Aggregation.ToString().ToLowerInvariant());
            }

            await xml.WriteAttributeStringAsync(null, "baseField", null, "0");
            await xml.WriteAttributeStringAsync(null, "baseItem", null, "0");

            if (value.NumberFormat is not null)
            {
                await xml.WriteAttributeStringAsync(null, "numFmtId", null, Text(formatId(value.NumberFormat)));
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();

        await xml.WriteStartElementAsync(null, "pivotTableStyleInfo", null);
        await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(definition.Style));
        await xml.WriteAttributeStringAsync(null, "showRowHeaders", null, "1");
        await xml.WriteAttributeStringAsync(null, "showColHeaders", null, "1");
        await xml.WriteAttributeStringAsync(null, "showRowStripes", null, "0");
        await xml.WriteAttributeStringAsync(null, "showColStripes", null, "0");
        await xml.WriteAttributeStringAsync(null, "showLastColumn", null, "1");
        await xml.WriteEndElementAsync();

        // Excel 2010 and later hide the row with the word Values when the values are the only column field.
        await xml.WriteStartElementAsync(null, "extLst", null);
        await xml.WriteStartElementAsync(null, "ext", null);
        await xml.WriteAttributeStringAsync(null, "uri", null, "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}");
        await xml.WriteAttributeStringAsync("xmlns", "x14", null, X14Namespace);
        await xml.WriteStartElementAsync("x14", "pivotTableDefinition", X14Namespace);
        await xml.WriteAttributeStringAsync(null, "hideValuesRow", null, "1");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>A slicer cache: the field, the pivot tables it filters (by the sheetId of their sheet) and the items with s="1" for selected.</summary>
    public static async Task WriteSlicerCacheAsync(XmlWriter xml, SlicerCacheModel cache, PivotCacheData data, Func<Worksheet, int> sheetId, long pivotCacheId)
    {
        int field = data.FieldIndex(cache.Field);
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "slicerCacheDefinition", X14Namespace);
        await xml.WriteAttributeStringAsync(null, "xmlns", null, X14Namespace);
        await xml.WriteAttributeStringAsync("xmlns", "mc", null, MarkupCompatibilityNamespace);
        await xml.WriteAttributeStringAsync("mc", "Ignorable", MarkupCompatibilityNamespace, "x");
        await xml.WriteAttributeStringAsync("xmlns", "x", null, MainNamespace);
        await xml.WriteAttributeStringAsync(null, "name", null, cache.Name);
        await xml.WriteAttributeStringAsync(null, "sourceName", null, CellText.Encode(cache.Field));
        await xml.WriteStartElementAsync(null, "pivotTables", null);

        foreach (CreatedPivot pivot in cache.PivotTables)
        {
            await xml.WriteStartElementAsync(null, "pivotTable", null);
            await xml.WriteAttributeStringAsync(null, "tabId", null, Text(sheetId(pivot.Sheet)));
            await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(pivot.Definition.Name));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync(null, "data", null);
        await xml.WriteStartElementAsync(null, "tabular", null);
        await xml.WriteAttributeStringAsync(null, "pivotCacheId", null, pivotCacheId.ToString(CultureInfo.InvariantCulture));
        await xml.WriteStartElementAsync(null, "items", null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(data.SharedItems[field].Count));

        foreach (int shared in data.SortedItems(field))
        {
            await xml.WriteStartElementAsync(null, "i", null);
            await xml.WriteAttributeStringAsync(null, "x", null, Text(shared));

            if (cache.Selected is null || cache.Selected.Any(item => PivotRegistry.Matches(data.SharedItems[field][shared], item) == true) == true)
            {
                await xml.WriteAttributeStringAsync(null, "s", null, "1");
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>The slicers part of a sheet: one &lt;slicer&gt; per slicer the sheet shows.</summary>
    public static async Task WriteSlicersAsync(XmlWriter xml, IEnumerable<CreatedSlicer> slicers)
    {
        await xml.WriteStartDocumentAsync(true);
        await xml.WriteStartElementAsync(null, "slicers", X14Namespace);
        await xml.WriteAttributeStringAsync(null, "xmlns", null, X14Namespace);
        await xml.WriteAttributeStringAsync("xmlns", "mc", null, MarkupCompatibilityNamespace);
        await xml.WriteAttributeStringAsync("mc", "Ignorable", MarkupCompatibilityNamespace, "x");
        await xml.WriteAttributeStringAsync("xmlns", "x", null, MainNamespace);

        foreach (CreatedSlicer slicer in slicers)
        {
            Slicer definition = slicer.Definition;
            await xml.WriteStartElementAsync(null, "slicer", null);
            await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(slicer.Name));
            await xml.WriteAttributeStringAsync(null, "cache", null, slicer.Cache.Name);
            await xml.WriteAttributeStringAsync(null, "caption", null, CellText.Encode(definition.Caption ?? slicer.Cache.Field));

            if (definition.Columns != 1)
            {
                await xml.WriteAttributeStringAsync(null, "columnCount", null, Text(definition.Columns));
            }

            if (definition.ShowCaption == false)
            {
                await xml.WriteAttributeStringAsync(null, "showCaption", null, "0");
            }

            if (string.Equals(definition.Style, "SlicerStyleLight1", StringComparison.Ordinal) == false)
            {
                await xml.WriteAttributeStringAsync(null, "style", null, CellText.Encode(definition.Style));
            }

            long rowHeight = DefaultSlicerRowHeight;

            if (definition.ItemHeight is double points)
            {
                rowHeight = (long)Math.Round(points * 12700);
            }

            await xml.WriteAttributeStringAsync(null, "rowHeight", null, rowHeight.ToString(CultureInfo.InvariantCulture));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndDocumentAsync();
    }

    /// <summary>The ext element that lists a sheet's slicers, for a worksheet's extLst.</summary>
    public static async Task WriteSheetSlicerListAsync(XmlWriter xml, string relationshipId)
    {
        await xml.WriteStartElementAsync(null, "ext", null);
        await xml.WriteAttributeStringAsync(null, "uri", null, SheetSlicerListUri);
        await xml.WriteAttributeStringAsync("xmlns", "x14", null, X14Namespace);
        await xml.WriteStartElementAsync("x14", "slicerList", X14Namespace);
        await xml.WriteStartElementAsync("x14", "slicer", X14Namespace);
        await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, relationshipId);
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>The ext element that lists the workbook's slicer caches, for the workbook's extLst.</summary>
    public static async Task WriteWorkbookSlicerCachesAsync(XmlWriter xml, IEnumerable<string> relationshipIds)
    {
        await xml.WriteStartElementAsync(null, "ext", null);
        await xml.WriteAttributeStringAsync(null, "uri", null, WorkbookSlicerCachesUri);
        await xml.WriteAttributeStringAsync("xmlns", "x14", null, X14Namespace);
        await xml.WriteStartElementAsync("x14", "slicerCaches", X14Namespace);

        foreach (string id in relationshipIds)
        {
            await xml.WriteStartElementAsync("x14", "slicerCache", X14Namespace);
            await xml.WriteAttributeStringAsync("r", "id", Ooxml.RelationshipsNamespace, id);
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    /// <summary>
    /// The graphic frame of a slicer in a drawing, inside mc:AlternateContent as Excel writes it,
    /// with the rectangle that versions before 2010 show in its place.
    /// </summary>
    public static async Task WriteSlicerFrameAsync(XmlWriter xml, string name, int shapeId, long left, long top, long width, long height)
    {
        const string Drawing = DrawingXml.SpreadsheetDrawingNamespace;
        const string Main = DrawingXml.DrawingMainNamespace;

        await xml.WriteStartElementAsync("mc", "AlternateContent", MarkupCompatibilityNamespace);
        await xml.WriteStartElementAsync("mc", "Choice", MarkupCompatibilityNamespace);
        await xml.WriteAttributeStringAsync("xmlns", "a14", null, A14Namespace);
        await xml.WriteAttributeStringAsync(null, "Requires", null, "a14");
        await xml.WriteStartElementAsync("xdr", "graphicFrame", Drawing);
        await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
        await xml.WriteStartElementAsync("xdr", "nvGraphicFramePr", Drawing);
        await xml.WriteStartElementAsync("xdr", "cNvPr", Drawing);
        await xml.WriteAttributeStringAsync(null, "id", null, Text(shapeId));
        await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(name));
        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("xdr", "cNvGraphicFramePr", Drawing, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "xfrm", Drawing);
        await WriteOffsetAsync(xml, 0, 0, 0, 0);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "graphic", Main);
        await xml.WriteStartElementAsync("a", "graphicData", Main);
        await xml.WriteAttributeStringAsync(null, "uri", null, SlicerDrawingNamespace);
        await xml.WriteStartElementAsync("sle", "slicer", SlicerDrawingNamespace);
        await xml.WriteAttributeStringAsync(null, "name", null, CellText.Encode(name));
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();

        await xml.WriteStartElementAsync("mc", "Fallback", MarkupCompatibilityNamespace);
        await xml.WriteStartElementAsync("xdr", "sp", Drawing);
        await xml.WriteAttributeStringAsync(null, "macro", null, string.Empty);
        await xml.WriteAttributeStringAsync(null, "textlink", null, string.Empty);
        await xml.WriteStartElementAsync("xdr", "nvSpPr", Drawing);
        await xml.WriteStartElementAsync("xdr", "cNvPr", Drawing);

        // Excel writes 0 here for every slicer, which the schema's unique ids do not allow; the frame's id is the shape's.
        await xml.WriteAttributeStringAsync(null, "id", null, Text(shapeId));
        await xml.WriteAttributeStringAsync(null, "name", null, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "cNvSpPr", Drawing);
        await xml.WriteStartElementAsync("a", "spLocks", Main);
        await xml.WriteAttributeStringAsync(null, "noTextEdit", null, "1");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "spPr", Drawing);
        await xml.WriteStartElementAsync("a", "xfrm", Main);
        await WriteOffsetAsync(xml, left, top, width, height);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "prstGeom", Main);
        await xml.WriteAttributeStringAsync(null, "prst", null, "rect");
        await xml.WriteElementStringAsync("a", "avLst", Main, string.Empty);
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "solidFill", Main);
        await WritePresetColorAsync(xml, "white");
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "ln", Main);
        await xml.WriteAttributeStringAsync(null, "w", null, "1");
        await xml.WriteStartElementAsync("a", "solidFill", Main);
        await WritePresetColorAsync(xml, "green");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("xdr", "txBody", Drawing);
        await xml.WriteStartElementAsync("a", "bodyPr", Main);
        await xml.WriteAttributeStringAsync(null, "vertOverflow", null, "clip");
        await xml.WriteAttributeStringAsync(null, "horzOverflow", null, "clip");
        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "lstStyle", Main, string.Empty);
        await xml.WriteStartElementAsync("a", "p", Main);
        await xml.WriteStartElementAsync("a", "r", Main);
        await xml.WriteStartElementAsync("a", "rPr", Main);
        await xml.WriteAttributeStringAsync(null, "lang", null, "en-US");
        await xml.WriteAttributeStringAsync(null, "sz", null, "1100");
        await xml.WriteEndElementAsync();
        await xml.WriteElementStringAsync("a", "t", Main, "This shape represents a slicer. Slicers are supported in Excel 2010 or later.\n\nIf the shape was modified in an earlier version of Excel, or if the workbook was saved in Excel 2003 or earlier, the slicer cannot be used.");
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteOffsetAsync(XmlWriter xml, long left, long top, long width, long height)
    {
        await xml.WriteStartElementAsync("a", "off", DrawingXml.DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "x", null, left.ToString(CultureInfo.InvariantCulture));
        await xml.WriteAttributeStringAsync(null, "y", null, top.ToString(CultureInfo.InvariantCulture));
        await xml.WriteEndElementAsync();
        await xml.WriteStartElementAsync("a", "ext", DrawingXml.DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "cx", null, width.ToString(CultureInfo.InvariantCulture));
        await xml.WriteAttributeStringAsync(null, "cy", null, height.ToString(CultureInfo.InvariantCulture));
        await xml.WriteEndElementAsync();
    }

    private static async Task WritePresetColorAsync(XmlWriter xml, string color)
    {
        await xml.WriteStartElementAsync("a", "prstClr", DrawingXml.DrawingMainNamespace);
        await xml.WriteAttributeStringAsync(null, "val", null, color);
        await xml.WriteEndElementAsync();
    }

    private static async Task WriteFieldsAsync(XmlWriter xml, string name, IReadOnlyList<int> fields)
    {
        if (fields.Count == 0)
        {
            return;
        }

        await xml.WriteStartElementAsync(null, name, null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(fields.Count));

        foreach (int field in fields)
        {
            await xml.WriteStartElementAsync(null, "field", null);
            await xml.WriteAttributeStringAsync(null, "x", null, Text(field));
            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>rowItems or colItems: per line its type, how many items it repeats from the line before, its value field and its items.</summary>
    private static async Task WriteLinesAsync(XmlWriter xml, string name, IReadOnlyList<PivotLine> lines, bool hasValues)
    {
        await xml.WriteStartElementAsync(null, name, null);
        await xml.WriteAttributeStringAsync(null, "count", null, Text(lines.Count));

        foreach (PivotLine line in lines)
        {
            await xml.WriteStartElementAsync(null, "i", null);

            if (line.Kind == PivotLineKind.Subtotal)
            {
                await xml.WriteAttributeStringAsync(null, "t", null, "default");
            }
            else if (line.Kind == PivotLineKind.Grand)
            {
                await xml.WriteAttributeStringAsync(null, "t", null, "grand");
            }

            if (line.Repeat > 0)
            {
                await xml.WriteAttributeStringAsync(null, "r", null, Text(line.Repeat));
            }

            if (hasValues == true && line.DataIndex > 0)
            {
                await xml.WriteAttributeStringAsync(null, "i", null, Text(line.DataIndex));
            }

            if (line.Kind == PivotLineKind.Grand)
            {
                await xml.WriteElementStringAsync(null, "x", null, string.Empty);
            }
            else
            {
                for (int index = line.Repeat; index < line.Members.Length; index++)
                {
                    await xml.WriteStartElementAsync(null, "x", null);

                    if (line.Members[index] != 0)
                    {
                        await xml.WriteAttributeStringAsync(null, "v", null, Text(line.Members[index]));
                    }

                    await xml.WriteEndElementAsync();
                }
            }

            await xml.WriteEndElementAsync();
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>sharedItems: what kinds of values the field holds, the range of its numbers and dates, and the items when the field lists them.</summary>
    private static async Task WriteSharedItemsAsync(XmlWriter xml, List<CellValue> items, bool listItems)
    {
        bool text = false;
        bool number = false;
        bool date = false;
        bool blank = false;
        bool other = false;
        bool integers = true;
        double minimum = double.MaxValue;
        double maximum = double.MinValue;
        DateTime earliest = DateTime.MaxValue;
        DateTime latest = DateTime.MinValue;

        foreach (CellValue value in items)
        {
            switch (value.Kind)
            {
                case CellKind.Text:
                    text = true;

                    break;
                case CellKind.Number when value.TryGetNumber(out double plain) == true:
                    number = true;
                    integers &= Math.Floor(plain) == plain && Math.Abs(plain) < 1e15;
                    minimum = Math.Min(minimum, plain);
                    maximum = Math.Max(maximum, plain);

                    break;
                case CellKind.Date when value.TryGetDate(out DateTime day) == true:
                    date = true;

                    if (day < earliest)
                    {
                        earliest = day;
                    }

                    if (day > latest)
                    {
                        latest = day;
                    }

                    break;
                case CellKind.Empty:
                    blank = true;

                    break;
                default:
                    other = true;

                    break;
            }
        }

        int kinds = Convert.ToInt32(text) + Convert.ToInt32(number) + Convert.ToInt32(date) + Convert.ToInt32(other);
        await xml.WriteStartElementAsync(null, "sharedItems", null);

        if (text == false && blank == false)
        {
            await xml.WriteAttributeStringAsync(null, "containsSemiMixedTypes", null, "0");
        }

        if (date == true && text == false && number == false && other == false)
        {
            await xml.WriteAttributeStringAsync(null, "containsNonDate", null, "0");
        }

        if (date == true)
        {
            await xml.WriteAttributeStringAsync(null, "containsDate", null, "1");
        }

        if (text == false)
        {
            await xml.WriteAttributeStringAsync(null, "containsString", null, "0");
        }

        if (blank == true)
        {
            await xml.WriteAttributeStringAsync(null, "containsBlank", null, "1");
        }

        if (kinds > 1)
        {
            await xml.WriteAttributeStringAsync(null, "containsMixedTypes", null, "1");
        }

        if (number == true)
        {
            await xml.WriteAttributeStringAsync(null, "containsNumber", null, "1");

            if (integers == true)
            {
                await xml.WriteAttributeStringAsync(null, "containsInteger", null, "1");
            }

            await xml.WriteAttributeStringAsync(null, "minValue", null, Number(minimum));
            await xml.WriteAttributeStringAsync(null, "maxValue", null, Number(maximum));
        }

        if (date == true)
        {
            await xml.WriteAttributeStringAsync(null, "minDate", null, DateText(earliest));
            await xml.WriteAttributeStringAsync(null, "maxDate", null, DateText(latest));
        }

        if (listItems == true)
        {
            await xml.WriteAttributeStringAsync(null, "count", null, Text(items.Count));

            foreach (CellValue value in items)
            {
                await WriteItemAsync(xml, value);
            }
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>A value as the cache stores it: s, n, d, b, e or m for a blank.</summary>
    private static async Task WriteItemAsync(XmlWriter xml, CellValue value)
    {
        switch (value.Kind)
        {
            case CellKind.Text when value.TryGetText(out string? text) == true:
                await xml.WriteStartElementAsync(null, "s", null);
                await xml.WriteAttributeStringAsync(null, "v", null, CellText.Encode(text));

                break;
            case CellKind.Number when value.TryGetNumber(out double number) == true:
                await xml.WriteStartElementAsync(null, "n", null);
                await xml.WriteAttributeStringAsync(null, "v", null, Number(number));

                break;
            case CellKind.Date when value.TryGetDate(out DateTime date) == true:
                await xml.WriteStartElementAsync(null, "d", null);
                await xml.WriteAttributeStringAsync(null, "v", null, DateText(date));

                break;
            case CellKind.Bool when value.TryGetBool(out bool flag) == true:
                await xml.WriteStartElementAsync(null, "b", null);
                await xml.WriteAttributeStringAsync(null, "v", null, Text(Convert.ToInt32(flag)));

                break;
            case CellKind.Error when value.TryGetError(out string? code) == true:
                await xml.WriteStartElementAsync(null, "e", null);
                await xml.WriteAttributeStringAsync(null, "v", null, code);

                break;
            default:
                await xml.WriteStartElementAsync(null, "m", null);

                break;
        }

        await xml.WriteEndElementAsync();
    }

    private static string Text(int number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }

    private static string Number(double number)
    {
        return number.ToString("R", CultureInfo.InvariantCulture);
    }

    private static string DateText(DateTime date)
    {
        return date.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
    }
}
