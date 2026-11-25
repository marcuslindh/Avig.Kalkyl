using System.Text;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// The pivot tables and slicers of a workbook: those the opened file had (read-only, for
/// <see cref="Workbook.PivotTables"/> and GETPIVOTDATA), and those added since, which the
/// package renders into their sheets and writes when the workbook is saved.
/// </summary>
internal sealed class PivotRegistry
{
    private readonly Workbook book;
    private readonly List<CreatedPivot> created = [];
    private readonly List<SlicerCacheModel> slicerCaches = [];
    private readonly List<CreatedSlicer> slicers = [];

    public PivotRegistry(Workbook book)
    {
        this.book = book;
    }

    /// <summary>What the opened file had.</summary>
    public LoadedPivots Loaded { get; set; } = LoadedPivots.Empty;

    public IReadOnlyList<CreatedPivot> Created => created;

    public IReadOnlyList<SlicerCacheModel> SlicerCaches => slicerCaches;

    public IReadOnlyList<CreatedSlicer> Slicers => slicers;

    /// <summary>Every pivot table GETPIVOTDATA can read: the file's and the rendered new ones.</summary>
    public IEnumerable<PivotGrid> Grids
    {
        get
        {
            foreach (PivotGrid grid in Loaded.Grids)
            {
                yield return grid;
            }

            foreach (CreatedPivot pivot in created)
            {
                if (pivot.Rendering is not null)
                {
                    yield return pivot.Rendering.Grid;
                }
            }
        }
    }

    /// <summary>The new caches, one per source, in the order their first pivot table was added.</summary>
    public IReadOnlyList<(int CacheId, PivotSource Source, List<CreatedPivot> PivotTables)> CacheGroups
    {
        get
        {
            List<(int CacheId, PivotSource Source, List<CreatedPivot> PivotTables)> groups = [];

            foreach (CreatedPivot pivot in created)
            {
                int index = groups.FindIndex(group => string.Equals(group.Source.Key, pivot.Source.Key, StringComparison.Ordinal) == true);

                if (index < 0)
                {
                    groups.Add((Loaded.MaxCacheId + 1 + groups.Count, pivot.Source, [pivot]));

                    continue;
                }

                groups[index].PivotTables.Add(pivot);
            }

            return groups;
        }
    }

    public Result Add(Worksheet sheet, PivotTable definition, CellAddress topLeft)
    {
        if (definition.Values.Count == 0)
        {
            throw new ArgumentException($"The pivot table '{definition.Name}' has no values. Set PivotTable.Values.", nameof(definition));
        }

        List<string> placed = [.. definition.Rows, .. definition.Columns, .. definition.Filters];
        string? twice = placed.GroupBy(name => name.Trim(), StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1).Select(group => group.Key).FirstOrDefault();

        if (twice is not null)
        {
            throw new ArgumentException($"The field '{twice}' is in the pivot table '{definition.Name}' more than once. A field goes in one of Rows, Columns and Filters.", nameof(definition));
        }

        if (IsStrict() == true)
        {
            return Result.Invalid("The workbook is a Strict OOXML file. Pivot tables can be added to transitional files only; save it as an ordinary .xlsx in Excel first.");
        }

        if (TryFindByName(sheet, definition.Name) == true)
        {
            return Result.Conflict($"The sheet '{sheet.Name}' already has a pivot table named '{definition.Name}'.");
        }

        Result<PivotSource> resolved = ResolveSource(sheet, definition.Source);

        if (resolved.IsFailure(out string? error, out ResultStatus status, out PivotSource? source) == true)
        {
            return Result.Failure(error, status);
        }

        CreatedPivot pivot = new(definition, sheet, topLeft, source);
        Result<PivotCacheData> data = PivotCacheData.Read(source);

        if (data.IsFailure(out error, out status, out PivotCacheData? cache) == true)
        {
            return Result.Failure(error, status);
        }

        Result<PivotRendering> rendered = PivotRender.Render(pivot, cache, new Dictionary<int, HashSet<int>>(), new HashSet<int>(), book.Uses1904Dates);

        if (rendered.IsFailure(out error, out status, out PivotRendering? rendering) == true)
        {
            return Result.Failure(error, status);
        }

        foreach (CreatedPivot other in created.Where(other => ReferenceEquals(other.Sheet, sheet) == true && other.Rendering is not null))
        {
            if (Overlaps(other.Rendering!.Area, rendering.Area) == true)
            {
                return Result.Conflict($"The pivot table '{definition.Name}' would cover {rendering.Area}, which overlaps the pivot table '{other.Definition.Name}' at {other.Rendering.Area}.");
            }
        }

        foreach (PivotGrid other in Loaded.Grids.Where(other => ReferenceEquals(other.Sheet, sheet) == true))
        {
            if (Overlaps(other.Location, rendering.Area) == true)
            {
                return Result.Conflict($"The pivot table '{definition.Name}' would cover {rendering.Area}, which overlaps the pivot table '{other.Name}' at {other.Location}.");
            }
        }

        if (source.TableName is null && ReferenceEquals(source.Sheet, sheet) == true && Overlaps(source.Range, rendering.Area) == true)
        {
            return Result.Conflict($"The pivot table '{definition.Name}' would cover {rendering.Area}, which overlaps its own source {source.Range}.");
        }

        created.Add(pivot);
        Result renderedAll = RenderAll(out _);

        if (renderedAll.IsFailure(out error, out status) == true)
        {
            created.Remove(pivot);

            return Result.Failure(error, status);
        }

        return Result.Success();
    }

    public Result AddSlicer(Worksheet sheet, Slicer definition, CellRange area)
    {
        List<CreatedPivot> pivots = [];

        foreach (string name in definition.PivotTables)
        {
            CreatedPivot? found = created.Find(pivot => string.Equals(pivot.Definition.Name, name, StringComparison.OrdinalIgnoreCase) == true);

            if (found is null)
            {
                string known = string.Join(", ", created.Select(pivot => "'" + pivot.Definition.Name + "'"));

                if (known.Length == 0)
                {
                    known = "none";
                }

                return Result.NotFound($"There is no pivot table named '{name}' added with AddPivotTable; a slicer filters those only. The workbook has {known}.");
            }

            if (pivots.Contains(found) == false)
            {
                pivots.Add(found);
            }
        }

        if (pivots.Any(pivot => string.Equals(pivot.Source.Key, pivots[0].Source.Key, StringComparison.Ordinal) == false) == true)
        {
            return Result.Invalid($"The pivot tables {string.Join(", ", pivots.Select(pivot => "'" + pivot.Definition.Name + "'"))} have different sources. A slicer filters pivot tables on one source.");
        }

        Result<PivotCacheData> data = PivotCacheData.Read(pivots[0].Source);

        if (data.IsFailure(out string? error, out ResultStatus status, out PivotCacheData? cache) == true)
        {
            return Result.Failure(error, status);
        }

        int field = cache.FieldIndex(definition.Field.Trim());

        if (field < 0)
        {
            string known = string.Join(", ", cache.FieldNames.Select(name => "'" + name + "'"));

            return Result.NotFound($"The source {pivots[0].Source.Display} has no column headed '{definition.Field}'. Its headings are {known}.");
        }

        string fieldName = cache.FieldNames[field];

        if (definition.Selected is IReadOnlyList<string> selected)
        {
            foreach (string item in selected)
            {
                if (cache.SharedItems[field].Exists(value => Matches(value, item) == true) == false)
                {
                    return Result.NotFound($"The field '{fieldName}' has no item '{item}'.");
                }
            }
        }

        string shapeName = definition.Name ?? fieldName;

        if (IsSlicerNameTaken(shapeName) == true)
        {
            if (definition.Name is not null)
            {
                return Result.Conflict($"The workbook already has a slicer named '{definition.Name}'.");
            }

            int number = 1;

            while (IsSlicerNameTaken(fieldName + " " + number.ToString(System.Globalization.CultureInfo.InvariantCulture)) == true)
            {
                number++;
            }

            shapeName = fieldName + " " + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        SlicerCacheModel? slicerCache = slicerCaches.Find(candidate => string.Equals(candidate.SourceKey, pivots[0].Source.Key, StringComparison.Ordinal) == true && string.Equals(candidate.Field, fieldName, StringComparison.Ordinal) == true);

        if (slicerCache is not null && definition.Selected is not null && slicerCache.Selected is not null && SameSelection(slicerCache.Selected, definition.Selected) == false)
        {
            return Result.Conflict($"The field '{fieldName}' already has a slicer with other items selected. Slicers on the same field share their selection.");
        }

        Result drawable = sheet.CheckDrawable();

        if (drawable.IsFailure(out error, out status) == true)
        {
            return Result.Failure(error, status);
        }

        if (slicerCache is null)
        {
            slicerCache = new SlicerCacheModel(UniqueCacheName(fieldName), pivots[0].Source.Key, fieldName);
            slicerCaches.Add(slicerCache);
        }

        foreach (CreatedPivot pivot in pivots.Where(pivot => slicerCache.PivotTables.Contains(pivot) == false))
        {
            slicerCache.PivotTables.Add(pivot);
        }

        if (definition.Selected is not null)
        {
            slicerCache.Selected = definition.Selected;
        }

        slicers.Add(new CreatedSlicer(definition, shapeName, sheet, area, slicerCache));
        sheet.Extras.AddSlicer(shapeName, area);

        return RenderAll(out _);
    }

    /// <summary>
    /// Renders every new pivot table from its source as it is now and writes it into its sheet,
    /// clearing what an earlier rendering covered. Changed is true when a cell changed.
    /// </summary>
    public Result RenderAll(out bool changed)
    {
        changed = false;
        Dictionary<string, PivotCacheData> caches = new(StringComparer.Ordinal);

        foreach ((int cacheId, PivotSource source, List<CreatedPivot> pivots) in CacheGroups)
        {
            Result<PivotCacheData> data = PivotCacheData.Read(source);

            if (data.IsFailure(out string? error, out ResultStatus status, out PivotCacheData? cache) == true)
            {
                return Result.Failure(error, status);
            }

            caches[source.Key] = cache;

            foreach (CreatedPivot pivot in pivots)
            {
                pivot.CacheId = cacheId;
                (Dictionary<int, HashSet<int>> hidden, HashSet<int> slicerFields) = Filters(pivot, cache);
                Result<PivotRendering> rendered = PivotRender.Render(pivot, cache, hidden, slicerFields, book.Uses1904Dates);

                if (rendered.IsFailure(out error, out status, out PivotRendering? rendering) == true)
                {
                    return Result.Failure(error, status);
                }

                changed |= Write(pivot, rendering);
                pivot.Rendering = rendering;
            }
        }

        return Result.Success();
    }

    /// <summary>Per field, the shared indexes the slicers on a pivot table hide, and the fields with slicers.</summary>
    public (Dictionary<int, HashSet<int>> Hidden, HashSet<int> SlicerFields) Filters(CreatedPivot pivot, PivotCacheData cache)
    {
        Dictionary<int, HashSet<int>> hidden = [];
        HashSet<int> fields = [];

        foreach (SlicerCacheModel slicerCache in slicerCaches.Where(candidate => candidate.PivotTables.Contains(pivot) == true))
        {
            int field = cache.FieldIndex(slicerCache.Field);

            if (field < 0)
            {
                continue;
            }

            fields.Add(field);

            if (slicerCache.Selected is not IReadOnlyList<string> selected)
            {
                continue;
            }

            HashSet<int> hide = [];

            for (int index = 0; index < cache.SharedItems[field].Count; index++)
            {
                CellValue value = cache.SharedItems[field][index];

                if (selected.Any(item => Matches(value, item) == true) == false)
                {
                    hide.Add(index);
                }
            }

            hidden[field] = hide;
        }

        return (hidden, fields);
    }

    /// <summary>True when an item is the one a slicer button shows as the text given.</summary>
    public static bool Matches(CellValue value, string text)
    {
        return string.Equals(ItemText(value), text.Trim(), StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>An item as a slicer button and a pivot table show it.</summary>
    public static string ItemText(CellValue value)
    {
        if (value.IsEmpty == true)
        {
            return PivotRender.Blank;
        }

        if (value.TryGetDate(out DateTime date) == true)
        {
            if (date.TimeOfDay == TimeSpan.Zero)
            {
                return date.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
            }

            return date.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
        }

        return value.ToString();
    }

    private static bool Write(CreatedPivot pivot, PivotRendering rendering)
    {
        bool changed = false;
        Worksheet sheet = pivot.Sheet;
        HashSet<(int Row, int Column)> written = [];

        foreach ((int row, int column, Cell cell) in rendering.Cells)
        {
            written.Add((row, column));
            Cell current = sheet[row, column];

            if (current.Value == cell.Value && current.Formula is null && Equals(current.Style, cell.Style) == true)
            {
                continue;
            }

            sheet[row, column] = cell;
            changed = true;
        }

        List<CellRange> areas = [rendering.Area];

        if (pivot.Rendering is not null)
        {
            areas.Add(pivot.Rendering.Area);
        }

        foreach (CellRange area in areas)
        {
            foreach ((int row, RowEntry entry) in sheet.RowEntries.Between(area.First.Row, area.Last.Row).ToList())
            {
                foreach (int column in entry.Cells.Between(area.First.Column, area.Last.Column).Select(pair => pair.Key).ToList())
                {
                    if (written.Contains((row, column)) == false)
                    {
                        sheet.Clear(row, column);
                        changed = true;
                    }
                }
            }
        }

        return changed;
    }

    private Result<PivotSource> ResolveSource(Worksheet owner, string text)
    {
        if (text.Contains('!', StringComparison.Ordinal) == true || CellRange.TryParse(text, out _) == true)
        {
            SheetReference reference = SheetReference.Parse(text, nameof(PivotTable.Source));
            Worksheet sheet = owner;

            if (reference.Sheet is not null)
            {
                Result<Worksheet> found = book.Sheet(reference.Sheet);

                if (found.IsFailure(out string? error, out ResultStatus status, out Worksheet? named) == true)
                {
                    return Result<PivotSource>.Failure(error, status);
                }

                sheet = named;
            }

            return Trim(sheet, reference.Range, null, text);
        }

        if (Loaded.Tables.TryGetValue(text, out (string Sheet, CellRange Range) table) == true && book.Sheet(table.Sheet).IsSuccess(out Worksheet? tableSheet) == true)
        {
            return Trim(tableSheet, table.Range, Loaded.Tables.Keys.First(name => string.Equals(name, text, StringComparison.OrdinalIgnoreCase) == true), text);
        }

        return Result<PivotSource>.NotFound($"There is no table named '{text}', and it is not a range such as Data!A1:F49.");
    }

    private static Result<PivotSource> Trim(Worksheet sheet, CellRange range, string? tableName, string text)
    {
        int last = Math.Min(range.Last.Row, sheet.RowCount - 1);

        if (last < range.First.Row)
        {
            return Result<PivotSource>.Invalid($"The source {text} is empty: a pivot table's source needs a row of headings.");
        }

        return new PivotSource(sheet, new CellRange(range.First, new CellAddress(last, range.Last.Column)), tableName);
    }

    /// <summary>True for a Strict OOXML file, whose parts have namespaces of their own: transitional pivot parts would make a package that is neither.</summary>
    private bool IsStrict()
    {
        return book.Sheets.Any(sheet => sheet.Source is not null && string.Equals(sheet.Source.Namespace, Ooxml.SpreadsheetNamespace, StringComparison.Ordinal) == false);
    }

    private bool TryFindByName(Worksheet sheet, string name)
    {
        return created.Exists(pivot => ReferenceEquals(pivot.Sheet, sheet) == true && string.Equals(pivot.Definition.Name, name, StringComparison.OrdinalIgnoreCase) == true)
            || Loaded.Grids.Any(grid => ReferenceEquals(grid.Sheet, sheet) == true && string.Equals(grid.Name, name, StringComparison.OrdinalIgnoreCase) == true);
    }

    private bool IsSlicerNameTaken(string name)
    {
        return slicers.Exists(slicer => string.Equals(slicer.Name, name, StringComparison.OrdinalIgnoreCase) == true)
            || Loaded.SlicerNames.Contains(name);
    }

    /// <summary>Slicer_ and the field name with what a defined name cannot hold made _, numbered when taken, as Excel names them.</summary>
    private string UniqueCacheName(string field)
    {
        StringBuilder builder = new("Slicer_");

        foreach (char character in field)
        {
            if (char.IsLetterOrDigit(character) == true || character is '_' or '.')
            {
                builder.Append(character);
            }
            else
            {
                builder.Append('_');
            }
        }

        string name = builder.ToString();
        string unique = name;
        int number = 1;

        while (IsNameTaken(unique) == true)
        {
            unique = name + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
            number++;
        }

        return unique;
    }

    private bool IsNameTaken(string name)
    {
        return slicerCaches.Exists(cache => string.Equals(cache.Name, name, StringComparison.OrdinalIgnoreCase) == true)
            || book.DefinedNames.Exists(defined => string.Equals(defined.Name, name, StringComparison.OrdinalIgnoreCase) == true)
            || Loaded.SlicerCacheNames.Contains(name);
    }

    private static bool SameSelection(IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        HashSet<string> first = new(left.Select(item => item.Trim()), StringComparer.OrdinalIgnoreCase);

        return first.SetEquals(right.Select(item => item.Trim()));
    }

    private static bool Overlaps(CellRange left, CellRange right)
    {
        return left.First.Row <= right.Last.Row && right.First.Row <= left.Last.Row && left.First.Column <= right.Last.Column && right.First.Column <= left.Last.Column;
    }
}

/// <summary>What an opened file holds of pivot tables, pivot caches, slicers and tables.</summary>
internal sealed class LoadedPivots
{
    public static LoadedPivots Empty { get; } = new();

    public List<PivotGrid> Grids { get; } = [];

    public List<PivotCacheInfo> Caches { get; } = [];

    /// <summary>The paths of the pivot cache definition parts, which <see cref="Workbook.RefreshPivotTablesOnOpen"/> marks.</summary>
    public List<string> CacheDefinitionPaths { get; } = [];

    /// <summary>The Excel tables by name: their sheet and range.</summary>
    public Dictionary<string, (string Sheet, CellRange Range)> Tables { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The x14 pivotCacheId of each cache, which a new cache must not reuse.</summary>
    public HashSet<long> PivotCacheIds { get; } = [];

    public HashSet<string> SlicerNames { get; } = new(StringComparer.OrdinalIgnoreCase);

    public HashSet<string> SlicerCacheNames { get; } = new(StringComparer.OrdinalIgnoreCase);

    public int MaxCacheId { get; set; } = -1;
}
