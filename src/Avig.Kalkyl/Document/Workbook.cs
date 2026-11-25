using System.Globalization;
using System.IO.Compression;
using Avig.Kalkyl.Document;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;
using Avig.OpenXml;

namespace Avig.Kalkyl;

/// <summary>
/// A workbook held in memory: open a file, read and change any cell in any order, add sheets,
/// and save. Everything the package does not understand — charts, pictures, conditional
/// formatting, comments, data validation, macros — is kept as it was and written back.
/// </summary>
/// <example>
/// <code>
/// Result&lt;Workbook&gt; opened = await Workbook.OpenAsync("budget.xlsx");
///
/// if (opened.IsSuccess(out Workbook? book) &amp;&amp; book.Sheet("2026").IsSuccess(out Worksheet? sheet))
/// {
///     sheet.SetValue(4, 2, 125_000);                 // C5, keeps its formatting
///     sheet["D5"] = Cell.FromFormula("C5*1.25").Bold();
///     await book.SaveAsync("budget.xlsx");
/// }
/// </code>
/// </example>
/// <remarks>
/// Not safe to change from several threads at once; separate instances are independent. For a
/// file from outside, see <see cref="KalkylLimits"/> and <see cref="ActiveContent"/>. Error
/// messages name paths, sheets and cell contents: log them, but do not show them to whoever sent the file.
/// </remarks>
public sealed class Workbook
{
    private readonly List<Worksheet> sheets = [];
    private readonly HashSet<(Worksheet Sheet, int Row, int Column)> changedCells = [];

    private Workbook(byte[]? package, StyleTable styles, bool uses1904Dates, string? workbookPath)
    {
        Package = package;
        Styles = styles;
        Uses1904Dates = uses1904Dates;
        WorkbookPath = workbookPath;
        Pivots = new PivotRegistry(this);
    }

    /// <summary>The worksheets, in tab order.</summary>
    public IReadOnlyList<Worksheet> Sheets => sheets;

    /// <summary>True when the workbook counts dates from 1904, as old Mac files do. Dates are converted either way.</summary>
    public bool Uses1904Dates { get; }

    /// <summary>
    /// The culture formulas read text in, when text has to become a number or a date: the
    /// separators in "1,5" and "1 000", and the order in "3/2/2013" (3 February in en-GB,
    /// 2 March in en-US). ISO dates such as 2026-02-01 are read in every culture. Excel uses
    /// the culture of the machine that calculates, and so does this, unless it is set; set it
    /// to the culture the workbook was made in to get the results its author saw. Formatting
    /// (TEXT, DOLLAR, FIXED) is not affected. Recalculate with RecalculationScope.All after
    /// changing it.
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

    /// <summary>The bytes of the file that was opened, or null for a new workbook.</summary>
    internal byte[]? Package { get; }

    internal StyleTable Styles { get; }

    internal string? WorkbookPath { get; }

    /// <summary>The clock NOW() and TODAY() read. Replaceable in tests.</summary>
    internal TimeProvider Clock { get; set; } = TimeProvider.System;

    /// <summary>The workbook's defined names, read when the file is opened. Used by the formula engine.</summary>
    internal List<Formulas.DefinedName> DefinedNames { get; } = [];

    /// <summary>The pivot tables and slicers: those the file had and those added.</summary>
    internal PivotRegistry Pivots { get; }

    /// <summary>
    /// The pivot tables: those the file had when it was opened, then those added with
    /// <see cref="Worksheet.AddPivotTable(PivotTable, CellAddress)"/>.
    /// </summary>
    public IReadOnlyList<PivotTableInfo> PivotTables
    {
        get
        {
            List<PivotTableInfo> tables = [];

            foreach (PivotGrid grid in Pivots.Grids)
            {
                tables.Add(new PivotTableInfo(grid.Name, grid.Sheet.Name, grid.Location, grid.CacheId));
            }

            return tables;
        }
    }

    /// <summary>
    /// The pivot caches: those the file had, then one per source of the pivot tables added.
    /// <see cref="PivotCacheInfo.RefreshOnLoad"/> is what the saved file will say.
    /// </summary>
    public IReadOnlyList<PivotCacheInfo> PivotCaches
    {
        get
        {
            List<PivotCacheInfo> caches = [];

            foreach (PivotCacheInfo cache in Pivots.Loaded.Caches)
            {
                caches.Add(cache with { RefreshOnLoad = cache.RefreshOnLoad == true || RefreshPivotTablesOnOpen == true });
            }

            foreach ((int cacheId, PivotSource source, List<CreatedPivot> _) in Pivots.CacheGroups)
            {
                if (PivotCacheData.Read(source).IsSuccess(out PivotCacheData? data) == true)
                {
                    caches.Add(new PivotCacheInfo(cacheId, source.Display, data.FieldNames, data.Records.Count, true));
                }
            }

            return caches;
        }
    }

    /// <summary>
    /// Makes Excel refresh every pivot cache from its source when it opens the saved file, so
    /// that the pivot tables of a template show the data the package put in. The save sets
    /// refreshOnLoad on each cache the file has; the caches of pivot tables the package adds
    /// always have it. Excel refreshes quietly, except for a cache on a data connection, which
    /// it may ask about. False when the workbook is opened, whatever the file says.
    /// </summary>
    public bool RefreshPivotTablesOnOpen { get; set; }

    /// <summary>Cells set since the last recalculation. Read and cleared by the formula engine.</summary>
    internal HashSet<(Worksheet Sheet, int Row, int Column)> ChangedCells => changedCells;

    /// <summary>Creates an empty workbook. Add sheets with <see cref="AddSheet"/>.</summary>
    /// <returns>The workbook.</returns>
    public static Workbook Create()
    {
        return new Workbook(null, StyleTable.Empty, false, null);
    }

    /// <summary>Opens a file.</summary>
    /// <param name="path">The file.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>
    /// The workbook; NotFound when the file does not exist, Forbidden when it may not be read,
    /// Conflict when another program holds it locked, and Invalid when it is not an .xlsx workbook.
    /// </returns>
    public static Task<Result<Workbook>> OpenAsync(string path, CancellationToken cancellationToken = default)
    {
        return OpenAsync(path, KalkylLimits.Default, cancellationToken);
    }

    /// <summary>Opens a file, with limits on how much it may unpack to and how many cells it may hold.</summary>
    /// <param name="path">The file.</param>
    /// <param name="limits">The limits; <see cref="KalkylLimits.Default"/> unless the file is trusted. Kept in <see cref="Limits"/>.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>As <see cref="OpenAsync(string, CancellationToken)"/>, and Invalid when the file goes past a limit.</returns>
    public static async Task<Result<Workbook>> OpenAsync(string path, KalkylLimits limits, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(limits);
        byte[] bytes;

        try
        {
            if (new FileInfo(path).Length > limits.MaxUncompressedBytes)
            {
                return Result<Workbook>.Invalid($"'{path}' is larger than {limits.MaxUncompressedBytes:N0} bytes, the limit set in KalkylLimits.MaxUncompressedBytes.");
            }

            bytes = await File.ReadAllBytesAsync(path, cancellationToken);
        }
        catch (FileNotFoundException)
        {
            return Result<Workbook>.NotFound($"'{path}' does not exist.");
        }
        catch (DirectoryNotFoundException)
        {
            return Result<Workbook>.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result<Workbook>.Forbidden($"'{path}' may not be read.");
        }
        catch (IOException exception)
        {
            return Result<Workbook>.Conflict($"'{path}' could not be opened: {exception.Message}");
        }

        return await OpenBytesAsync(bytes, limits, cancellationToken);
    }

    /// <summary>Reads a workbook from a stream. The whole stream is read into memory.</summary>
    /// <param name="input">The stream. It is not closed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The workbook, or Invalid when the content is not an .xlsx workbook.</returns>
    public static Task<Result<Workbook>> OpenAsync(Stream input, CancellationToken cancellationToken = default)
    {
        return OpenAsync(input, KalkylLimits.Default, cancellationToken);
    }

    /// <summary>Reads a workbook from a stream, with limits on how much it may unpack to and how many cells it may hold.</summary>
    /// <param name="input">The stream. It is not closed.</param>
    /// <param name="limits">The limits; <see cref="KalkylLimits.Default"/> unless the content is trusted. Kept in <see cref="Limits"/>.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The workbook, or Invalid when the content is not an .xlsx workbook or goes past a limit.</returns>
    public static async Task<Result<Workbook>> OpenAsync(Stream input, KalkylLimits limits, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(limits);
        using MemoryStream copy = new();

        if (await XlsxReader.CopyLimitedAsync(input, copy, limits.MaxUncompressedBytes, cancellationToken) == false)
        {
            return Result<Workbook>.Invalid($"The stream is longer than {limits.MaxUncompressedBytes:N0} bytes, the limit set in KalkylLimits.MaxUncompressedBytes.");
        }

        return await OpenBytesAsync(copy.ToArray(), limits, cancellationToken);
    }

    /// <summary>Adds an empty sheet after the others.</summary>
    /// <param name="name">The name on the tab; the same rules as <see cref="XlsxWriter.AddSheetAsync"/>.</param>
    /// <returns>The sheet, Invalid for a name Excel would refuse, or Conflict when the name is taken.</returns>
    public Result<Worksheet> AddSheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Result validation = XlsxWriter.ValidateName(name);

        if (validation.IsFailure(out string? error, out ResultStatus status) == true)
        {
            return Result<Worksheet>.Failure(error, status);
        }

        if (sheets.Exists(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase) == true))
        {
            return Result<Worksheet>.Conflict($"A sheet named '{name}' already exists. Excel compares sheet names without regard to case.");
        }

        Worksheet added = new(this, name, sheets.Count, null);
        sheets.Add(added);

        return added;
    }

    /// <summary>Gets a worksheet by name, without regard to case.</summary>
    /// <param name="name">The name on the tab.</param>
    /// <returns>The sheet, or NotFound.</returns>
    public Result<Worksheet> Sheet(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        Worksheet? found = sheets.Find(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase) == true);

        if (found is null)
        {
            string names = string.Join(", ", sheets.Select(sheet => "'" + sheet.Name + "'"));

            return Result<Worksheet>.NotFound($"There is no sheet named '{name}'. The workbook has {names}.");
        }

        return found;
    }

    /// <summary>Gets a worksheet by position.</summary>
    /// <param name="index">The position, 0-based.</param>
    /// <returns>The sheet, or NotFound.</returns>
    public Result<Worksheet> Sheet(int index)
    {
        if (index < 0 || index >= sheets.Count)
        {
            return Result<Worksheet>.NotFound($"There is no sheet {index}. The workbook has {sheets.Count} worksheet(s), numbered from 0.");
        }

        return sheets[index];
    }

    /// <summary>Saves to a file, replacing it if it exists. Saving over the file that was opened is fine.</summary>
    /// <param name="path">The file.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>Success, or NotFound, Forbidden or Conflict when the file cannot be written.</returns>
    public async Task<Result> SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        // Build the file in memory first, so a failure half-way never leaves a broken file behind.
        using MemoryStream buffer = new();
        await SaveAsync(buffer, cancellationToken);

        try
        {
            await File.WriteAllBytesAsync(path, buffer.ToArray(), cancellationToken);
        }
        catch (DirectoryNotFoundException)
        {
            return Result.NotFound($"The folder for '{path}' does not exist.");
        }
        catch (UnauthorizedAccessException)
        {
            return Result.Forbidden($"'{path}' may not be written.");
        }
        catch (IOException exception)
        {
            return Result.Conflict($"'{path}' could not be written: {exception.Message}");
        }

        return Result.Success();
    }

    /// <summary>Saves to a stream.</summary>
    /// <param name="output">The stream. It is not closed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>A task that completes when the workbook is written.</returns>
    public async Task SaveAsync(Stream output, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        await DocumentSaver.SaveAsync(this, output, cancellationToken);
    }

    /// <summary>
    /// Calculates formulas and stores their results in the cells, so that they can be read and
    /// are saved with the file. With <see cref="RecalculationScope.Changed"/> (the default) only
    /// formulas affected by cells set since the last recalculation are calculated, together with
    /// volatile functions; <see cref="RecalculationScope.All"/> calculates every formula.
    /// A function the engine does not know gives #NAME?, as in Excel; the report lists them.
    /// Pivot tables added with <see cref="Worksheet.AddPivotTable(PivotTable, CellAddress)"/>
    /// are rendered again from their sources first, and once more after the formulas when a
    /// source holds formulas, so that GETPIVOTDATA reads them as they are.
    /// </summary>
    /// <param name="scope">Which formulas to calculate.</param>
    /// <param name="cancellationToken">Cancels the calculation.</param>
    /// <returns>What was calculated, and circular references, unknown functions and broken formulas.</returns>
    public async Task<RecalculationReport> RecalculateAsync(RecalculationScope scope = RecalculationScope.Changed, CancellationToken cancellationToken = default)
    {
        DateTime now = Clock.GetLocalNow().DateTime;
        RenderPivotTables();
        RecalculationReport report = await Task.Run(() => Formulas.Recalculator.Run(this, scope, now, cancellationToken), cancellationToken);

        if (RenderPivotTables() == true)
        {
            RecalculationReport again = await Task.Run(() => Formulas.Recalculator.Run(this, RecalculationScope.Changed, now, cancellationToken), cancellationToken);
            report = again with { Calculated = report.Calculated + again.Calculated };
        }

        return report;
    }

    /// <summary>
    /// Evaluates a formula against the workbook as it is, without storing anything. Formula cells
    /// it reads give their stored results, so recalculate first if cells have changed.
    /// </summary>
    /// <param name="formula">The formula, with or without a leading <c>=</c>, such as <c>SUM(Budget!B2:B9)</c>.</param>
    /// <param name="sheet">The sheet that references without a sheet name point to; the first sheet when null.</param>
    /// <param name="cancellationToken">Stops a formula that takes too long; it throws OperationCanceledException.</param>
    /// <returns>The result, or Invalid when the formula cannot be parsed.</returns>
    public Result<CellValue> Evaluate(string formula, Worksheet? sheet = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(formula);
        Worksheet? context = sheet ?? sheets.FirstOrDefault();

        if (context is null)
        {
            return Result<CellValue>.Invalid("The workbook has no sheets to evaluate the formula against.");
        }

        Result<Formulas.Node> parsed = Formulas.FormulaParser.Parse(formula);

        if (parsed.IsSuccess(out Formulas.Node? node) == false)
        {
            return Result<CellValue>.Invalid(parsed.Error!);
        }

        Formulas.Evaluator evaluator = new(this, Clock.GetLocalNow().DateTime) { CancellationToken = cancellationToken };

        Formulas.EvaluationContext evaluationContext = new(context, 0, 0);

        using (Formulas.ValueArray.LimitTo(Limits.MaxArrayCells))
        {
            Formulas.Value result = evaluator.Dereference(evaluator.Evaluate(node, evaluationContext), evaluationContext);

            return Formulas.Recalculator.ToCellValue(result, Style.Default, Uses1904Dates);
        }
    }

    /// <summary>
    /// The limits the workbook was opened with; <see cref="KalkylLimits.Default"/> for a new one.
    /// <see cref="KalkylLimits.MaxArrayCells"/> also applies to recalculation and <see cref="Evaluate"/>.
    /// </summary>
    public KalkylLimits Limits { get; set; } = KalkylLimits.Default;

    /// <summary>
    /// True when the file has active content: a VBA project (.xlsm), Excel 4.0 macro sheets,
    /// dialog sheets or ActiveX controls. The package never runs any of it, but it keeps it when
    /// the file is saved, as it keeps everything it does not understand: opening and saving a
    /// file does not make it safe. <see cref="RemoveMacros"/> drops the VBA project, and
    /// <see cref="RemoveActiveContent"/> all of <see cref="ActiveContent"/>.
    /// </summary>
    public bool HasMacros { get; private set; }

    /// <summary>True when <see cref="RemoveMacros"/> was called and the next save drops the VBA project.</summary>
    internal bool RemovingMacros { get; private set; }

    /// <summary>
    /// Makes the next save drop the VBA project (vbaProject.bin and its signatures) and turn the
    /// workbook into a plain .xlsx. Save it under a name ending in .xlsx, since Excel refuses an
    /// .xlsm without macros less than an .xlsx with them.
    /// </summary>
    /// <returns>
    /// Success, also when there is nothing to remove; Invalid when the file has Excel 4.0 macro
    /// sheets, dialog sheets or ActiveX controls, which are woven into the sheets and cannot be
    /// taken out without changing them.
    /// </returns>
    public Result RemoveMacros()
    {
        if (Package is not null)
        {
            using ZipArchive archive = new(new MemoryStream(Package, false), ZipArchiveMode.Read);

            if (archive.Entries.Any(entry => IsWovenActiveContent(entry.FullName) == true) == true)
            {
                return Result.Invalid("The workbook has Excel 4.0 macro sheets, dialog sheets or ActiveX controls, which the package cannot remove. Remove them in Excel.");
            }
        }

        RemovingMacros = true;

        return Result.Success();
    }

    /// <summary>A part of the VBA project, which <see cref="RemoveMacros"/> drops: vbaProject.bin, its signatures and their relationships.</summary>
    internal static bool IsVbaPart(string path)
    {
        string name = path.Replace('\\', '/');
        string file = name[(name.LastIndexOf('/') + 1)..];

        return file.StartsWith("vbaProject", StringComparison.OrdinalIgnoreCase) == true || file.StartsWith("vbaData", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsWovenActiveContent(string path)
    {
        string name = path.Replace('\\', '/');

        return name.Contains("/macrosheets/", StringComparison.OrdinalIgnoreCase) == true
            || name.Contains("/dialogsheets/", StringComparison.OrdinalIgnoreCase) == true
            || name.Contains("/activeX/", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool HasActiveContent(ZipArchive archive)
    {
        return archive.Entries.Any(entry => IsVbaPart(entry.FullName) == true || IsWovenActiveContent(entry.FullName) == true);
    }

    /// <summary>
    /// What in the file can act on its own or fetch something when Excel opens it: macros (VBA,
    /// Excel 4.0 macro sheets, dialog sheets), ActiveX controls, embedded OLE objects, external
    /// references (links to other workbooks, DDE links and other external relationships) and data
    /// connections (connections, query tables, Power Query). Found when the file is opened. The
    /// package never runs, follows or refreshes any of it, but it keeps it when the file is saved,
    /// so opening and saving a file does not make it safe; <see cref="RemoveActiveContent"/> does.
    /// <see cref="ActiveContent.None"/> for a new workbook, and after <see cref="RemoveActiveContent"/>.
    /// </summary>
    public ActiveContent ActiveContent
    {
        get
        {
            if (RemovingActiveContent == true)
            {
                return ActiveContent.None;
            }

            return Scan.Found;
        }
    }

    /// <summary>
    /// The relationships to something outside the file that Excel follows on its own, unlike
    /// hyperlinks: the path of each linked workbook, linked pictures and objects, a chart's
    /// external data. A DDE link has no relationship; it is listed with the external link part it
    /// is in, <see cref="OpcExternalRelationship.Kind"/> ddeLink and the target service|topic,
    /// such as cmd|/c calc. Data connections are not listed; see <see cref="ActiveContent"/>.
    /// Empty after <see cref="RemoveActiveContent"/>.
    /// </summary>
    public IReadOnlyList<OpcExternalRelationship> ExternalReferences
    {
        get
        {
            if (RemovingActiveContent == true)
            {
                return [];
            }

            return Scan.ExternalReferences;
        }
    }

    /// <summary>True when <see cref="RemoveActiveContent"/> was called and the next save leaves the active content out.</summary>
    internal bool RemovingActiveContent { get; private set; }

    /// <summary>True when the next save drops the VBA project, after <see cref="RemoveMacros"/> or <see cref="RemoveActiveContent"/>.</summary>
    internal bool DropsVba => RemovingMacros == true || RemovingActiveContent == true;

    /// <summary>What was found in the package when it was opened.</summary>
    internal PackageScan Scan { get; private init; } = PackageScan.Empty;

    /// <summary>The opened file's sheet list, with the chart, macro and dialog sheets the model does not hold.</summary>
    internal IReadOnlyList<DeclaredSheet> DeclaredSheets { get; private init; } = [];

    /// <summary>
    /// Makes the next save leave out all <see cref="ActiveContent"/> and write a plain .xlsx:
    /// the VBA project (as <see cref="RemoveMacros"/> does), Excel 4.0 macro sheets and dialog
    /// sheets with the defined names that start them (Auto_Open and the like), ActiveX controls
    /// and OLE objects with their pictures in the sheets, data connections, query tables and
    /// Power Query queries (a table a query filled becomes an ordinary table with the same data),
    /// links to other workbooks and DDE links, and every other external relationship except
    /// hyperlinks. A formula that refers to another workbook is replaced by the value it last
    /// had, now and in the saved file, and a defined name that does becomes #REF!. What the
    /// active content pointed to and nothing else used goes with it; so does a part of the file
    /// that nothing pointed to. Save under a name ending in .xlsx.
    /// </summary>
    /// <returns>
    /// Success, also when there is nothing to remove; Invalid when taking something out would
    /// break the workbook or change what it shows: a pivot table or the data model (Power Pivot)
    /// fed by a data connection, a chart drawn from another workbook, or an external relationship
    /// in a part the package does not clean, such as a pivot table on another workbook.
    /// </returns>
    public Result RemoveActiveContent()
    {
        if (Scan.TryGetRemovalBlocker(out string reason) == true)
        {
            return Result.Invalid($"The workbook has {reason}, which the package cannot take out without breaking the workbook. Remove it in Excel.");
        }

        if (RemovingActiveContent == false && Scan.Found.HasFlag(ActiveContent.ExternalReferences) == true)
        {
            foreach (Worksheet sheet in sheets)
            {
                sheet.ReplaceExternalFormulas();
            }

            for (int index = 0; index < DefinedNames.Count; index++)
            {
                Formulas.DefinedName name = DefinedNames[index];

                if (ActiveContentRules.HasExternalReference(name.Formula) == true)
                {
                    DefinedNames[index] = name with { Formula = "#REF!" };
                }
            }
        }

        RemovingActiveContent = true;

        return Result.Success();
    }

    /// <summary>True once a sheet's settings (widths, panes, filter, visibility) have changed.</summary>
    internal bool SettingsChanged { get; private set; }

    internal void MarkChanged()
    {
        SettingsChanged = true;
    }

    internal void MarkChanged(Worksheet sheet, int row, int column)
    {
        changedCells.Add((sheet, row, column));
    }

    /// <summary>Renders the added pivot tables into their sheets; true when a cell changed.</summary>
    internal bool RenderPivotTables()
    {
        Result rendered = Pivots.RenderAll(out bool changed);

        // A source that worked when the pivot table was added can stop working only through a
        // change to its headings; the table then keeps its last rendering, as Excel keeps a
        // pivot table whose refresh fails.
        if (rendered.IsFailure(out _, out _) == true)
        {
            return false;
        }

        return changed;
    }

    internal void AddLoadedSheet(Worksheet sheet)
    {
        sheets.Add(sheet);
    }

    private static async Task<Result<Workbook>> OpenBytesAsync(byte[] bytes, KalkylLimits limits, CancellationToken cancellationToken)
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(new MemoryStream(bytes, false), limits, false, cancellationToken);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            return Result<Workbook>.Failure(opened.Error!, opened.Status);
        }

        await using (reader)
        {
            PackageScan scan;

            try
            {
                scan = await reader.GetScanAsync(cancellationToken);
            }
            catch (InvalidDataException exception)
            {
                return Result<Workbook>.Invalid(KalkylError.Describe(exception));
            }

            Workbook book = new(bytes, reader.Styles, reader.Uses1904Dates, reader.WorkbookPath)
            {
                Limits = limits,
                HasMacros = HasActiveContent(reader.Archive),
                Scan = scan,
                DeclaredSheets = reader.DeclaredSheets,
            };

            Result loaded = await DocumentLoader.LoadAsync(book, reader, cancellationToken);

            if (loaded.IsFailure(out string? error, out ResultStatus status) == true)
            {
                return Result<Workbook>.Failure(error, status);
            }

            return book;
        }
    }
}
