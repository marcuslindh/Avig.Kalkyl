using Avig.Kalkyl.Document;
using Avig.Kalkyl.Reading;

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
    }

    /// <summary>The worksheets, in tab order.</summary>
    public IReadOnlyList<Worksheet> Sheets => sheets;

    /// <summary>True when the workbook counts dates from 1904, as old Mac files do. Dates are converted either way.</summary>
    public bool Uses1904Dates { get; }

    /// <summary>The bytes of the file that was opened, or null for a new workbook.</summary>
    internal byte[]? Package { get; }

    internal StyleTable Styles { get; }

    internal string? WorkbookPath { get; }

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
    public static async Task<Result<Workbook>> OpenAsync(string path, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        byte[] bytes;

        try
        {
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

        return await OpenBytesAsync(bytes, cancellationToken);
    }

    /// <summary>Reads a workbook from a stream. The whole stream is read into memory.</summary>
    /// <param name="input">The stream. It is not closed.</param>
    /// <param name="cancellationToken">Cancels the operation.</param>
    /// <returns>The workbook, or Invalid when the content is not an .xlsx workbook.</returns>
    public static async Task<Result<Workbook>> OpenAsync(Stream input, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        using MemoryStream copy = new();
        await input.CopyToAsync(copy, cancellationToken);

        return await OpenBytesAsync(copy.ToArray(), cancellationToken);
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

    internal void AddLoadedSheet(Worksheet sheet)
    {
        sheets.Add(sheet);
    }

    private static async Task<Result<Workbook>> OpenBytesAsync(byte[] bytes, CancellationToken cancellationToken)
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(new MemoryStream(bytes, false), false, cancellationToken);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            return Result<Workbook>.Failure(opened.Error!, opened.Status);
        }

        await using (reader)
        {
            Workbook book = new(bytes, reader.Styles, reader.Uses1904Dates, reader.WorkbookPath);
            Result loaded = await DocumentLoader.LoadAsync(book, reader, cancellationToken);

            if (loaded.IsFailure(out string? error, out ResultStatus status) == true)
            {
                return Result<Workbook>.Failure(error, status);
            }

            return book;
        }
    }
}
