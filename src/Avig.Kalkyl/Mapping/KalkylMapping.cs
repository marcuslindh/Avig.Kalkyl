using System.Runtime.CompilerServices;

namespace Avig.Kalkyl;

/// <summary>Writing and reading lists of <see cref="KalkylRowAttribute"/> types.</summary>
public static class KalkylMapping
{
    /// <summary>
    /// Writes a header row and one row per item. On an empty sheet the header is bold and
    /// frozen, and column widths from <see cref="KalkylColumnAttribute.Width"/> are applied.
    /// </summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <param name="sheet">The sheet.</param>
    /// <param name="items">The items.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    /// <returns>A task that completes when the rows are written.</returns>
    public static async Task WriteAsync<T>(this SheetWriter sheet, IEnumerable<T> items, CancellationToken cancellationToken = default)
        where T : IKalkylRow<T>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(items);
        KalkylMap<T> map = T.KalkylMap;
        await WriteHeaderAsync(sheet, map, cancellationToken);

        foreach (T item in items)
        {
            await sheet.WriteRowAsync(map.ToCells(item), cancellationToken);
        }
    }

    /// <summary>Writes a header row and one row per item, as the items arrive.</summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <param name="sheet">The sheet.</param>
    /// <param name="items">The items.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    /// <returns>A task that completes when the rows are written.</returns>
    public static async Task WriteAsync<T>(this SheetWriter sheet, IAsyncEnumerable<T> items, CancellationToken cancellationToken = default)
        where T : IKalkylRow<T>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(items);
        KalkylMap<T> map = T.KalkylMap;
        await WriteHeaderAsync(sheet, map, cancellationToken);

        await foreach (T item in items.WithCancellation(cancellationToken))
        {
            await sheet.WriteRowAsync(map.ToCells(item), cancellationToken);
        }
    }

    /// <summary>
    /// Reads the sheet as items. The first row that is not empty is the header; columns are
    /// matched by header without regard to case or surrounding spaces, so their order does not
    /// matter. Empty rows are skipped. A row that cannot become an item gives an Invalid result
    /// naming the cell, and reading goes on with the next row.
    /// </summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <param name="sheet">The sheet.</param>
    /// <param name="cancellationToken">Cancels the reading.</param>
    /// <returns>
    /// One result per row. When a required column is missing from the header, a single Invalid
    /// result comes back and nothing more.
    /// </returns>
    public static async IAsyncEnumerable<Result<T>> ReadAsync<T>(this SheetReader sheet, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : IKalkylRow<T>
    {
        ArgumentNullException.ThrowIfNull(sheet);
        KalkylMap<T> map = T.KalkylMap;
        int[]? sheetColumns = null;

        await foreach (Row row in sheet.ReadRowsAsync(cancellationToken))
        {
            if (row.IsEmpty == true)
            {
                continue;
            }

            if (sheetColumns is null)
            {
                Result<int[]> matched = MatchHeaders(row, map.Columns, sheet.Info.Name);

                if (matched.IsFailure(out string? error, out ResultStatus status) == true)
                {
                    yield return Result<T>.Failure(error, status);

                    yield break;
                }

                matched.IsSuccess(out sheetColumns);

                continue;
            }

            yield return map.FromRow(new MappedRow(row, map.Columns, sheetColumns, sheet.Info.Name));
        }
    }

    private static async Task WriteHeaderAsync<T>(SheetWriter sheet, KalkylMap<T> map, CancellationToken cancellationToken)
    {
        if (sheet.RowCount == 0)
        {
            for (int column = 0; column < map.Columns.Count; column++)
            {
                if (map.Columns[column].Width > 0)
                {
                    sheet.SetColumnWidth(column, map.Columns[column].Width);
                }
            }

            sheet.FreezePanes(1);
        }

        Cell[] header = new Cell[map.Columns.Count];

        for (int column = 0; column < header.Length; column++)
        {
            header[column] = Cell.Of(map.Columns[column].Header).Bold();
        }

        await sheet.WriteRowAsync(header, cancellationToken);
    }

    private static Result<int[]> MatchHeaders(Row header, IReadOnlyList<KalkylColumnInfo> columns, string sheetName)
    {
        Dictionary<string, int> found = new(StringComparer.OrdinalIgnoreCase);

        for (int column = 0; column < header.ColumnCount; column++)
        {
            string text = header[column].Value.ToString().Trim();

            if (text.Length > 0)
            {
                found.TryAdd(text, column);
            }
        }

        int[] sheetColumns = new int[columns.Count];
        List<string> missing = [];

        for (int index = 0; index < columns.Count; index++)
        {
            if (found.TryGetValue(columns[index].Header.Trim(), out int column) == true)
            {
                sheetColumns[index] = column;

                continue;
            }

            sheetColumns[index] = -1;

            if (columns[index].Required == true)
            {
                missing.Add("'" + columns[index].Header + "'");
            }
        }

        if (missing.Count > 0)
        {
            string headers = string.Join(", ", found.Keys.Select(key => "'" + key + "'"));

            return Result<int[]>.Invalid($"The header in row {header.Index + 1} of '{sheetName}' lacks the required column(s) {string.Join(", ", missing)}. It has {headers}.");
        }

        return sheetColumns;
    }
}

/// <summary>Shortcuts that write or read a whole file of <see cref="KalkylRowAttribute"/> items.</summary>
public static class Xlsx
{
    /// <summary>Writes the items to a new file with one sheet.</summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <param name="path">The file.</param>
    /// <param name="items">The items.</param>
    /// <param name="sheetName">The sheet's name.</param>
    /// <param name="cancellationToken">Cancels the write.</param>
    /// <returns>Success, or why the file could not be written.</returns>
    public static async Task<Result> WriteAsync<T>(string path, IEnumerable<T> items, string sheetName = "Sheet1", CancellationToken cancellationToken = default)
        where T : IKalkylRow<T>
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path, cancellationToken);

        if (created.IsSuccess(out XlsxWriter? writer) == false)
        {
            return Result.Failure(created.Error!, created.Status);
        }

        await using (writer)
        {
            Result<SheetWriter> added = await writer.AddSheetAsync(sheetName, cancellationToken);

            if (added.IsSuccess(out SheetWriter? sheet) == false)
            {
                return Result.Failure(added.Error!, added.Status);
            }

            await sheet.WriteAsync(items, cancellationToken);
            await writer.CompleteAsync(cancellationToken);
        }

        return Result.Success();
    }

    /// <summary>Reads the items from a file.</summary>
    /// <typeparam name="T">The row type.</typeparam>
    /// <param name="path">The file.</param>
    /// <param name="sheetName">The sheet, or null for the first.</param>
    /// <param name="cancellationToken">Cancels the reading.</param>
    /// <returns>One result per row, or a single failure when the file or sheet cannot be opened.</returns>
    public static async IAsyncEnumerable<Result<T>> ReadAsync<T>(string path, string? sheetName = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        where T : IKalkylRow<T>
    {
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(path, cancellationToken);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            yield return Result<T>.Failure(opened.Error!, opened.Status);

            yield break;
        }

        await using (reader)
        {
            Result<SheetReader> found = FindSheet(reader, sheetName);

            if (found.IsSuccess(out SheetReader? sheet) == false)
            {
                yield return Result<T>.Failure(found.Error!, found.Status);

                yield break;
            }

            await foreach (Result<T> item in sheet.ReadAsync<T>(cancellationToken))
            {
                yield return item;
            }
        }
    }

    private static Result<SheetReader> FindSheet(XlsxReader reader, string? sheetName)
    {
        if (sheetName is null)
        {
            return reader.Sheet(0);
        }

        return reader.Sheet(sheetName);
    }
}
