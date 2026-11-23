using System.Globalization;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Xml;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl;

/// <summary>
/// Reads the rows of one worksheet. Each call to <see cref="ReadRowsAsync"/> starts over from
/// the top of the sheet.
/// </summary>
public sealed class SheetReader
{
    private readonly ZipArchiveEntry entry;
    private readonly IReadOnlyList<string> sharedStrings;
    private readonly StyleTable styles;
    private readonly bool uses1904;

    internal SheetReader(SheetInfo info, ZipArchiveEntry entry, IReadOnlyList<string> sharedStrings, StyleTable styles, bool uses1904)
    {
        Info = info;
        this.entry = entry;
        this.sharedStrings = sharedStrings;
        this.styles = styles;
        this.uses1904 = uses1904;
    }

    /// <summary>The sheet's name, position and visibility.</summary>
    public SheetInfo Info { get; }

    /// <summary>
    /// Reads the rows from the top. Only rows the file has are returned, so <see cref="Row.Index"/>
    /// can jump. A file damaged in the middle of a sheet throws <see cref="InvalidDataException"/>
    /// when the damaged row is reached.
    /// </summary>
    /// <param name="cancellationToken">Cancels the reading.</param>
    /// <returns>The rows.</returns>
    public async IAsyncEnumerable<Row> ReadRowsAsync([EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await using Stream stream = await entry.OpenAsync(cancellationToken);
        using XmlReader reader = XmlReader.Create(stream, XmlParts.Settings);
        int previousRow = -1;

        while (await ReadOrThrowAsync(reader) == true)
        {
            if (XmlParts.IsElement(reader, "row") == false)
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            int rowIndex = previousRow + 1;

            if (int.TryParse(reader.GetAttribute("r"), NumberStyles.None, CultureInfo.InvariantCulture, out int rowNumber) == true && rowNumber > 0)
            {
                rowIndex = rowNumber - 1;
            }

            previousRow = rowIndex;

            if (reader.IsEmptyElement == true)
            {
                yield return new Row(rowIndex, []);

                continue;
            }

            yield return await ReadRowAsync(reader, rowIndex);
        }
    }

    private async Task<Row> ReadRowAsync(XmlReader reader, int rowIndex)
    {
        int depth = reader.Depth;
        List<(int Column, Cell Cell)> cells = [];
        int previousColumn = -1;
        int lastColumn = -1;
        await ReadOrThrowAsync(reader);

        while (XmlParts.IsEnd(reader, depth) == false)
        {
            if (XmlParts.IsElement(reader, "c") == true)
            {
                (int column, Cell cell) = await ReadCellAsync(reader, rowIndex, previousColumn);
                previousColumn = column;

                if (cell.IsEmpty == false)
                {
                    cells.Add((column, cell));
                    lastColumn = Math.Max(lastColumn, column);
                }

                continue;
            }

            if (await ReadOrThrowAsync(reader) == false)
            {
                break;
            }
        }

        Cell[] dense = new Cell[lastColumn + 1];

        foreach ((int column, Cell cell) in cells)
        {
            dense[column] = cell;
        }

        return new Row(rowIndex, dense);
    }

    private async Task<(int Column, Cell Cell)> ReadCellAsync(XmlReader reader, int rowIndex, int previousColumn)
    {
        int column = previousColumn + 1;
        string? reference = reader.GetAttribute("r");

        if (reference is not null)
        {
            int letters = 0;

            while (letters < reference.Length && char.IsAsciiLetter(reference[letters]) == true)
            {
                letters++;
            }

            if (CellAddress.TryParseColumn(reference.AsSpan(0, letters), out int parsed) == true)
            {
                column = parsed;
            }
        }

        if (column >= CellAddress.MaxColumns)
        {
            throw Damaged(rowIndex, CellAddress.MaxColumns - 1, "the row has more cells than a sheet has columns");
        }

        string? type = reader.GetAttribute("t");
        int.TryParse(reader.GetAttribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out int styleIndex);
        string? stored = null;
        string? formula = null;
        string? inline = null;

        if (reader.IsEmptyElement == false)
        {
            int depth = reader.Depth;
            await ReadOrThrowAsync(reader);

            while (XmlParts.IsEnd(reader, depth) == false)
            {
                if (XmlParts.IsElement(reader, "v") == true)
                {
                    stored = await reader.ReadElementContentAsStringAsync();

                    continue;
                }

                if (XmlParts.IsElement(reader, "f") == true)
                {
                    // En delad formel (t="shared") står bara i första cellen; de andra har ett
                    // tomt <f si="0"/>. Formeltexten för dem räknas fram av formelmotorn senare.
                    string text = await reader.ReadElementContentAsStringAsync();

                    if (text.Length > 0)
                    {
                        formula = CellText.Decode(text);
                    }

                    continue;
                }

                if (XmlParts.IsElement(reader, "is") == true)
                {
                    inline = await XmlParts.ReadRichTextAsync(reader);
                }

                if (await ReadOrThrowAsync(reader) == false)
                {
                    break;
                }
            }
        }

        await ReadOrThrowAsync(reader);
        CellValue value = ToValue(type, stored, inline, styleIndex, rowIndex, column);

        return (column, Cell.Create(value, formula, styles.StyleAt(styleIndex)));
    }

    private CellValue ToValue(string? type, string? stored, string? inline, int styleIndex, int rowIndex, int column)
    {
        switch (type)
        {
            case "s":
                if (int.TryParse(stored, NumberStyles.None, CultureInfo.InvariantCulture, out int index) == false || index >= sharedStrings.Count)
                {
                    throw Damaged(rowIndex, column, $"it points to shared text {stored}, which does not exist");
                }

                return CellValue.Text(sharedStrings[index]);

            case "inlineStr":
                return CellValue.Text(inline ?? string.Empty);

            case "str":
                return CellValue.Text(CellText.Decode(stored ?? string.Empty));

            case "b":
                return CellValue.Bool(stored is "1" or "true");

            case "e":
                return ToError(stored);

            case "d":
                if (DateTime.TryParse(stored, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime isoDate) == true)
                {
                    return CellValue.Date(isoDate);
                }

                throw Damaged(rowIndex, column, $"'{stored}' is not a date");

            default:
                if (string.IsNullOrEmpty(stored) == true)
                {
                    return CellValue.Empty;
                }

                if (double.TryParse(stored, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == false)
                {
                    throw Damaged(rowIndex, column, $"'{stored}' is not a number");
                }

                if (styles.IsDate(styleIndex) == true && ExcelDate.TryFromSerial(number, uses1904, out DateTime date) == true)
                {
                    return CellValue.Date(date);
                }

                return CellValue.Number(number);
        }
    }

    private static CellValue ToError(string? stored)
    {
        try
        {
            return CellValue.Error(stored ?? "#N/A");
        }
        catch (ArgumentException)
        {
            // Ett felvärde från en nyare Excel-version som paketet inte känner till.
            return CellValue.Error("#N/A");
        }
    }

    private InvalidDataException Damaged(int rowIndex, int column, string problem)
    {
        CellAddress address = new(Math.Min(rowIndex, CellAddress.MaxRows - 1), column);

        return new InvalidDataException($"The file is damaged: cell {address} (row {rowIndex}, column {column}) in sheet '{Info.Name}' cannot be read, because {problem}.");
    }

    private async Task<bool> ReadOrThrowAsync(XmlReader reader)
    {
        try
        {
            return await reader.ReadAsync();
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"The file is damaged: the sheet '{Info.Name}' has broken XML at line {exception.LineNumber}.", exception);
        }
    }
}
