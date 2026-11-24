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
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        CellContext context = new(sharedStrings, styles, uses1904, Info.Name);
        int previousRow = -1;

        while (await CellXml.ReadAsync(reader, context) == true)
        {
            if (XmlParts.IsElement(reader, "row") == false)
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            int rowIndex = CellXml.RowIndex(reader, previousRow, context);

            previousRow = rowIndex;

            if (reader.IsEmptyElement == true)
            {
                yield return new Row(rowIndex, []);

                continue;
            }

            yield return await ReadRowAsync(reader, rowIndex, context);
        }
    }

    private static async Task<Row> ReadRowAsync(XmlReader reader, int rowIndex, CellContext context)
    {
        int depth = reader.Depth;
        List<(int Column, Cell Cell)> cells = [];
        int previousColumn = -1;
        int lastColumn = -1;
        await CellXml.ReadAsync(reader, context);

        while (XmlParts.IsEnd(reader, depth) == false)
        {
            if (XmlParts.IsElement(reader, "c") == true)
            {
                ParsedCell parsed = await CellXml.ReadCellAsync(reader, rowIndex, previousColumn, context);
                previousColumn = parsed.Column;

                if (parsed.Cell.IsEmpty == false)
                {
                    cells.Add((parsed.Column, parsed.Cell));
                    lastColumn = Math.Max(lastColumn, parsed.Column);
                }

                continue;
            }

            if (await CellXml.ReadAsync(reader, context) == false)
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
}
