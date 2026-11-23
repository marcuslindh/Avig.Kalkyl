using System.Globalization;
using System.Xml;
using Avig.Kalkyl.Formulas;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Internal;

/// <summary>What is needed to interpret the cells of a sheet.</summary>
internal sealed class CellContext
{
    public CellContext(IReadOnlyList<string> sharedStrings, StyleTable styles, bool uses1904, string sheetName)
    {
        SharedStrings = sharedStrings;
        Styles = styles;
        Uses1904 = uses1904;
        SheetName = sheetName;
    }

    public IReadOnlyList<string> SharedStrings { get; }

    public StyleTable Styles { get; }

    public bool Uses1904 { get; }

    public string SheetName { get; }

    /// <summary>Shared formulas seen so far in the sheet: shared index to (formula, row, column) of the first cell.</summary>
    public Dictionary<string, (string Formula, int Row, int Column)> SharedFormulas { get; } = new(StringComparer.Ordinal);
}

/// <summary>A parsed cell, and whether it was part of a shared formula.</summary>
internal readonly record struct ParsedCell(int Column, Cell Cell, int StyleIndex, bool SharedFormula);

/// <summary>Writes and parses &lt;c&gt; elements. Shared by the streaming mode and the document model.</summary>
internal static class CellXml
{
    /// <summary>The number format a date without a format of its own gets; empty for anything but a date.</summary>
    public static string DefaultFormat(CellValue value)
    {
        if (value.TryGetDate(out DateTime date) == false)
        {
            return string.Empty;
        }

        if (date.TimeOfDay != TimeSpan.Zero)
        {
            return NumberFormats.DateTimeFormat;
        }

        return NumberFormats.DateFormat;
    }

    public static async Task WriteCellAsync(XmlWriter xml, Cell cell, string reference, int styleIndex, bool uses1904)
    {
        CellValue value = cell.Value;
        await xml.WriteStartElementAsync(null, "c", null);
        await xml.WriteAttributeStringAsync(null, "r", null, reference);

        if (styleIndex != 0)
        {
            await xml.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture));
        }

        if (cell.Formula is not null)
        {
            if (value.Kind == CellKind.Text)
            {
                await xml.WriteAttributeStringAsync(null, "t", null, "str");
            }
            else
            {
                await WriteTypeAttributeAsync(xml, value.Kind);
            }

            await xml.WriteElementStringAsync(null, "f", null, CellText.Encode(cell.Formula));

            if (value.IsEmpty == false)
            {
                await xml.WriteElementStringAsync(null, "v", null, StoredValue(value, uses1904));
            }

            await xml.WriteEndElementAsync();

            return;
        }

        switch (value.Kind)
        {
            case CellKind.Text:
                value.TryGetText(out string? text);
                await xml.WriteAttributeStringAsync(null, "t", null, "inlineStr");
                await xml.WriteStartElementAsync(null, "is", null);
                await xml.WriteStartElementAsync(null, "t", null);
                await xml.WriteAttributeStringAsync("xml", "space", null, "preserve");
                await xml.WriteStringAsync(CellText.Encode(text!));
                await xml.WriteEndElementAsync();
                await xml.WriteEndElementAsync();

                break;

            case CellKind.Empty:
                break;

            default:
                await WriteTypeAttributeAsync(xml, value.Kind);
                await xml.WriteElementStringAsync(null, "v", null, StoredValue(value, uses1904));

                break;
        }

        await xml.WriteEndElementAsync();
    }

    /// <summary>Parses the &lt;c&gt; the reader is on. Afterwards the reader is on the node after the element.</summary>
    public static async Task<ParsedCell> ReadCellAsync(XmlReader reader, int rowIndex, int previousColumn, CellContext context)
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
            throw Damaged(context, rowIndex, CellAddress.MaxColumns - 1, "the row has more cells than a sheet has columns");
        }

        string? type = reader.GetAttribute("t");
        int.TryParse(reader.GetAttribute("s"), NumberStyles.None, CultureInfo.InvariantCulture, out int styleIndex);
        string? stored = null;
        string? formula = null;
        string? inline = null;
        bool shared = false;

        if (reader.IsEmptyElement == false)
        {
            int depth = reader.Depth;
            await ReadAsync(reader, context);

            while (XmlParts.IsEnd(reader, depth) == false)
            {
                if (XmlParts.IsElement(reader, "v") == true)
                {
                    stored = await reader.ReadElementContentAsStringAsync();

                    continue;
                }

                if (XmlParts.IsElement(reader, "f") == true)
                {
                    string? formulaType = reader.GetAttribute("t");
                    string? sharedIndex = reader.GetAttribute("si");
                    string text = CellText.Decode(await reader.ReadElementContentAsStringAsync());
                    shared = string.Equals(formulaType, "shared", StringComparison.Ordinal);
                    TryResolveFormula(text, shared, sharedIndex, rowIndex, column, context, out formula);

                    continue;
                }

                if (XmlParts.IsElement(reader, "is") == true)
                {
                    inline = await XmlParts.ReadRichTextAsync(reader);
                }

                if (await ReadAsync(reader, context) == false)
                {
                    break;
                }
            }
        }

        await ReadAsync(reader, context);
        CellValue value = ToValue(type, stored, inline, styleIndex, rowIndex, column, context);

        return new ParsedCell(column, Cell.Create(value, formula, context.Styles.StyleAt(styleIndex)), styleIndex, shared);
    }

    public static async Task<bool> ReadAsync(XmlReader reader, CellContext context)
    {
        try
        {
            return await reader.ReadAsync();
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"The file is damaged: the sheet '{context.SheetName}' has broken XML at line {exception.LineNumber}.", exception);
        }
    }

    /// <summary>
    /// The formula text for a cell. A shared formula is stored once, in its first cell; the other
    /// cells only carry the shared index, and their text is the first cell's formula moved by the
    /// distance between the two cells.
    /// </summary>
    private static bool TryResolveFormula(string text, bool shared, string? sharedIndex, int rowIndex, int column, CellContext context, out string? formula)
    {
        formula = null;

        if (shared == false || sharedIndex is null)
        {
            if (text.Length > 0)
            {
                formula = text;
            }

            return formula is not null;
        }

        if (text.Length > 0)
        {
            context.SharedFormulas[sharedIndex] = (text, rowIndex, column);
            formula = text;

            return true;
        }

        if (context.SharedFormulas.TryGetValue(sharedIndex, out (string Formula, int Row, int Column) master) == false)
        {
            return false;
        }

        Result<string> shifted = FormulaShifter.Shift(master.Formula, rowIndex - master.Row, column - master.Column);

        return shifted.IsSuccess(out formula);
    }

    private static CellValue ToValue(string? type, string? stored, string? inline, int styleIndex, int rowIndex, int column, CellContext context)
    {
        switch (type)
        {
            case "s":
                if (int.TryParse(stored, NumberStyles.None, CultureInfo.InvariantCulture, out int index) == false || index >= context.SharedStrings.Count)
                {
                    throw Damaged(context, rowIndex, column, $"it points to shared text {stored}, which does not exist");
                }

                return CellValue.Text(context.SharedStrings[index]);

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

                throw Damaged(context, rowIndex, column, $"'{stored}' is not a date");

            default:
                if (string.IsNullOrEmpty(stored) == true)
                {
                    return CellValue.Empty;
                }

                if (double.TryParse(stored, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == false)
                {
                    throw Damaged(context, rowIndex, column, $"'{stored}' is not a number");
                }

                if (context.Styles.IsDate(styleIndex) == true && ExcelDate.TryFromSerial(number, context.Uses1904, out DateTime date) == true)
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
            // An error value from a newer Excel version that the package does not know.
            return CellValue.Error("#N/A");
        }
    }

    private static InvalidDataException Damaged(CellContext context, int rowIndex, int column, string problem)
    {
        CellAddress address = new(Math.Min(rowIndex, CellAddress.MaxRows - 1), column);

        return new InvalidDataException($"The file is damaged: cell {address} (row {rowIndex}, column {column}) in sheet '{context.SheetName}' cannot be read, because {problem}.");
    }

    private static async Task WriteTypeAttributeAsync(XmlWriter xml, CellKind kind)
    {
        if (kind == CellKind.Bool)
        {
            await xml.WriteAttributeStringAsync(null, "t", null, "b");
        }
        else if (kind == CellKind.Error)
        {
            await xml.WriteAttributeStringAsync(null, "t", null, "e");
        }
    }

    private static string StoredValue(CellValue value, bool uses1904)
    {
        switch (value.Kind)
        {
            case CellKind.Number:
                value.TryGetNumber(out double number);

                return number.ToString("R", CultureInfo.InvariantCulture);

            case CellKind.Date:
                value.TryGetDate(out DateTime date);

                return ExcelDate.ToSerial(date, uses1904).ToString("R", CultureInfo.InvariantCulture);

            case CellKind.Bool:
                value.TryGetBool(out bool flag);

                if (flag == true)
                {
                    return "1";
                }

                return "0";

            case CellKind.Text:
                value.TryGetText(out string? text);

                return CellText.Encode(text!);

            case CellKind.Error:
                value.TryGetError(out string? code);

                return code!;

            default:
                return string.Empty;
        }
    }
}
