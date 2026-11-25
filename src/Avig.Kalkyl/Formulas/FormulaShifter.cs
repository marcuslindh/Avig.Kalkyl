using System.Text;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// Moves the relative references in a formula, as Excel does when a formula is copied to
/// another cell. Used for shared formulas: Excel stores the formula once, in the first cell,
/// and the other cells only refer to it.
/// </summary>
internal static class FormulaShifter
{
    /// <summary>
    /// The formula as it becomes when moved <paramref name="rows"/> rows and
    /// <paramref name="columns"/> columns. A reference that ends up outside the sheet becomes #REF!.
    /// </summary>
    public static Result<string> Shift(string formula, int rows, int columns)
    {
        Result<List<Token>> tokenized = FormulaLexer.Tokenize(formula);

        if (tokenized.IsSuccess(out List<Token>? tokens) == false)
        {
            return Result<string>.Invalid(tokenized.Error!);
        }

        string source = formula;
        int offset = 0;

        if (source.StartsWith('=') == true)
        {
            offset = 1;
        }

        StringBuilder builder = new();
        int copied = offset;

        foreach (Token token in tokens)
        {
            if (token.Kind != TokenKind.Reference || token.Reference is null)
            {
                continue;
            }

            builder.Append(source, copied, token.Start - copied);
            // A sheet name straight after a colon is quoted: A1:Data!B2 would read as the 3D
            // reference 'A1:Data'!B2 the next time, where A1:'Data'!B2 is the area from A1 to Data!B2.
            bool afterColon = token.Start > offset && source[token.Start - 1] == ':';
            builder.Append(ShiftReference(token.Reference, rows, columns, afterColon));
            copied = token.Start + token.Text.Length;
        }

        builder.Append(source, copied, source.Length - copied);

        return builder.ToString();
    }

    private static string ShiftReference(ReferenceToken reference, int rows, int columns, bool quoteSheet)
    {
        int firstRow = Move(reference.FirstRow, reference.FirstRowAbsolute, rows);
        int lastRow = Move(reference.LastRow, reference.LastRowAbsolute, rows);
        int firstColumn = Move(reference.FirstColumn, reference.FirstColumnAbsolute, columns);
        int lastColumn = Move(reference.LastColumn, reference.LastColumnAbsolute, columns);
        bool rowsOutside = reference.Kind != ReferenceKind.Columns && (Outside(firstRow, CellAddress.MaxRows) || Outside(lastRow, CellAddress.MaxRows));
        bool columnsOutside = reference.Kind != ReferenceKind.Rows && (Outside(firstColumn, CellAddress.MaxColumns) || Outside(lastColumn, CellAddress.MaxColumns));

        if (rowsOutside == true || columnsOutside == true)
        {
            if (reference.Sheet is not null)
            {
                return FormulaLexer.QuoteSheet(reference.Sheet, quoteSheet) + "!#REF!";
            }

            return "#REF!";
        }

        ReferenceToken moved = reference with
        {
            FirstRow = firstRow,
            LastRow = lastRow,
            FirstColumn = firstColumn,
            LastColumn = lastColumn,
        };

        return moved.Format(quoteSheet);
    }

    private static int Move(int value, bool absolute, int delta)
    {
        if (absolute == true || value < 0)
        {
            return value;
        }

        return value + delta;
    }

    private static bool Outside(int value, int count)
    {
        return value < 0 || value >= count;
    }
}
