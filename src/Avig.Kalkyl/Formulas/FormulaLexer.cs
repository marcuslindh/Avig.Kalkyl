using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Formulas;

/// <summary>The kind of a token.</summary>
internal enum TokenKind
{
    Number,
    Text,
    Bool,
    Error,
    Reference,
    Function,
    Name,
    Operator,
    Comma,
    Semicolon,
    OpenParen,
    CloseParen,
    OpenBrace,
    CloseBrace,
    StructuredReference,
    End,
}

/// <summary>The kind of a reference.</summary>
internal enum ReferenceKind
{
    Cell,
    Area,
    Columns,
    Rows,
}

/// <summary>
/// A reference in a formula: a cell, an area, whole columns or whole rows, with or without
/// a sheet name. Indexes are 0-based. For whole columns the rows are -1, for whole rows the columns.
/// </summary>
internal sealed record ReferenceToken(
    string? Sheet,
    ReferenceKind Kind,
    int FirstRow,
    int FirstColumn,
    bool FirstRowAbsolute,
    bool FirstColumnAbsolute,
    int LastRow,
    int LastColumn,
    bool LastRowAbsolute,
    bool LastColumnAbsolute)
{
    public string Format(bool quoteSheet = false)
    {
        StringBuilder builder = new();

        if (Sheet is not null)
        {
            builder.Append(FormulaLexer.QuoteSheet(Sheet, quoteSheet)).Append('!');
        }

        switch (Kind)
        {
            case ReferenceKind.Cell:
                AppendCell(builder, FirstRow, FirstColumn, FirstRowAbsolute, FirstColumnAbsolute);

                break;

            case ReferenceKind.Area:
                AppendCell(builder, FirstRow, FirstColumn, FirstRowAbsolute, FirstColumnAbsolute);
                builder.Append(':');
                AppendCell(builder, LastRow, LastColumn, LastRowAbsolute, LastColumnAbsolute);

                break;

            case ReferenceKind.Columns:
                AppendColumn(builder, FirstColumn, FirstColumnAbsolute);
                builder.Append(':');
                AppendColumn(builder, LastColumn, LastColumnAbsolute);

                break;

            case ReferenceKind.Rows:
                AppendRow(builder, FirstRow, FirstRowAbsolute);
                builder.Append(':');
                AppendRow(builder, LastRow, LastRowAbsolute);

                break;
        }

        return builder.ToString();
    }

    private static void AppendCell(StringBuilder builder, int row, int column, bool rowAbsolute, bool columnAbsolute)
    {
        AppendColumn(builder, column, columnAbsolute);
        AppendRow(builder, row, rowAbsolute);
    }

    private static void AppendColumn(StringBuilder builder, int column, bool absolute)
    {
        if (absolute == true)
        {
            builder.Append('$');
        }

        builder.Append(CellAddress.ColumnName(column));
    }

    private static void AppendRow(StringBuilder builder, int row, bool absolute)
    {
        if (absolute == true)
        {
            builder.Append('$');
        }

        builder.Append((row + 1).ToString(CultureInfo.InvariantCulture));
    }
}

/// <summary>A token with its text as it stood in the formula.</summary>
internal sealed record Token(TokenKind Kind, string Text, int Start)
{
    public ReferenceToken? Reference { get; init; }

    public double Number { get; init; }
}

/// <summary>
/// Splits a formula into tokens. Covers Excel's formula language in A1 style: numbers, text,
/// booleans, errors, references with sheet names (quoted too), whole columns and rows,
/// functions (with dots and the _xlfn. prefix too), names, operators and array constants.
/// Structured references (Table[Column]) become a token of their own that evaluates to #REF!.
/// </summary>
internal static class FormulaLexer
{
    private const string RefError = "#REF!";

    private static readonly string[] ErrorLiterals = ["#GETTING_DATA", "#DIV/0!", "#VALUE!", "#NULL!", "#NAME?", "#SPILL!", "#CALC!", "#REF!", "#NUM!", "#N/A"];

    public static string QuoteSheet(string sheet, bool always = false)
    {
        if (always == true)
        {
            return "'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'";
        }

        // A 3D range of two plain names stays as it is (Jan:Mar); otherwise it is quoted whole.
        string[] ends = sheet.Split(':');

        if (ends.Length == 2 && IsPlainSheet(ends[0]) == true && IsPlainSheet(ends[1]) == true)
        {
            return sheet;
        }

        if (IsPlainSheet(sheet) == true)
        {
            return sheet;
        }

        return "'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'";
    }

    /// <summary>A sheet name that needs no quotes: letters, digits, _ and ., starting with a letter or _, and not a cell address or TRUE or FALSE.</summary>
    private static bool IsPlainSheet(string sheet)
    {
        bool plain = sheet.Length > 0 && (char.IsLetter(sheet[0]) == true || sheet[0] == '_');

        foreach (char letter in sheet)
        {
            if (char.IsLetterOrDigit(letter) == false && letter != '_' && letter != '.')
            {
                plain = false;
            }
        }

        return plain == true && CellAddress.TryParse(sheet, out _) == false && string.Equals(sheet, "TRUE", StringComparison.OrdinalIgnoreCase) == false && string.Equals(sheet, "FALSE", StringComparison.OrdinalIgnoreCase) == false;
    }

    public static Result<List<Token>> Tokenize(string formula)
    {
        List<Token> tokens = [];
        int position = 0;

        if (formula.StartsWith('=') == true)
        {
            position = 1;
        }

        while (position < formula.Length)
        {
            char current = formula[position];

            if (char.IsWhiteSpace(current) == true || current == '@')
            {
                position++;

                continue;
            }

            int start = position;

            switch (current)
            {
                case '"':
                    Result<int> textEnd = ReadText(formula, position, out string text);

                    if (textEnd.IsSuccess(out int afterText) == false)
                    {
                        return Result<List<Token>>.Invalid(textEnd.Error!);
                    }

                    tokens.Add(new Token(TokenKind.Text, text, start));
                    position = afterText;

                    continue;

                case '#':
                    string? error = ErrorLiterals.FirstOrDefault(literal => string.Compare(formula, position, literal, 0, literal.Length, StringComparison.OrdinalIgnoreCase) == 0);

                    if (error is null)
                    {
                        return Result<List<Token>>.Invalid($"Unknown error value at position {position} in '{formula}'.");
                    }

                    tokens.Add(new Token(TokenKind.Error, error, start));
                    position += error.Length;

                    continue;

                case '\'':
                    Result<int> quoted = ReadQuotedReference(formula, position, tokens);

                    if (quoted.IsSuccess(out int afterQuoted) == false)
                    {
                        return Result<List<Token>>.Invalid(quoted.Error!);
                    }

                    position = afterQuoted;

                    continue;

                case '(':
                    tokens.Add(new Token(TokenKind.OpenParen, "(", start));
                    position++;

                    continue;

                case ')':
                    tokens.Add(new Token(TokenKind.CloseParen, ")", start));
                    position++;

                    continue;

                case '{':
                    tokens.Add(new Token(TokenKind.OpenBrace, "{", start));
                    position++;

                    continue;

                case '}':
                    tokens.Add(new Token(TokenKind.CloseBrace, "}", start));
                    position++;

                    continue;

                case ',':
                    tokens.Add(new Token(TokenKind.Comma, ",", start));
                    position++;

                    continue;

                case ';':
                    tokens.Add(new Token(TokenKind.Semicolon, ";", start));
                    position++;

                    continue;
            }

            if (TryReadOperator(formula, position, out string? op) == true)
            {
                tokens.Add(new Token(TokenKind.Operator, op, start));
                position += op.Length;

                continue;
            }

            if (char.IsAsciiDigit(current) == true || (current == '.' && position + 1 < formula.Length && char.IsAsciiDigit(formula[position + 1]) == true))
            {
                if (TryReadRowRange(formula, position, null, out ReferenceToken? rows, out int afterRows) == true)
                {
                    tokens.Add(new Token(TokenKind.Reference, formula[start..afterRows], start) { Reference = rows });
                    position = afterRows;

                    continue;
                }

                int afterNumber = ReadNumber(formula, position, out double number);

                if (afterNumber == position)
                {
                    return Result<List<Token>>.Invalid($"'{formula}' has a malformed number at position {position}.");
                }

                tokens.Add(new Token(TokenKind.Number, formula[start..afterNumber], start) { Number = number });
                position = afterNumber;

                continue;
            }

            if (char.IsLetter(current) == true || current == '_' || current == '\\' || current == '$')
            {
                Result<int> word = ReadWord(formula, position, tokens);

                if (word.IsSuccess(out int afterWord) == false)
                {
                    return Result<List<Token>>.Invalid(word.Error!);
                }

                position = afterWord;

                continue;
            }

            return Result<List<Token>>.Invalid($"'{formula}' has an unexpected character '{current}' at position {position}.");
        }

        tokens.Add(new Token(TokenKind.End, string.Empty, formula.Length));

        return tokens;
    }

    private static bool TryReadOperator(string formula, int position, out string op)
    {
        op = string.Empty;
        char current = formula[position];
        char next = '\0';

        if (position + 1 < formula.Length)
        {
            next = formula[position + 1];
        }

        if ((current == '<' && (next == '=' || next == '>')) || (current == '>' && next == '='))
        {
            op = formula.Substring(position, 2);

            return true;
        }

        if ("+-*/^&=<>%:".Contains(current, StringComparison.Ordinal) == true)
        {
            op = current.ToString();

            return true;
        }

        return false;
    }

    private static Result<int> ReadText(string formula, int position, out string text)
    {
        StringBuilder builder = new();
        int index = position + 1;

        while (index < formula.Length)
        {
            if (formula[index] == '"')
            {
                if (index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    builder.Append('"');
                    index += 2;

                    continue;
                }

                text = builder.ToString();

                return index + 1;
            }

            builder.Append(formula[index]);
            index++;
        }

        text = string.Empty;

        return Result<int>.Invalid($"'{formula}' has text without a closing quote.");
    }

    private static int ReadNumber(string formula, int position, out double number)
    {
        int index = position;

        while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
        {
            index++;
        }

        if (index < formula.Length && formula[index] == '.')
        {
            index++;

            while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
            {
                index++;
            }
        }

        if (index < formula.Length && (formula[index] == 'E' || formula[index] == 'e'))
        {
            int exponent = index + 1;

            if (exponent < formula.Length && (formula[exponent] == '+' || formula[exponent] == '-'))
            {
                exponent++;
            }

            if (exponent < formula.Length && char.IsAsciiDigit(formula[exponent]) == true)
            {
                index = exponent;

                while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
                {
                    index++;
                }
            }
        }

        double.TryParse(formula.AsSpan(position, index - position), NumberStyles.Float, CultureInfo.InvariantCulture, out number);

        return index;
    }

    private static Result<int> ReadQuotedReference(string formula, int position, List<Token> tokens)
    {
        StringBuilder sheet = new();
        int index = position + 1;

        while (index < formula.Length)
        {
            if (formula[index] == '\'')
            {
                if (index + 1 < formula.Length && formula[index + 1] == '\'')
                {
                    sheet.Append('\'');
                    index += 2;

                    continue;
                }

                break;
            }

            sheet.Append(formula[index]);
            index++;
        }

        if (index + 1 >= formula.Length || formula[index + 1] != '!')
        {
            return Result<int>.Invalid($"'{formula}' has a quoted sheet name without a reference after it.");
        }

        int referenceStart = index + 2;

        if (IsRefError(formula, referenceStart) == true)
        {
            tokens.Add(new Token(TokenKind.Error, "#REF!", position));

            return referenceStart + RefError.Length;
        }

        if (TryReadReference(formula, referenceStart, sheet.ToString(), out ReferenceToken? reference, out int after) == false)
        {
            return Result<int>.Invalid($"'{formula}' has no valid reference after the sheet name '{sheet}'.");
        }

        tokens.Add(new Token(TokenKind.Reference, formula[position..after], position) { Reference = reference });

        return after;
    }

    /// <summary>
    /// True when <c>#REF!</c> starts at <paramref name="position"/>. Excel writes a reference
    /// whose cells were deleted as <c>Sheet!#REF!</c>, and moving a formula off the sheet does
    /// the same, so it reads as the error rather than as a broken formula.
    /// </summary>
    private static bool IsRefError(string formula, int position)
    {
        return position + RefError.Length <= formula.Length
            && string.Compare(formula, position, RefError, 0, RefError.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static Result<int> ReadWord(string formula, int position, List<Token> tokens)
    {
        int index = position;

        while (index < formula.Length && (char.IsLetterOrDigit(formula[index]) == true || formula[index] is '_' or '.' or '\\' or '$' or '?'))
        {
            index++;
        }

        string word = formula[position..index];

        // A 3D reference, Sheet1:Sheet3!A1, spans the sheets from the first to the last. Its sheet
        // name is the whole range, so that it is written back as Excel writes it, quoted whole
        // when it has to be ('A2:A9'!A1): quoting only one end would make the other a cell or a name.
        if (TryReadSheetRangeEnd(formula, index, out int bang) == true)
        {
            word = formula[position..bang];
            index = bang;
        }

        if (index < formula.Length && formula[index] == '!')
        {
            if (IsRefError(formula, index + 1) == true)
            {
                tokens.Add(new Token(TokenKind.Error, "#REF!", position));

                return index + 1 + RefError.Length;
            }

            if (TryReadReference(formula, index + 1, word, out ReferenceToken? sheetReference, out int afterSheet) == false)
            {
                return Result<int>.Invalid($"'{formula}' has no valid reference after the sheet name '{word}'.");
            }

            tokens.Add(new Token(TokenKind.Reference, formula[position..afterSheet], position) { Reference = sheetReference });

            return afterSheet;
        }

        if (index < formula.Length && formula[index] == '(')
        {
            tokens.Add(new Token(TokenKind.Function, word, position));

            return index;
        }

        if (index < formula.Length && formula[index] == '[')
        {
            int depth = 0;
            int end = index;

            while (end < formula.Length)
            {
                if (formula[end] == '[')
                {
                    depth++;
                }
                else if (formula[end] == ']')
                {
                    depth--;

                    if (depth == 0)
                    {
                        break;
                    }
                }

                end++;
            }

            tokens.Add(new Token(TokenKind.StructuredReference, formula[position..Math.Min(end + 1, formula.Length)], position));

            return Math.Min(end + 1, formula.Length);
        }

        if (string.Equals(word, "TRUE", StringComparison.OrdinalIgnoreCase) == true || string.Equals(word, "FALSE", StringComparison.OrdinalIgnoreCase) == true)
        {
            tokens.Add(new Token(TokenKind.Bool, word.ToUpperInvariant(), position));

            return index;
        }

        if (TryReadReference(formula, position, null, out ReferenceToken? reference, out int after) == true)
        {
            tokens.Add(new Token(TokenKind.Reference, formula[position..after], position) { Reference = reference });

            return after;
        }

        if (word.Contains('$', StringComparison.Ordinal) == true)
        {
            return Result<int>.Invalid($"'{formula}' has a malformed reference '{word}'.");
        }

        tokens.Add(new Token(TokenKind.Name, word, position));

        return index;
    }

    /// <summary>A reference starting at a position: a cell, an area, columns (A:C) or rows (1:3).</summary>
    /// <summary>
    /// True when a sheet name ending at <paramref name="colon"/> is the first of a 3D range: a
    /// colon, a second sheet name and the !. <paramref name="bang"/> is where the ! is.
    /// </summary>
    private static bool TryReadSheetRangeEnd(string formula, int colon, out int bang)
    {
        bang = colon;

        if (colon >= formula.Length || formula[colon] != ':')
        {
            return false;
        }

        int index = colon + 1;

        while (index < formula.Length && (char.IsLetterOrDigit(formula[index]) == true || formula[index] is '_' or '.'))
        {
            index++;
        }

        if (index == colon + 1 || index >= formula.Length || formula[index] != '!')
        {
            return false;
        }

        bang = index;

        return true;
    }

    private static bool TryReadReference(string formula, int position, string? sheet, out ReferenceToken? reference, out int after)
    {
        reference = null;
        after = position;

        if (TryReadRowRange(formula, position, sheet, out reference, out after) == true)
        {
            return true;
        }

        if (TryReadCellPart(formula, position, out int firstRow, out int firstColumn, out bool firstRowAbsolute, out bool firstColumnAbsolute, out int afterFirst) == true)
        {
            if (afterFirst < formula.Length && formula[afterFirst] == ':'
                && TryReadCellPart(formula, afterFirst + 1, out int lastRow, out int lastColumn, out bool lastRowAbsolute, out bool lastColumnAbsolute, out int afterLast) == true
                && lastRow >= 0 && firstRow >= 0)
            {
                reference = new ReferenceToken(sheet, ReferenceKind.Area, firstRow, firstColumn, firstRowAbsolute, firstColumnAbsolute, lastRow, lastColumn, lastRowAbsolute, lastColumnAbsolute);
                after = afterLast;

                return IsBoundary(formula, after);
            }

            if (firstRow >= 0)
            {
                reference = new ReferenceToken(sheet, ReferenceKind.Cell, firstRow, firstColumn, firstRowAbsolute, firstColumnAbsolute, firstRow, firstColumn, firstRowAbsolute, firstColumnAbsolute);
                after = afterFirst;

                return IsBoundary(formula, after);
            }

            // Column letters only: must be a column range such as A:C.
            if (afterFirst < formula.Length && formula[afterFirst] == ':'
                && TryReadCellPart(formula, afterFirst + 1, out int endRow, out int endColumn, out _, out bool endColumnAbsolute, out int afterEnd) == true
                && endRow < 0)
            {
                reference = new ReferenceToken(sheet, ReferenceKind.Columns, -1, firstColumn, false, firstColumnAbsolute, -1, endColumn, false, endColumnAbsolute);
                after = afterEnd;

                return IsBoundary(formula, after);
            }
        }

        return false;
    }

    private static bool TryReadRowRange(string formula, int position, string? sheet, out ReferenceToken? reference, out int after)
    {
        reference = null;
        after = position;

        if (TryReadRowPart(formula, position, out int firstRow, out bool firstAbsolute, out int afterFirst) == false
            || afterFirst >= formula.Length || formula[afterFirst] != ':'
            || TryReadRowPart(formula, afterFirst + 1, out int lastRow, out bool lastAbsolute, out int afterLast) == false
            || IsBoundary(formula, afterLast) == false)
        {
            return false;
        }

        reference = new ReferenceToken(sheet, ReferenceKind.Rows, firstRow, -1, firstAbsolute, false, lastRow, -1, lastAbsolute, false);
        after = afterLast;

        return true;
    }

    private static bool TryReadRowPart(string formula, int position, out int row, out bool absolute, out int after)
    {
        row = -1;
        absolute = false;
        after = position;
        int index = position;

        if (index < formula.Length && formula[index] == '$')
        {
            absolute = true;
            index++;
        }

        int digits = index;

        while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
        {
            index++;
        }

        if (index == digits || formula[digits] == '0' || (index < formula.Length && (char.IsLetter(formula[index]) == true || formula[index] == '.')))
        {
            return false;
        }

        if (int.TryParse(formula.AsSpan(digits, index - digits), NumberStyles.None, CultureInfo.InvariantCulture, out int number) == false || number > CellAddress.MaxRows)
        {
            return false;
        }

        row = number - 1;
        after = index;

        return true;
    }

    /// <summary>Column letters with or without a row number. The row is -1 when there is none.</summary>
    private static bool TryReadCellPart(string formula, int position, out int row, out int column, out bool rowAbsolute, out bool columnAbsolute, out int after)
    {
        row = -1;
        column = -1;
        rowAbsolute = false;
        columnAbsolute = false;
        after = position;
        int index = position;

        if (index < formula.Length && formula[index] == '$')
        {
            columnAbsolute = true;
            index++;
        }

        int letters = index;

        while (index < formula.Length && char.IsAsciiLetter(formula[index]) == true)
        {
            index++;
        }

        if (CellAddress.TryParseColumn(formula.AsSpan(letters, index - letters), out column) == false)
        {
            return false;
        }

        int rowStart = index;

        if (index < formula.Length && formula[index] == '$')
        {
            rowAbsolute = true;
            index++;
        }

        int digits = index;

        while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
        {
            index++;
        }

        if (index == digits)
        {
            // No digits: a column only. A $ with no digits after it is not valid.
            if (rowAbsolute == true)
            {
                return false;
            }

            after = rowStart;

            return true;
        }

        if (formula[digits] == '0' || int.TryParse(formula.AsSpan(digits, index - digits), NumberStyles.None, CultureInfo.InvariantCulture, out int number) == false || number > CellAddress.MaxRows)
        {
            return false;
        }

        row = number - 1;
        after = index;

        return true;
    }

    private static bool IsBoundary(string formula, int position)
    {
        if (position >= formula.Length)
        {
            return true;
        }

        char next = formula[position];

        return char.IsLetterOrDigit(next) == false && next != '_' && next != '(' && next != '.' && next != '!' && next != '$';
    }
}
