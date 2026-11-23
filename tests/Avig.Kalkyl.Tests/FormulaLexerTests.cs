using Avig.Kalkyl.Formulas;

namespace Avig.Kalkyl.Tests;

public class FormulaLexerTests
{
    [Theory]
    [InlineData("A1*2", 1, 0, "A2*2")]
    [InlineData("$A$1+A1", 3, 2, "$A$1+C4")]
    [InlineData("A$1+$A1", 1, 1, "B$1+$A2")]
    [InlineData("SUM(A1:B2)", 2, 0, "SUM(A3:B4)")]
    [InlineData("SUM(A:A)", 0, 1, "SUM(B:B)")]
    [InlineData("SUM(1:1)", 4, 0, "SUM(5:5)")]
    [InlineData("Data!B2+'Min flik'!C3", 1, 0, "Data!B3+'Min flik'!C4")]
    [InlineData("\"A1\"&A1", 1, 0, "\"A1\"&A2")]
    [InlineData("LOG10(A1)+STDEV.S(B1:B3)", 1, 0, "LOG10(A2)+STDEV.S(B2:B4)")]
    [InlineData("_xlfn.XLOOKUP(A1,B:B,C:C)", 1, 0, "_xlfn.XLOOKUP(A2,B:B,C:C)")]
    [InlineData("A2-1", -1, 0, "A1-1")]
    [InlineData("A1+1", -1, 0, "#REF!+1")]
    [InlineData("IF(A1>=10%,TRUE,#N/A)", 1, 0, "IF(A2>=10%,TRUE,#N/A)")]
    public void Relative_references_move_and_absolute_ones_stay(string formula, int rows, int columns, string expected)
    {
        Result<string> shifted = FormulaShifter.Shift(formula, rows, columns);

        Assert.True(shifted.IsSuccess(out string? moved), shifted.Error);
        Assert.Equal(expected, moved);
    }

    [Fact]
    public void Tokens_cover_the_formula_language()
    {
        Result<List<Token>> tokenized = FormulaLexer.Tokenize("=SUM('Q1 2026'!$A$1:B10, {1,2;3,4}) * -2.5E3 & \"x\"\"y\" <> Total");

        Assert.True(tokenized.IsSuccess(out List<Token>? tokens), tokenized.Error);
        Assert.Equal(
            new[] { TokenKind.Function, TokenKind.OpenParen, TokenKind.Reference, TokenKind.Comma, TokenKind.OpenBrace, TokenKind.Number, TokenKind.Comma, TokenKind.Number, TokenKind.Semicolon, TokenKind.Number, TokenKind.Comma, TokenKind.Number, TokenKind.CloseBrace, TokenKind.CloseParen, TokenKind.Operator, TokenKind.Operator, TokenKind.Number, TokenKind.Operator, TokenKind.Text, TokenKind.Operator, TokenKind.Name, TokenKind.End },
            tokens.Select(token => token.Kind));
        Assert.Equal("Q1 2026", tokens[2].Reference!.Sheet);
        Assert.Equal(ReferenceKind.Area, tokens[2].Reference!.Kind);
        Assert.Equal("x\"y", tokens[18].Text);
        Assert.Equal(2500, tokens[16].Number);
    }

    [Theory]
    [InlineData("\"open")]
    [InlineData("#OOPS!")]
    [InlineData("A1 ~ 2")]
    public void Malformed_formulas_are_Invalid(string formula)
    {
        Assert.True(FormulaLexer.Tokenize(formula).IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
    }
}
