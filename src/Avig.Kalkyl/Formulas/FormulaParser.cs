using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>A node in a parsed formula.</summary>
internal abstract record Node;

internal sealed record LiteralNode(Value Value) : Node;

internal sealed record ReferenceNode(ReferenceToken Reference) : Node;

internal sealed record NameNode(string Name) : Node;

/// <summary>A function call. The name is upper case without the _xlfn./_xlws. prefixes Excel adds to newer functions.</summary>
internal sealed record FunctionNode(string Name, IReadOnlyList<Node> Arguments) : Node;

internal sealed record UnaryNode(string Operator, Node Operand) : Node;

internal sealed record BinaryNode(string Operator, Node Left, Node Right) : Node;

internal sealed record PercentNode(Node Operand) : Node;

internal sealed record ArrayNode(IReadOnlyList<IReadOnlyList<Value>> Rows) : Node;

/// <summary>An argument left out, as in IF(A1,,2).</summary>
internal sealed record MissingNode : Node;

/// <summary>
/// Parses a formula into a tree, with Excel's operator precedence: range (:), negation,
/// percent, ^, * and /, + and -, &amp;, then comparisons. As in Excel, negation binds tighter
/// than ^, so -2^2 is 4.
/// </summary>
internal sealed class FormulaParser
{
    /// <summary>Excel's limit on the length of a formula, not counting the leading '='.</summary>
    private const int MaxFormulaLength = 8192;

    /// <summary>Excel's limit on functions inside functions.</summary>
    private const int MaxFunctionNesting = 64;

    /// <summary>
    /// Parentheses, signs and functions inside each other, together. The parser and the
    /// evaluator go one level deeper on the stack for each, and a stack overflow ends the
    /// process instead of failing the formula: a thousand '(' did that. This is well below
    /// what the stack holds and well above what a workbook uses.
    /// </summary>
    private const int MaxNesting = 255;

    private readonly List<Token> tokens;
    private readonly string formula;
    private int position;
    private int nesting;
    private int functionNesting;

    private FormulaParser(List<Token> tokens, string formula)
    {
        this.tokens = tokens;
        this.formula = formula;
    }

    public static Result<Node> Parse(string formula)
    {
        if (formula.Length > MaxFormulaLength + 1)
        {
            return Result<Node>.Invalid($"The formula is {formula.Length:N0} characters long; Excel allows {MaxFormulaLength:N0}, and so does this.");
        }

        Result<List<Token>> tokenized = FormulaLexer.Tokenize(formula);

        if (tokenized.IsSuccess(out List<Token>? tokens) == false)
        {
            return Result<Node>.Invalid(tokenized.Error!);
        }

        FormulaParser parser = new(tokens, formula);

        try
        {
            Node node = parser.ParseComparison();

            if (parser.Current.Kind != TokenKind.End)
            {
                return Result<Node>.Invalid($"'{formula}' has unexpected '{parser.Current.Text}' at position {parser.Current.Start}.");
            }

            return node;
        }
        catch (FormatException exception)
        {
            return Result<Node>.Invalid(exception.Message);
        }
    }

    private Token Current => tokens[position];

    private Node ParseComparison()
    {
        Node left = ParseConcatenation();

        while (Current.Kind == TokenKind.Operator && Current.Text is "=" or "<>" or "<" or ">" or "<=" or ">=")
        {
            string op = Next().Text;
            left = new BinaryNode(op, left, ParseConcatenation());
        }

        return left;
    }

    private Node ParseConcatenation()
    {
        Node left = ParseAdditive();

        while (IsOperator("&") == true)
        {
            Next();
            left = new BinaryNode("&", left, ParseAdditive());
        }

        return left;
    }

    private Node ParseAdditive()
    {
        Node left = ParseMultiplicative();

        while (IsOperator("+") == true || IsOperator("-") == true)
        {
            string op = Next().Text;
            left = new BinaryNode(op, left, ParseMultiplicative());
        }

        return left;
    }

    private Node ParseMultiplicative()
    {
        Node left = ParsePower();

        while (IsOperator("*") == true || IsOperator("/") == true)
        {
            string op = Next().Text;
            left = new BinaryNode(op, left, ParsePower());
        }

        return left;
    }

    private Node ParsePower()
    {
        Node left = ParsePercent();

        while (IsOperator("^") == true)
        {
            Next();
            left = new BinaryNode("^", left, ParsePercent());
        }

        return left;
    }

    private Node ParsePercent()
    {
        Node operand = ParseUnary();

        while (IsOperator("%") == true)
        {
            Next();
            operand = new PercentNode(operand);
        }

        return operand;
    }

    private Node ParseUnary()
    {
        if (IsOperator("-") == true || IsOperator("+") == true)
        {
            string op = Next().Text;
            Enter();
            Node operand = ParseUnary();
            nesting--;

            return new UnaryNode(op, operand);
        }

        return ParseRange();
    }

    private void Enter()
    {
        nesting++;

        if (nesting > MaxNesting)
        {
            throw new FormatException($"'{Shorten(formula)}' nests parentheses, signs and functions more than {MaxNesting} levels deep.");
        }
    }

    private static string Shorten(string text)
    {
        if (text.Length <= 60)
        {
            return text;
        }

        return text[..60] + "...";
    }

    private Node ParseRange()
    {
        Node left = ParsePrimary();

        while (IsOperator(":") == true)
        {
            Next();
            left = new BinaryNode(":", left, ParsePrimary());
        }

        return left;
    }

    private Node ParsePrimary()
    {
        Token token = Next();

        switch (token.Kind)
        {
            case TokenKind.Number:
                return new LiteralNode(Value.FromNumber(token.Number));

            case TokenKind.Text:
                return new LiteralNode(Value.FromText(token.Text));

            case TokenKind.Bool:
                return new LiteralNode(Value.FromBool(token.Text == "TRUE"));

            case TokenKind.Error:
                return new LiteralNode(Value.Error(token.Text.ToUpperInvariant()));

            case TokenKind.Reference:
                return new ReferenceNode(token.Reference!);

            case TokenKind.Name:
                return new NameNode(token.Text);

            case TokenKind.StructuredReference:
                // Table references need the table definitions, which the engine does not read.
                return new LiteralNode(Value.Ref);

            case TokenKind.Function:
                Enter();
                functionNesting++;

                if (functionNesting > MaxFunctionNesting)
                {
                    throw new FormatException($"'{Shorten(formula)}' has more than {MaxFunctionNesting} levels of functions inside functions, which Excel does not allow either.");
                }

                FunctionNode function = ParseFunction(token);
                functionNesting--;
                nesting--;

                return function;

            case TokenKind.OpenParen:
                Enter();
                Node inner = ParseComparison();
                Expect(TokenKind.CloseParen, ")");
                nesting--;

                return inner;

            case TokenKind.OpenBrace:
                return ParseArray();

            default:
                throw new FormatException($"'{formula}' has unexpected '{token.Text}' at position {token.Start}.");
        }
    }

    private FunctionNode ParseFunction(Token name)
    {
        Expect(TokenKind.OpenParen, "(");
        List<Node> arguments = [];

        if (Current.Kind == TokenKind.CloseParen)
        {
            Next();

            return new FunctionNode(NormalizeName(name.Text), arguments);
        }

        while (true)
        {
            if (Current.Kind is TokenKind.Comma or TokenKind.CloseParen)
            {
                arguments.Add(new MissingNode());
            }
            else
            {
                arguments.Add(ParseComparison());
            }

            if (Current.Kind == TokenKind.Comma)
            {
                Next();

                continue;
            }

            Expect(TokenKind.CloseParen, ")");

            return new FunctionNode(NormalizeName(name.Text), arguments);
        }
    }

    private ArrayNode ParseArray()
    {
        List<IReadOnlyList<Value>> rows = [];
        List<Value> row = [];

        while (true)
        {
            row.Add(ParseArrayConstant());

            if (Current.Kind == TokenKind.Comma)
            {
                Next();

                continue;
            }

            if (Current.Kind == TokenKind.Semicolon)
            {
                Next();
                rows.Add(row);
                row = [];

                continue;
            }

            Expect(TokenKind.CloseBrace, "}");
            rows.Add(row);

            if (rows.Any(candidate => candidate.Count != rows[0].Count) == true)
            {
                throw new FormatException($"'{formula}' has an array constant whose rows differ in length.");
            }

            return new ArrayNode(rows);
        }
    }

    private Value ParseArrayConstant()
    {
        bool negative = false;

        if (IsOperator("-") == true)
        {
            Next();
            negative = true;
        }

        Token token = Next();

        switch (token.Kind)
        {
            case TokenKind.Number when negative == true:
                return Value.FromNumber(-token.Number);

            case TokenKind.Number:
                return Value.FromNumber(token.Number);

            case TokenKind.Text when negative == false:
                return Value.FromText(token.Text);

            case TokenKind.Bool when negative == false:
                return Value.FromBool(token.Text == "TRUE");

            case TokenKind.Error when negative == false:
                return Value.Error(token.Text.ToUpperInvariant());

            default:
                throw new FormatException($"'{formula}' has '{token.Text}' in an array constant, which only holds constants.");
        }
    }

    private bool IsOperator(string op)
    {
        return Current.Kind == TokenKind.Operator && Current.Text == op;
    }

    private Token Next()
    {
        Token token = tokens[position];

        if (position < tokens.Count - 1)
        {
            position++;
        }

        return token;
    }

    private void Expect(TokenKind kind, string text)
    {
        if (Current.Kind != kind)
        {
            throw new FormatException($"'{formula}' is missing '{text}' at position {Current.Start.ToString(CultureInfo.InvariantCulture)}.");
        }

        Next();
    }

    /// <summary>Upper case, without the prefixes Excel writes in front of functions newer than 2007.</summary>
    public static string NormalizeName(string name)
    {
        string upper = name.ToUpperInvariant();

        foreach (string prefix in new[] { "_XLFN._XLWS.", "_XLFN.", "_XLWS.", "_XLUDF." })
        {
            if (upper.StartsWith(prefix, StringComparison.Ordinal) == true)
            {
                return upper[prefix.Length..];
            }
        }

        return upper;
    }
}
