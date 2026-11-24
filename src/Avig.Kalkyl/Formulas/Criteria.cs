using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// The criteria of COUNTIF, SUMIF, AVERAGEIF and their …IFS forms: "&gt;5", "&lt;&gt;x", "=",
/// "a*", 10, TRUE. A criterion text can start with =, &lt;&gt;, &lt;, &gt;, &lt;= or &gt;=; the
/// rest is compared as a number when it looks like one (dates too), otherwise as text without
/// regard to case, where * matches any run of characters, ? one character, and ~ escapes them.
/// "=" alone matches blank cells, "&lt;&gt;" alone matches cells that are not blank.
/// </summary>
internal sealed class Criteria
{
    private readonly string op;
    private readonly Value operand;
    private readonly Regex? pattern;
    private readonly TextConversion conversion;

    private Criteria(string op, Value operand, Regex? pattern, TextConversion conversion)
    {
        this.op = op;
        this.operand = operand;
        this.pattern = pattern;
        this.conversion = conversion;
    }

    public static Criteria Parse(Value criterion, TextConversion conversion)
    {
        switch (criterion.Kind)
        {
            case ValueKind.Number:
            case ValueKind.Bool:
            case ValueKind.Error:
                return new Criteria("=", criterion, null, conversion);

            case ValueKind.Blank:
                // A criterion that points to an empty cell is read as 0, as in Excel.
                return new Criteria("=", Value.FromNumber(0), null, conversion);
        }

        string text = criterion.Text ?? string.Empty;
        string prefix = "=";

        foreach (string candidate in new[] { "<=", ">=", "<>", "=", "<", ">" })
        {
            if (text.StartsWith(candidate, StringComparison.Ordinal) == true)
            {
                prefix = candidate;
                text = text[candidate.Length..];

                break;
            }
        }

        if (Coerce.TryParseNumber(text, conversion, out double number) == true)
        {
            return new Criteria(prefix, Value.FromNumber(number), null, conversion);
        }

        if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase) == true || string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new Criteria(prefix, Value.FromBool(string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase)), null, conversion);
        }

        if (IsKnownError(text.ToUpperInvariant()) == true)
        {
            return new Criteria(prefix, Value.Error(text.ToUpperInvariant()), null, conversion);
        }

        Regex? regex = null;

        if ((prefix == "=" || prefix == "<>") && HasWildcard(text) == true)
        {
            regex = new Regex("^" + WildcardToRegex(text) + "$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, TimeSpan.FromSeconds(1));
        }

        return new Criteria(prefix, Value.FromText(text), regex, conversion);
    }

    /// <summary>Whether a cell's value meets the criterion.</summary>
    public bool Matches(Value value)
    {
        if (operand.Kind == ValueKind.Text && operand.Text!.Length == 0)
        {
            bool blank = value.IsBlank == true || (value.Kind == ValueKind.Text && value.Text!.Length == 0);

            if (op == "=")
            {
                return blank;
            }

            if (op == "<>")
            {
                return blank == false;
            }
        }

        if (pattern is not null)
        {
            bool found = value.Kind == ValueKind.Text && pattern.IsMatch(value.Text!);

            return (op == "=") == found;
        }

        if (op == "<>")
        {
            return Same(value) == false;
        }

        if (op == "=")
        {
            return Same(value);
        }

        // Ordering comparisons only apply between values of the same kind: ">5" never matches text.
        if (value.Kind != operand.Kind)
        {
            return false;
        }

        int comparison = Coerce.Compare(value, operand);

        return op switch
        {
            "<" => comparison < 0,
            ">" => comparison > 0,
            "<=" => comparison <= 0,
            _ => comparison >= 0,
        };
    }

    private bool Same(Value value)
    {
        switch (operand.Kind)
        {
            case ValueKind.Number:
                if (value.Kind == ValueKind.Number)
                {
                    return value.Number == operand.Number;
                }

                // A number criterion also matches text that reads as the same number.
                return value.Kind == ValueKind.Text && Coerce.TryParseNumber(value.Text!, conversion, out double parsed) == true && parsed == operand.Number;

            case ValueKind.Bool:
                return value.Kind == ValueKind.Bool && value.Bool == operand.Bool;

            case ValueKind.Error:
                return value.Kind == ValueKind.Error && string.Equals(value.Text, operand.Text, StringComparison.Ordinal);

            default:
                return value.Kind == ValueKind.Text && string.Equals(value.Text, operand.Text, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>Whether text has an unescaped * or ?. Shared with MATCH, VLOOKUP, SEARCH and friends.</summary>
    public static bool HasWildcard(string text)
    {
        for (int index = 0; index < text.Length; index++)
        {
            if (text[index] == '~')
            {
                index++;

                continue;
            }

            if (text[index] is '*' or '?')
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>Excel's wildcards as a regular expression, without anchors.</summary>
    public static string WildcardToRegex(string text)
    {
        StringBuilder builder = new();

        for (int index = 0; index < text.Length; index++)
        {
            char current = text[index];

            if (current == '~' && index + 1 < text.Length)
            {
                index++;
                builder.Append(Regex.Escape(text[index].ToString()));

                continue;
            }

            if (current == '*')
            {
                builder.Append(".*");
            }
            else if (current == '?')
            {
                builder.Append('.');
            }
            else
            {
                builder.Append(Regex.Escape(current.ToString()));
            }
        }

        return builder.ToString();
    }

    private static bool IsKnownError(string code)
    {
        return code is "#NULL!" or "#DIV/0!" or "#VALUE!" or "#REF!" or "#NAME?" or "#NUM!" or "#N/A" or "#SPILL!" or "#CALC!" or "#GETTING_DATA";
    }
}
