using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Compatibility;

/// <summary>What a cell holds, in a form every library's answer can be brought to. Empty cells are not stored.</summary>
internal enum ValueKind
{
    Number,
    Text,
    Bool,
    Date,
    Error,
}

/// <summary>A cell value from any library, normalised so that the answers can be compared.</summary>
internal readonly record struct Value(ValueKind Kind, double Number, string Text, DateTime Date)
{
    public static Value Of(double number)
    {
        return new Value(ValueKind.Number, number, string.Empty, default);
    }

    public static Value Of(string text)
    {
        return new Value(ValueKind.Text, 0, text, default);
    }

    public static Value Of(bool flag)
    {
        double stored = 0;

        if (flag == true)
        {
            stored = 1;
        }

        return new Value(ValueKind.Bool, stored, string.Empty, default);
    }

    public static Value Of(DateTime date)
    {
        return new Value(ValueKind.Date, 0, string.Empty, date);
    }

    public static Value Error(string code)
    {
        return new Value(ValueKind.Error, 0, code, default);
    }

    /// <summary>
    /// Equal within what the libraries can be expected to agree on: numbers to 12 significant
    /// digits (some libraries parse with float shortcuts), dates to the millisecond, and an
    /// error equal to the same code read as text (MiniExcel does not know error cells).
    /// </summary>
    public bool Matches(Value other)
    {
        if (Kind == ValueKind.Error && other.Kind == ValueKind.Text)
        {
            return Text == other.Text;
        }

        if (Kind == ValueKind.Text && other.Kind == ValueKind.Error)
        {
            return Text == other.Text;
        }

        if (Kind != other.Kind)
        {
            return false;
        }

        return Kind switch
        {
            ValueKind.Number => NumbersMatch(Number, other.Number),
            ValueKind.Bool => Number == other.Number,
            ValueKind.Date => Math.Abs((Date - other.Date).TotalMilliseconds) < 1.5,
            _ => string.Equals(Text, other.Text, StringComparison.Ordinal),
        };
    }

    public static bool NumbersMatch(double left, double right)
    {
        if (left == right)
        {
            return true;
        }

        double scale = Math.Max(Math.Abs(left), Math.Abs(right));

        return Math.Abs(left - right) <= scale * 1e-12;
    }

    public override string ToString()
    {
        return Kind switch
        {
            ValueKind.Number => Number.ToString("R", CultureInfo.InvariantCulture),
            ValueKind.Bool => (Number == 1).ToString(CultureInfo.InvariantCulture).ToUpperInvariant(),
            ValueKind.Date => Date.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
            ValueKind.Error => Text,
            _ => "\"" + Shorten(Text) + "\"",
        };
    }

    private static string Shorten(string text)
    {
        StringBuilder escaped = new();

        foreach (char current in text)
        {
            if (current < ' ')
            {
                escaped.Append(CultureInfo.InvariantCulture, $"\\x{(int)current:X2}");
            }
            else
            {
                escaped.Append(current);
            }
        }

        string line = escaped.ToString();

        // Both ends: a difference is as likely to be a lost trailing character as a leading one.
        if (line.Length <= 70)
        {
            return line;
        }

        return line[..35] + "..." + line[^30..];
    }
}

/// <summary>Every non-empty cell of a workbook as one library read it: sheet name, then (row, column), both 0-based.</summary>
internal sealed class WorkbookValues
{
    public WorkbookValues(string library)
    {
        Library = library;
    }

    public string Library { get; }

    public Dictionary<string, Dictionary<(int Row, int Column), Value>> Sheets { get; } = new(StringComparer.Ordinal);

    /// <summary>Set when the library could not read the file; it then takes no part in the comparison.</summary>
    public string Failure { get; set; } = string.Empty;

    public void Add(string sheet, int row, int column, Value value)
    {
        // Every library stops at the same place, so that a huge sheet does not make the test take minutes.
        if (row >= Readers.MaxRows)
        {
            return;
        }

        // An empty string and an empty cell are the same to a reader of the sheet, and the
        // libraries disagree on which one a cell with <v></v> is.
        if (value.Kind == ValueKind.Text && value.Text.Length == 0)
        {
            return;
        }

        if (Sheets.TryGetValue(sheet, out Dictionary<(int Row, int Column), Value>? cells) == false)
        {
            cells = [];
            Sheets[sheet] = cells;
        }

        cells[(row, column)] = value;
    }

    public void AddSheet(string sheet)
    {
        if (Sheets.ContainsKey(sheet) == false)
        {
            Sheets[sheet] = [];
        }
    }

    public bool TryGet(string sheet, int row, int column, out Value value)
    {
        value = default;

        return Sheets.TryGetValue(sheet, out Dictionary<(int Row, int Column), Value>? cells) == true && cells.TryGetValue((row, column), out value) == true;
    }
}
