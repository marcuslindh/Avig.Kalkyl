using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>The kind of a formula value.</summary>
internal enum ValueKind
{
    Blank,
    Number,
    Text,
    Bool,
    Error,
    Reference,
    Array,
}

/// <summary>
/// A value while a formula is evaluated: a scalar, an error, a reference to cells, or an array.
/// Dates are numbers here, as they are in Excel; they become dates again when a result is stored
/// in a cell with a date format.
/// </summary>
internal readonly struct Value
{
    public static Value Blank => default;
    public static readonly Value True = FromBool(true);
    public static readonly Value False = FromBool(false);
    public static readonly Value Div0 = Error("#DIV/0!");
    public static readonly Value ValueError = Error("#VALUE!");
    public static readonly Value Ref = Error("#REF!");
    public static readonly Value Name = Error("#NAME?");
    public static readonly Value Num = Error("#NUM!");
    public static readonly Value NA = Error("#N/A");
    public static readonly Value Null = Error("#NULL!");
    public static readonly Value Calc = Error("#CALC!");

    private Value(ValueKind kind, double number, string? text, RangeReference? reference, ValueArray? array)
    {
        Kind = kind;
        Number = number;
        Text = text;
        Reference = reference;
        Array = array;
    }

    public ValueKind Kind { get; }

    /// <summary>The number, or 1/0 for a boolean.</summary>
    public double Number { get; }

    /// <summary>The text, or the error code for an error.</summary>
    public string? Text { get; }

    public RangeReference? Reference { get; }

    public ValueArray? Array { get; }

    public bool IsError => Kind == ValueKind.Error;

    public bool IsBlank => Kind == ValueKind.Blank;

    public bool Bool => Number != 0;

    public static Value FromNumber(double number)
    {
        if (double.IsFinite(number) == false)
        {
            return Num;
        }

        // Excel has no negative zero.
        if (number == 0)
        {
            number = 0;
        }

        return new Value(ValueKind.Number, number, null, null, null);
    }

    public static Value FromText(string text)
    {
        return new Value(ValueKind.Text, 0, text, null, null);
    }

    public static Value FromBool(bool value)
    {
        double number = 0;

        if (value == true)
        {
            number = 1;
        }

        return new Value(ValueKind.Bool, number, null, null, null);
    }

    public static Value Error(string code)
    {
        return new Value(ValueKind.Error, 0, code, null, null);
    }

    public static Value FromReference(RangeReference reference)
    {
        return new Value(ValueKind.Reference, 0, null, reference, null);
    }

    public static Value FromArray(ValueArray array)
    {
        return new Value(ValueKind.Array, 0, null, null, array);
    }

    /// <summary>The number as Excel's General format shows it in text: up to 15 significant digits.</summary>
    public static string NumberText(double number)
    {
        return number.ToString("G15", CultureInfo.InvariantCulture);
    }

    public override string ToString()
    {
        return Kind switch
        {
            ValueKind.Number => NumberText(Number),
            ValueKind.Text => Text!,
            ValueKind.Bool when Bool == true => "TRUE",
            ValueKind.Bool => "FALSE",
            ValueKind.Error => Text!,
            ValueKind.Reference => Reference!.ToString(),
            ValueKind.Array => "{array " + Array!.Rows.ToString(CultureInfo.InvariantCulture) + "x" + Array.Columns.ToString(CultureInfo.InvariantCulture) + "}",
            _ => string.Empty,
        };
    }
}

/// <summary>A rectangle of cells on a sheet. Whole columns and rows are rectangles too.</summary>
internal sealed record RangeReference(Worksheet Sheet, int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    public int Rows => LastRow - FirstRow + 1;

    public int Columns => LastColumn - FirstColumn + 1;

    public bool IsSingleCell => FirstRow == LastRow && FirstColumn == LastColumn;

    public bool Contains(Worksheet sheet, int row, int column)
    {
        return ReferenceEquals(sheet, Sheet) == true && row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
    }

    /// <summary>The last row worth visiting: the range's last row or the sheet's last used row, whichever is smaller.</summary>
    public int UsedLastRow => Math.Min(LastRow, Sheet.RowCount - 1);

    /// <summary>The last column worth visiting.</summary>
    public int UsedLastColumn => Math.Min(LastColumn, Math.Max(Sheet.ColumnCount - 1, FirstColumn - 1));

    public override string ToString()
    {
        string start = new CellAddress(FirstRow, FirstColumn).ToString();

        if (IsSingleCell == true)
        {
            return Sheet.Name + "!" + start;
        }

        return Sheet.Name + "!" + start + ":" + new CellAddress(LastRow, LastColumn).ToString();
    }
}

/// <summary>A rectangular array of values, from an array constant, a range or a function.</summary>
internal sealed class ValueArray
{
    private readonly Value[] items;

    public ValueArray(int rows, int columns)
    {
        Rows = rows;
        Columns = columns;
        items = new Value[rows * columns];
    }

    public int Rows { get; }

    public int Columns { get; }

    public Value this[int row, int column]
    {
        get => items[(row * Columns) + column];
        set => items[(row * Columns) + column] = value;
    }

    public IEnumerable<Value> All()
    {
        return items;
    }
}
