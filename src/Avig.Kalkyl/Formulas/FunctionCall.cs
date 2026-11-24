namespace Avig.Kalkyl.Formulas;

/// <summary>
/// One call of a worksheet function, handed to its body. Arguments are evaluated when they are
/// first asked for, so IF, IFERROR, CHOOSE and the like only evaluate what they use.
/// <para>
/// The usual pattern: <c>if (call.TryNumber(0, out double x, out Value error) == false) { return error; }</c>.
/// Every <c>Try…</c> method gives the error to return when it fails, following Excel's rules
/// for that conversion.
/// </para>
/// </summary>
internal sealed class FunctionCall
{
    private readonly IReadOnlyList<Node> arguments;
    private readonly Value?[] evaluated;

    public FunctionCall(Evaluator evaluator, EvaluationContext context, IReadOnlyList<Node> arguments)
    {
        Evaluator = evaluator;
        Context = context;
        this.arguments = arguments;
        evaluated = new Value?[arguments.Count];
    }

    public Evaluator Evaluator { get; }

    /// <summary>The cell the formula belongs to.</summary>
    public EvaluationContext Context { get; }

    /// <summary>The number of arguments written, including left-out ones such as the middle one in IF(A1,,2).</summary>
    public int Count => arguments.Count;

    public bool Uses1904 => Evaluator.Uses1904;

    public TextConversion Conversion => Evaluator.Conversion;

    /// <summary>True when the argument was not written at all, or left empty.</summary>
    public bool IsMissing(int index)
    {
        return index >= arguments.Count || arguments[index] is MissingNode;
    }

    /// <summary>The argument as evaluated: a reference stays a reference, an array stays an array. A missing argument is blank.</summary>
    public Value Raw(int index)
    {
        if (IsMissing(index) == true)
        {
            return Value.Blank;
        }

        Value? cached = evaluated[index];

        if (cached is null)
        {
            cached = Evaluator.Evaluate(arguments[index], Context);
            evaluated[index] = cached;
        }

        return cached.Value;
    }

    /// <summary>The argument as a single value (see <see cref="Evaluator.Dereference"/>).</summary>
    public Value Scalar(int index)
    {
        return Evaluator.Dereference(Raw(index), Context);
    }

    /// <summary>The argument as a number. Text that looks like a number, a date or a time is converted; other text gives #VALUE!.</summary>
    public bool TryNumber(int index, out double number, out Value error)
    {
        Value converted = Coerce.ToNumber(Scalar(index), Conversion);

        return Unpack(converted, out number, out error);
    }

    /// <summary>An optional number: <paramref name="fallback"/> when the argument is missing.</summary>
    public bool TryNumber(int index, double fallback, out double number, out Value error)
    {
        if (IsMissing(index) == true)
        {
            number = fallback;
            error = Value.Blank;

            return true;
        }

        return TryNumber(index, out number, out error);
    }

    /// <summary>The argument as a whole number, truncated toward zero as Excel does.</summary>
    public bool TryInteger(int index, out long integer, out Value error)
    {
        integer = 0;

        if (TryNumber(index, out double number, out error) == false)
        {
            return false;
        }

        if (Math.Abs(number) > 1e15)
        {
            error = Value.Num;

            return false;
        }

        integer = (long)Math.Truncate(number);

        return true;
    }

    /// <summary>An optional whole number.</summary>
    public bool TryInteger(int index, long fallback, out long integer, out Value error)
    {
        if (IsMissing(index) == true)
        {
            integer = fallback;
            error = Value.Blank;

            return true;
        }

        return TryInteger(index, out integer, out error);
    }

    /// <summary>The argument as text: numbers as General format shows them, TRUE/FALSE, blank as "".</summary>
    public bool TryText(int index, out string text, out Value error)
    {
        Value converted = Coerce.ToText(Scalar(index));
        text = string.Empty;
        error = Value.Blank;

        if (converted.IsError == true)
        {
            error = converted;

            return false;
        }

        text = converted.Text!;

        return true;
    }

    /// <summary>An optional text argument.</summary>
    public bool TryText(int index, string fallback, out string text, out Value error)
    {
        if (IsMissing(index) == true)
        {
            text = fallback;
            error = Value.Blank;

            return true;
        }

        return TryText(index, out text, out error);
    }

    /// <summary>The argument as TRUE or FALSE. Numbers are TRUE unless 0; text must be "TRUE" or "FALSE".</summary>
    public bool TryBool(int index, out bool flag, out Value error)
    {
        Value converted = Coerce.ToBool(Scalar(index));
        flag = false;
        error = Value.Blank;

        if (converted.IsError == true)
        {
            error = converted;

            return false;
        }

        flag = converted.Bool;

        return true;
    }

    /// <summary>An optional boolean argument.</summary>
    public bool TryBool(int index, bool fallback, out bool flag, out Value error)
    {
        if (IsMissing(index) == true)
        {
            flag = fallback;
            error = Value.Blank;

            return true;
        }

        return TryBool(index, out flag, out error);
    }

    /// <summary>The argument as a reference to cells. False (with #VALUE!) when it is not a reference.</summary>
    public bool TryReference(int index, out RangeReference reference, out Value error)
    {
        Value raw = Raw(index);
        reference = null!;
        error = Value.Blank;

        if (raw.IsError == true)
        {
            error = raw;

            return false;
        }

        if (raw.Kind != ValueKind.Reference)
        {
            error = Value.ValueError;

            return false;
        }

        reference = raw.Reference!;

        return true;
    }

    /// <summary>The argument as an array: a reference's cells, an array constant, or a scalar as 1×1.</summary>
    public bool TryArray(int index, out ValueArray array, out Value error)
    {
        Value raw = Raw(index);
        array = null!;
        error = Value.Blank;

        if (raw.IsError == true)
        {
            error = raw;

            return false;
        }

        array = Evaluator.ToValueArray(raw);

        return true;
    }

    /// <summary>
    /// Every value in the argument, row by row: a reference's cells (blank ones included), an
    /// array's items, or the scalar itself. Whole columns stop at the sheet's last used row.
    /// </summary>
    public IEnumerable<Value> Values(int index)
    {
        Value raw = Raw(index);

        if (raw.Kind is ValueKind.Reference or ValueKind.Array)
        {
            return Evaluator.ToValueArray(raw).All();
        }

        return [raw];
    }

    /// <summary>True when the argument is a reference or an array rather than a value written in the formula.</summary>
    public bool IsRangeOrArray(int index)
    {
        Value raw = Raw(index);

        return raw.Kind is ValueKind.Reference or ValueKind.Array;
    }

    /// <summary>
    /// The argument's values without the blanks, in the order <see cref="Values"/> gives them.
    /// A large, mostly empty range is read from the cells the sheet holds rather than position
    /// by position, so SUM(A:XFD) does not visit 17 billion cells for one value in XFD1048576.
    /// </summary>
    public IEnumerable<Value> NonBlankValues(int index)
    {
        Value raw = Raw(index);

        if (raw.Kind == ValueKind.Reference)
        {
            RangeReference range = raw.Reference!;
            int lastRow = Math.Min(range.LastRow, range.UsedLastRow);
            int lastColumn = Math.Min(range.LastColumn, range.UsedLastColumn);
            long area = (long)Math.Max(0, lastRow - range.FirstRow + 1) * Math.Max(0, lastColumn - range.FirstColumn + 1);

            if (Evaluator.PreferStored(range.Sheet, area) == true)
            {
                return SparseValues(range, lastRow, lastColumn);
            }
        }

        return Values(index).Where(value => value.IsBlank == false);
    }

    private IEnumerable<Value> SparseValues(RangeReference range, int lastRow, int lastColumn)
    {
        foreach ((_, _, Value value) in Evaluator.StoredValues(range.Sheet, range.FirstRow, lastRow, range.FirstColumn, lastColumn))
        {
            if (value.IsBlank == false)
            {
                yield return value;
            }
        }
    }

    /// <summary>
    /// The numbers in the arguments from <paramref name="first"/> on, with Excel's rule for SUM,
    /// AVERAGE, MAX and the like: in references and arrays only numbers count (text, booleans and
    /// blanks are skipped); a value written directly is converted, so =SUM("2",TRUE) is 3 but
    /// =SUM("x") is #VALUE!. An error anywhere stops the collection.
    /// <para>
    /// With <paramref name="includeTextAndBooleans"/> (the …A functions such as AVERAGEA), text
    /// and booleans in references count as 0 and 1/0 instead of being skipped.
    /// </para>
    /// </summary>
    public bool TryCollectNumbers(int first, bool includeTextAndBooleans, out List<double> numbers, out Value error)
    {
        numbers = [];
        error = Value.Blank;

        for (int index = first; index < Count; index++)
        {
            if (IsMissing(index) == true)
            {
                continue;
            }

            if (IsRangeOrArray(index) == true)
            {
                foreach (Value item in NonBlankValues(index))
                {
                    switch (item.Kind)
                    {
                        case ValueKind.Number:
                            numbers.Add(item.Number);

                            break;

                        case ValueKind.Error:
                            error = item;

                            return false;

                        case ValueKind.Bool when includeTextAndBooleans == true:
                            numbers.Add(item.Number);

                            break;

                        case ValueKind.Text when includeTextAndBooleans == true:
                            numbers.Add(0);

                            break;
                    }
                }

                continue;
            }

            Value scalar = Scalar(index);

            if (scalar.IsBlank == true)
            {
                continue;
            }

            Value converted = Coerce.ToNumber(scalar, Conversion);

            if (converted.IsError == true)
            {
                error = converted;

                return false;
            }

            numbers.Add(converted.Number);
        }

        return true;
    }

    private static bool Unpack(Value converted, out double number, out Value error)
    {
        number = 0;
        error = Value.Blank;

        if (converted.IsError == true)
        {
            error = converted;

            return false;
        }

        number = converted.Number;

        return true;
    }
}
