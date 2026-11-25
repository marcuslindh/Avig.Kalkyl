using System.Text;
using Avig.Kalkyl.Formulas;

namespace Avig.Kalkyl.Fuzz;

/// <summary>
/// The formula engine. <c>formula</c> is a crash target, but a strict one: the engine's contract
/// is that any formula text gives a value, an Excel error or an Invalid result, so <em>every</em>
/// exception is a finding. <c>formula-shift</c> is an invariant: moving a formula there and back
/// gives the formula it started as.
/// </summary>
internal static class FormulaTargets
{
    /// <summary>Excel's limit on the length of a formula.</summary>
    private const int MaxFormulaLength = 8192;

    private static readonly string[] FunctionNames = [.. FunctionLibrary.Names.Order(StringComparer.Ordinal)];

    private static readonly string[] References =
    [
        "A1", "B2", "$C$3", "C$4", "$D5", "A1:D8", "B2:B8", "$A$1:$D$8", "A:A", "B:D", "1:1", "2:8", "A1:A1",
        "D8:A1", "XFD1048576", "A1048576", "XFD1", "Data!A1", "Data!B2:C8", "'Data'!A1:A3", "Other!A1", "'Other'!B:B",
        "Missing!A1", "Rate", "Bogus", "A1:XFD1048576", "Data!A:XFD", "_xlfn.X", "#REF!", "Data!#REF!",
    ];

    private static readonly string[] Literals =
    [
        "0", "1", "-1", "2", "0.5", "1E+300", "1E-300", "9.99999999999999E+307", "4.9E-324", "2147483648", "-2147483649",
        "45000", "2958465", "2958466", "60", "1E15", "123456789012345678", ".5", "1%", "TRUE", "FALSE",
        "\"\"", "\"a\"", "\"10\"", "\" 1,5 \"", "\"2026-01-02\"", "\"x\"\"y\"", "\"Öst\"", "\"#N/A\"", "\"*\"", "\"?\"", "\"~*\"",
        "\"<>\"", "\">=2\"", "\"=a\"", "\"0x1F\"", "\"1e999\"", "\"TRUE\"",
        "#N/A", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "#GETTING_DATA",
        "{1,2,3}", "{1;2;3}", "{1,\"a\";TRUE,#N/A}", "{}", "{-1,0.5}",
    ];

    private static readonly string[] BinaryOperators = ["+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">=", ":", " ", ","];

    /// <summary>Tokens for breaking a generated formula: the parser's error paths need visiting too.</summary>
    private static readonly string[] BreakingTokens = ["(", ")", ",", ";", "\"", "'", "!", "{", "}", "$", "#", "%", "=", " ", "@", "[", "]", "\0", "\uD800"];

    public static FuzzTarget Evaluate()
    {
        return new FuzzTarget("formula", "evaluate any formula text; every exception is a finding", EvaluateAsync, GenerateFormula);
    }

    public static FuzzTarget Shift()
    {
        return new FuzzTarget("formula-shift", "a formula moved there and back is unchanged", ShiftAsync, GenerateShift);
    }

    /// <summary>
    /// <c>formula-sweep</c> — the same harness as <c>formula</c>, with every function called on
    /// every pair of edge values in turn instead of at random.
    ///
    /// Random formulas spread their luck over 333 functions and some 85 operands, so a bug that
    /// needs one function and two particular arguments is a needle: REPT("",2147483648) threw
    /// once in the first run and then not again in twenty thousand inputs. A sweep visits every
    /// function with every pair, which is about 110 000 inputs (run it with --iter 110000).
    /// </summary>
    public static FuzzTarget Sweep()
    {
        return new FuzzTarget("formula-sweep", "every function with every pair of edge arguments, in order", EvaluateAsync, GenerateSweep);
    }

    /// <summary>The edge values of the sweep, the most likely to break something first.</summary>
    private static readonly string[] EdgeArguments =
    [
        "\"\"", "2147483648", "-1", "0", "1E+15", "-2147483649", "1E+300", "0.5", "\"abc\"", "TRUE", "#N/A", "A1:D8",
        "{1,2;3,4}", "Data!A:A", "1", "-1E+300", "\" \"", "999999999999999",
    ];

    private static int sweepPosition;

    private static byte[] GenerateSweep(Random random)
    {
        int position = sweepPosition;
        sweepPosition++;
        string name = FunctionNames[position % FunctionNames.Length];
        int pair = (position / FunctionNames.Length) % (EdgeArguments.Length * EdgeArguments.Length);
        FunctionLibrary.TryGet(name, out FunctionDefinition? definition);
        int count = Math.Max(definition!.MinArguments, Math.Min(2, definition.MaxArguments));
        string[] arguments = new string[count];

        for (int index = 0; index < count; index++)
        {
            arguments[index] = "1";
        }

        if (count > 0)
        {
            arguments[0] = EdgeArguments[pair % EdgeArguments.Length];
        }

        if (count > 1)
        {
            arguments[1] = EdgeArguments[pair / EdgeArguments.Length];
        }

        return Encoding.UTF8.GetBytes($"{name}({string.Join(',', arguments)})");
    }

    private static byte[] GenerateFormula(Random random)
    {
        return Encoding.UTF8.GetBytes(Formula(random));
    }

    private static byte[] GenerateShift(Random random)
    {
        byte[] formula = Encoding.UTF8.GetBytes(Formula(random));
        byte[] input = new byte[formula.Length + 2];
        input[0] = (byte)(sbyte)random.Next(-12, 13);
        input[1] = (byte)(sbyte)random.Next(-12, 13);

        if (random.Next(8) == 0)
        {
            input[0] = (byte)random.Next(256);
            input[1] = (byte)random.Next(256);
        }

        formula.CopyTo(input, 2);

        return input;
    }

    /// <summary>A formula, mostly well formed; one in six has a token thrown in or a piece cut out.</summary>
    private static string Formula(Random random)
    {
        StringBuilder builder = new();

        if (random.Next(2) == 0)
        {
            builder.Append('=');
        }

        Expression(random, builder, random.Next(1, 6));

        // Depth, for the recursion in the parser and the evaluator. A stack overflow ends the
        // process instead of being caught, so these must stay within what a formula may hold
        // (8 192 characters) and get past the parser's nesting limit (255) and Excel's limit on
        // functions inside functions (64). The first version stopped at 400 parentheses and
        // never built a long chain, and so missed that 820 parentheses or 700 terms of
        // 1+1+…+1 killed the process.
        int shape = random.Next(40);

        if (shape == 0)
        {
            int depth = random.Next(50, 4000);
            builder.Insert(0, new string('(', depth)).Append(')', depth);
        }
        else if (shape == 1)
        {
            string[] operators = ["+", "-", "*", "&", "^", "=", ":"];
            string op = Inputs.Pick(random, operators);
            int terms = random.Next(200, 4000);

            for (int term = 0; term < terms; term++)
            {
                builder.Append(op).Append('1');
            }
        }
        else if (shape == 2)
        {
            int depth = random.Next(20, 1500);
            builder.Insert(0, string.Concat(Enumerable.Repeat("ABS(", depth))).Append(')', depth);
        }
        else if (shape == 3)
        {
            builder.Insert(0, new string('-', random.Next(100, 4000)));
        }

        string formula = builder.ToString();

        if (random.Next(6) == 0)
        {
            formula = Inputs.Mutate(random, formula, BreakingTokens);
        }

        return formula;
    }

    private static void Expression(Random random, StringBuilder builder, int depth)
    {
        int choice = random.Next(10);

        if (depth <= 0 || choice < 3)
        {
            Operand(random, builder);

            return;
        }

        if (choice < 7)
        {
            string name = Inputs.Pick(random, FunctionNames);

            if (random.Next(20) == 0)
            {
                name = "NOSUCHFUNCTION";
            }

            builder.Append(name).Append('(');
            int arguments = random.Next(0, 5);

            for (int argument = 0; argument < arguments; argument++)
            {
                if (argument > 0)
                {
                    builder.Append(',');
                }

                // An empty argument (SUM(1,,2)) is legal and easy to mishandle.
                if (random.Next(12) != 0)
                {
                    Expression(random, builder, depth - 1);
                }
            }

            builder.Append(')');

            return;
        }

        if (choice < 9)
        {
            Expression(random, builder, depth - 1);
            builder.Append(Inputs.Pick(random, BinaryOperators));
            Expression(random, builder, depth - 1);

            return;
        }

        string[] wrappers = ["-", "+", "--", "("];
        string wrapper = Inputs.Pick(random, wrappers);
        builder.Append(wrapper);
        Expression(random, builder, depth - 1);

        if (wrapper == "(")
        {
            builder.Append(')');
        }

        if (random.Next(4) == 0)
        {
            builder.Append('%');
        }
    }

    private static void Operand(Random random, StringBuilder builder)
    {
        if (random.Next(2) == 0)
        {
            builder.Append(Inputs.Pick(random, References));
        }
        else
        {
            builder.Append(Inputs.Pick(random, Literals));
        }
    }

    /// <summary>A workbook with a bit of everything in it for formulas to point at.</summary>
    private static Workbook DataWorkbook(out Worksheet data)
    {
        Workbook book = Workbook.Create();
        book.AddSheet("Data").IsSuccess(out Worksheet? sheet);
        book.AddSheet("Other").IsSuccess(out Worksheet? other);
        data = sheet!;
        CellValue[] values =
        [
            1, 2.5, -3, 0, 1e300, "text", "", " 12 ", "2026-01-02", true, false,
            new DateTime(2026, 1, 2, 12, 0, 0), new DateTime(1900, 2, 28), CellValue.Error("#N/A"),
            CellValue.Error("#DIV/0!"), CellValue.Empty, 45000, -0.0001, "Öst", "a*b",
        ];

        for (int index = 0; index < values.Length; index++)
        {
            data[index % 8, index / 8] = values[index];
        }

        data["D1"] = Cell.FromFormula("A1+B1");
        data["D2"] = Cell.FromFormula("D1*2");
        other!["A1"] = 7;
        other["B1"] = "x";
        book.DefinedNames.Add(new DefinedName("Rate", -1, "Data!$A$2"));
        book.DefinedNames.Add(new DefinedName("Bogus", -1, "SUM("));

        return book;
    }

    private static async Task EvaluateAsync(byte[] input)
    {
        string formula = Encoding.UTF8.GetString(input);

        if (formula.Length > MaxFormulaLength)
        {
            return;
        }

        // A formula of arrays just under KalkylLimits.MaxArrayCells can take seconds, and time is
        // the caller's to limit with the token, as the documentation says. So each step gets two
        // seconds: a cancelled one is not a finding, but one that runs on past the token is, as
        // libFuzzer's timeout (10 s) then catches it.
        Workbook book = DataWorkbook(out Worksheet data);

        if (await WithinTimeAsync(token => Task.FromResult(book.Evaluate(formula, data, token))) == false || string.IsNullOrWhiteSpace(formula) == true)
        {
            return;
        }

        // The same formula in a cell: recalculation, conversion to a cell value and saving.
        data["F1"] = Cell.FromFormula(formula);
        data["F2"] = Cell.FromFormula("F1");

        if (await WithinTimeAsync(token => book.RecalculateAsync(RecalculationScope.Changed, token)) == false)
        {
            return;
        }

        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
    }

    /// <summary>Runs a step with a two-second token; false when the token stopped it.</summary>
    private static async Task<bool> WithinTimeAsync<T>(Func<CancellationToken, Task<T>> step)
    {
        using CancellationTokenSource time = new(TimeSpan.FromSeconds(2));

        try
        {
            await step(time.Token);

            return true;
        }
        catch (OperationCanceledException) when (time.IsCancellationRequested == true)
        {
            return false;
        }
    }

    private static Task ShiftAsync(byte[] input)
    {
        if (input.Length < 2)
        {
            return Task.CompletedTask;
        }

        int rows = (sbyte)input[0];
        int columns = (sbyte)input[1];
        string formula = Encoding.UTF8.GetString(input, 2, input.Length - 2);

        // The shifter takes one leading '=' off, as files store formulas without it. Text that
        // still starts with '=' after that is not a formula a file can hold.
        if (formula.StartsWith('=') == true)
        {
            formula = formula[1..];
        }

        if (formula.StartsWith('=') == true)
        {
            return Task.CompletedTask;
        }

        if (FormulaShifter.Shift(formula, 0, 0).IsSuccess(out string? canonical) == false)
        {
            return Task.CompletedTask;
        }

        bool parses = FormulaParser.Parse(formula).IsSuccess(out _);

        // Moving by nothing only rewrites references in their canonical form, so doing it again
        // changes nothing. Only for formulas that parse: in A:XFD'Other'!B:B, which has no
        // operator between the references, dropping the quotes runs XFD and Other together.
        if (parses == true && (FormulaShifter.Shift(canonical, 0, 0).IsSuccess(out string? again) == false || again != canonical))
        {
            throw new InvariantViolationException($"Moving by (0,0) is not stable: '{formula}' -> '{canonical}'.");
        }

        if (FormulaShifter.Shift(formula, rows, columns).IsSuccess(out string? moved) == false)
        {
            throw new InvariantViolationException($"'{formula}' can be moved by (0,0) but not by ({rows},{columns}).");
        }

        // A moved formula can grow (row 1 becomes row 11) past the 8 192 characters Excel
        // allows, and then it is rightly refused.
        if (parses == true && moved.Length <= 8192 && FormulaParser.Parse(moved).IsSuccess(out _) == false)
        {
            throw new InvariantViolationException($"'{formula}' parses but moved by ({rows},{columns}) it is '{moved}', which does not.");
        }

        // A reference pushed off the sheet becomes #REF! and cannot come back. Otherwise the way
        // back is exact, for formulas that parse: in text that does not (2:8A1) a moved reference
        // can run into its neighbour and lex differently (2:8E5 is 2: and the number 8E5), and
        // the shifter only ever sees formulas from a file.
        if (parses == false || CountRefErrors(moved) != CountRefErrors(formula))
        {
            return Task.CompletedTask;
        }

        if (FormulaShifter.Shift(moved, -rows, -columns).IsSuccess(out string? back) == false || back != canonical)
        {
            throw new InvariantViolationException($"'{formula}' moved by ({rows},{columns}) is '{moved}', and moved back it is '{back}', not '{canonical}'.");
        }

        return Task.CompletedTask;
    }

    private static int CountRefErrors(string formula)
    {
        int count = 0;
        int index = formula.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase);

        while (index >= 0)
        {
            count++;
            index = formula.IndexOf("#REF!", index + 1, StringComparison.OrdinalIgnoreCase);
        }

        return count;
    }
}
