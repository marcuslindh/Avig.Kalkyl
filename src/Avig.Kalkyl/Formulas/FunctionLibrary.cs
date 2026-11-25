using System.Diagnostics.CodeAnalysis;
using Avig.Kalkyl.Formulas.Functions;

namespace Avig.Kalkyl.Formulas;

/// <summary>The body of a worksheet function.</summary>
internal delegate Value FunctionBody(FunctionCall call);

/// <summary>A worksheet function: its name, how many arguments it takes, and whether it is volatile (recalculated every time).</summary>
internal sealed record FunctionDefinition(string Name, int MinArguments, int MaxArguments, FunctionBody Body, bool Volatile);

/// <summary>Collects function definitions while the library is built. A name can only be added once.</summary>
internal sealed class FunctionTable
{
    /// <summary>The most arguments Excel allows in a call.</summary>
    public const int Many = 255;

    private readonly Dictionary<string, FunctionDefinition> functions = new(StringComparer.Ordinal);

    public IReadOnlyDictionary<string, FunctionDefinition> Functions => functions;

    /// <summary>Adds a function. The name is upper case, without the _xlfn. prefix, for example "STDEV.S".</summary>
    public void Add(string name, int minArguments, int maxArguments, FunctionBody body, bool isVolatile = false)
    {
        if (functions.ContainsKey(name) == true)
        {
            throw new InvalidOperationException($"The function {name} is added twice.");
        }

        functions[name] = new FunctionDefinition(name, minArguments, maxArguments, body, isVolatile);
    }
}

/// <summary>
/// Every worksheet function the engine knows. Each category lives in its own class under
/// Formulas/Functions and adds its functions in a Register method.
/// </summary>
internal static class FunctionLibrary
{
    private static readonly IReadOnlyDictionary<string, FunctionDefinition> Functions = Build();

    public static IEnumerable<string> Names => Functions.Keys;

    public static bool TryGet(string name, [NotNullWhen(true)] out FunctionDefinition? definition)
    {
        return Functions.TryGetValue(name, out definition);
    }

    private static IReadOnlyDictionary<string, FunctionDefinition> Build()
    {
        FunctionTable table = new();
        LogicalFunctions.Register(table);
        MathFunctions.Register(table);
        StatisticalFunctions.Register(table);
        TextFunctions.Register(table);
        DateFunctions.Register(table);
        LookupFunctions.Register(table);
        InformationFunctions.Register(table);
        FinancialFunctions.Register(table);
        EngineeringFunctions.Register(table);
        DatabaseFunctions.Register(table);
        PivotFunctions.Register(table);

        return table.Functions;
    }
}
