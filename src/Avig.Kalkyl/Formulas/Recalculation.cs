namespace Avig.Kalkyl;

/// <summary>Which formulas <see cref="Workbook.RecalculateAsync"/> calculates.</summary>
public enum RecalculationScope
{
    /// <summary>
    /// Formulas affected by cells set since the last recalculation (or since the file was opened),
    /// everything that depends on them, and volatile functions such as NOW, RAND and INDIRECT.
    /// </summary>
    Changed,

    /// <summary>Every formula in the workbook.</summary>
    All,
}

/// <summary>What a recalculation did.</summary>
/// <param name="Calculated">The number of formula cells calculated.</param>
/// <param name="CircularReferences">Cells in a circular reference, such as <c>Budget!C5</c>. They get 0, as in Excel without iterative calculation.</param>
/// <param name="UnknownFunctions">Functions used that the engine does not know. Their cells show #NAME?.</param>
/// <param name="InvalidFormulas">Formulas that could not be parsed, with the reason. Their cells show #NAME?.</param>
public sealed record RecalculationReport(
    int Calculated,
    IReadOnlyList<string> CircularReferences,
    IReadOnlyList<string> UnknownFunctions,
    IReadOnlyList<string> InvalidFormulas);
