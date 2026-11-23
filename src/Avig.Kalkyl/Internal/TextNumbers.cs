using System.Globalization;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Numbers stored as text in a cell, which is common in files from other systems. Both a dot
/// and a comma are accepted as decimal separators, since Swedish files often have "1,5". A single
/// comma without a dot is a decimal separator; otherwise invariant parsing with thousands separators applies.
/// </summary>
internal static class TextNumbers
{
    public static bool TryParse(string text, out double value)
    {
        string candidate = text.Replace(" ", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);

        if (candidate.Contains(',', StringComparison.Ordinal) == true
            && candidate.Contains('.', StringComparison.Ordinal) == false
            && candidate.IndexOf(',', StringComparison.Ordinal) == candidate.LastIndexOf(',') )
        {
            candidate = candidate.Replace(',', '.');
        }

        return double.TryParse(candidate, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value)
            && double.IsFinite(value);
    }
}
