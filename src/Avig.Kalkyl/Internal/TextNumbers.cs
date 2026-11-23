using System.Globalization;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Tal som står som text i en cell, vilket är vanligt i filer från andra system. Både punkt och
/// komma godtas som decimaltecken, eftersom svenska filer ofta har "1,5". Ett ensamt komma
/// utan punkt tolkas som decimaltecken; annars gäller invariant tolkning med tusentalsavgränsare.
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
