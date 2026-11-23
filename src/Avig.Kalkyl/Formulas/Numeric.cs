using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>Number helpers that follow Excel rather than IEEE 754 where the two differ.</summary>
internal static class Numeric
{
    /// <summary>
    /// Rounds half away from zero, as ROUND does. The number is first taken at the 15 significant
    /// digits Excel shows, so ROUND(2.675,2) is 2.68 even though the double is 2.67499999...
    /// </summary>
    public static double Round(double number, int digits)
    {
        if (digits > 15)
        {
            digits = 15;
        }

        if (digits < -15)
        {
            return 0;
        }

        double shown = double.Parse(number.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        double scale = Math.Pow(10, Math.Abs(digits));

        if (digits >= 0)
        {
            if (Math.Abs(shown) < 7.9e13 && digits <= 15)
            {
                decimal exact = (decimal)shown;

                return (double)Math.Round(exact, Math.Min(digits, 28), MidpointRounding.AwayFromZero);
            }

            return Math.Round(shown * scale, MidpointRounding.AwayFromZero) / scale;
        }

        return Math.Round(shown / scale, MidpointRounding.AwayFromZero) * scale;
    }

    /// <summary>Rounds toward zero at a number of digits, as ROUNDDOWN and TRUNC do.</summary>
    public static double RoundDown(double number, int digits)
    {
        double shown = double.Parse(number.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        double scale = Math.Pow(10, digits);

        return Math.Truncate(Math.Round(shown * scale, 9)) / scale;
    }

    /// <summary>Rounds away from zero at a number of digits, as ROUNDUP does.</summary>
    public static double RoundUp(double number, int digits)
    {
        double shown = double.Parse(number.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        double scale = Math.Pow(10, digits);
        double scaled = Math.Round(shown * scale, 9);

        if (scaled < 0)
        {
            return Math.Floor(scaled) / scale;
        }

        return Math.Ceiling(scaled) / scale;
    }
}
