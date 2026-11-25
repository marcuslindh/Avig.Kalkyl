using System.Globalization;

namespace Avig.Kalkyl.Internal;

/// <summary>Checks and conversions shared by the drawing types: colours, units and ranges.</summary>
internal static class DrawingValues
{
    /// <summary>EMU (English Metric Units) per point.</summary>
    public const long EmuPerPoint = 12_700;

    /// <summary>The largest coordinate DrawingML accepts (ST_Coordinate).</summary>
    public const long MaxCoordinate = 27_273_042_316_900;

    /// <summary>A colour as #RRGGBB in upper case, or an ArgumentException naming the parameter.</summary>
    public static string Color(string value, string parameter)
    {
        ArgumentNullException.ThrowIfNull(value, parameter);
        string hex = value;

        if (hex.StartsWith('#') == true)
        {
            hex = hex[1..];
        }

        if (hex.Length != 6 || int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _) == false)
        {
            throw new ArgumentException($"'{value}' is not a colour. Expected #RRGGBB.", parameter);
        }

        return "#" + hex.ToUpperInvariant();
    }

    /// <summary>A transparency from 0 (opaque) to 1 (invisible).</summary>
    public static double Transparency(double value, string parameter)
    {
        if (double.IsFinite(value) == false || value < 0 || value > 1)
        {
            throw new ArgumentOutOfRangeException(parameter, value, "A transparency goes from 0 (opaque) to 1 (invisible).");
        }

        return value;
    }

    /// <summary>A finite number in a range, inclusive.</summary>
    public static double Range(double value, double minimum, double maximum, string parameter, string what)
    {
        if (double.IsFinite(value) == false || value < minimum || value > maximum)
        {
            throw new ArgumentOutOfRangeException(parameter, value, $"{what} goes from {Text(minimum)} to {Text(maximum)}.");
        }

        return value;
    }

    /// <summary>An integer in a range, inclusive.</summary>
    public static int Range(int value, int minimum, int maximum, string parameter, string what)
    {
        if (value < minimum || value > maximum)
        {
            throw new ArgumentOutOfRangeException(parameter, value, $"{what} goes from {Text(minimum)} to {Text(maximum)}.");
        }

        return value;
    }

    /// <summary>Points to EMU, rounded.</summary>
    public static long Emu(double points)
    {
        return (long)Math.Round(points * EmuPerPoint);
    }

    /// <summary>Degrees to DrawingML's 60 000ths of a degree, turned into 0 to 360.</summary>
    public static long Angle(double degrees)
    {
        double turned = degrees % 360;

        if (turned < 0)
        {
            turned += 360;
        }

        return (long)Math.Round(turned * 60_000) % 21_600_000;
    }

    /// <summary>A fraction (0 to 1) to DrawingML's 100 000ths.</summary>
    public static long Percent(double fraction)
    {
        return (long)Math.Round(fraction * 100_000);
    }

    public static string Text(double number)
    {
        return number.ToString("R", CultureInfo.InvariantCulture);
    }

    public static string Text(long number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}
