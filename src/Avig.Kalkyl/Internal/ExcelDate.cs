namespace Avig.Kalkyl.Internal;

/// <summary>
/// Excel dates are days since a starting point. In the 1900 system day 1 is 1900-01-01, and Excel
/// counts 1900 as a leap year (inherited from Lotus 1-2-3), so day 60 is the made-up 1900-02-29.
/// From day 61 (1900-03-01) the count is right again with 1899-12-30 as day 0. In the 1904 system
/// (old Mac files) day 0 is 1904-01-01 and there is no made-up date.
/// </summary>
internal static class ExcelDate
{
    /// <summary>10000-01-01 in the 1900 system: every serial number below it is a date DateTime can hold.</summary>
    private const double EndSerial1900 = 2_958_466;

    /// <summary>10000-01-01 in the 1904 system, which counts 1 462 days fewer.</summary>
    private const double EndSerial1904 = EndSerial1900 - 1_462;

    private static readonly DateTime Base1900 = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime DayZero1900 = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime FirstMarch1900 = new(1900, 3, 1, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime Base1904 = new(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

    /// <summary>A date to a serial number in the 1900 or 1904 system.</summary>
    public static double ToSerial(DateTime date, bool uses1904)
    {
        if (uses1904 == true)
        {
            return (date - Base1904).TotalDays;
        }

        return ToSerial(date);
    }

    /// <summary>A date to a serial number in the 1900 system, which is what new files always use.</summary>
    public static double ToSerial(DateTime date)
    {
        if (date < FirstMarch1900)
        {
            return (date - DayZero1900).TotalDays;
        }

        return (date - Base1900).TotalDays;
    }

    /// <summary>A serial number to a date, rounded to whole milliseconds so that 12:00 does not become 11:59:59.999.</summary>
    public static bool TryFromSerial(double serial, bool uses1904, out DateTime date)
    {
        date = DateTime.MinValue;

        double end = EndSerial1900;

        if (uses1904 == true)
        {
            end = EndSerial1904;
        }

        if (double.IsFinite(serial) == false || serial < 0 || serial >= end)
        {
            return false;
        }

        // Rounded to whole milliseconds first, and the side of the made-up day chosen after: a
        // serial a hair below 60 is day 60 once rounded, and must not become 1900-03-01.
        const long Day = 86_400_000;
        long milliseconds = (long)Math.Round(serial * Day, MidpointRounding.AwayFromZero);
        DateTime start;

        if (uses1904 == true)
        {
            start = Base1904;
        }
        else if (milliseconds < 60 * Day)
        {
            start = DayZero1900;
        }
        else if (milliseconds < 61 * Day)
        {
            // The made-up 1900-02-29 does not exist in DateTime. The nearest real day is the 28th.
            start = DayZero1900;
            milliseconds -= Day;
        }
        else
        {
            start = Base1900;
        }

        // The last fraction of a millisecond before 10000-01-01 rounds up past what DateTime can hold.
        if (milliseconds > (DateTime.MaxValue - start).Ticks / TimeSpan.TicksPerMillisecond)
        {
            return false;
        }

        date = start.AddMilliseconds(milliseconds);

        return true;
    }
}
