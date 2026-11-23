namespace Avig.Kalkyl.Internal;

/// <summary>
/// Excels datum är dagar sedan en startpunkt. I 1900-systemet är dag 1 1900-01-01, och Excel
/// räknar 1900 som skottår (ett arv från Lotus 1-2-3), så dag 60 är den påhittade 1900-02-29.
/// Från dag 61 (1900-03-01) stämmer räkningen igen med 1899-12-30 som dag 0. I 1904-systemet
/// (gamla Mac-filer) är dag 0 1904-01-01 och det finns inget påhittat datum.
/// </summary>
internal static class ExcelDate
{
    private const double LastSerial = 2_958_465.999_999_9;

    private static readonly DateTime Base1900 = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime DayZero1900 = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime FirstMarch1900 = new(1900, 3, 1, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime Base1904 = new(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

    /// <summary>Datum till serienummer i 1900-systemet, som är det paketet alltid skriver.</summary>
    public static double ToSerial(DateTime date)
    {
        if (date < FirstMarch1900)
        {
            return (date - DayZero1900).TotalDays;
        }

        return (date - Base1900).TotalDays;
    }

    /// <summary>Serienummer till datum. Avrundas till hela millisekunder så att 12:00 inte blir 11:59:59.999.</summary>
    public static bool TryFromSerial(double serial, bool uses1904, out DateTime date)
    {
        date = DateTime.MinValue;

        if (double.IsFinite(serial) == false || serial < 0 || serial > LastSerial)
        {
            return false;
        }

        DateTime start;
        double days = serial;

        if (uses1904 == true)
        {
            start = Base1904;
        }
        else if (serial < 60)
        {
            start = DayZero1900;
        }
        else if (serial < 61)
        {
            // Den påhittade 1900-02-29 finns inte i DateTime. Närmaste riktiga dag är den 28:e.
            start = DayZero1900;
            days = serial - 1;
        }
        else
        {
            start = Base1900;
        }

        long milliseconds = (long)Math.Round(days * 86_400_000d, MidpointRounding.AwayFromZero);
        date = start.AddMilliseconds(milliseconds);

        return true;
    }
}
