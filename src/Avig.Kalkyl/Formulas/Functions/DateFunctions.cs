using System.Globalization;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Date and time functions. Dates are serial numbers in the workbook's date system (1900 or 1904).
/// Day parts are worked out from the serial, so the 1900 system's day 0 (1900-01-00) and its
/// made-up day 60 (1900-02-29) give what Excel gives.
/// </summary>
internal static class DateFunctions
{
    private static readonly DateTime Base1900 = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime DayZero1900 = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime Base1904 = new(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
    private static readonly DateTime LastDate = new(9999, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);

    private static readonly string[] ExtraDateFormats =
    [
        "d-MMM-yyyy", "d-MMM-yy", "d MMM yy", "MMMM d, yyyy", "MMMM d yyyy", "MMM d yyyy", "yyyy-M-d", "yyyy/M/d", "M/d/yy", "d-MMMM-yyyy",
    ];

    public static void Register(FunctionTable table)
    {
        table.Add("TODAY", 0, 0, call => Value.FromNumber(Math.Floor(ToSerial(call, call.Evaluator.Now))), true);
        table.Add("NOW", 0, 0, call => Value.FromNumber(ToSerial(call, call.Evaluator.Now)), true);
        table.Add("DATE", 3, 3, Date);
        table.Add("DATEVALUE", 1, 1, DateValue);
        table.Add("TIMEVALUE", 1, 1, TimeValue);
        table.Add("YEAR", 1, 1, call => DatePart(call, parts => parts.Year));
        table.Add("MONTH", 1, 1, call => DatePart(call, parts => parts.Month));
        table.Add("DAY", 1, 1, call => DatePart(call, parts => parts.Day));
        table.Add("HOUR", 1, 1, call => TimePart(call, seconds => seconds / 3600));
        table.Add("MINUTE", 1, 1, call => TimePart(call, seconds => seconds / 60 % 60));
        table.Add("SECOND", 1, 1, call => TimePart(call, seconds => seconds % 60));
        table.Add("TIME", 3, 3, Time);
        table.Add("WEEKDAY", 1, 2, Weekday);
        table.Add("WEEKNUM", 1, 2, WeekNumber);
        table.Add("ISOWEEKNUM", 1, 1, IsoWeekNumber);
        table.Add("EDATE", 2, 2, call => MonthsLater(call, false));
        table.Add("EOMONTH", 2, 2, call => MonthsLater(call, true));
        table.Add("DAYS", 2, 2, Days);
        table.Add("DAYS360", 2, 3, Days360);
        table.Add("DATEDIF", 3, 3, DateDif);
        table.Add("NETWORKDAYS", 2, 3, call => NetworkDays(call, false));
        table.Add("NETWORKDAYS.INTL", 2, 4, call => NetworkDays(call, true));
        table.Add("WORKDAY", 2, 3, call => Workday(call, false));
        table.Add("WORKDAY.INTL", 2, 4, call => Workday(call, true));
        table.Add("YEARFRAC", 2, 3, YearFrac);
    }

    /// <summary>A date as a serial number in the workbook's date system.</summary>
    internal static double ToSerial(FunctionCall call, DateTime date)
    {
        return ExcelDate.ToSerial(date, call.Uses1904);
    }

    /// <summary>A serial number as a date in the workbook's date system. False (with #NUM!) outside Excel's dates.</summary>
    internal static bool TryToDate(FunctionCall call, double serial, out DateTime date, out Value error)
    {
        error = Value.Blank;

        if (ExcelDate.TryFromSerial(serial, call.Uses1904, out date) == false)
        {
            error = Value.Num;

            return false;
        }

        return true;
    }

    /// <summary>DATE(year, month, day): months and days overflow into the next year or month; years 0–1899 mean 1900 + year.</summary>
    private static Value Date(FunctionCall call)
    {
        if (call.TryInteger(0, out long year, out Value error) == false
            || call.TryInteger(1, out long month, out error) == false
            || call.TryInteger(2, out long day, out error) == false)
        {
            return error;
        }

        if (year < 0 || year >= 10000)
        {
            return Value.Num;
        }

        if (year < 1900)
        {
            year += 1900;
        }

        if (TrySerialOf(year, month, day, call.Uses1904, out long serial) == false || IsValidDay(serial, call.Uses1904) == false)
        {
            return Value.Num;
        }

        return Value.FromNumber(serial);
    }

    private static Value DateValue(FunctionCall call)
    {
        if (TryParseDateText(call, out double serial, out Value error) == false)
        {
            return error;
        }

        double day = Math.Floor(serial);

        if (IsValidDay(day, call.Uses1904) == false)
        {
            return Value.ValueError;
        }

        return Value.FromNumber(day);
    }

    private static Value TimeValue(FunctionCall call)
    {
        if (TryParseDateText(call, out double serial, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(serial - Math.Floor(serial));
    }

    /// <summary>The text argument of DATEVALUE and TIMEVALUE as a serial. A plain number is not a date (#VALUE!).</summary>
    private static bool TryParseDateText(FunctionCall call, out double serial, out Value error)
    {
        serial = 0;
        error = Value.ValueError;
        Value scalar = call.Scalar(0);

        if (scalar.IsError == true)
        {
            error = scalar;

            return false;
        }

        if (scalar.Kind != ValueKind.Text)
        {
            return false;
        }

        return TryParseDate(scalar.Text!, call.Uses1904, out serial);
    }

    /// <summary>Date or time text as a serial, in the invariant formats. A plain number, a percentage or an amount is not a date.</summary>
    private static bool TryParseDate(string source, bool uses1904, out double serial)
    {
        serial = 0;
        string text = source.Trim();

        if (text.Contains('%', StringComparison.Ordinal) == true || text.StartsWith('$') == true)
        {
            return false;
        }

        if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _) == true)
        {
            return false;
        }

        if (Coerce.TryParseNumber(text, uses1904, out serial) == true)
        {
            return serial >= 0;
        }

        if (DateTime.TryParseExact(text, ExtraDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime date) == true)
        {
            serial = ExcelDate.ToSerial(date, uses1904);

            return serial >= 0;
        }

        return false;
    }

    private static Value DatePart(FunctionCall call, Func<DateParts, int> part)
    {
        if (TryDay(call, 0, out long day, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(part(PartsOf(day, call.Uses1904)));
    }

    /// <summary>HOUR, MINUTE and SECOND: the time of day rounded to the nearest second.</summary>
    private static Value TimePart(FunctionCall call, Func<long, long> part)
    {
        if (call.TryNumber(0, out double serial, out Value error) == false)
        {
            return error;
        }

        if (serial < 0 || serial >= MaxSerial(call.Uses1904) + 1)
        {
            return Value.Num;
        }

        long seconds = (long)Math.Round((serial - Math.Floor(serial)) * 86_400, MidpointRounding.AwayFromZero) % 86_400;

        return Value.FromNumber(part(seconds));
    }

    /// <summary>TIME(hour, minute, second): wraps past 24 hours; a negative total or a part above 32767 is #NUM!.</summary>
    private static Value Time(FunctionCall call)
    {
        if (call.TryInteger(0, out long hour, out Value error) == false
            || call.TryInteger(1, out long minute, out error) == false
            || call.TryInteger(2, out long second, out error) == false)
        {
            return error;
        }

        if (hour > 32767 || minute > 32767 || second > 32767)
        {
            return Value.Num;
        }

        long total = (hour * 3600) + (minute * 60) + second;

        if (total < 0)
        {
            return Value.Num;
        }

        return Value.FromNumber((total % 86_400) / 86_400d);
    }

    /// <summary>WEEKDAY(serial, [type]): types 1, 2, 3 and 11–17 as documented; others are #NUM!.</summary>
    private static Value Weekday(FunctionCall call)
    {
        if (TryDay(call, 0, out long day, out Value error) == false || call.TryInteger(1, 1, out long type, out error) == false)
        {
            return error;
        }

        int sunday = SundayIndex(day, call.Uses1904);

        switch (type)
        {
            case 1:
            case 17:
                return Value.FromNumber(sunday + 1);

            case 2:
                return Value.FromNumber(((sunday + 6) % 7) + 1);

            case 3:
                return Value.FromNumber((sunday + 6) % 7);

            case >= 11 and <= 16:
                int start = (int)(type - 10);

                return Value.FromNumber(((sunday - start + 7) % 7) + 1);

            default:
                return Value.Num;
        }
    }

    /// <summary>WEEKNUM(serial, [type]): the week that holds January 1 is week 1; type 21 is the ISO week.</summary>
    private static Value WeekNumber(FunctionCall call)
    {
        if (TryDay(call, 0, out long day, out Value error) == false || call.TryInteger(1, 1, out long type, out error) == false)
        {
            return error;
        }

        int start;

        switch (type)
        {
            case 1:
            case 17:
                start = 0;

                break;

            case 2:
                start = 1;

                break;

            case >= 11 and <= 16:
                start = (int)(type - 10);

                break;

            case 21:
                return Value.FromNumber(IsoWeek(day, call.Uses1904));

            default:
                return Value.Num;
        }

        DateParts parts = PartsOf(day, call.Uses1904);
        TrySerialOf(parts.Year, 1, 1, call.Uses1904, out long january);
        int offset = (SundayIndex(january, call.Uses1904) - start + 7) % 7;
        long week = (long)Math.Floor((day - january + offset) / 7d) + 1;

        return Value.FromNumber(week);
    }

    private static Value IsoWeekNumber(FunctionCall call)
    {
        if (TryDay(call, 0, out long day, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(IsoWeek(day, call.Uses1904));
    }

    private static long IsoWeek(long day, bool uses1904)
    {
        int isoDay = ((SundayIndex(day, uses1904) + 6) % 7) + 1;
        long thursday = day - isoDay + 4;
        DateParts parts = PartsOf(thursday, uses1904);
        TrySerialOf(parts.Year, 1, 1, uses1904, out long january);

        return ((thursday - january) / 7) + 1;
    }

    /// <summary>EDATE and EOMONTH: the same day (or the month's last day) a number of months away.</summary>
    private static Value MonthsLater(FunctionCall call, bool endOfMonth)
    {
        if (TryDay(call, 0, out long day, out Value error) == false || call.TryInteger(1, out long months, out error) == false)
        {
            return error;
        }

        DateParts parts = PartsOf(day, call.Uses1904);
        long total = (parts.Year * 12L) + (parts.Month - 1) + months;
        long year = FloorDivide(total, 12);
        int month = (int)(total - (year * 12)) + 1;

        if (year < 1 || year > 9999)
        {
            return Value.Num;
        }

        int last = DaysInMonth((int)year, month, call.Uses1904);
        int target = Math.Min(parts.Day, last);

        if (endOfMonth == true)
        {
            target = last;
        }

        if (TrySerialOf(year, month, target, call.Uses1904, out long serial) == false || IsValidDay(serial, call.Uses1904) == false)
        {
            return Value.Num;
        }

        return Value.FromNumber(serial);
    }

    private static Value Days(FunctionCall call)
    {
        if (TryDay(call, 0, out long end, out Value error) == false || TryDay(call, 1, out long start, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(end - start);
    }

    /// <summary>
    /// DAYS360(start, end, [european]): the US (NASD) method moves a start on the 31st or on the
    /// last day of February to the 30th, and an end on the 31st to the 30th only when the start is
    /// then the 30th. The European method moves every 31st to the 30th.
    /// </summary>
    private static Value Days360(FunctionCall call)
    {
        if (TryDay(call, 0, out long start, out Value error) == false
            || TryDay(call, 1, out long end, out error) == false
            || call.TryBool(2, false, out bool european, out error) == false)
        {
            return error;
        }

        DateParts first = PartsOf(start, call.Uses1904);
        DateParts second = PartsOf(end, call.Uses1904);
        int firstDay = first.Day;
        int secondDay = second.Day;

        if (european == true)
        {
            firstDay = Math.Min(firstDay, 30);
            secondDay = Math.Min(secondDay, 30);
        }
        else
        {
            if (firstDay == 31 || (first.Month == 2 && firstDay == DaysInMonth(first.Year, 2, call.Uses1904)))
            {
                firstDay = 30;
            }

            if (secondDay == 31 && firstDay == 30)
            {
                secondDay = 30;
            }
        }

        long days = ((second.Year - first.Year) * 360L) + ((second.Month - first.Month) * 30L) + secondDay - firstDay;

        return Value.FromNumber(days);
    }

    /// <summary>DATEDIF(start, end, unit) with units Y, M, D, MD, YM and YD, including Excel's MD quirk at month ends.</summary>
    private static Value DateDif(FunctionCall call)
    {
        if (TryDay(call, 0, out long start, out Value error) == false
            || TryDay(call, 1, out long end, out error) == false
            || call.TryText(2, out string unit, out error) == false)
        {
            return error;
        }

        if (start > end)
        {
            return Value.Num;
        }

        DateParts first = PartsOf(start, call.Uses1904);
        DateParts second = PartsOf(end, call.Uses1904);
        long months = ((second.Year - first.Year) * 12L) + second.Month - first.Month;

        if (second.Day < first.Day)
        {
            months--;
        }

        switch (unit.ToUpperInvariant())
        {
            case "Y":
                return Value.FromNumber(months / 12);

            case "M":
                return Value.FromNumber(months);

            case "D":
                return Value.FromNumber(end - start);

            case "YM":
                return Value.FromNumber(months % 12);

            case "MD":
                if (second.Day >= first.Day)
                {
                    return Value.FromNumber(second.Day - first.Day);
                }

                TrySerialOf(second.Year, second.Month - 1, first.Day, call.Uses1904, out long shifted);

                return Value.FromNumber(end - shifted);

            case "YD":
                long year = second.Year;

                if (second.Month < first.Month || (second.Month == first.Month && second.Day < first.Day))
                {
                    year--;
                }

                TrySerialOf(year, first.Month, first.Day, call.Uses1904, out long anniversary);

                return Value.FromNumber(end - anniversary);

            default:
                return Value.Num;
        }
    }

    /// <summary>NETWORKDAYS and NETWORKDAYS.INTL: working days from start to end, both counted; negative when end comes first.</summary>
    private static Value NetworkDays(FunctionCall call, bool international)
    {
        int holidaysIndex = 2;
        bool[] weekend = DefaultWeekend();
        Value error;

        if (international == true)
        {
            holidaysIndex = 3;

            if (TryWeekend(call, 2, out weekend, out error) == false)
            {
                return error;
            }
        }

        if (TryDay(call, 0, out long start, out error) == false
            || TryDay(call, 1, out long end, out error) == false
            || TryHolidays(call, holidaysIndex, out HashSet<long> holidays, out error) == false)
        {
            return error;
        }

        long sign = 1;

        if (start > end)
        {
            (start, end) = (end, start);
            sign = -1;
        }

        long total = end - start + 1;
        long weeks = total / 7;
        long perWeek = weekend.Count(day => day == false);
        long count = weeks * perWeek;

        for (long day = start + (weeks * 7); day <= end; day++)
        {
            if (weekend[MondayIndex(day, call.Uses1904)] == false)
            {
                count++;
            }
        }

        foreach (long holiday in holidays)
        {
            if (holiday >= start && holiday <= end && weekend[MondayIndex(holiday, call.Uses1904)] == false)
            {
                count--;
            }
        }

        return Value.FromNumber(sign * count);
    }

    /// <summary>WORKDAY and WORKDAY.INTL: the working day a number of working days before or after start.</summary>
    private static Value Workday(FunctionCall call, bool international)
    {
        int holidaysIndex = 2;
        bool[] weekend = DefaultWeekend();
        Value error;

        if (international == true)
        {
            holidaysIndex = 3;

            if (TryWeekend(call, 2, out weekend, out error) == false)
            {
                return error;
            }
        }

        if (TryDay(call, 0, out long start, out error) == false
            || call.TryInteger(1, out long days, out error) == false
            || TryHolidays(call, holidaysIndex, out HashSet<long> holidays, out error) == false)
        {
            return error;
        }

        double maximum = MaxSerial(call.Uses1904);

        if (Math.Abs(days) > maximum)
        {
            return Value.Num;
        }

        long step = Math.Sign(days);
        long remaining = Math.Abs(days);
        long current = start;

        while (remaining > 0)
        {
            current += step;

            if (current < 0 || current > maximum)
            {
                return Value.Num;
            }

            if (weekend[MondayIndex(current, call.Uses1904)] == false && holidays.Contains(current) == false)
            {
                remaining--;
            }
        }

        return Value.FromNumber(current);
    }

    /// <summary>
    /// YEARFRAC(start, end, [basis]): 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365,
    /// 4 European 30/360, following Excel's own rules for each (basis 1 averages the year length
    /// over the years spanned when the dates are more than a year apart).
    /// </summary>
    private static Value YearFrac(FunctionCall call)
    {
        if (TryDay(call, 0, out long start, out Value error) == false
            || TryDay(call, 1, out long end, out error) == false
            || call.TryInteger(2, 0, out long basis, out error) == false)
        {
            return error;
        }

        if (basis < 0 || basis > 4)
        {
            return Value.Num;
        }

        if (start > end)
        {
            (start, end) = (end, start);
        }

        bool uses1904 = call.Uses1904;
        DateParts first = PartsOf(start, uses1904);
        DateParts second = PartsOf(end, uses1904);
        double days = end - start;

        switch (basis)
        {
            case 0:
                return Value.FromNumber(UsThirty360(first, second, uses1904) / 360d);

            case 2:
                return Value.FromNumber(days / 360);

            case 3:
                return Value.FromNumber(days / 365);

            case 4:
                long europeanDays = ((second.Year - first.Year) * 360L) + ((second.Month - first.Month) * 30L) + Math.Min(second.Day, 30) - Math.Min(first.Day, 30);

                return Value.FromNumber(europeanDays / 360d);
        }

        if (start == end)
        {
            return Value.FromNumber(0);
        }

        bool withinYear = first.Year == second.Year
            || (first.Year + 1 == second.Year && (first.Month > second.Month || (first.Month == second.Month && first.Day >= second.Day)));

        if (withinYear == true)
        {
            double yearLength = 365;

            if (first.Year == second.Year && DateTime.IsLeapYear(first.Year) == true)
            {
                yearLength = 366;
            }
            else if (HasLeapDay(first, start, second, end, uses1904) == true || (second.Month == 2 && second.Day == 29))
            {
                yearLength = 366;
            }

            return Value.FromNumber(days / yearLength);
        }

        TrySerialOf(first.Year, 1, 1, uses1904, out long firstJanuary);
        TrySerialOf(second.Year + 1L, 1, 1, uses1904, out long afterLast);
        double average = (afterLast - firstJanuary) / (double)(second.Year - first.Year + 1);

        return Value.FromNumber(days / average);
    }

    /// <summary>YEARFRAC basis 0, which treats the last day of February differently from DAYS360.</summary>
    private static long UsThirty360(DateParts first, DateParts second, bool uses1904)
    {
        int firstDay = first.Day;
        int secondDay = second.Day;
        bool firstIsFebruaryEnd = first.Month == 2 && firstDay == DaysInMonth(first.Year, 2, uses1904);
        bool secondIsFebruaryEnd = second.Month == 2 && secondDay == DaysInMonth(second.Year, 2, uses1904);

        if (firstDay == 31 && secondDay == 31)
        {
            firstDay = 30;
            secondDay = 30;
        }
        else if (firstDay == 31)
        {
            firstDay = 30;
        }
        else if (firstDay == 30 && secondDay == 31)
        {
            secondDay = 30;
        }
        else if (firstIsFebruaryEnd == true && secondIsFebruaryEnd == true)
        {
            firstDay = 30;
            secondDay = 30;
        }
        else if (firstIsFebruaryEnd == true)
        {
            firstDay = 30;
        }

        return ((second.Year - first.Year) * 360L) + ((second.Month - first.Month) * 30L) + secondDay - firstDay;
    }

    /// <summary>Whether a February 29 lies between two dates at most a year apart.</summary>
    private static bool HasLeapDay(DateParts first, long start, DateParts second, long end, bool uses1904)
    {
        TrySerialOf(first.Year, 3, 1, uses1904, out long firstMarch);

        if (DateTime.IsLeapYear(first.Year) == true && start < firstMarch)
        {
            return true;
        }

        TrySerialOf(second.Year, 3, 1, uses1904, out long secondMarch);

        return DateTime.IsLeapYear(second.Year) == true && end >= secondMarch;
    }

    /// <summary>The weekend of the .INTL functions: a number 1–7 or 11–17, or seven 0/1 characters starting on Monday.</summary>
    private static bool TryWeekend(FunctionCall call, int index, out bool[] weekend, out Value error)
    {
        weekend = DefaultWeekend();
        error = Value.Blank;

        if (call.IsMissing(index) == true)
        {
            return true;
        }

        Value scalar = call.Scalar(index);

        if (scalar.IsError == true)
        {
            error = scalar;

            return false;
        }

        if (scalar.Kind == ValueKind.Text)
        {
            string text = scalar.Text!;
            weekend = new bool[7];

            if (text.Length != 7 || text.All(letter => letter is '0' or '1') == false || text.All(letter => letter == '1') == true)
            {
                error = Value.ValueError;

                return false;
            }

            for (int day = 0; day < 7; day++)
            {
                weekend[day] = text[day] == '1';
            }

            return true;
        }

        if (call.TryInteger(index, out long number, out error) == false)
        {
            return false;
        }

        weekend = new bool[7];

        if (number >= 1 && number <= 7)
        {
            weekend[(int)((number + 4) % 7)] = true;
            weekend[(int)((number + 5) % 7)] = true;

            return true;
        }

        if (number >= 11 && number <= 17)
        {
            weekend[(int)((number - 11 + 6) % 7)] = true;

            return true;
        }

        error = Value.Num;

        return false;
    }

    private static bool[] DefaultWeekend()
    {
        return [false, false, false, false, false, true, true];
    }

    private static bool TryHolidays(FunctionCall call, int index, out HashSet<long> holidays, out Value error)
    {
        holidays = [];
        error = Value.Blank;

        if (call.IsMissing(index) == true)
        {
            return true;
        }

        List<double> numbers = [];

        foreach (Value item in call.Values(index))
        {
            Value converted = item;

            if (item.Kind == ValueKind.Text)
            {
                if (TryParseDate(item.Text!, call.Uses1904, out double parsed) == false)
                {
                    error = Value.ValueError;

                    return false;
                }

                converted = Value.FromNumber(parsed);
            }

            switch (converted.Kind)
            {
                case ValueKind.Blank:
                    continue;

                case ValueKind.Error:
                    error = converted;

                    return false;

                case ValueKind.Number:
                    numbers.Add(converted.Number);

                    break;

                default:
                    error = Value.ValueError;

                    return false;
            }
        }

        foreach (double number in numbers)
        {
            if (IsValidDay(number, call.Uses1904) == false)
            {
                error = Value.Num;

                return false;
            }

            holidays.Add((long)Math.Floor(number));
        }

        return true;
    }

    /// <summary>A date argument as a whole day: the time is dropped; outside Excel's dates is #NUM!.</summary>
    private static bool TryDay(FunctionCall call, int index, out long day, out Value error)
    {
        day = 0;

        if (call.TryNumber(index, out double serial, out error) == false)
        {
            return false;
        }

        if (IsValidDay(serial, call.Uses1904) == false)
        {
            error = Value.Num;

            return false;
        }

        day = (long)Math.Floor(serial);

        return true;
    }

    private static bool IsValidDay(double serial, bool uses1904)
    {
        return serial >= 0 && Math.Floor(serial) <= MaxSerial(uses1904);
    }

    private static double MaxSerial(bool uses1904)
    {
        return ExcelDate.ToSerial(LastDate, uses1904);
    }

    /// <summary>The year, month and day of a whole day number, Excel's way.</summary>
    private static DateParts PartsOf(long day, bool uses1904)
    {
        DateTime date;

        if (uses1904 == true)
        {
            date = Base1904.AddDays(day);
        }
        else if (day == 0)
        {
            return new DateParts(1900, 1, 0);
        }
        else if (day == 60)
        {
            return new DateParts(1900, 2, 29);
        }
        else if (day < 60)
        {
            date = DayZero1900.AddDays(day);
        }
        else
        {
            date = Base1900.AddDays(day);
        }

        return new DateParts(date.Year, date.Month, date.Day);
    }

    /// <summary>The day number of a year, month and day, where the month and day may overflow. False outside the years 1–9999.</summary>
    private static bool TrySerialOf(long year, long month, long day, bool uses1904, out long serial)
    {
        serial = 0;
        long total = (year * 12) + month - 1;
        long normalizedYear = FloorDivide(total, 12);
        int normalizedMonth = (int)(total - (normalizedYear * 12)) + 1;

        if (normalizedYear < 1 || normalizedYear > 9999 || Math.Abs(day) > 10_000_000)
        {
            return false;
        }

        long first = (long)ExcelDate.ToSerial(new DateTime((int)normalizedYear, normalizedMonth, 1, 0, 0, 0, DateTimeKind.Unspecified), uses1904);
        serial = first + day - 1;

        return true;
    }

    /// <summary>Days in a month; February 1900 has 29 in the 1900 system, as Excel believes.</summary>
    private static int DaysInMonth(int year, int month, bool uses1904)
    {
        if (uses1904 == false && year == 1900 && month == 2)
        {
            return 29;
        }

        return DateTime.DaysInMonth(year, month);
    }

    /// <summary>The day of the week, 0 for Sunday. In the 1900 system day 1 is a Sunday, as Excel has it.</summary>
    private static int SundayIndex(long day, bool uses1904)
    {
        long shifted = day + 6;

        if (uses1904 == true)
        {
            shifted = day + 1462 + 6;
        }

        return (int)(((shifted % 7) + 7) % 7);
    }

    /// <summary>The day of the week, 0 for Monday.</summary>
    private static int MondayIndex(long day, bool uses1904)
    {
        return (SundayIndex(day, uses1904) + 6) % 7;
    }

    private static long FloorDivide(long value, long divisor)
    {
        long quotient = value / divisor;

        if (value % divisor != 0 && (value < 0) != (divisor < 0))
        {
            quotient--;
        }

        return quotient;
    }

    private readonly record struct DateParts(int Year, int Month, int Day);
}
