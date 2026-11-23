using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>Date and time functions. Dates are serial numbers in the workbook's date system (1900 or 1904).</summary>
internal static class DateFunctions
{
    public static void Register(FunctionTable table)
    {
        table.Add("TODAY", 0, 0, call => Value.FromNumber(Math.Floor(ToSerial(call, call.Evaluator.Now))), true);
        table.Add("NOW", 0, 0, call => Value.FromNumber(ToSerial(call, call.Evaluator.Now)), true);
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
}
