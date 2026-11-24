namespace Avig.Kalkyl.Formulas.Functions;

/// <summary>
/// Financial functions: loans, annuities, investments and depreciation. They follow Excel's sign
/// convention (money paid out is negative) and its type argument (0 = payments at the end of each
/// period, 1 = at the beginning).
/// </summary>
internal static class FinancialFunctions
{
    private const double RateTolerance = 1e-7;

    private const int RateIterations = 20;

    private const int IrrIterations = 100;

    public static void Register(FunctionTable table)
    {
        ArgumentNullException.ThrowIfNull(table);

        table.Add("PMT", 3, 5, Pmt);
        table.Add("IPMT", 4, 6, Ipmt);
        table.Add("PPMT", 4, 6, Ppmt);
        table.Add("PV", 3, 5, Pv);
        table.Add("FV", 3, 5, Fv);
        table.Add("NPER", 3, 5, Nper);
        table.Add("RATE", 3, 6, Rate);
        table.Add("NPV", 2, FunctionTable.Many, Npv);
        table.Add("IRR", 1, 2, Irr);
        table.Add("XNPV", 3, 3, Xnpv);
        table.Add("XIRR", 2, 3, Xirr);
        table.Add("MIRR", 3, 3, Mirr);
        table.Add("SLN", 3, 3, Sln);
        table.Add("SYD", 4, 4, Syd);
        table.Add("DB", 4, 5, Db);
        table.Add("DDB", 4, 5, Ddb);
        table.Add("VDB", 5, 7, Vdb);
        table.Add("EFFECT", 2, 2, Effect);
        table.Add("NOMINAL", 2, 2, Nominal);
        table.Add("CUMIPMT", 6, 6, call => Cumulative(call, true));
        table.Add("CUMPRINC", 6, 6, call => Cumulative(call, false));
        table.Add("FVSCHEDULE", 2, 2, FvSchedule);
        table.Add("PDURATION", 3, 3, PDuration);
        table.Add("RRI", 3, 3, Rri);
        table.Add("ISPMT", 4, 4, IsPmt);
        table.Add("DOLLARDE", 2, 2, call => Dollar(call, true));
        table.Add("DOLLARFR", 2, 2, call => Dollar(call, false));
    }

    private static Value Pmt(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double periods, out error) == false
            || call.TryNumber(2, out double present, out error) == false
            || call.TryNumber(3, 0, out double future, out error) == false
            || call.TryNumber(4, 0, out double type, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(Payment(rate, periods, present, future, Timing(type)));
    }

    private static Value Ipmt(FunctionCall call)
    {
        if (TryPeriodPayment(call, out double rate, out double period, out double periods, out double present, out double future, out int timing, out Value error) == false)
        {
            return error;
        }

        return Value.FromNumber(InterestPayment(rate, period, periods, present, future, timing));
    }

    private static Value Ppmt(FunctionCall call)
    {
        if (TryPeriodPayment(call, out double rate, out double period, out double periods, out double present, out double future, out int timing, out Value error) == false)
        {
            return error;
        }

        double payment = Payment(rate, periods, present, future, timing);

        return Value.FromNumber(payment - InterestPayment(rate, period, periods, present, future, timing));
    }

    /// <summary>The arguments of IPMT and PPMT: the period must lie between 1 and nper.</summary>
    private static bool TryPeriodPayment(FunctionCall call, out double rate, out double period, out double periods, out double present, out double future, out int timing, out Value error)
    {
        timing = 0;
        period = 0;
        periods = 0;
        present = 0;
        future = 0;

        if (call.TryNumber(0, out rate, out error) == false
            || call.TryNumber(1, out period, out error) == false
            || call.TryNumber(2, out periods, out error) == false
            || call.TryNumber(3, out present, out error) == false
            || call.TryNumber(4, 0, out future, out error) == false
            || call.TryNumber(5, 0, out double type, out error) == false)
        {
            return false;
        }

        if (period < 1 || period > periods)
        {
            error = Value.Num;

            return false;
        }

        timing = Timing(type);

        return true;
    }

    private static Value Pv(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double periods, out error) == false
            || call.TryNumber(2, out double payment, out error) == false
            || call.TryNumber(3, 0, out double future, out error) == false
            || call.TryNumber(4, 0, out double type, out error) == false)
        {
            return error;
        }

        int timing = Timing(type);

        if (rate == 0)
        {
            return Value.FromNumber(-(future + (payment * periods)));
        }

        double growth = Math.Pow(1 + rate, periods);

        return Value.FromNumber(-(future + (payment * (1 + (rate * timing)) * (growth - 1) / rate)) / growth);
    }

    private static Value Fv(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double periods, out error) == false
            || call.TryNumber(2, out double payment, out error) == false
            || call.TryNumber(3, 0, out double present, out error) == false
            || call.TryNumber(4, 0, out double type, out error) == false)
        {
            return error;
        }

        return Value.FromNumber(FutureValue(rate, periods, payment, present, Timing(type)));
    }

    private static Value Nper(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double payment, out error) == false
            || call.TryNumber(2, out double present, out error) == false
            || call.TryNumber(3, 0, out double future, out error) == false
            || call.TryNumber(4, 0, out double type, out error) == false)
        {
            return error;
        }

        int timing = Timing(type);

        if (rate == 0)
        {
            if (payment == 0)
            {
                return Value.Num;
            }

            return Value.FromNumber(-(present + future) / payment);
        }

        double numerator = (payment * (1 + (rate * timing))) - (future * rate);
        double denominator = (present * rate) + (payment * (1 + (rate * timing)));

        if (rate <= -1 || denominator == 0 || numerator / denominator <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Log(numerator / denominator) / Math.Log(1 + rate));
    }

    /// <summary>RATE: Newton's method from the guess (0.1), at most 20 steps, #NUM! when it does not converge.</summary>
    private static Value Rate(FunctionCall call)
    {
        if (call.TryNumber(0, out double periods, out Value error) == false
            || call.TryNumber(1, out double payment, out error) == false
            || call.TryNumber(2, out double present, out error) == false
            || call.TryNumber(3, 0, out double future, out error) == false
            || call.TryNumber(4, 0, out double type, out error) == false
            || call.TryNumber(5, 0.1, out double guess, out error) == false)
        {
            return error;
        }

        if (periods <= 0)
        {
            return Value.Num;
        }

        int timing = Timing(type);
        double rate = guess;

        for (int iteration = 0; iteration < RateIterations; iteration++)
        {
            if (rate <= -1)
            {
                return Value.Num;
            }

            double value;
            double slope;

            if (Math.Abs(rate) < 1e-10)
            {
                value = present + (payment * periods) + future;
                slope = (present * periods) + (payment * ((timing * periods) + (periods * (periods - 1) / 2)));
            }
            else
            {
                double growth = Math.Pow(1 + rate, periods);
                double growthSlope = periods * Math.Pow(1 + rate, periods - 1);
                double annuity = (growth - 1) / rate;
                double annuitySlope = ((growthSlope * rate) - (growth - 1)) / (rate * rate);
                value = (present * growth) + (payment * (1 + (rate * timing)) * annuity) + future;
                slope = (present * growthSlope) + (payment * ((timing * annuity) + ((1 + (rate * timing)) * annuitySlope)));
            }

            if (TryNewtonStep(ref rate, value, slope, out bool converged) == false)
            {
                return Value.Num;
            }

            if (converged == true)
            {
                return Value.FromNumber(rate);
            }
        }

        return Value.Num;
    }

    /// <summary>One step of Newton's method: false when it breaks down, <paramref name="converged"/> when the step was below the tolerance.</summary>
    private static bool TryNewtonStep(ref double rate, double value, double slope, out bool converged)
    {
        converged = false;

        if (slope == 0 || double.IsFinite(value) == false || double.IsFinite(slope) == false)
        {
            return false;
        }

        double next = rate - (value / slope);

        if (double.IsFinite(next) == false)
        {
            return false;
        }

        converged = Math.Abs(next - rate) < RateTolerance;
        rate = next;

        return true;
    }

    private static Value Npv(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false || call.TryCollectNumbers(1, false, out List<double> values, out error) == false)
        {
            return error;
        }

        if (rate == -1)
        {
            return Value.Div0;
        }

        double sum = 0;

        for (int index = 0; index < values.Count; index++)
        {
            sum += values[index] / Math.Pow(1 + rate, index + 1);
        }

        return Value.FromNumber(sum);
    }

    /// <summary>IRR: Newton's method on the NPV from the guess (0.1); the values need a positive and a negative one.</summary>
    private static Value Irr(FunctionCall call)
    {
        if (TryRangeNumbers(call, 0, out List<double> values, out Value error) == false || call.TryNumber(1, 0.1, out double guess, out error) == false)
        {
            return error;
        }

        if (HasBothSigns(values) == false)
        {
            return Value.Num;
        }

        double rate = guess;

        for (int iteration = 0; iteration < IrrIterations; iteration++)
        {
            if (rate <= -1)
            {
                return Value.Num;
            }

            double value = 0;
            double slope = 0;

            for (int index = 0; index < values.Count; index++)
            {
                double discount = Math.Pow(1 + rate, index);
                value += values[index] / discount;
                slope -= index * values[index] / (discount * (1 + rate));
            }

            if (TryNewtonStep(ref rate, value, slope, out bool converged) == false)
            {
                return Value.Num;
            }

            if (converged == true)
            {
                return Value.FromNumber(rate);
            }
        }

        return Value.Num;
    }

    private static Value Xnpv(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false || TryCashFlows(call, 1, out List<double> values, out List<double> dates, out error) == false)
        {
            return error;
        }

        if (rate <= -1)
        {
            return Value.Num;
        }

        return Value.FromNumber(DatedValue(rate, values, dates));
    }

    /// <summary>XIRR: Newton's method on XNPV from the guess (0.1), with the dates as serials.</summary>
    private static Value Xirr(FunctionCall call)
    {
        if (TryCashFlows(call, 0, out List<double> values, out List<double> dates, out Value error) == false || call.TryNumber(2, 0.1, out double guess, out error) == false)
        {
            return error;
        }

        if (HasBothSigns(values) == false)
        {
            return Value.Num;
        }

        double rate = guess;

        for (int iteration = 0; iteration < IrrIterations; iteration++)
        {
            if (rate <= -1)
            {
                return Value.Num;
            }

            double value = DatedValue(rate, values, dates);
            double slope = 0;

            for (int index = 0; index < values.Count; index++)
            {
                double years = (dates[index] - dates[0]) / 365;
                slope -= years * values[index] / Math.Pow(1 + rate, years + 1);
            }

            if (TryNewtonStep(ref rate, value, slope, out bool converged) == false)
            {
                return Value.Num;
            }

            if (converged == true)
            {
                return Value.FromNumber(rate);
            }
        }

        return Value.Num;
    }

    private static double DatedValue(double rate, List<double> values, List<double> dates)
    {
        double sum = 0;

        for (int index = 0; index < values.Count; index++)
        {
            sum += values[index] / Math.Pow(1 + rate, (dates[index] - dates[0]) / 365);
        }

        return sum;
    }

    /// <summary>
    /// The values and dates of XNPV and XIRR: every item must be a number (#VALUE! otherwise), the
    /// counts must match and no date may come before the first (#NUM!). Dates are truncated to days.
    /// </summary>
    private static bool TryCashFlows(FunctionCall call, int first, out List<double> values, out List<double> dates, out Value error)
    {
        dates = [];

        if (TryStrictNumbers(call, first, out values, out error) == false || TryStrictNumbers(call, first + 1, out dates, out error) == false)
        {
            return false;
        }

        if (values.Count != dates.Count || values.Count == 0)
        {
            error = Value.Num;

            return false;
        }

        for (int index = 0; index < dates.Count; index++)
        {
            dates[index] = Math.Truncate(dates[index]);

            if (dates[index] < 0 || dates[index] < dates[0])
            {
                error = Value.Num;

                return false;
            }
        }

        return true;
    }

    private static bool TryStrictNumbers(FunctionCall call, int index, out List<double> numbers, out Value error)
    {
        numbers = [];
        error = Value.Blank;

        foreach (Value item in call.Values(index))
        {
            if (item.IsError == true)
            {
                error = item;

                return false;
            }

            if (item.Kind is ValueKind.Blank or ValueKind.Bool)
            {
                error = Value.ValueError;

                return false;
            }

            Value converted = Coerce.ToNumber(item, call.Conversion);

            if (converted.IsError == true)
            {
                error = converted;

                return false;
            }

            numbers.Add(converted.Number);
        }

        return true;
    }

    /// <summary>The numbers of a range or array argument, skipping text, booleans and blanks as IRR and MIRR do.</summary>
    private static bool TryRangeNumbers(FunctionCall call, int index, out List<double> numbers, out Value error)
    {
        numbers = [];
        error = Value.Blank;

        foreach (Value item in call.Values(index))
        {
            if (item.IsError == true)
            {
                error = item;

                return false;
            }

            if (item.Kind == ValueKind.Number)
            {
                numbers.Add(item.Number);
            }
        }

        return true;
    }

    private static bool HasBothSigns(List<double> values)
    {
        return values.Exists(item => item > 0) == true && values.Exists(item => item < 0) == true;
    }

    private static Value Mirr(FunctionCall call)
    {
        if (TryRangeNumbers(call, 0, out List<double> values, out Value error) == false
            || call.TryNumber(1, out double financeRate, out error) == false
            || call.TryNumber(2, out double reinvestRate, out error) == false)
        {
            return error;
        }

        int count = values.Count;
        double positive = 0;
        double negative = 0;

        for (int index = 0; index < count; index++)
        {
            if (values[index] > 0)
            {
                positive += values[index] / Math.Pow(1 + reinvestRate, index);
            }
            else
            {
                negative += values[index] / Math.Pow(1 + financeRate, index);
            }
        }

        if (positive == 0 || negative == 0 || count < 2)
        {
            return Value.Div0;
        }

        double ratio = -positive * Math.Pow(1 + reinvestRate, count - 1) / negative;

        return Value.FromNumber(Math.Pow(ratio, 1.0 / (count - 1)) - 1);
    }

    private static Value Sln(FunctionCall call)
    {
        if (call.TryNumber(0, out double cost, out Value error) == false
            || call.TryNumber(1, out double salvage, out error) == false
            || call.TryNumber(2, out double life, out error) == false)
        {
            return error;
        }

        if (life == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber((cost - salvage) / life);
    }

    private static Value Syd(FunctionCall call)
    {
        if (call.TryNumber(0, out double cost, out Value error) == false
            || call.TryNumber(1, out double salvage, out error) == false
            || call.TryNumber(2, out double life, out error) == false
            || call.TryNumber(3, out double period, out error) == false)
        {
            return error;
        }

        if (life <= 0 || period <= 0 || period > life)
        {
            return Value.Num;
        }

        return Value.FromNumber((cost - salvage) * (life - period + 1) * 2 / (life * (life + 1)));
    }

    /// <summary>DB: fixed-declining balance with the rate rounded to three decimals, as Excel does; month is the first year's months (default 12).</summary>
    private static Value Db(FunctionCall call)
    {
        if (call.TryNumber(0, out double cost, out Value error) == false
            || call.TryNumber(1, out double salvage, out error) == false
            || call.TryNumber(2, out double life, out error) == false
            || call.TryNumber(3, out double periodNumber, out error) == false
            || call.TryNumber(4, 12, out double monthNumber, out error) == false)
        {
            return error;
        }

        double period = Math.Truncate(periodNumber);
        double month = Math.Truncate(monthNumber);

        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || month < 1 || month > 12 || period > life + 1)
        {
            return Value.Num;
        }

        if (cost == 0)
        {
            return Value.FromNumber(0);
        }

        double rate = Numeric.Round(1 - Math.Pow(salvage / cost, 1 / life), 3);
        double total = 0;
        double depreciation = 0;

        for (int current = 1; current <= period; current++)
        {
            if (current == 1)
            {
                depreciation = cost * rate * month / 12;
            }
            else if (current > life)
            {
                depreciation = (cost - total) * rate * (12 - month) / 12;
            }
            else
            {
                depreciation = (cost - total) * rate;
            }

            total += depreciation;
        }

        return Value.FromNumber(depreciation);
    }

    private static Value Ddb(FunctionCall call)
    {
        if (call.TryNumber(0, out double cost, out Value error) == false
            || call.TryNumber(1, out double salvage, out error) == false
            || call.TryNumber(2, out double life, out error) == false
            || call.TryNumber(3, out double period, out error) == false
            || call.TryNumber(4, 2, out double factor, out error) == false)
        {
            return error;
        }

        if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(DecliningBalance(cost, salvage, life, period, factor));
    }

    /// <summary>The declining-balance depreciation of one period; fractional periods are allowed.</summary>
    private static double DecliningBalance(double cost, double salvage, double life, double period, double factor)
    {
        double rate = factor / life;
        double oldValue;

        if (rate >= 1)
        {
            rate = 1;
            oldValue = 0;

            if (period == 1)
            {
                oldValue = cost;
            }
        }
        else
        {
            oldValue = cost * Math.Pow(1 - rate, period - 1);
        }

        double newValue = cost * Math.Pow(1 - rate, period);
        double depreciation = oldValue - newValue;

        if (newValue < salvage)
        {
            depreciation = oldValue - salvage;
        }

        return Math.Max(depreciation, 0);
    }

    /// <summary>VDB: declining balance between two (possibly fractional) periods, switching to straight line unless no_switch is TRUE.</summary>
    private static Value Vdb(FunctionCall call)
    {
        if (call.TryNumber(0, out double cost, out Value error) == false
            || call.TryNumber(1, out double salvage, out error) == false
            || call.TryNumber(2, out double life, out error) == false
            || call.TryNumber(3, out double start, out error) == false
            || call.TryNumber(4, out double end, out error) == false
            || call.TryNumber(5, 2, out double factor, out error) == false
            || call.TryBool(6, false, out bool noSwitch, out error) == false)
        {
            return error;
        }

        if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor <= 0)
        {
            return Value.Num;
        }

        double intStart = Math.Floor(start);
        double intEnd = Math.Ceiling(end);

        if (noSwitch == true)
        {
            double sum = 0;

            for (double current = intStart + 1; current <= intEnd; current++)
            {
                double term = DecliningBalance(cost, salvage, life, current, factor);

                if (current == intStart + 1)
                {
                    term *= Math.Min(end, intStart + 1) - start;
                }
                else if (current == intEnd)
                {
                    term *= end + 1 - intEnd;
                }

                sum += term;
            }

            return Value.FromNumber(sum);
        }

        double part = 0;

        if (start != intStart)
        {
            double remaining = cost - InterVdb(cost, salvage, life, life, intStart, factor);
            part += (start - intStart) * InterVdb(remaining, salvage, life, life - intStart, 1, factor);
        }

        if (end != intEnd)
        {
            double tempStart = intEnd - 1;
            double remaining = cost - InterVdb(cost, salvage, life, life, tempStart, factor);
            part += (intEnd - end) * InterVdb(remaining, salvage, life, life - tempStart, intEnd - tempStart, factor);
        }

        double reduced = cost - InterVdb(cost, salvage, life, life, intStart, factor);

        return Value.FromNumber(InterVdb(reduced, salvage, life, life - intStart, intEnd - intStart, factor) - part);
    }

    /// <summary>The depreciation over the first <paramref name="period"/> periods, switching to straight line when that is larger.</summary>
    private static double InterVdb(double cost, double salvage, double life, double remainingLife, double period, double factor)
    {
        double sum = 0;
        double intEnd = Math.Ceiling(period);
        double value = cost - salvage;
        bool straight = false;
        double straightTerm = 0;

        for (double current = 1; current <= intEnd; current++)
        {
            double term;

            if (straight == false)
            {
                double declining = DecliningBalance(cost, salvage, life, current, factor);
                straightTerm = value / (remainingLife - (current - 1));

                if (straightTerm > declining)
                {
                    term = straightTerm;
                    straight = true;
                }
                else
                {
                    term = declining;
                    value -= declining;
                }
            }
            else
            {
                term = straightTerm;
            }

            if (current == intEnd)
            {
                term *= period + 1 - intEnd;
            }

            sum += term;
        }

        return sum;
    }

    private static Value Effect(FunctionCall call)
    {
        if (call.TryNumber(0, out double nominal, out Value error) == false || call.TryNumber(1, out double perYear, out error) == false)
        {
            return error;
        }

        double periods = Math.Truncate(perYear);

        if (nominal <= 0 || periods < 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Pow(1 + (nominal / periods), periods) - 1);
    }

    private static Value Nominal(FunctionCall call)
    {
        if (call.TryNumber(0, out double effective, out Value error) == false || call.TryNumber(1, out double perYear, out error) == false)
        {
            return error;
        }

        double periods = Math.Truncate(perYear);

        if (effective <= 0 || periods < 1)
        {
            return Value.Num;
        }

        return Value.FromNumber(periods * (Math.Pow(1 + effective, 1 / periods) - 1));
    }

    /// <summary>CUMIPMT and CUMPRINC: the interest or principal paid between two periods; type must be 0 or 1.</summary>
    private static Value Cumulative(FunctionCall call, bool interest)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double periods, out error) == false
            || call.TryNumber(2, out double present, out error) == false
            || call.TryNumber(3, out double startNumber, out error) == false
            || call.TryNumber(4, out double endNumber, out error) == false
            || call.TryNumber(5, out double type, out error) == false)
        {
            return error;
        }

        double start = Math.Truncate(startNumber);
        double end = Math.Truncate(endNumber);

        if (rate <= 0 || periods <= 0 || present <= 0 || start < 1 || end < 1 || start > end || end > periods || (type != 0 && type != 1))
        {
            return Value.Num;
        }

        int timing = (int)type;
        double payment = Payment(rate, periods, present, 0, timing);
        double sum = 0;

        for (double period = start; period <= end; period++)
        {
            double interestPart = InterestPayment(rate, period, periods, present, 0, timing);

            if (interest == true)
            {
                sum += interestPart;
            }
            else
            {
                sum += payment - interestPart;
            }
        }

        return Value.FromNumber(sum);
    }

    private static Value FvSchedule(FunctionCall call)
    {
        if (call.TryNumber(0, out double principal, out Value error) == false)
        {
            return error;
        }

        double result = principal;

        foreach (Value item in call.Values(1))
        {
            if (item.IsError == true)
            {
                return item;
            }

            if (item.Kind == ValueKind.Bool)
            {
                return Value.ValueError;
            }

            Value converted = Coerce.ToNumber(item, call.Conversion);

            if (converted.IsError == true)
            {
                return converted;
            }

            result *= 1 + converted.Number;
        }

        return Value.FromNumber(result);
    }

    private static Value PDuration(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double present, out error) == false
            || call.TryNumber(2, out double future, out error) == false)
        {
            return error;
        }

        if (rate <= 0 || present <= 0 || future <= 0)
        {
            return Value.Num;
        }

        return Value.FromNumber((Math.Log(future) - Math.Log(present)) / Math.Log(1 + rate));
    }

    private static Value Rri(FunctionCall call)
    {
        if (call.TryNumber(0, out double periods, out Value error) == false
            || call.TryNumber(1, out double present, out error) == false
            || call.TryNumber(2, out double future, out error) == false)
        {
            return error;
        }

        if (periods <= 0 || present == 0)
        {
            return Value.Num;
        }

        return Value.FromNumber(Math.Pow(future / present, 1 / periods) - 1);
    }

    private static Value IsPmt(FunctionCall call)
    {
        if (call.TryNumber(0, out double rate, out Value error) == false
            || call.TryNumber(1, out double period, out error) == false
            || call.TryNumber(2, out double periods, out error) == false
            || call.TryNumber(3, out double present, out error) == false)
        {
            return error;
        }

        if (periods == 0)
        {
            return Value.Div0;
        }

        return Value.FromNumber(present * rate * ((period / periods) - 1));
    }

    /// <summary>
    /// DOLLARDE and DOLLARFR: the digits after the point are a count of 1/fraction parts, scaled by
    /// the power of ten that holds the fraction (1.02 in sixteenths is 1 + 2/16).
    /// </summary>
    private static Value Dollar(FunctionCall call, bool toDecimal)
    {
        if (call.TryNumber(0, out double dollar, out Value error) == false || call.TryNumber(1, out double fractionNumber, out error) == false)
        {
            return error;
        }

        double fraction = Math.Truncate(fractionNumber);

        if (fraction < 0)
        {
            return Value.Num;
        }

        if (fraction == 0)
        {
            return Value.Div0;
        }

        double scale = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
        double whole = Math.Truncate(dollar);
        double part = dollar - whole;

        if (toDecimal == true)
        {
            return Value.FromNumber(whole + (part * scale / fraction));
        }

        return Value.FromNumber(whole + (part * fraction / scale));
    }

    /// <summary>Excel's type argument: 0 for payments at the end of a period, anything else for the beginning.</summary>
    private static int Timing(double type)
    {
        if (type == 0)
        {
            return 0;
        }

        return 1;
    }

    private static double Payment(double rate, double periods, double present, double future, int timing)
    {
        if (rate == 0)
        {
            return -(present + future) / periods;
        }

        double growth = Math.Pow(1 + rate, periods);

        return -(rate * (future + (present * growth))) / ((1 + (rate * timing)) * (growth - 1));
    }

    private static double FutureValue(double rate, double periods, double payment, double present, int timing)
    {
        if (rate == 0)
        {
            return -(present + (payment * periods));
        }

        double growth = Math.Pow(1 + rate, periods);

        return -((present * growth) + (payment * (1 + (rate * timing)) * (growth - 1) / rate));
    }

    private static double InterestPayment(double rate, double period, double periods, double present, double future, int timing)
    {
        double payment = Payment(rate, periods, present, future, timing);
        double balance;

        if (period == 1)
        {
            if (timing == 1)
            {
                return 0;
            }

            balance = -present;
        }
        else if (timing == 1)
        {
            balance = FutureValue(rate, period - 2, payment, present, 1) - payment;
        }
        else
        {
            balance = FutureValue(rate, period - 1, payment, present, 0);
        }

        return balance * rate;
    }
}
