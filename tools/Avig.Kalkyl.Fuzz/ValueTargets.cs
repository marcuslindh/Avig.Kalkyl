using System.Globalization;
using System.Text;
using System.Xml;
using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Fuzz;

/// <summary>
/// Invariants of the small conversions everything else is built on. None of them can crash in
/// an interesting way; they can only answer wrong, so each target states what the answer must be.
/// </summary>
internal static class ValueTargets
{
    /// <summary>Half a millisecond as a fraction of a day: <see cref="ExcelDate.TryFromSerial"/> rounds to whole milliseconds.</summary>
    private const double HalfMillisecond = 0.5 / 86_400_000;

    private static readonly double[] EdgeSerials =
    [
        0, 1, 59, 59.999999, 60, 60.5, 61, 61.000001, 1462, 2_958_465, 2_958_465.999_999_9, 2_958_466, -1, -0.0,
        0.5, 0.999999994, 0.9999999999, 45_000.000_005_787, 45_000.999_994_213, double.NaN, double.PositiveInfinity, 1e-300,
    ];

    private static readonly string[] AddressPieces =
    [
        "A", "Z", "AA", "AZ", "BA", "ZZ", "AAA", "XFD", "XFE", "ZZZ", "AAAA", "a", "xfd", "$", "$$",
        "0", "1", "01", "9", "10", "1048576", "1048577", "0001", "99999999999", "-1", "+1", " ", "١", "１", "Ａ", ":",
    ];

    public static FuzzTarget CellText()
    {
        return new FuzzTarget("cell-text", "_xHHHH_ encoding: decode(encode(s)) == s, and the result is valid XML text", CellTextAsync, GenerateText);
    }

    public static FuzzTarget ExcelDate()
    {
        return new FuzzTarget("excel-date", "serial <-> date conversions agree in both directions", ExcelDateAsync, GenerateDate);
    }

    public static FuzzTarget Address()
    {
        return new FuzzTarget("cell-address", "what CellAddress parses, it formats back to the same address", AddressAsync, GenerateAddress);
    }

    private static byte[] GenerateText(Random random)
    {
        return Inputs.Utf16(Inputs.Text(random, 16));
    }

    private static Task CellTextAsync(byte[] input)
    {
        string text = Inputs.FromUtf16(input);
        string encoded = Internal.CellText.Encode(text);

        for (int index = 0; index < encoded.Length; index++)
        {
            char current = encoded[index];

            if (char.IsHighSurrogate(current) == true && index + 1 < encoded.Length && char.IsLowSurrogate(encoded[index + 1]) == true)
            {
                index++;

                continue;
            }

            if (XmlConvert.IsXmlChar(current) == false)
            {
                throw new InvariantViolationException($"Encode left U+{(int)current:X4} at {index} in its output, which XML cannot hold: {Escape(text)}");
            }
        }

        string decoded = Internal.CellText.Decode(encoded);

        if (decoded != text)
        {
            throw new InvariantViolationException($"Decode(Encode(s)) != s: {Escape(text)} -> {Escape(encoded)} -> {Escape(decoded)}");
        }

        return Task.CompletedTask;
    }

    private static byte[] GenerateDate(Random random)
    {
        byte[] input = new byte[9];
        input[0] = (byte)random.Next(4);
        long payload;

        if ((input[0] & 2) == 0)
        {
            double serial = Inputs.Pick(random, EdgeSerials);

            if (random.Next(2) == 0)
            {
                serial = random.NextDouble() * 3_000_000;
            }

            payload = BitConverter.DoubleToInt64Bits(serial);
        }
        else
        {
            payload = random.NextInt64(DateTime.MaxValue.Ticks);
        }

        for (int index = 0; index < 8; index++)
        {
            input[index + 1] = (byte)(payload >> (56 - (index * 8)));
        }

        return input;
    }

    /// <summary>
    /// The first byte holds two flags: bit 0 is the 1904 system, bit 1 says whether the rest is
    /// a serial number (read it, write it back) or date ticks (write it, read it back).
    /// </summary>
    private static Task ExcelDateAsync(byte[] input)
    {
        ByteSource source = new(input);
        byte flags = source.NextByte();
        bool uses1904 = (flags & 1) == 1;

        if ((flags & 2) == 0)
        {
            SerialFirst(source.NextDouble(), uses1904);
        }
        else
        {
            DateFirst(source.NextLong(), uses1904);
        }

        return Task.CompletedTask;
    }

    private static void SerialFirst(double serial, bool uses1904)
    {
        if (Internal.ExcelDate.TryFromSerial(serial, uses1904, out DateTime date) == false)
        {
            return;
        }

        // Day 60 in the 1900 system is 1900-02-29, which never existed; it has no date to come back
        // from. A serial a hair below 60 is day 60 once rounded to the millisecond, as ExcelDate rounds.
        if (uses1904 == false && IsMadeUpDay(serial) == true)
        {
            return;
        }

        double back = Internal.ExcelDate.ToSerial(date, uses1904);

        if (Math.Abs(back - serial) > HalfMillisecond + (Math.Abs(serial) * 1e-15))
        {
            throw new InvariantViolationException(string.Create(
                CultureInfo.InvariantCulture,
                $"Serial {serial:R} (1904: {uses1904}) reads as {date:O}, which writes as {back:R}."));
        }

        // A later serial number is never an earlier date. The made-up day reads as 1900-02-28
        // again (on purpose, see ExcelDate), so the order breaks there and only there.
        double later = serial + (1.0 / 86_400);
        bool intoMadeUpDay = uses1904 == false && IsMadeUpDay(later) == true;

        if (intoMadeUpDay == false && Internal.ExcelDate.TryFromSerial(later, uses1904, out DateTime laterDate) == true && laterDate < date)
        {
            throw new InvariantViolationException(string.Create(
                CultureInfo.InvariantCulture,
                $"Serial {serial:R} is {date:O} but one second later, {later:R}, is {laterDate:O}."));
        }
    }

    private static bool IsMadeUpDay(double serial)
    {
        double rounded = Math.Round(serial * 86_400_000d, MidpointRounding.AwayFromZero) / 86_400_000d;

        return rounded >= 60 && rounded < 61;
    }

    private static void DateFirst(long payload, bool uses1904)
    {
        long ticks = Math.Abs(payload % DateTime.MaxValue.Ticks);
        DateTime date = new(ticks - (ticks % TimeSpan.TicksPerMillisecond), DateTimeKind.Unspecified);
        DateTime earliest = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);

        if (uses1904 == true)
        {
            earliest = new DateTime(1904, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        }

        // CellValue.Date takes dates up to the end of 9999, and the writer writes every one of them.
        if (date < earliest || date.Year > 9999)
        {
            return;
        }

        double serial = Internal.ExcelDate.ToSerial(date, uses1904);

        if (Internal.ExcelDate.TryFromSerial(serial, uses1904, out DateTime back) == false)
        {
            throw new InvariantViolationException(string.Create(
                CultureInfo.InvariantCulture,
                $"{date:O} (1904: {uses1904}) writes as serial {serial:R}, which does not read back as a date."));
        }

        if (back != date)
        {
            throw new InvariantViolationException(string.Create(
                CultureInfo.InvariantCulture,
                $"{date:O} (1904: {uses1904}) writes as serial {serial:R} and reads back as {back:O}."));
        }
    }

    private static byte[] GenerateAddress(Random random)
    {
        StringBuilder builder = new();

        if (random.Next(3) == 0)
        {
            int row = random.Next(CellAddress.MaxRows);
            int column = random.Next(CellAddress.MaxColumns);

            if (random.Next(2) == 0)
            {
                row = Inputs.Pick(random, [0, 1, 8, 9, 99, CellAddress.MaxRows - 1]);
                column = Inputs.Pick(random, [0, 25, 26, 701, 702, 16_382, CellAddress.MaxColumns - 1]);
            }

            builder.Append(new CellAddress(row, column).ToString());
        }
        else
        {
            int pieces = random.Next(1, 5);

            for (int piece = 0; piece < pieces; piece++)
            {
                builder.Append(Inputs.Pick(random, AddressPieces));
            }
        }

        return Encoding.UTF8.GetBytes(builder.ToString());
    }

    private static Task AddressAsync(byte[] input)
    {
        string text = Encoding.UTF8.GetString(input);
        bool parsed = CellAddress.TryParse(text, out CellAddress address);
        Result<CellAddress> result = CellAddress.Parse(text);

        if (result.IsSuccess(out CellAddress resultAddress) != parsed || (parsed == true && resultAddress != address))
        {
            throw new InvariantViolationException($"Parse and TryParse disagree on '{text}'.");
        }

        if (parsed == false)
        {
            return Task.CompletedTask;
        }

        if (address.Row < 0 || address.Row >= CellAddress.MaxRows || address.Column < 0 || address.Column >= CellAddress.MaxColumns)
        {
            throw new InvariantViolationException($"'{text}' parses to row {address.Row}, column {address.Column}, outside the sheet.");
        }

        string formatted = address.ToString();

        if (CellAddress.TryParse(formatted, out CellAddress again) == false || again != address)
        {
            throw new InvariantViolationException($"'{text}' parses to {formatted}, which does not parse back to the same address.");
        }

        string absolute = "$" + CellAddress.ColumnName(address.Column) + "$" + (address.Row + 1).ToString(CultureInfo.InvariantCulture);

        if (CellAddress.TryParse(absolute, out CellAddress fromAbsolute) == false || fromAbsolute != address)
        {
            throw new InvariantViolationException($"{formatted} parses but {absolute} does not give the same address.");
        }

        return Task.CompletedTask;
    }

    private static string Escape(string text)
    {
        StringBuilder builder = new("\"");

        foreach (char current in text)
        {
            if (current < ' ' || current > '~')
            {
                builder.Append(CultureInfo.InvariantCulture, $"\\u{(int)current:X4}");
            }
            else
            {
                builder.Append(current);
            }
        }

        return builder.Append('"').ToString();
    }
}
