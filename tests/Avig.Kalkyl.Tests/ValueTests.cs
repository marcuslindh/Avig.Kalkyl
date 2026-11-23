using Avig.Kalkyl.Internal;

namespace Avig.Kalkyl.Tests;

public class ValueTests
{
    [Theory]
    [InlineData("A1", 0, 0)]
    [InlineData("C2", 1, 2)]
    [InlineData("$C$2", 1, 2)]
    [InlineData("aa10", 9, 26)]
    [InlineData("ZZ1", 0, 701)]
    [InlineData("AAA1", 0, 702)]
    [InlineData("XFD1048576", 1_048_575, 16_383)]
    public void Parse_reads_A1_notation_as_zero_based(string reference, int row, int column)
    {
        Result<CellAddress> parsed = CellAddress.Parse(reference);

        Assert.True(parsed.IsSuccess(out CellAddress address));
        Assert.Equal(row, address.Row);
        Assert.Equal(column, address.Column);
        Assert.Equal(reference.Replace("$", string.Empty, StringComparison.Ordinal).ToUpperInvariant(), address.ToString());
    }

    [Theory]
    [InlineData("")]
    [InlineData("1A")]
    [InlineData("A")]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("A1B")]
    public void Parse_rejects_what_is_not_an_address(string reference)
    {
        Result<CellAddress> parsed = CellAddress.Parse(reference);

        Assert.True(parsed.IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    [InlineData(16_383, "XFD")]
    public void ColumnName_counts_like_Excel(int column, string name)
    {
        Assert.Equal(name, CellAddress.ColumnName(column));
    }

    [Fact]
    public void TryGet_only_succeeds_for_the_kind_the_value_holds()
    {
        CellValue number = 42.5;

        Assert.True(number.TryGetNumber(out double value));
        Assert.Equal(42.5, value);
        Assert.False(number.TryGetText(out _));
        Assert.False(number.TryGetDate(out _));
        Assert.Equal(CellKind.Empty, default(CellValue).Kind);
        Assert.Equal(CellKind.Empty, ((CellValue)(string?)null).Kind);
    }

    [Fact]
    public void Values_Excel_cannot_store_are_refused_when_created()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.Number(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.Text(new string('x', CellValue.MaxTextLength + 1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => CellValue.Date(new DateTime(1850, 1, 1)));
        Assert.Throws<ArgumentException>(() => CellValue.Error("#OOPS"));
        Assert.Throws<ArgumentException>(() => Cell.Of("x").Background("red"));
    }

    [Fact]
    public void Styling_returns_a_new_cell_and_equal_styles_compare_equal()
    {
        Cell plain = "Total";
        Cell styled = plain.Bold().Background("#eeeeee");

        Assert.True(plain.Style.IsDefault);
        Assert.Equal("#EEEEEE", styled.Style.Background);
        Assert.Equal(Cell.Of("Total").Background("#EEEEEE").Bold(), styled);
    }

    [Theory]
    [InlineData(1900, 1, 1, 1)]
    [InlineData(1900, 2, 28, 59)]
    [InlineData(1900, 3, 1, 61)]
    [InlineData(2026, 11, 23, 46349)]
    [InlineData(9999, 12, 31, 2958465)]
    public void Dates_use_Excels_1900_serial_numbers_including_the_leap_year_bug(int year, int month, int day, double serial)
    {
        DateTime date = new(year, month, day);

        Assert.Equal(serial, ExcelDate.ToSerial(date));
        Assert.True(ExcelDate.TryFromSerial(serial, false, out DateTime back));
        Assert.Equal(date, back);
    }

    [Fact]
    public void The_made_up_29_February_1900_reads_as_the_28th()
    {
        Assert.True(ExcelDate.TryFromSerial(60.5, false, out DateTime date));
        Assert.Equal(new DateTime(1900, 2, 28, 12, 0, 0), date);
    }

    [Fact]
    public void A_time_without_a_date_is_day_zero_plus_the_time()
    {
        Assert.True(ExcelDate.TryFromSerial(0.75, false, out DateTime date));
        Assert.Equal(new DateTime(1899, 12, 31, 18, 0, 0), date);
        Assert.Equal(0.75, ExcelDate.ToSerial(date));
    }

    [Fact]
    public void The_1904_system_counts_from_1_January_1904()
    {
        Assert.True(ExcelDate.TryFromSerial(0, true, out DateTime start));
        Assert.Equal(new DateTime(1904, 1, 1), start);
        Assert.True(ExcelDate.TryFromSerial(366, true, out DateTime later));
        Assert.Equal(new DateTime(1905, 1, 1), later);
    }

    [Fact]
    public void Times_are_rounded_to_whole_milliseconds()
    {
        DateTime noon = new(2026, 11, 23, 12, 0, 0);
        double serial = ExcelDate.ToSerial(noon) - 1e-10;

        Assert.True(ExcelDate.TryFromSerial(serial, false, out DateTime back));
        Assert.Equal(noon, back);
    }

    [Theory]
    [InlineData("yyyy-mm-dd", true)]
    [InlineData("yyyy\\-mm\\-dd", true)]
    [InlineData("[$-41D]d mmmm yyyy", true)]
    [InlineData("h:mm AM/PM", true)]
    [InlineData("dd/mm/yyyy;@", true)]
    [InlineData("General", false)]
    [InlineData("0.00", false)]
    [InlineData("#,##0 \"kr\"", false)]
    [InlineData("0.00E+00", false)]
    [InlineData("[Red]0.00", false)]
    [InlineData("[h]:mm:ss", false)]
    [InlineData("\"day\" 0", false)]
    [InlineData("@", false)]
    public void Date_formats_are_told_apart_from_number_formats(string code, bool isDate)
    {
        Assert.Equal(isDate, NumberFormats.IsDateFormat(code));
    }

    [Theory]
    [InlineData("plain", "plain")]
    [InlineData("a\u0001b", "a_x0001_b")]
    [InlineData("_x0041_", "_x005F_x0041_")]
    [InlineData("snake_case_x", "snake_case_x")]
    public void Text_XML_cannot_hold_is_escaped_the_way_Excel_does_it(string text, string encoded)
    {
        Assert.Equal(encoded, CellText.Encode(text));
        Assert.Equal(text, CellText.Decode(encoded));
    }
}
