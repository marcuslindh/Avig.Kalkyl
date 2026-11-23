namespace Avig.Kalkyl.Tests;

public class DateFunctionTests
{
    [Theory]
    [InlineData("=DATE(2008,1,1)", 39448)]
    [InlineData("=DATE(108,1,2)", 39449)]
    [InlineData("=DATE(2008,14,2)", 39846)]
    [InlineData("=DATE(2008,1,-15)", 39432)]
    [InlineData("=DATE(2008,0,1)", 39417)]
    [InlineData("=DATE(1900,2,29)", 60)]
    [InlineData("=DATE(1900,3,1)", 61)]
    [InlineData("=DATE(1900,1,0)", 0)]
    [InlineData("=DATE(9999,12,31)", 2958465)]
    [InlineData("=DAY(60)", 29)]
    [InlineData("=MONTH(60)", 2)]
    [InlineData("=YEAR(60)", 1900)]
    [InlineData("=DAY(61)", 1)]
    [InlineData("=MONTH(61)", 3)]
    [InlineData("=DAY(0)", 0)]
    [InlineData("=MONTH(0)", 1)]
    [InlineData("=YEAR(0)", 1900)]
    [InlineData("=DAY(59)", 28)]
    [InlineData("=DAY(DATE(2026,11,23)+0.75)", 23)]
    [InlineData("=MONTH(\"2026-11-23\")", 11)]
    [InlineData("=YEAR(46349)", 2026)]
    [InlineData("=DATEVALUE(\"2008-08-22\")", 39682)]
    [InlineData("=DATEVALUE(\"22-Aug-2008\")", 39682)]
    [InlineData("=DATEVALUE(\"August 22, 2008\")", 39682)]
    [InlineData("=DATEVALUE(\"2008-08-22 18:00\")", 39682)]
    [InlineData("=HOUR(0.75)", 18)]
    [InlineData("=HOUR(\"3:30:30 PM\")", 15)]
    [InlineData("=MINUTE(\"12:45\")", 45)]
    [InlineData("=SECOND(\"4:48:18 PM\")", 18)]
    [InlineData("=SECOND(TIME(1,2,3)+39448)", 3)]
    [InlineData("=HOUR(TIME(23,59,59))", 23)]
    [InlineData("=WEEKDAY(\"2008-02-14\")", 5)]
    [InlineData("=WEEKDAY(\"2008-02-14\",2)", 4)]
    [InlineData("=WEEKDAY(\"2008-02-14\",3)", 3)]
    [InlineData("=WEEKDAY(\"2008-02-14\",11)", 4)]
    [InlineData("=WEEKDAY(\"2008-02-14\",14)", 1)]
    [InlineData("=WEEKDAY(\"2008-02-14\",16)", 6)]
    [InlineData("=WEEKDAY(\"2008-02-14\",17)", 5)]
    [InlineData("=WEEKDAY(1)", 1)]
    [InlineData("=WEEKDAY(0)", 7)]
    [InlineData("=WEEKDAY(60)", 4)]
    [InlineData("=WEEKDAY(61)", 5)]
    [InlineData("=WEEKNUM(\"2012-03-09\")", 10)]
    [InlineData("=WEEKNUM(\"2012-03-09\",2)", 11)]
    [InlineData("=WEEKNUM(DATE(2026,1,1))", 1)]
    [InlineData("=WEEKNUM(DATE(2026,1,4),1)", 2)]
    [InlineData("=WEEKNUM(DATE(2026,1,4),2)", 1)]
    [InlineData("=WEEKNUM(DATE(2026,1,4),11)", 1)]
    [InlineData("=WEEKNUM(DATE(2026,1,4),17)", 2)]
    [InlineData("=WEEKNUM(DATE(2021,1,1),21)", 53)]
    [InlineData("=ISOWEEKNUM(\"2012-03-09\")", 10)]
    [InlineData("=ISOWEEKNUM(DATE(2026,11,23))", 48)]
    [InlineData("=ISOWEEKNUM(DATE(2021,1,3))", 53)]
    [InlineData("=ISOWEEKNUM(DATE(2024,12,30))", 1)]
    [InlineData("=EDATE(\"2011-01-15\",-1)", 40527)]
    [InlineData("=EDATE(DATE(2011,1,31),1)", 40602)]
    [InlineData("=EDATE(DATE(2011,1,31),13)", 40968)]
    [InlineData("=EOMONTH(DATE(2011,1,1),1)", 40602)]
    [InlineData("=EOMONTH(DATE(2011,1,1),-3)", 40482)]
    [InlineData("=EOMONTH(DATE(2012,1,15),1)", 40968)]
    [InlineData("=DAYS(\"2021-03-15\",\"2021-02-01\")", 42)]
    [InlineData("=DAYS(DATE(2021,12,31),DATE(2021,1,1))", 364)]
    [InlineData("=DAYS(DATE(2021,1,1),DATE(2021,12,31))", -364)]
    [InlineData("=DAYS360(\"2011-01-30\",\"2011-12-31\")", 330)]
    [InlineData("=DAYS360(DATE(2011,1,1),DATE(2011,1,31))", 30)]
    [InlineData("=DAYS360(DATE(2011,2,28),DATE(2011,3,31))", 30)]
    [InlineData("=DAYS360(DATE(2011,1,30),DATE(2011,2,28))", 28)]
    [InlineData("=DAYS360(DATE(2011,1,31),DATE(2011,3,31),TRUE)", 60)]
    [InlineData("=DAYS360(DATE(2011,2,28),DATE(2011,3,31),TRUE)", 32)]
    [InlineData("=DAYS360(DATE(2011,12,31),DATE(2011,1,1))", -359)]
    [InlineData("=DATEDIF(\"2001-01-01\",\"2003-01-01\",\"Y\")", 2)]
    [InlineData("=DATEDIF(\"2001-06-01\",\"2002-08-15\",\"D\")", 440)]
    [InlineData("=DATEDIF(\"2001-06-01\",\"2002-08-15\",\"YD\")", 75)]
    [InlineData("=DATEDIF(\"2001-06-01\",\"2002-08-15\",\"MD\")", 14)]
    [InlineData("=DATEDIF(\"2001-06-01\",\"2002-08-15\",\"M\")", 14)]
    [InlineData("=DATEDIF(\"2001-06-01\",\"2002-08-15\",\"ym\")", 2)]
    [InlineData("=DATEDIF(DATE(2011,1,31),DATE(2011,3,1),\"MD\")", -2)]
    [InlineData("=DATEDIF(DATE(2020,5,20),DATE(2021,5,19),\"Y\")", 0)]
    [InlineData("=DATEDIF(DATE(2020,5,20),DATE(2021,5,19),\"YD\")", 364)]
    [InlineData("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))", 110)]
    [InlineData("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))", 109)]
    [InlineData("=NETWORKDAYS(DATE(2013,3,1),DATE(2012,10,1))", -110)]
    [InlineData("=NETWORKDAYS(DATE(2026,11,21),DATE(2026,11,22))", 0)]
    [InlineData("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))", 22)]
    [InlineData("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{\"2006/1/2\",\"2006/1/16\"})", 22)]
    [InlineData("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",{\"2006/1/2\",\"2006/1/16\"})", 20)]
    [InlineData("=NETWORKDAYS.INTL(DATE(2026,11,23),DATE(2026,11,29),11)", 6)]
    [InlineData("=WORKDAY(DATE(2008,10,1),151)", 39933)]
    [InlineData("=WORKDAY(DATE(2008,10,1),-1)", 39721)]
    [InlineData("=WORKDAY(DATE(2026,11,21),0)", 46347)]
    [InlineData("=WORKDAY(DATE(2026,11,20),1)", 46349)]
    [InlineData("=WORKDAY.INTL(DATE(2012,1,1),90,11)", 41013)]
    [InlineData("=WORKDAY.INTL(DATE(2012,1,1),30,17)", 40944)]
    [InlineData("=WORKDAY.INTL(DATE(2026,11,20),1,\"0000011\")", 46349)]
    public void Date_functions_give_Excels_results(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=TIME(12,0,0)", 0.5)]
    [InlineData("=TIME(16,48,10)", 0.700115740740741)]
    [InlineData("=TIME(27,0,0)", 0.125)]
    [InlineData("=TIME(0,750,0)", 0.520833333333333)]
    [InlineData("=TIME(1,-30,0)", 0.0208333333333333)]
    [InlineData("=TIMEVALUE(\"6:35 AM\")", 0.274305555555556)]
    [InlineData("=TIMEVALUE(\"2008-08-22 18:35\")", 0.774305555555556)]
    [InlineData("=TIMEVALUE(\"2008-08-22\")", 0)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))", 0.580555555555556)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)", 0.576502732240437)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)", 0.586111111111111)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)", 0.578082191780822)]
    [InlineData("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),4)", 0.580555555555556)]
    [InlineData("=YEARFRAC(DATE(2012,7,30),DATE(2012,1,1),3)", 0.578082191780822)]
    [InlineData("=YEARFRAC(DATE(2000,1,1),DATE(2003,6,30),1)", 3.49349760438056)]
    [InlineData("=YEARFRAC(DATE(2011,6,1),DATE(2012,3,1),1)", 0.748633879781421)]
    [InlineData("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),1)", 0.997260273972603)]
    [InlineData("=YEARFRAC(DATE(2011,2,28),DATE(2012,2,29),0)", 1)]
    [InlineData("=YEARFRAC(DATE(2011,1,31),DATE(2011,3,31),4)", 0.166666666666667)]
    [InlineData("=YEARFRAC(DATE(2012,5,5),DATE(2012,5,5),1)", 0)]
    public void Fractional_date_results(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=DATE(-1,1,1)", "#NUM!")]
    [InlineData("=DATE(10000,1,1)", "#NUM!")]
    [InlineData("=DATE(1900,0,1)", "#NUM!")]
    [InlineData("=DATE(\"x\",1,1)", "#VALUE!")]
    [InlineData("=DAY(-1)", "#NUM!")]
    [InlineData("=YEAR(\"not a date\")", "#VALUE!")]
    [InlineData("=DATEVALUE(\"5\")", "#VALUE!")]
    [InlineData("=DATEVALUE(39448)", "#VALUE!")]
    [InlineData("=DATEVALUE(\"hello\")", "#VALUE!")]
    [InlineData("=TIMEVALUE(\"25 apples\")", "#VALUE!")]
    [InlineData("=HOUR(-0.5)", "#NUM!")]
    [InlineData("=TIME(0,-1,0)", "#NUM!")]
    [InlineData("=TIME(32768,0,0)", "#NUM!")]
    [InlineData("=WEEKDAY(1,4)", "#NUM!")]
    [InlineData("=WEEKNUM(1,3)", "#NUM!")]
    [InlineData("=EDATE(DATE(9999,12,1),1)", "#NUM!")]
    [InlineData("=EOMONTH(-1,1)", "#NUM!")]
    [InlineData("=DATEDIF(DATE(2003,1,1),DATE(2001,1,1),\"Y\")", "#NUM!")]
    [InlineData("=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"W\")", "#NUM!")]
    [InlineData("=NETWORKDAYS.INTL(1,10,0)", "#NUM!")]
    [InlineData("=NETWORKDAYS.INTL(1,10,\"1111111\")", "#VALUE!")]
    [InlineData("=NETWORKDAYS.INTL(1,10,\"001\")", "#VALUE!")]
    [InlineData("=NETWORKDAYS(1,10,\"holiday\")", "#VALUE!")]
    [InlineData("=WORKDAY.INTL(DATE(2012,1,1),30,0)", "#NUM!")]
    [InlineData("=WORKDAY(DATE(9999,12,30),5)", "#NUM!")]
    [InlineData("=YEARFRAC(1,2,5)", "#NUM!")]
    [InlineData("=YEARFRAC(-1,2)", "#NUM!")]
    public void Date_errors_are_Excels_errors(string formula, string error)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(error, code);
    }

    [Fact]
    public void Dates_and_holidays_can_come_from_cells()
    {
        (string, CellValue)[] cells =
        [
            ("A1", new DateTime(2008, 10, 1)),
            ("A2", new DateTime(2008, 11, 26)),
            ("A3", new DateTime(2008, 12, 4)),
            ("A4", new DateTime(2009, 1, 21)),
            ("B1", new DateTime(2012, 10, 1)),
            ("B2", new DateTime(2013, 3, 1)),
            ("C1", new DateTime(2012, 11, 22)),
            ("C2", new DateTime(2012, 12, 4)),
            ("C3", new DateTime(2013, 1, 21)),
        ];

        Assert.Equal(39938, Formula.Number("=WORKDAY(A1,151,A2:A4)", cells));
        Assert.Equal(107, Formula.Number("=NETWORKDAYS(B1,B2,C1:C3)", cells));
        Assert.Equal(107, Formula.Number("=NETWORKDAYS.INTL(B1,B2,1,C1:C5)", cells));
        Assert.Equal(2008, Formula.Number("=YEAR(A1)", cells));
        Assert.Equal(4, Formula.Number("=WEEKDAY(A1)", cells));
    }
}
