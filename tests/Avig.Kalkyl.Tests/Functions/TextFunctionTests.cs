namespace Avig.Kalkyl.Tests;

public class TextFunctionTests
{
    private static readonly (string, CellValue)[] Cells =
    [
        ("A1", new DateTime(2026, 3, 5, 14, 7, 9)),
        ("A2", "a"),
        ("A3", 2),
        ("B1", "x"),
        ("B3", "y"),
    ];

    private static string Text(string formula)
    {
        CellValue value = Formula.Evaluate(formula, Cells);
        Assert.True(value.TryGetText(out string? text) || value.IsEmpty, $"{formula} gave {value} ({value.Kind}), not text.");

        return text ?? string.Empty;
    }

    private static string Error(string formula)
    {
        CellValue value = Formula.Evaluate(formula, Cells);
        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value} ({value.Kind}), not an error.");

        return code!;
    }

    [Theory]
    [InlineData("=CONCATENATE(\"a\",1,TRUE)", "a1TRUE")]
    [InlineData("=CONCAT(A2:B3)", "a2y")]
    [InlineData("=CONCAT(\"x\",{1,2})", "x12")]
    [InlineData("=TEXTJOIN(\", \",TRUE,\"a\",\"\",\"b\")", "a, b")]
    [InlineData("=TEXTJOIN(\", \",FALSE,\"a\",\"\",\"b\")", "a, , b")]
    [InlineData("=TEXTJOIN({\"-\",\"+\"},TRUE,\"a\",\"b\",\"c\",\"d\")", "a-b+c-d")]
    [InlineData("=TEXTJOIN(\"|\",TRUE,A2:B3)", "a|2|y")]
    [InlineData("=LEFT(\"Sale Price\",4)", "Sale")]
    [InlineData("=LEFT(\"abc\")", "a")]
    [InlineData("=LEFT(\"abc\",10)", "abc")]
    [InlineData("=RIGHT(\"Stock Number\",6)", "Number")]
    [InlineData("=RIGHT(\"abc\")", "c")]
    [InlineData("=MID(\"Fluid Flow\",1,5)", "Fluid")]
    [InlineData("=MID(\"Fluid Flow\",7,20)", "Flow")]
    [InlineData("=MID(\"Fluid Flow\",20,5)", "")]
    [InlineData("=LOWER(\"E. E. Cummings\")", "e. e. cummings")]
    [InlineData("=UPPER(\"total\")", "TOTAL")]
    [InlineData("=PROPER(\"this is a TITLE\")", "This Is A Title")]
    [InlineData("=PROPER(\"2-way street\")", "2-Way Street")]
    [InlineData("=PROPER(\"76BudGet\")", "76Budget")]
    [InlineData("=TRIM(\"  First   Quarter  Earnings \")", "First Quarter Earnings")]
    [InlineData("=TRIM(CHAR(9)&\" a \")", "\t a")]
    [InlineData("=CLEAN(CHAR(9)&\"Monthly report\"&CHAR(10))", "Monthly report")]
    [InlineData("=SUBSTITUTE(\"Sales Data\",\"Sales\",\"Cost\")", "Cost Data")]
    [InlineData("=SUBSTITUTE(\"Quarter 1, 2008\",\"1\",\"2\",1)", "Quarter 2, 2008")]
    [InlineData("=SUBSTITUTE(\"Quarter 1, 2011\",\"1\",\"2\",3)", "Quarter 1, 2012")]
    [InlineData("=SUBSTITUTE(\"abc\",\"x\",\"y\",1)", "abc")]
    [InlineData("=SUBSTITUTE(\"aAa\",\"a\",\"b\")", "bAb")]
    [InlineData("=REPLACE(\"abcdefghijk\",6,5,\"*\")", "abcde*k")]
    [InlineData("=REPLACE(\"2009\",3,2,\"10\")", "2010")]
    [InlineData("=REPLACE(\"abc\",10,1,\"d\")", "abcd")]
    [InlineData("=REPT(\"*-\",3)", "*-*-*-")]
    [InlineData("=REPT(\"a\",0)", "")]
    [InlineData("=CHAR(65)", "A")]
    [InlineData("=CHAR(128)", "\u20AC")]
    [InlineData("=CHAR(150)", "\u2013")]
    [InlineData("=CHAR(233)", "\u00E9")]
    [InlineData("=UNICHAR(66)", "B")]
    [InlineData("=UNICHAR(128512)", "\U0001F600")]
    [InlineData("=FIXED(1234.567,1)", "1,234.6")]
    [InlineData("=FIXED(1234.567,-1)", "1,230")]
    [InlineData("=FIXED(-1234.567,-1,TRUE)", "-1230")]
    [InlineData("=FIXED(44.332)", "44.33")]
    [InlineData("=FIXED(1234567.891,2,FALSE)", "1,234,567.89")]
    [InlineData("=DOLLAR(1234.567,2)", "$1,234.57")]
    [InlineData("=DOLLAR(1234.567,-2)", "$1,200")]
    [InlineData("=DOLLAR(-1234.567,4)", "($1,234.5670)")]
    [InlineData("=DOLLAR(-0.0123,4)", "($0.0123)")]
    [InlineData("=DOLLAR(99.888)", "$99.89")]
    [InlineData("=T(\"Rainfall\")", "Rainfall")]
    [InlineData("=T(19)", "")]
    [InlineData("=T(TRUE)", "")]
    [InlineData("=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\")", "Red riding ")]
    [InlineData("=TEXTBEFORE(\"Red riding hood's, red hood\",\"hood\",-1)", "Red riding hood's, red ")]
    [InlineData("=TEXTBEFORE(\"Red riding hood's, red hood\",\"red\",2,1)", "Red riding hood's, ")]
    [InlineData("=TEXTBEFORE(\"abc\",\"x\",1,0,1)", "abc")]
    [InlineData("=TEXTBEFORE(\"abc\",\"x\",,,,\"none\")", "none")]
    [InlineData("=TEXTBEFORE(\"a-b+c\",{\"+\",\"-\"})", "a")]
    [InlineData("=TEXTBEFORE(\"abc\",\"\")", "")]
    [InlineData("=TEXTBEFORE(\"abc\",\"\",-1)", "abc")]
    [InlineData("=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\")", "'s, red hood")]
    [InlineData("=TEXTAFTER(\"Red riding hood's, red hood\",\"hood\",-1)", "")]
    [InlineData("=TEXTAFTER(\"Red riding hood's, red hood\",\"HOOD\",1,1)", "'s, red hood")]
    [InlineData("=TEXTAFTER(\"a-b+c\",{\"+\",\"-\"},2)", "c")]
    [InlineData("=TEXTAFTER(\"abc\",\"x\",1,0,1)", "")]
    [InlineData("=TEXTAFTER(\"abc\",\"x\",-1,0,1)", "abc")]
    [InlineData("=TEXTAFTER(\"abc\",\"\")", "abc")]
    [InlineData("=VALUETOTEXT(\"a\")", "a")]
    [InlineData("=VALUETOTEXT(\"a\",1)", "\"a\"")]
    [InlineData("=VALUETOTEXT(1.5)", "1.5")]
    [InlineData("=VALUETOTEXT(TRUE,1)", "TRUE")]
    [InlineData("=VALUETOTEXT(NA())", "#N/A")]
    public void Text_functions(string formula, string expected)
    {
        Assert.Equal(expected, Text(formula));
    }

    [Theory]
    [InlineData("=LEN(\"Phoenix, AZ\")", 11)]
    [InlineData("=LEN(\"\")", 0)]
    [InlineData("=LEN(123.5)", 5)]
    [InlineData("=FIND(\"M\",\"Miriam McGovern\")", 1)]
    [InlineData("=FIND(\"m\",\"Miriam McGovern\")", 6)]
    [InlineData("=FIND(\"M\",\"Miriam McGovern\",3)", 8)]
    [InlineData("=FIND(\"\",\"abc\",2)", 2)]
    [InlineData("=SEARCH(\"e\",\"Statements\",6)", 7)]
    [InlineData("=SEARCH(\"margin\",\"Profit Margin\")", 8)]
    [InlineData("=SEARCH(\"p?o\",\"Profit\")", 1)]
    [InlineData("=SEARCH(\"f*t\",\"Profit\")", 4)]
    [InlineData("=SEARCH(\"~?\",\"a?b\")", 2)]
    [InlineData("=VALUE(\"$1,000\")", 1000)]
    [InlineData("=VALUE(\"16:48:00\")-VALUE(\"12:00:00\")", 0.2)]
    [InlineData("=VALUE(\"50%\")", 0.5)]
    [InlineData("=VALUE(D9)", 0)]
    [InlineData("=NUMBERVALUE(\"2.500,27\",\",\",\".\")", 2500.27)]
    [InlineData("=NUMBERVALUE(\"3.5%\")", 0.035)]
    [InlineData("=NUMBERVALUE(\"\")", 0)]
    [InlineData("=NUMBERVALUE(\" 1 000 \")", 1000)]
    [InlineData("=NUMBERVALUE(\"1,000.5\")", 1000.5)]
    [InlineData("=CODE(\"A\")", 65)]
    [InlineData("=CODE(\"Alphabet\")", 65)]
    [InlineData("=CODE(\"\u20AC\")", 128)]
    [InlineData("=CODE(\"\u00E9\")", 233)]
    [InlineData("=CODE(\"\u4E2D\")", 63)]
    [InlineData("=UNICODE(\"B\")", 66)]
    [InlineData("=UNICODE(\"\U0001F600\")", 128512)]
    public void Number_results(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Cells), 9);
    }

    [Theory]
    [InlineData("=EXACT(\"word\",\"word\")", true)]
    [InlineData("=EXACT(\"Word\",\"word\")", false)]
    [InlineData("=EXACT(1,\"1\")", true)]
    public void Exact_is_case_sensitive(string formula, bool expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula, Cells));
    }

    [Theory]
    [InlineData("=LEFT(\"abc\",-1)", "#VALUE!")]
    [InlineData("=RIGHT(\"abc\",-1)", "#VALUE!")]
    [InlineData("=MID(\"abc\",0,1)", "#VALUE!")]
    [InlineData("=MID(\"abc\",1,-1)", "#VALUE!")]
    [InlineData("=LEN(1/0)", "#DIV/0!")]
    [InlineData("=CONCATENATE(\"a\",NA())", "#N/A")]
    [InlineData("=TEXTJOIN(\",\",TRUE,\"a\",1/0)", "#DIV/0!")]
    [InlineData("=SUBSTITUTE(\"a\",\"a\",\"b\",0)", "#VALUE!")]
    [InlineData("=REPLACE(\"abc\",0,1,\"x\")", "#VALUE!")]
    [InlineData("=FIND(\"x\",\"abc\")", "#VALUE!")]
    [InlineData("=FIND(\"A\",\"abc\")", "#VALUE!")]
    [InlineData("=FIND(\"a\",\"abc\",0)", "#VALUE!")]
    [InlineData("=FIND(\"a\",\"abc\",5)", "#VALUE!")]
    [InlineData("=SEARCH(\"z\",\"abc\")", "#VALUE!")]
    [InlineData("=SEARCH(\"a\",\"abc\",0)", "#VALUE!")]
    [InlineData("=REPT(\"a\",-1)", "#VALUE!")]
    [InlineData("=REPT(\"ab\",20000)", "#VALUE!")]
    [InlineData("=VALUE(\"abc\")", "#VALUE!")]
    [InlineData("=VALUE(TRUE)", "#VALUE!")]
    [InlineData("=NUMBERVALUE(\"1.2.3\")", "#VALUE!")]
    [InlineData("=NUMBERVALUE(\"1.5,3\")", "#VALUE!")]
    [InlineData("=CHAR(0)", "#VALUE!")]
    [InlineData("=CHAR(256)", "#VALUE!")]
    [InlineData("=CODE(\"\")", "#VALUE!")]
    [InlineData("=UNICHAR(0)", "#VALUE!")]
    [InlineData("=UNICHAR(55296)", "#N/A")]
    [InlineData("=UNICODE(\"\")", "#VALUE!")]
    [InlineData("=T(NA())", "#N/A")]
    [InlineData("=FIXED(1,128)", "#VALUE!")]
    [InlineData("=TEXTBEFORE(\"Red riding hood's, red hood\",\"basket\")", "#N/A")]
    [InlineData("=TEXTBEFORE(\"Red riding hood's, red hood\",\"red\",2)", "#N/A")]
    [InlineData("=TEXTAFTER(\"abc\",\"b\",0)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",\"b\",4)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",\"b\",1,2)", "#VALUE!")]
    [InlineData("=VALUETOTEXT(1,2)", "#VALUE!")]
    [InlineData("=TEXT(-1,\"yyyy\")", "#VALUE!")]
    [InlineData("=TEXT(1/0,\"0\")", "#DIV/0!")]
    public void Errors(string formula, string expected)
    {
        Assert.Equal(expected, Error(formula));
    }

    [Theory]
    [InlineData("=TEXT(1234.567,\"$#,##0.00\")", "$1,234.57")]
    [InlineData("=TEXT(0.285,\"0.0%\")", "28.5%")]
    [InlineData("=TEXT(1234.5678,\"0.00\")", "1234.57")]
    [InlineData("=TEXT(2.675,\"0.00\")", "2.68")]
    [InlineData("=TEXT(5,\"000\")", "005")]
    [InlineData("=TEXT(1234,\"#,##0\")", "1,234")]
    [InlineData("=TEXT(1234567,\"#,##0\")", "1,234,567")]
    [InlineData("=TEXT(1234567,\"0,000\")", "1,234,567")]
    [InlineData("=TEXT(0,\"#,##0\")", "0")]
    [InlineData("=TEXT(0,\"#\")", "")]
    [InlineData("=TEXT(1234567,\"#,##0,\")", "1,235")]
    [InlineData("=TEXT(1234567,\"0.0,,\")", "1.2")]
    [InlineData("=TEXT(1234567,\"0.0,,\"\" M\"\"\")", "1.2 M")]
    [InlineData("=TEXT(0.5,\"#.00\")", ".50")]
    [InlineData("=TEXT(12.5,\".00\")", "12.50")]
    [InlineData("=TEXT(1.5,\"#.##\")", "1.5")]
    [InlineData("=TEXT(1,\"#.##\")", "1.")]
    [InlineData("=TEXT(1.5,\"0.0?\")", "1.5 ")]
    [InlineData("=TEXT(5,\"??0\")", "  5")]
    [InlineData("=TEXT(-1.5,\"0.0\")", "-1.5")]
    [InlineData("=TEXT(-5,\"0;(0)\")", "(5)")]
    [InlineData("=TEXT(5,\"0;(0)\")", "5")]
    [InlineData("=TEXT(0,\"0;-0;\"\"zero\"\"\")", "zero")]
    [InlineData("=TEXT(12200000,\"0.00E+00\")", "1.22E+07")]
    [InlineData("=TEXT(0.00012,\"0.0E-00\")", "1.2E-04")]
    [InlineData("=TEXT(12345,\"0.0E-00\")", "1.2E04")]
    [InlineData("=TEXT(12345,\"##0.0E+0\")", "12.3E+3")]
    [InlineData("=TEXT(0,\"0.00E+00\")", "0.00E+00")]
    [InlineData("=TEXT(123,\"\\$0\")", "$123")]
    [InlineData("=TEXT(0.5,\"0 \"\"units\"\"\")", "1 units")]
    [InlineData("=TEXT(1,\"[Red]0.00\")", "1.00")]
    [InlineData("=TEXT(1.5,\"[$\u20AC-407]0.00\")", "\u20AC1.50")]
    [InlineData("=TEXT(1.5,\"0.00_)\")", "1.50 ")]
    [InlineData("=TEXT(1234.5,\"General\")", "1234.5")]
    [InlineData("=TEXT(5,\"@\")", "5")]
    [InlineData("=TEXT(TRUE,\"0\")", "TRUE")]
    [InlineData("=TEXT(\"12\",\"0.00\")", "12.00")]
    [InlineData("=TEXT(\"abc\",\"0.00\")", "abc")]
    [InlineData("=TEXT(\"abc\",\"0;0;0;\"\"[\"\"@\"\"]\"\"\")", "[abc]")]
    [InlineData("=TEXT(\"abc\",\"@ \"\"x\"\"\")", "abc x")]
    [InlineData("=TEXT(A1,\"yyyy-mm-dd\")", "2026-03-05")]
    [InlineData("=TEXT(A1,\"dddd, mmmm d, yyyy\")", "Thursday, March 5, 2026")]
    [InlineData("=TEXT(A1,\"ddd d mmm yy\")", "Thu 5 Mar 26")]
    [InlineData("=TEXT(A1,\"m/d/yyyy\")", "3/5/2026")]
    [InlineData("=TEXT(A1,\"mmmmm\")", "M")]
    [InlineData("=TEXT(A1,\"dd.mm.yyyy hh:mm\")", "05.03.2026 14:07")]
    [InlineData("=TEXT(A1,\"h:mm AM/PM\")", "2:07 PM")]
    [InlineData("=TEXT(A1,\"hh:mm:ss\")", "14:07:09")]
    [InlineData("=TEXT(A1,\"h:mm:ss a/p\")", "2:07:09 p")]
    [InlineData("=TEXT(A1,\"mm:ss\")", "07:09")]
    [InlineData("=TEXT(0.75,\"h AM/PM\")", "6 PM")]
    [InlineData("=TEXT(0.25,\"hh:mm AM/PM\")", "06:00 AM")]
    [InlineData("=TEXT(0,\"h:mm AM/PM\")", "12:00 AM")]
    [InlineData("=TEXT(1.5,\"[h]:mm\")", "36:00")]
    [InlineData("=TEXT(0.5,\"[mm]:ss\")", "720:00")]
    [InlineData("=TEXT(0.5,\"hh:mm:ss.00\")", "12:00:00.00")]
    [InlineData("=TEXT(0.000005,\"ss.000\")", "00.432")]
    [InlineData("=TEXT(0.99999999,\"hh:mm:ss\")", "00:00:00")]
    [InlineData("=TEXT(0,\"hh:mm\")", "00:00")]
    [InlineData("=TEXT(60,\"yyyy-mm-dd\")", "1900-02-29")]
    [InlineData("=TEXT(1,\"dddd\")", "Sunday")]
    [InlineData("=TEXT(61,\"dddd d mmmm\")", "Thursday 1 March")]
    [InlineData("=TEXT(\"2026-03-05\",\"d mmm\")", "5 Mar")]
    public void Text_applies_number_formats(string formula, string expected)
    {
        Assert.Equal(expected, Text(formula));
    }
}
