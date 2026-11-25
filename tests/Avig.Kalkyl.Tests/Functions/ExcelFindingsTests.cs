namespace Avig.Kalkyl.Tests;

/// <summary>
/// Regression tests for the differences found by evaluating the formulas of the tests, and every argument of them
/// replaced by an error or by text, in Excel 365 and in Avig.Kalkyl. Every expected value is the one Excel gave.
/// </summary>
public class ExcelFindingsTests
{
    private static readonly (string Address, CellValue Value)[] Table =
    [
        ("A1", 101), ("A2", 102), ("A3", 103), ("A4", 104), ("A5", 105),
        ("B1", "a"), ("B2", "b"), ("B3", "c"), ("B4", "d"), ("B5", "e"),
        ("C1", 11), ("C2", 12), ("C3", 13), ("C4", 14), ("C5", 15),
        ("D1", 21), ("D2", 22), ("D3", 23), ("D4", 24), ("D5", 25),
    ];

    private static readonly (string Address, CellValue Value)[] Words =
    [
        ("A1", 101), ("A2", 102), ("A3", 1234),
        ("B1", "apple"), ("B2", "Banana"), ("B3", "cherry"),
        ("C1", 1), ("C2", 2), ("C3", 3), ("C4", 4),
        ("D1", true),
    ];

    [Theory]
    [InlineData("=AND(\"abc\",TRUE)", "TRUE")]
    [InlineData("=AND(\"abc\",FALSE)", "FALSE")]
    [InlineData("=AND(TRUE,1,\"abc\")", "TRUE")]
    [InlineData("=OR(\"abc\",TRUE)", "TRUE")]
    [InlineData("=OR(FALSE,\"abc\")", "FALSE")]
    [InlineData("=XOR(\"abc\",TRUE,FALSE)", "TRUE")]
    [InlineData("=AND(\"abc\")", "#VALUE!")]
    [InlineData("=AND(\"abc\",#DIV/0!)", "#DIV/0!")]
    [InlineData("=MODE(\"abc\",1,1)", "1")]
    [InlineData("=MODE(\"1\",1,2,2)", "2")]
    [InlineData("=MODE(\"abc\")", "#N/A")]
    [InlineData("=MODE.SNGL(\"abc\",1,1)", "1")]
    [InlineData("=MODE.MULT(\"abc\",1,1)", "1")]
    [InlineData("=MEDIAN(\"3\",1,2)", "2")]
    [InlineData("=MEDIAN(\"abc\",1,2)", "#VALUE!")]
    public void Logical_functions_and_MODE_skip_text_written_as_an_argument(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    [Theory]
    [InlineData("=SMALL({1,2,3},1.9)", "1")]
    [InlineData("=SMALL({1,2,3},0.5)", "#NUM!")]
    [InlineData("=LARGE({1,2,3},1.2)", "2")]
    [InlineData("=LARGE({1,2,3},0.5)", "#NUM!")]
    [InlineData("=PERMUT(0,0)", "1")]
    [InlineData("=NPER(0,0,1000,0,1)", "#DIV/0!")]
    [InlineData("=T.INV.2T(1,5)", "0")]
    [InlineData("=T.INV.2T(2,5)", "#NUM!")]
    [InlineData("=REPT(\"\",4294967296)", "#VALUE!")]
    [InlineData("=NETWORKDAYS.INTL(1,10,\"1111111\")", "0")]
    [InlineData("=NETWORKDAYS.INTL(10,1,\"1111111\",{3})", "0")]
    [InlineData("=WORKDAY.INTL(1,0,\"1111111\")", "1")]
    [InlineData("=NETWORKDAYS.INTL(1,10,#DIV/0!)", "#VALUE!")]
    [InlineData("=WORKDAY.INTL(1,5,#DIV/0!)", "#NUM!")]
    [InlineData("=WORKDAY.INTL(\"abc\",30,0)", "#VALUE!")]
    [InlineData("=WORKDAY.INTL(1,0,\"abc\")", "1")]
    [InlineData("=WORKDAY.INTL(1,0,#DIV/0!)", "1")]
    [InlineData("=WORKDAY.INTL(1,0,1,{1})", "1")]
    [InlineData("=MATCH(\"a\",\"a\",#DIV/0!)", "#VALUE!")]
    [InlineData("=MATCH(1,1,#DIV/0!)", "#DIV/0!")]
    [InlineData("=VLOOKUP(1,\"abc\",#DIV/0!,FALSE)", "#VALUE!")]
    [InlineData("=VLOOKUP(1,1,\"abc\",FALSE)", "#VALUE!")]
    [InlineData("=XNPV(0.1,\"abc\",#DIV/0!)", "#VALUE!")]
    [InlineData("=XNPV(0.1,{-100,\"x\"},#DIV/0!)", "#DIV/0!")]
    [InlineData("=YEARFRAC(#DIV/0!,2,5)", "#NUM!")]
    [InlineData("=IRR({-100,#DIV/0!,110})", "#VALUE!")]
    [InlineData("=SUMPRODUCT({1,#N/A},\"abc\")", "#N/A")]
    [InlineData("=SUMPRODUCT({1,#N/A},{1,2,3})", "#N/A")]
    [InlineData("=SUMPRODUCT({1,2},{1,2,3})", "#VALUE!")]
    public void Single_values_and_errors_are_Excels(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    [Theory]
    [InlineData("=T.INV.2T(1.5,5)", -0.7266868438004216)]
    [InlineData("=T.INV.2T(1.99,5)", -4.032142983555227)]
    [InlineData("=YEARFRAC(1,400,1)", 1.093150684931507)]
    [InlineData("=YEARFRAC(1,800,1)", 2.1890410958904107)]
    [InlineData("=YEARFRAC(DATE(1899,12,31),DATE(1900,12,31),1)", 1899.0006643024606)]
    [InlineData("=AGGREGATE(14,6,{1,#N/A,3},1)", 3)]
    [InlineData("=AGGREGATE(15,6,{3,1,2},2)", 2)]
    [InlineData("=AGGREGATE(15,6,{3,1,2},1.9)", 1)]
    [InlineData("=AGGREGATE(16,6,{1,2},0.5)", 1.5)]
    [InlineData("=AGGREGATE(17,6,{1,2,3,4},1.5)", 1.75)]
    [InlineData("=AGGREGATE(18,6,{1,2,3,4,5},0.5)", 3)]
    [InlineData("=AGGREGATE(19,6,{1,2,3,4,5,6,7},1.5)", 2)]
    public void Numbers_are_Excels(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=AGGREGATE(4,6,{1,2})", "#VALUE!")]
    [InlineData("=AGGREGATE(9,0,1,2)", "#VALUE!")]
    [InlineData("=AGGREGATE(19,6,{1,#N/A,3},0.5)", "#NUM!")]
    [InlineData("=AGGREGATE(14,6,#DIV/0!,1)", "#DIV/0!")]
    [InlineData("=AGGREGATE(0,6,#DIV/0!)", "#DIV/0!")]
    [InlineData("=AGGREGATE(4,6,A1:A3)", "3")]
    [InlineData("=AGGREGATE(4,0,A1:A3)", "#N/A")]
    public void Aggregate_takes_only_references_for_functions_1_to_13(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula, ("A1", 1), ("A2", CellValue.Error("#N/A")), ("A3", 3)).ToString());
    }

    [Theory]
    [InlineData("=VLOOKUP(102,A1:D5,5,FALSE)", "#REF!")]
    [InlineData("=VLOOKUP(102,A1:D5,0,FALSE)", "#VALUE!")]
    [InlineData("=VLOOKUP(999,A1:D5,5,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(999,A1:D5,0,FALSE)", "#N/A")]
    [InlineData("=HLOOKUP(999,A1:E1,3,FALSE)", "#N/A")]
    [InlineData("=VLOOKUP(1,\"abc\",1,FALSE)", "#VALUE!")]
    [InlineData("=VLOOKUP(1,1,1,FALSE)", "#N/A")]
    [InlineData("=HLOOKUP(1,1,1,FALSE)", "#N/A")]
    [InlineData("=MATCH(\"a\",\"a\",0)", "#VALUE!")]
    [InlineData("=MATCH(1,1,0)", "#N/A")]
    [InlineData("=MATCH(TRUE,TRUE,0)", "#N/A")]
    [InlineData("=XLOOKUP(999,A1:A5,C1:C4)", "#VALUE!")]
    [InlineData("=XLOOKUP(101,A1:A5,\"abc\")", "#VALUE!")]
    [InlineData("=XLOOKUP(999,\"abc\",C1:C5,\"none\")", "none")]
    [InlineData("=XLOOKUP(102.5,\"abc\",C1:C5,,1)", "11")]
    [InlineData("=XLOOKUP(101,A1:A5,#DIV/0!,,0,0)", "#DIV/0!")]
    [InlineData("=XLOOKUP(102.5,A1:A5,C1:C5,,-1.5)", "#VALUE!")]
    [InlineData("=XLOOKUP(102.5,A1:A5,C1:C5,,1.5)", "13")]
    [InlineData("=XLOOKUP(999,A1:A6,C1:C6,,1)", "0")]
    [InlineData("=XLOOKUP(999,A1:A6,C1:C6,\"none\",1)", "0")]
    [InlineData("=XLOOKUP(0,A1:A6,C1:C6,,-1)", "#N/A")]
    [InlineData("=XMATCH(999,A1:A6,1)", "6")]
    [InlineData("=XMATCH(999,A1:A6,1,2)", "6")]
    [InlineData("=XMATCH(6,\"abc\",1)", "1")]
    [InlineData("=XMATCH(\"abc\",{5,4,3,2,1},-1)", "1")]
    [InlineData("=XMATCH(\"abc\",{1,2,3,4,5},-1,2)", "5")]
    [InlineData("=XLOOKUP(5000,{1,2},{3,4},,1)", "#N/A")]
    public void Lookups_check_and_compare_as_Excel_does(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula, Table).ToString());
    }

    [Theory]
    [InlineData("=XLOOKUP(\"an\",B1:B3,C1:C3,,3)", "2")]
    [InlineData("=XLOOKUP(\"^b\",B1:B3,C1:C3,,3)", "#N/A")]
    [InlineData("=XLOOKUP(\"^B\",B1:B3,C1:C3,,3)", "2")]
    [InlineData("=XLOOKUP(\"(?i)^b\",B1:B3,C1:C3,,3)", "2")]
    [InlineData("=XLOOKUP(\"e$\",B1:B3,C1:C3,,3)", "1")]
    [InlineData("=XLOOKUP(\"y$\",B1:B3,C1:C3,,3,-1)", "3")]
    [InlineData("=XLOOKUP(\"a.p\",B1:B3,C1:C3,,3)", "1")]
    [InlineData("=XLOOKUP(23,A1:A3,C1:C3,,3)", "3")]
    [InlineData("=XLOOKUP(\"TRUE\",D1:D3,C1:C3,,3)", "#N/A")]
    [InlineData("=XLOOKUP(\"\",B1:B3,C1:C3,,3)", "1")]
    [InlineData("=XLOOKUP(\"x\",B1:B3,C1:C3,\"none\",3)", "none")]
    [InlineData("=XLOOKUP(\"[\",B1:B3,C1:C3,,3)", "#VALUE!")]
    [InlineData("=XLOOKUP(\"an\",B1:B3,C1:C3,,3,2)", "#VALUE!")]
    [InlineData("=XMATCH(\"an\",B1:B3,3,-1)", "2")]
    [InlineData("=XLOOKUP(\"^$\",B1:B4,C1:C4,,3)", "4")]
    [InlineData("=XLOOKUP(\"\",B1:B4,C1:C4,,3,-1)", "4")]
    public void Match_mode_3_is_a_case_sensitive_regular_expression(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula, Words).ToString());
    }

    [Fact]
    public void A_regular_expression_that_would_backtrack_forever_is_answered_at_once()
    {
        (string, CellValue)[] cells = [("A1", new string('a', 5000) + "!"), ("B1", 1)];

        Assert.Equal("#N/A", Formula.Evaluate("=XLOOKUP(\"^(a+)+$\",A1:A1,B1:B1,,3)", cells).ToString());
    }

    [Theory]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "A2=", "#NUM!")]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "A2=x", "#NUM!")]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "A2=TRUE", "#NUM!")]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "B2=40100x", "#NUM!")]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "A2=#N/A", "#N/A")]
    [InlineData("=XIRR(A1:A3,B1:B3)", "A2=x", "#VALUE!")]
    [InlineData("=XIRR(A1:A3,B1:B3)", "A2=TRUE", "#VALUE!")]
    [InlineData("=XIRR(A1:A3,B1:B3)", "B2=x", "#VALUE!")]
    [InlineData("=XIRR(A1:A3,B1:B3)", "B2=", "#NUM!")]
    [InlineData("=XIRR(A1:A3,B1:B2)", "", "#NUM!")]
    [InlineData("=XIRR(A1:A3,B1:B3,\"abc\")", "", "#VALUE!")]
    public void Xnpv_and_xirr_read_their_cells_as_Excel_does(string formula, string change, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula, Flows(change)).ToString());
    }

    [Theory]
    [InlineData("=XNPV(0.1,A1:A3,B1:B3)", "", 5.65821149956642)]
    [InlineData("=XIRR(A1:A3,B1:B3)", "", 0.253860110044479)]
    [InlineData("=XIRR(A1:A3,B1:B3)", "A2=", -0.606335703283549)]
    [InlineData("=XIRR(A1:A3,B1:B3)", "B2=40100x", 0.253860110044479)]
    public void Xirr_skips_a_blank_value_and_reads_a_date_written_as_text(string formula, string change, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Flows(change)), 6);
    }

    [Theory]
    [InlineData("=XNPV(0.1,\"x\",40000)", "#VALUE!")]
    [InlineData("=XNPV(0.1,-100,40000)", "-100")]
    [InlineData("=XNPV(0.1,\"abc\",{40000,40100})", "#VALUE!")]
    [InlineData("=XNPV(0.09,A1:A5,B1:B5)", "#VALUE!")]
    [InlineData("=XIRR(A1:A5,B1:B5)", "#NUM!")]
    public void Xnpv_and_xirr_convert_a_value_written_in_place_of_their_ranges_first(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    [Theory]
    [InlineData("=SERIESSUM(2,0,1,A1:A3)", 7)]
    [InlineData("=SERIESSUM(2,0,1,B1:B3)", 0)]
    public void Seriessum_skips_blank_coefficients(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, ("A1", 1), ("A3", 3)));
    }

    [Theory]
    [InlineData("=DSUM(A1:B2,2,{\"x\";1})", "#VALUE!")]
    [InlineData("=DSUM(A1:B2,2,5)", "#VALUE!")]
    [InlineData("=DCOUNT(A1:B2,,\"abc\")", "#VALUE!")]
    [InlineData("=DSUM(A1:B2,2,D1:D2)", "2")]
    [InlineData("=DGET(A1:B2,#DIV/0!,{\"x\";1})", "#DIV/0!")]
    public void Database_criteria_must_be_cells(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula, ("A1", "x"), ("A2", 1), ("B1", "y"), ("B2", 2), ("D1", "x"), ("D2", 1)).ToString());
    }

    [Theory]
    [InlineData("=TEXTAFTER(\"abc\",\"hood\")", "#N/A")]
    [InlineData("=TEXTAFTER(\"abc\",\"hood\",1)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",\"hood\",-1)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",\"hood\",1,,,\"none\")", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abcd\",\"hood\",1)", "#N/A")]
    [InlineData("=TEXTAFTER(\"abcde\",\"hood\",2)", "#N/A")]
    [InlineData("=TEXTAFTER(\"abcde\",\"hood\",3)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",\"x\",3)", "#N/A")]
    [InlineData("=TEXTAFTER(\"abc\",\"x\",-4)", "#VALUE!")]
    [InlineData("=TEXTAFTER(\"abc\",{\"hood\",\"b\"},1)", "c")]
    [InlineData("=TEXTAFTER(\"abc\",{\"hood\",\"zz\"},2)", "#N/A")]
    [InlineData("=TEXTBEFORE(\"abc\",\"red\",2,1)", "#VALUE!")]
    [InlineData("=TEXTBEFORE(\"abc\",\"red\",,,1)", "abc")]
    [InlineData("=TEXTBEFORE(\"ab\",\"red\",,,1)", "#N/A")]
    [InlineData("=TEXTBEFORE(\"ab\",\"red\",,,1,\"none\")", "none")]
    [InlineData("=TEXTBEFORE(\"ab\",\"rr\",,,1)", "ab")]
    [InlineData("=TEXTAFTER(\"ab\",\"rr\",,,1)", "")]
    [InlineData("=TEXTAFTER(\"ab\",\"rrr\",,,1)", "#N/A")]
    [InlineData("=TEXTBEFORE(\"ab\",\"r\",3,,1)", "#VALUE!")]
    [InlineData("=TEXTBEFORE(\"a-b\",\"-\",2,,1)", "a-b")]
    [InlineData("=TEXTAFTER(\"a-b\",\"-\",-3,,1)", "#N/A")]
    public void Textbefore_and_textafter_check_the_instance_against_the_delimiter(string formula, string expected)
    {
        Assert.Equal(expected, Formula.Evaluate(formula).ToString());
    }

    /// <summary>
    /// Cash flows of -100, 50 and 60 on the serials 40000, 40100 and 40200 in A1:B3, with one cell changed:
    /// "A2=x" puts text in A2, "A2=" empties it, "B2=40100x" writes the date as the text "40100".
    /// </summary>
    private static (string Address, CellValue Value)[] Flows(string change)
    {
        Dictionary<string, CellValue> cells = new()
        {
            ["A1"] = -100, ["A2"] = 50, ["A3"] = 60, ["B1"] = 40000, ["B2"] = 40100, ["B3"] = 40200,
        };

        if (change.Length > 0)
        {
            string[] parts = change.Split('=');
            cells.Remove(parts[0]);

            if (parts[1].Length > 0)
            {
                cells[parts[0]] = Parse(parts[1]);
            }
        }

        return [.. cells.Select(pair => (pair.Key, pair.Value))];
    }

    private static CellValue Parse(string text)
    {
        if (text == "TRUE")
        {
            return true;
        }

        if (text.StartsWith('#') == true)
        {
            return CellValue.Error(text);
        }

        if (text.EndsWith('x') == true && text.Length > 1)
        {
            return text[..^1];
        }

        return text;
    }
}
