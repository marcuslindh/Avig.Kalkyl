namespace Avig.Kalkyl.Tests;

public class EngineeringFunctionTests
{
    [Theory]
    [InlineData("=DEC2BIN(9)", "1001")]
    [InlineData("=DEC2BIN(9,4)", "1001")]
    [InlineData("=DEC2BIN(9,8)", "00001001")]
    [InlineData("=DEC2BIN(-1)", "1111111111")]
    [InlineData("=DEC2BIN(-100)", "1110011100")]
    [InlineData("=DEC2BIN(-512)", "1000000000")]
    [InlineData("=DEC2BIN(511)", "111111111")]
    [InlineData("=DEC2BIN(\"9\")", "1001")]
    [InlineData("=DEC2BIN(9.9)", "1001")]
    [InlineData("=DEC2BIN(0)", "0")]
    [InlineData("=DEC2HEX(100,4)", "0064")]
    [InlineData("=DEC2HEX(-54)", "FFFFFFFFCA")]
    [InlineData("=DEC2HEX(28)", "1C")]
    [InlineData("=DEC2HEX(549755813887)", "7FFFFFFFFF")]
    [InlineData("=DEC2OCT(58,3)", "072")]
    [InlineData("=DEC2OCT(-100)", "7777777634")]
    [InlineData("=BIN2HEX(11111011,4)", "00FB")]
    [InlineData("=BIN2HEX(1110)", "E")]
    [InlineData("=BIN2HEX(1111111111)", "FFFFFFFFFF")]
    [InlineData("=BIN2OCT(1001,3)", "011")]
    [InlineData("=BIN2OCT(1100100)", "144")]
    [InlineData("=BIN2OCT(1111111111)", "7777777777")]
    [InlineData("=HEX2BIN(\"F\",8)", "00001111")]
    [InlineData("=HEX2BIN(\"B7\")", "10110111")]
    [InlineData("=HEX2BIN(\"FFFFFFFFFF\")", "1111111111")]
    [InlineData("=HEX2OCT(\"F\",3)", "017")]
    [InlineData("=HEX2OCT(\"3B4E\")", "35516")]
    [InlineData("=HEX2OCT(\"FFFFFFFF00\")", "7777777400")]
    [InlineData("=OCT2BIN(3,3)", "011")]
    [InlineData("=OCT2BIN(7777777000)", "1000000000")]
    [InlineData("=OCT2HEX(100,4)", "0040")]
    [InlineData("=OCT2HEX(\"7777777533\")", "FFFFFFFF5B")]
    public void Base_conversions_give_Excels_text(string formula, string expected)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetText(out string? text), $"{formula} gave {value}.");
        Assert.Equal(expected, text);
    }

    [Theory]
    [InlineData("=BIN2DEC(1100100)", 100)]
    [InlineData("=BIN2DEC(1111111111)", -1)]
    [InlineData("=BIN2DEC(\"\")", 0)]
    [InlineData("=HEX2DEC(\"A5\")", 165)]
    [InlineData("=HEX2DEC(\"a5\")", 165)]
    [InlineData("=HEX2DEC(\"FFFFFFFF5B\")", -165)]
    [InlineData("=HEX2DEC(\"3DA408B9\")", 1034160313)]
    [InlineData("=OCT2DEC(54)", 44)]
    [InlineData("=OCT2DEC(7777777533)", -165)]
    [InlineData("=DELTA(5,4)", 0)]
    [InlineData("=DELTA(5,5)", 1)]
    [InlineData("=DELTA(0.5)", 0)]
    [InlineData("=DELTA(0)", 1)]
    [InlineData("=GESTEP(5,4)", 1)]
    [InlineData("=GESTEP(5,5)", 1)]
    [InlineData("=GESTEP(-4,-5)", 1)]
    [InlineData("=GESTEP(-1)", 0)]
    [InlineData("=BITAND(1,5)", 1)]
    [InlineData("=BITAND(13,25)", 9)]
    [InlineData("=BITOR(23,10)", 31)]
    [InlineData("=BITXOR(5,3)", 6)]
    [InlineData("=BITLSHIFT(4,2)", 16)]
    [InlineData("=BITLSHIFT(4,-2)", 1)]
    [InlineData("=BITRSHIFT(13,2)", 3)]
    [InlineData("=BITRSHIFT(13,-2)", 52)]
    [InlineData("=BITAND(281474976710655,281474976710655)", 281474976710655)]
    public void Engineering_functions_give_Excels_values(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 9);
    }

    [Theory]
    [InlineData("=ERF(0.745)", 0.70792892)]
    [InlineData("=ERF(1,2)", 0.15262147)]
    [InlineData("=ERF(0.5)", 0.5204998778130465)]
    [InlineData("=ERF(1)", 0.8427007929497149)]
    [InlineData("=ERF(-1)", -0.8427007929497149)]
    [InlineData("=ERF(2)", 0.9953222650189527)]
    [InlineData("=ERF(0)", 0)]
    [InlineData("=ERF(10)", 1)]
    [InlineData("=ERF.PRECISE(0.745)", 0.70792892)]
    [InlineData("=ERF.PRECISE(1)", 0.8427007929497149)]
    [InlineData("=ERFC(1)", 0.15729920705028513)]
    [InlineData("=ERFC(0.5)", 0.4795001221869535)]
    [InlineData("=ERFC(-1)", 1.8427007929497149)]
    [InlineData("=ERFC.PRECISE(1)", 0.15729920705028513)]
    public void Erf_and_erfc_give_Excels_values(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula), 8);
    }

    [Theory]
    [InlineData(2, 0.004677734981047266)]
    [InlineData(3, 2.209049699858544e-05)]
    [InlineData(4, 1.541725790028002e-08)]
    [InlineData(6, 2.1519736712498913e-17)]
    [InlineData(1.2, 0.0896860217703646)]
    public void Erfc_keeps_full_relative_precision_in_the_tail(double number, double expected)
    {
        double actual = Formula.Number("=ERFC(" + number.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")");

        Assert.True(Math.Abs(actual - expected) <= Math.Abs(expected) * 1e-14, $"ERFC({number}) gave {actual}, expected {expected}.");
    }

    [Theory]
    [InlineData("=DEC2BIN(512)", "#NUM!")]
    [InlineData("=DEC2BIN(-513)", "#NUM!")]
    [InlineData("=DEC2BIN(9,3)", "#NUM!")]
    [InlineData("=DEC2BIN(9,-1)", "#NUM!")]
    [InlineData("=DEC2BIN(9,11)", "#NUM!")]
    [InlineData("=DEC2BIN(9,\"x\")", "#VALUE!")]
    [InlineData("=DEC2BIN(\"x\")", "#VALUE!")]
    [InlineData("=DEC2OCT(536870912)", "#NUM!")]
    [InlineData("=DEC2HEX(549755813888)", "#NUM!")]
    [InlineData("=BIN2DEC(102)", "#NUM!")]
    [InlineData("=BIN2DEC(11111111111)", "#NUM!")]
    [InlineData("=BIN2DEC(TRUE)", "#VALUE!")]
    [InlineData("=HEX2DEC(\"G\")", "#NUM!")]
    [InlineData("=HEX2BIN(\"200\")", "#NUM!")]
    [InlineData("=HEX2BIN(\"FFFFFFFDFF\")", "#NUM!")]
    [InlineData("=HEX2OCT(\"FFDFFFFFFF\")", "#NUM!")]
    [InlineData("=OCT2DEC(8)", "#NUM!")]
    [InlineData("=OCT2BIN(1000)", "#NUM!")]
    [InlineData("=DELTA(\"x\",1)", "#VALUE!")]
    [InlineData("=GESTEP(\"x\")", "#VALUE!")]
    [InlineData("=BITAND(-1,5)", "#NUM!")]
    [InlineData("=BITAND(1.5,5)", "#NUM!")]
    [InlineData("=BITOR(281474976710656,1)", "#NUM!")]
    [InlineData("=BITXOR(\"x\",1)", "#VALUE!")]
    [InlineData("=BITLSHIFT(1,54)", "#NUM!")]
    [InlineData("=BITRSHIFT(1,-54)", "#NUM!")]
    [InlineData("=ERF(\"x\")", "#VALUE!")]
    public void Engineering_errors_are_Excels_errors(string formula, string expected)
    {
        CellValue value = Formula.Evaluate(formula);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(expected, code);
    }
}
