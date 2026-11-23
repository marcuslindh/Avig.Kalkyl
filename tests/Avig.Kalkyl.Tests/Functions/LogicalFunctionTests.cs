namespace Avig.Kalkyl.Tests;

public class LogicalFunctionTests
{
    private static readonly (string, CellValue)[] Cells =
    [
        ("A1", 85),
        ("A2", "b"),
    ];

    [Theory]
    [InlineData("=IFS(A1>89,\"A\",A1>79,\"B\",A1>69,\"C\")", "B")]
    [InlineData("=IFS(A1>89,\"A\",TRUE,\"other\")", "other")]
    [InlineData("=IFNA(NA(),\"missing\")", "missing")]
    [InlineData("=IFNA(\"ok\",\"missing\")", "ok")]
    [InlineData("=SWITCH(2,1,\"one\",2,\"two\",\"none\")", "two")]
    [InlineData("=SWITCH(3,1,\"one\",2,\"two\",\"none\")", "none")]
    [InlineData("=SWITCH(A2,\"a\",\"first\",\"B\",\"second\")", "second")]
    [InlineData("=SWITCH(\"1\",1,\"number\",\"text\")", "text")]
    [InlineData("=SWITCH(B9,0,\"zero\",\"other\")", "zero")]
    public void Text_results(string formula, string expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula, Cells));
    }

    [Theory]
    [InlineData("=XOR(3>0,2<9)", false)]
    [InlineData("=XOR(3>12,4>6)", false)]
    [InlineData("=XOR(3>12,4<6)", true)]
    [InlineData("=XOR(TRUE,TRUE,TRUE)", true)]
    [InlineData("=XOR({TRUE,FALSE,TRUE})", false)]
    [InlineData("=XOR(1,0)", true)]
    [InlineData("=IFS(1,TRUE)", true)]
    public void Booleans(string formula, bool expected)
    {
        Assert.Equal((CellValue)expected, Formula.Evaluate(formula, Cells));
    }

    [Theory]
    [InlineData("=IFS(A1>89,1,A1>79,2)", 2)]
    [InlineData("=IFNA(5,0)", 5)]
    [InlineData("=SWITCH(1+0,1,10,2,20)", 10)]
    public void Number_results(string formula, double expected)
    {
        Assert.Equal(expected, Formula.Number(formula, Cells), 9);
    }

    [Theory]
    [InlineData("=IFS(FALSE,1)", "#N/A")]
    [InlineData("=IFS(FALSE,1,TRUE)", "#N/A")]
    [InlineData("=IFS(\"x\",1)", "#VALUE!")]
    [InlineData("=IFS(1/0,1)", "#DIV/0!")]
    [InlineData("=XOR(\"x\")", "#VALUE!")]
    [InlineData("=XOR(1/0,TRUE)", "#DIV/0!")]
    [InlineData("=IFNA(1/0,\"x\")", "#DIV/0!")]
    [InlineData("=SWITCH(3,1,\"one\",2,\"two\")", "#N/A")]
    [InlineData("=SWITCH(1/0,1,\"one\")", "#DIV/0!")]
    public void Errors(string formula, string expected)
    {
        CellValue value = Formula.Evaluate(formula, Cells);

        Assert.True(value.TryGetError(out string? code), $"{formula} gave {value}.");
        Assert.Equal(expected, code);
    }
}
