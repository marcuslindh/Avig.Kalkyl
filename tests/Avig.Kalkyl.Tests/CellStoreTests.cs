namespace Avig.Kalkyl.Tests;

/// <summary>
/// The sheet's storage (RowTable and CellRow) and the recalculator's formula index, which both
/// replaced sorted dictionaries for speed: order, lookups, removal and ranges must work as before.
/// </summary>
public class CellStoreTests
{
    private static Worksheet NewSheet()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        return sheet;
    }

    [Fact]
    public void Cells_added_in_any_order_come_back_in_order()
    {
        Worksheet sheet = NewSheet();
        (int Row, int Column)[] positions = [(5, 3), (0, 9), (5, 0), (2, 2), (0, 1), (CellAddress.MaxRows - 1, CellAddress.MaxColumns - 1), (5, 1)];

        foreach ((int row, int column) in positions)
        {
            sheet[row, column] = Cell.Of($"{row},{column}");
        }

        List<string> read = [.. sheet.Rows.SelectMany(row => row.Cells.Where(cell => cell.IsEmpty == false).Select(cell => cell.Value.ToString()))];

        Assert.Equal(["0,1", "0,9", "2,2", "5,0", "5,1", "5,3", $"{CellAddress.MaxRows - 1},{CellAddress.MaxColumns - 1}"], read);
        Assert.Equal(CellAddress.MaxRows, sheet.RowCount);
        Assert.Equal(CellAddress.MaxColumns, sheet.ColumnCount);
    }

    [Fact]
    public void A_cleared_cell_is_gone_and_the_others_stay()
    {
        Worksheet sheet = NewSheet();

        for (int column = 0; column < 10; column++)
        {
            sheet[0, column] = column;
        }

        sheet.Clear(0, 0);
        sheet.Clear(0, 5);
        sheet.Clear(0, 9);
        sheet.Clear(0, 5);

        Assert.True(sheet[0, 0].IsEmpty);
        Assert.True(sheet[0, 5].IsEmpty);
        Assert.Equal(CellValue.Number(4), sheet[0, 4].Value);
        Assert.Equal(CellValue.Number(6), sheet[0, 6].Value);
        Assert.Equal(9, sheet.ColumnCount);
        Assert.Equal(7, sheet.Rows.Single().Cells.Count(cell => cell.IsEmpty == false));
    }

    [Fact]
    public void A_replaced_cell_keeps_its_place()
    {
        Worksheet sheet = NewSheet();
        sheet[3, 1] = 1;
        sheet[3, 2] = 2;
        sheet[3, 1] = "one";

        Assert.Equal(CellValue.Text("one"), sheet[3, 1].Value);
        Assert.Equal(CellValue.Number(2), sheet[3, 2].Value);
    }

    [Fact]
    public async Task Formulas_over_whole_columns_follow_formulas_in_those_columns()
    {
        Worksheet sheet = NewSheet();

        for (int row = 0; row < 50; row++)
        {
            sheet[row, 0] = row;
            sheet[row, 1] = Cell.FromFormula($"A{row + 1}*2");
        }

        sheet[0, 2] = Cell.FromFormula("SUM(B:B)");
        sheet[1, 2] = Cell.FromFormula("SUMIF(A:A,\">=25\",B:B)");
        sheet[2, 2] = Cell.FromFormula("C1+C2");

        await sheet.Workbook.RecalculateAsync(RecalculationScope.All);

        Assert.Equal(CellValue.Number(2450), sheet["C1"].Value);
        Assert.Equal(CellValue.Number(1850), sheet["C2"].Value);
        Assert.Equal(CellValue.Number(4300), sheet["C3"].Value);

        sheet[49, 0] = 100;
        await sheet.Workbook.RecalculateAsync(RecalculationScope.Changed);

        Assert.Equal(CellValue.Number(2552), sheet["C1"].Value);
        Assert.Equal(CellValue.Number(1952), sheet["C2"].Value);
        Assert.Equal(CellValue.Number(4504), sheet["C3"].Value);
    }

    [Fact]
    public async Task Formulas_over_whole_rows_follow_formulas_in_those_rows()
    {
        Worksheet sheet = NewSheet();

        for (int column = 0; column < 30; column++)
        {
            sheet[0, column] = column;
            sheet[1, column] = Cell.FromFormula($"{new CellAddress(0, column)}+1");
        }

        sheet[5, 0] = Cell.FromFormula("SUM(2:2)");
        sheet[5, 1] = Cell.FromFormula("SUM(B2:D2)");

        await sheet.Workbook.RecalculateAsync(RecalculationScope.All);

        Assert.Equal(CellValue.Number(465), sheet["A6"].Value);
        Assert.Equal(CellValue.Number(9), sheet["B6"].Value);
    }

    [Fact]
    public async Task A_formula_that_reads_its_own_column_is_circular()
    {
        Worksheet sheet = NewSheet();
        sheet[0, 0] = 1;
        sheet[1, 0] = Cell.FromFormula("SUM(A:A)");

        RecalculationReport report = await sheet.Workbook.RecalculateAsync(RecalculationScope.All);

        Assert.Equal(["Data!A2"], report.CircularReferences);
    }
}
