using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A large export from the streaming writer: 100 000 order lines written a row at a time, so
/// memory stays flat however many rows there are. A styled header, frozen and filtered, banded
/// rows, dates and amounts with number formats, a formula per row, a data bar on the amount and
/// the status coloured by conditional formatting. The styles are made once and shared by every
/// row; the writer stores each distinct style only once anyway.
/// </summary>
public static class LargeExport
{
    /// <summary>The sheet with the orders.</summary>
    public const string Sheet = "Orders";

    /// <summary>How many rows the gallery writes.</summary>
    public const int DefaultRows = 100_000;

    /// <summary>The cells the gallery picture covers: the top of the sheet.</summary>
    public const string PictureRange = "A1:K28";

    private static readonly string[] Customers = ["Alba Bygg AB", "Björkdalens Gård", "Cederqvist & Co", "Dalarnas El", "Ekström Design", "Fjällbacka Marin", "Granskogen IT", "Hallands Mejeri", "Isaksson Logistik", "Järnbrukets Café"];
    private static readonly string[] Regions = ["North", "South", "East", "West"];
    private static readonly (string Name, double Price)[] Products = [("Desk lamp", 495), ("Floor lamp", 1_290), ("Pendant", 2_150), ("LED strip", 389), ("Dimmer", 549), ("Spotlight", 329)];
    private static readonly string[] Statuses = ["Delivered", "Delivered", "Delivered", "Shipped", "Shipped", "Processing", "Cancelled"];

    /// <summary>Writes the export to a file.</summary>
    public static async Task WriteAsync(string path, int rows)
    {
        Result<XlsxWriter> created = await XlsxWriter.CreateAsync(path);

        if (created.IsFailure(out string? error, out _, out XlsxWriter? writer) == true)
        {
            throw new InvalidOperationException(error);
        }

        await using (writer)
        {
            await WriteAsync(writer, rows);
        }
    }

    /// <summary>Writes the export to a writer, for a file or a stream.</summary>
    public static async Task WriteAsync(XlsxWriter writer, int rows)
    {
        SheetWriter sheet = await AddSheetAsync(writer, Sheet);
        double[] widths = [10, 12, 22, 10, 14, 8, 13, 14, 12, 13, 2];

        for (int column = 0; column < widths.Length; column++)
        {
            sheet.SetColumnWidth(column, widths[column]);
        }

        // Everything about the layout is set before the first row.
        sheet.SetView(new SheetView { ShowGridLines = false, TabColor = Navy });
        sheet.FreezePanes(1);
        sheet.EnableAutoFilter();
        sheet.AddConditionalFormat(new CellRange(1, 7, rows, 7), ConditionalFormat.DataBar(LightBlue));
        sheet.AddConditionalFormat(new CellRange(1, 9, rows, 9), ConditionalFormat.WhenTextContains("Delivered", new Style { FontColor = "#067647", Bold = true }));
        sheet.AddConditionalFormat(new CellRange(1, 9, rows, 9), ConditionalFormat.WhenTextContains("Cancelled", new Style { FontColor = Red, Bold = true }));
        sheet.AddConditionalFormat(new CellRange(1, 9, rows, 9), ConditionalFormat.WhenTextContains("Processing", new Style { FontColor = "#B45309", Bold = true }));

        string[] headers = ["Order", "Date", "Customer", "Region", "Product", "Qty", "Unit price", "Amount", "VAT", "Status"];
        List<Cell> header = [];

        foreach (string heading in headers)
        {
            // Left-aligned, so the filter buttons do not cover the headings.
            header.Add(Text(heading, 10, "#FFFFFF", bold: true).Background(Navy));
        }

        await sheet.WriteRowAsync(header);

        // One style per kind of cell and band, made once.
        Style plain = new() { FontName = FontName, FontSize = 10, FontColor = Ink };

        Style[] bands = [plain, plain with { Background = Band }];
        Random random = new(100_000);
        DateOnly first = new(2025, 1, 1);
        List<Cell> cells = new(headers.Length);

        for (int index = 0; index < rows; index++)
        {
            Style style = bands[index % 2];
            (string product, double price) = Products[random.Next(Products.Length)];
            int quantity = 1 + random.Next(12);
            int excelRow = index + 2;
            cells.Clear();
            cells.Add(Cell.Of(100_001 + index).WithStyle(style with { NumberFormat = "0" }));
            cells.Add(Cell.Of(first.AddDays(index * 700 / rows)).WithStyle(style with { NumberFormat = "yyyy-mm-dd" }));
            cells.Add(Cell.Of(Customers[random.Next(Customers.Length)]).WithStyle(style));
            cells.Add(Cell.Of(Regions[random.Next(Regions.Length)]).WithStyle(style with { FontColor = Grey }));
            cells.Add(Cell.Of(product).WithStyle(style));
            cells.Add(Cell.Of(quantity).WithStyle(style with { NumberFormat = "0" }));
            cells.Add(Cell.Of(price).WithStyle(style with { NumberFormat = "#,##0.00" }));
            cells.Add(Cell.FromFormula($"F{excelRow}*G{excelRow}", quantity * price).WithStyle(style with { NumberFormat = "#,##0.00" }));
            cells.Add(Cell.FromFormula($"H{excelRow}*0.25", quantity * price * 0.25).WithStyle(style with { NumberFormat = "#,##0.00", FontColor = Grey }));
            cells.Add(Cell.Of(Statuses[random.Next(Statuses.Length)]).WithStyle(style));
            await sheet.WriteRowAsync(cells);
        }
    }
}
