using static Avig.Kalkyl.Samples.SampleKit;

namespace Avig.Kalkyl.Samples;

/// <summary>
/// A Swedish invoice: an SVG logo, the customer and the invoice details, line items whose
/// amounts, VAT, rounding and total are formulas, payment details and a footer with the
/// company's registration. Laid out to the width of a printed page, without gridlines. The
/// company and the customer are made up.
/// </summary>
public static class Invoice
{
    /// <summary>The sheet with the invoice.</summary>
    public const string Sheet = "Faktura";

    /// <summary>The cells the gallery picture covers.</summary>
    public const string PictureRange = "A1:H38";

    private const string Kronor = "#,##0.00 \"kr\"";
    private const int ItemHeaderRow = 12;

    private static readonly (string Description, double Quantity, string Unit, double Price)[] Items =
    [
        ("Designlampa Nord, mässing", 4, "st", 2_490),
        ("Pendel Linje 120 cm", 2, "st", 4_250),
        ("LED-list 5 m, varmvit", 6, "st", 389),
        ("Dimmer med app", 4, "st", 549),
        ("Montering", 6.5, "tim", 695),
        ("Ljusplanering", 3, "tim", 950),
        ("Frakt", 1, "st", 295),
    ];

    private const string Logo = """
        <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="128" height="128">
          <circle cx="32" cy="32" r="30" fill="#1E2A44"/>
          <circle cx="32" cy="30" r="11" fill="#F59E0B"/>
          <g stroke="#F59E0B" stroke-width="3" stroke-linecap="round">
            <line x1="32" y1="9" x2="32" y2="14"/><line x1="14" y1="30" x2="19" y2="30"/><line x1="45" y1="30" x2="50" y2="30"/>
            <line x1="19.3" y1="17.3" x2="22.8" y2="20.8"/><line x1="44.7" y1="17.3" x2="41.2" y2="20.8"/>
          </g>
          <rect x="18" y="45" width="28" height="4" rx="2" fill="#FFFFFF"/>
        </svg>
        """;

    /// <summary>Builds the workbook.</summary>
    public static Workbook Build()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = AddSheet(book, Sheet);
        sheet.View = new SheetView { ShowGridLines = false, ShowHeaders = false, TabColor = Amber, ActiveCell = new CellAddress(0, 0) };

        double[] widths = [3, 34, 8, 8, 12, 8, 17, 3];

        for (int column = 0; column < widths.Length; column++)
        {
            sheet.SetColumnWidth(column, widths[column]);
        }

        Paint(sheet, new CellRange(0, 0, 0, 7), Navy);
        WriteHeading(sheet);
        WriteItems(sheet);
        WriteTotals(sheet);
        WriteFooter(sheet);

        return book;
    }

    private static void WriteHeading(Worksheet sheet)
    {
        double left = ColumnPoints(3);
        Check(sheet.AddPicture(Svg(Logo, "Nordljus logo"), DrawingBounds.FromPoints(left, 24, 42, 42)));
        Shape company = new(ShapeKind.TextBox)
        {
            Text = new ShapeText(
            [
                new TextRun("Nordljus AB\n", new TextFont { Name = FontName, Size = 16, Bold = true, Color = Navy }),
                new TextRun("Belysning och ljusdesign", new TextFont { Name = FontName, Size = 9, Color = Grey }),
            ])
            {
                Horizontal = HorizontalAlignment.Left,
                MarginLeft = 0,
            },
        };

        Check(sheet.AddShape(company, DrawingBounds.FromPoints(left + 52, 24, 220, 42)));
        sheet["G3"] = Text("FAKTURA", 22, Navy, bold: true).Align(HorizontalAlignment.Right);

        sheet["B6"] = Text("FAKTURERAS TILL", 8, Grey, bold: true);
        sheet["B7"] = Text("Kvarnbackens Café AB", 10, Ink, bold: true);
        sheet["B8"] = Text("Att: Maria Holm");
        sheet["B9"] = Text("Storgatan 12");
        sheet["B10"] = Text("722 15 Västerås");

        (string Label, Cell Value)[] details =
        [
            ("Fakturanummer", Text("2026-1047", 10, Ink, bold: true)),
            ("Fakturadatum", Cell.Of(new DateOnly(2026, 11, 25))),
            ("Förfallodatum", Cell.Of(new DateOnly(2026, 12, 25))),
            ("Kundnummer", Number(10234, "0")),
            ("Er referens", Text("Maria Holm")),
        ];

        for (int index = 0; index < details.Length; index++)
        {
            int row = 5 + index;
            sheet[row, 4] = Text(details[index].Label, 9, Grey);
            Cell value = details[index].Value;

            if (value.Kind == CellKind.Date)
            {
                value = value.Font(FontName).FontSize(10).FontColor(Ink).NumberFormat("yyyy-mm-dd");
            }

            sheet[row, 6] = value.Align(HorizontalAlignment.Right);
        }
    }

    private static void WriteItems(Worksheet sheet)
    {
        string[] headers = ["Beskrivning", "Antal", "Enhet", "À-pris", "Moms", "Belopp"];

        for (int column = 0; column < headers.Length; column++)
        {
            Cell header = Text(headers[column], 9, "#FFFFFF", bold: true).Background(Navy);

            if (column == 1 || column >= 3)
            {
                header = header.Align(HorizontalAlignment.Right);
            }

            sheet[ItemHeaderRow, column + 1] = header;
        }

        for (int index = 0; index < Items.Length; index++)
        {
            int row = ItemHeaderRow + 1 + index;
            int excelRow = row + 1;
            sheet[row, 1] = Text(Items[index].Description);
            sheet[row, 2] = Number(Items[index].Quantity, "General");
            sheet[row, 3] = Text(Items[index].Unit, 10, Grey);
            sheet[row, 4] = Number(Items[index].Price, Kronor);
            sheet[row, 5] = Number(0.25, "0%", Grey);
            sheet[row, 6] = Formula($"C{excelRow}*E{excelRow}", Kronor);

            if (index % 2 == 1)
            {
                Paint(sheet, new CellRange(row, 1, row, 6), Band);
            }
        }
    }

    private static void WriteTotals(Worksheet sheet)
    {
        int first = ItemHeaderRow + 2;
        int last = ItemHeaderRow + 1 + Items.Length;
        int top = last + 1;

        sheet[top, 4] = Text("Summa exkl. moms", 9, Grey);
        sheet[top, 6] = Formula($"SUM(G{first}:G{last})", Kronor);
        sheet[top + 1, 4] = Text("Moms 25 %", 9, Grey);
        sheet[top + 1, 6] = Formula($"SUMPRODUCT(G{first}:G{last},F{first}:F{last})", Kronor);
        sheet[top + 2, 4] = Text("Öresavrundning", 9, Grey);
        sheet[top + 2, 6] = Formula($"ROUND(G{top + 1}+G{top + 2},0)-(G{top + 1}+G{top + 2})", Kronor, Grey);

        int total = top + 3;
        Paint(sheet, new CellRange(total, 4, total, 6), Navy);
        sheet[total, 4] = Text("ATT BETALA", 10, "#FFFFFF", bold: true).Background(Navy);
        sheet[total, 6] = Cell.FromFormula($"G{top + 1}+G{top + 2}+G{top + 3}").Font(FontName).FontSize(13).Bold().FontColor("#FFFFFF").NumberFormat(Kronor).Background(Navy);

        int payment = total + 2;
        sheet[payment, 1] = Text("BETALNING", 8, Grey, bold: true);
        (string Label, Cell Value)[] details =
        [
            ("Bankgiro", Text("5050-1055", 10, Ink, bold: true)),
            ("OCR-nummer", Cell.FromFormula("\"2026104735\"").Font(FontName).FontSize(10).Bold().FontColor(Ink)),
            ("Villkor", Text("30 dagar netto")),
            ("Dröjsmålsränta", Text("Enligt räntelagen")),
        ];

        for (int index = 0; index < details.Length; index++)
        {
            sheet[payment + 1 + index, 1] = Text(details[index].Label, 9, Grey);
            sheet[payment + 1 + index, 2] = details[index].Value;
        }

        Shape thanks = new(ShapeKind.RoundedRectangle)
        {
            Fill = Fill.Solid("#FEF3C7"),
            Outline = Outline.None,
            CornerRadius = 0.12,
            Text = new ShapeText(
            [
                new TextRun("Tack för din beställning!\n", new TextFont { Name = FontName, Size = 11, Bold = true, Color = Navy }),
                new TextRun("Ange OCR-numret när du betalar. Frågor om fakturan: faktura@nordljus.example", new TextFont { Name = FontName, Size = 8, Color = Ink }),
            ])
            {
                Horizontal = HorizontalAlignment.Left,
                MarginLeft = 10,
                MarginRight = 10,
            },
        };

        Check(sheet.AddShape(thanks, new CellRange(payment, 4, payment + 4, 6)));
    }

    private static void WriteFooter(Worksheet sheet)
    {
        int top = 34;
        string[][] columns =
        [
            ["Nordljus AB", "Ljusgatan 4", "113 49 Stockholm"],
            ["Org.nr 559876-5432", "Momsreg.nr SE559876543201", "Godkänd för F-skatt"],
            ["info@nordljus.example", "08-123 456 78", "nordljus.example"],
        ];

        int[] places = [1, 3, 6];

        for (int column = 0; column < columns.Length; column++)
        {
            for (int line = 0; line < columns[column].Length; line++)
            {
                Cell cell = Text(columns[column][line], 8, Grey, bold: line == 0 && column == 0);

                if (column == 2)
                {
                    cell = cell.Align(HorizontalAlignment.Right);
                }

                sheet[top + line, places[column]] = cell;
            }
        }

        Paint(sheet, new CellRange(top - 1, 0, top + 3, 7), Band);
    }
}
