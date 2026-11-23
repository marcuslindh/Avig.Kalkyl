# Avig.Kalkyl

Read, write, edit and calculate Excel `.xlsx` files in .NET. Three ways in, one package:

- **Streaming** — `XlsxReader` and `XlsxWriter` read and write a row at a time with flat memory use.
- **Typed** — `[KalkylRow]` classes are written and read as lists; a source generator writes the
  mapping, so there is no reflection.
- **Document** — `Workbook` opens a file, lets you change any cell, recalculates formulas with a
  built-in engine of 300+ functions, and saves it with everything it does not understand (charts,
  pictures, conditional formatting, comments, macros) left as it was.

Expected failures come back as [Avig.Result](https://www.nuget.org/packages/Avig.Result) values
instead of exceptions. No dependencies besides Avig.Result. Safe for trimming and Native AOT.

```
dotnet add package Avig.Kalkyl
```

## Typed lists

```csharp
using Avig;
using Avig.Kalkyl;

[KalkylRow]
public partial class Customer
{
    [KalkylColumn("Name", Width = 30)]
    public required string Name { get; init; }

    [KalkylColumn(NumberFormat = "# ##0.00")]
    public decimal Balance { get; init; }

    public DateOnly? Since { get; init; }

    public CustomerLevel Level { get; init; }

    [KalkylIgnore]
    public string Notes { get; set; } = "";
}

Result written = await Xlsx.WriteAsync("customers.xlsx", customers, "Customers");

await foreach (Result<Customer> row in Xlsx.ReadAsync<Customer>("customers.xlsx"))
{
    if (row.IsFailure(out string? error, out _))
    {
        Console.WriteLine(error);   // Cell C7 (row 6, column 2) in 'Customers', column 'Balance', holds 'abc', which is not a number.
        continue;
    }

    row.IsSuccess(out Customer? customer);
}
```

The type must be `partial` with a public parameterless constructor. Supported property types:
`string`, `int`, `long`, `short`, `byte`, `double`, `float`, `decimal`, `bool`, `DateTime`,
`DateOnly`, `TimeOnly`, `Guid`, enums (by name), `CellValue`, and their nullable forms. A
non-nullable property is a required column; a missing required column gives one Invalid result
and nothing more, a bad cell gives an Invalid result for that row and reading goes on. Headers are
matched without regard to order, case or surrounding spaces. Numbers, dates and booleans stored
as text (`"1,5"`, `"2026-01-02"`, `"ja"`) are understood.

The generator reports what it cannot map at compile time: KALK001 (type not partial), KALK002
(property type not supported), KALK003 (no public parameterless constructor), KALK004 (nested or
generic type).

## Editing a workbook

```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsSuccess(out Workbook? book) && book.Sheet("2026").IsSuccess(out Worksheet? sheet))
{
    sheet.SetValue(4, 2, 125_000);                    // C5 — keeps the cell's formatting
    sheet["D5"] = Cell.FromFormula("C5*1.25").Bold();  // replaces value, formula and style
    sheet.SetColumnWidth(3, 18);

    RecalculationReport report = await book.RecalculateAsync();
    Result saved = await book.SaveAsync("budget.xlsx");
}
```

When a file is saved, every part the model does not own is copied byte for byte, worksheets are
copied element by element with only the cell data rewritten, and a cell that was not changed is
written from its original XML. Styles that were in the file keep their index, so borders and theme
colours survive even though `Style` does not describe them. The calculation chain is dropped and
Excel is told to recalculate on open.

## Formulas

`RecalculateAsync()` calculates the formulas affected by cells set since the last recalculation,
everything that depends on them, and volatile functions (NOW, TODAY, RAND, INDIRECT, OFFSET).
`RecalculateAsync(RecalculationScope.All)` calculates every formula. The engine builds a
dependency graph and calculates in dependency order; circular references get 0, as in Excel
without iterative calculation, and are listed in the report together with unknown functions
(which give `#NAME?`) and formulas that cannot be parsed.

```csharp
Result<CellValue> total = book.Evaluate("SUM(Budget!B2:B13)*(1+Moms)");
```

The engine follows Excel's conversions, comparison rules, error propagation, implicit
intersection, array arithmetic (`SUMPRODUCT((A1:A9>0)*B1:B9)`), defined names, shared formulas
and the 1900 and 1904 date systems. Not supported: dynamic-array spilling (an array result shows
its first value), structured table references (`Table1[Column]` gives `#REF!`), R1C1 references
in INDIRECT, and external workbooks.

## Streaming

```csharp
Result<XlsxWriter> created = await XlsxWriter.CreateAsync("report.xlsx");

if (created.IsSuccess(out XlsxWriter? writer))
{
    await using (writer)
    {
        Result<SheetWriter> added = await writer.AddSheetAsync("Report");
        added.IsSuccess(out SheetWriter? sheet);
        sheet.FreezePanes(1);
        sheet.EnableAutoFilter();

        await sheet.WriteRowAsync([Cell.Of("Name").Bold(), Cell.Of("Since").Bold()]);
        await sheet.WriteRowAsync(["Ada", new DateTime(2026, 1, 2)]);
    }
}

Result<XlsxReader> opened = await XlsxReader.OpenAsync("report.xlsx");

if (opened.IsSuccess(out XlsxReader? reader))
{
    await using (reader)
    {
        if (reader.Sheet("Report").IsSuccess(out SheetReader? sheet))
        {
            await foreach (Row row in sheet.ReadRowsAsync())
            {
                if (row.TryGetText(0, out string? name) && row.TryGetDate(1, out DateTime since)) { }
            }
        }
    }
}
```

## Addresses are 0-based

`row[0]` is column A and `sheet[1, 2]` is the cell Excel calls **C2**. `sheet["C2"]` and
`CellAddress.Parse("C2")` take Excel's notation. Error messages give both: "cell C2 (row 1, column 2)".

## Values and styles

`CellValue` holds `Text`, `Number`, `Bool`, `Date`, `Error` or nothing; read it with
`TryGetText`, `TryGetNumber`, `TryGetBool`, `TryGetDate`, `TryGetError`, or switch on `Kind`.
Excel stores dates as numbers with a date format; the reader turns them back into `DateTime`,
including the made-up 29 February 1900 and workbooks that count from 1904.

`Style` is an immutable record set with chained methods —
`Cell.Of("Total").Bold().FontColor("#C00000").Background("#FFF2CC").NumberFormat("0.00")` — and
compared by value: equal styles share one entry in the file, so styling every cell on its own never
hits Excel's limit of about 64 000 styles. When reading, font, fill, number format and alignment
are filled in.

## What goes wrong, and how you hear about it

| Situation | You get |
| --- | --- |
| File or folder missing, sheet not found | `NotFound` |
| No permission | `Forbidden` |
| File locked by another program, sheet name taken | `Conflict` |
| Not an .xlsx, an old .xls, password-protected, damaged, bad sheet name, bad cell when mapping | `Invalid` |
| A sheet damaged half-way through a stream | `InvalidDataException` when the damaged row is reached |
| Calling something in the wrong order, values Excel cannot store | `InvalidOperationException`, `ArgumentException` |

## Building

```powershell
dotnet test Avig.Kalkyl.slnx
./publish.ps1          # tests, packs, checks the package shape, copies to the local feed
./release.ps1 -DryRun
```

## Licence

MIT.
