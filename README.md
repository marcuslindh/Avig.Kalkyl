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

Text that a formula has to read as a number or a date (`"1,5"*2`, `YEAR("3/2/2013")`,
`COUNTIFS(B:B,">=31/12/2018")`) is read in `Workbook.Culture`. Like Excel, that is the culture
of the machine that calculates, unless you set it. Set it to the culture the workbook was made
in to get the results its author saw:

```csharp
book.Culture = CultureInfo.GetCultureInfo("en-GB");   // "3/2/2013" is 3 February
await book.RecalculateAsync(RecalculationScope.All);
```

ISO dates (`2026-02-01`) are read in every culture. Formatting with TEXT, DOLLAR and FIXED is
not affected by the culture.

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
| A file past one of the limits below | `Invalid` |
| A formula past the array limit | `#NUM!` in the cell |
| A cancelled or timed-out call | `OperationCanceledException` |

## Files from outside

An .xlsx is a zip file: a few kilobytes of it can unpack to gigabytes, and a small file can
describe billions of cells. So limits are on by default, high enough for any workbook people
make by hand (`KalkylLimits.Default`):

| Limit | Default | What it stops |
| --- | --- | --- |
| `MaxUncompressedBytes` | 1 GiB | The parts unpacking to more than this together. This is checked from the zip directory before anything is read, and each part is also held to the size its header gives, so a header cannot lie. A stream that is not seekable is not copied past this either. |
| `MaxCells` | 20 million | The document model holding more cells. `XlsxReader` keeps one row at a time and is not affected. |
| `MaxSharedStrings` | 10 million | More shared strings. They are read into memory when the file is opened, by `XlsxReader` too. |
| `MaxArrayCells` | 10 million | One array in a formula holding more values, for example `A1:A100000*TRANSPOSE(A1:A100000)`. A larger one gives `#NUM!`. Sums, counts and the *IF functions over whole columns read only the cells that are there, so `SUM(A:XFD)` is quick however far away the last cell is. |

```csharp
Result<Workbook> opened = await Workbook.OpenAsync(upload, KalkylLimits.Default with { MaxCells = 1_000_000 });
Result<XlsxReader> trusted = await XlsxReader.OpenAsync(path, KalkylLimits.Unlimited);
```

Time is limited with the `CancellationToken` that every call takes, `Workbook.Evaluate`
included. Cancellation is checked for every row read and every part of a formula evaluated:

```csharp
using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(10));
Result<Workbook> opened = await Workbook.OpenAsync(upload, timeout.Token);
await book.RecalculateAsync(RecalculationScope.All, timeout.Token);
```

Some protections are always on and cannot be turned off:

- XML is read without DTDs, which stops XXE and "billion laughs".
- Rows after 1,048,576 and columns after XFD make the file damaged.
- A formula may be at most 8,192 characters long, as in Excel.
- Parentheses, signs and functions may nest at most 255 levels deep, and functions inside
  functions at most 64, as in Excel. A deeper formula gives `Invalid` and does not overflow
  the stack.
- XML elements may nest at most 100 levels deep. Excel writes about ten.
- A cell holds at most 32,767 characters, and longer text in a file makes the file damaged.
- The zip directory may not have two parts with the same name, a part without a name, or a
  name that climbs out of the package (`..`, `/x`, `C:/x`). An entry that cannot be unpacked
  makes the file damaged when it is opened, not when it is saved. Two sheets may not share a part.

The package never reaches the network, starts a process, loads code or deserializes objects,
and it reads and writes only the paths you give it. DDE, external workbooks, `WEBSERVICE`,
`CELL("filename")` and `INFO` give `#NAME?` or `#REF!`. A test checks the source for this.

**Macros.** A VBA project, Excel 4.0 macro sheets and ActiveX controls are never run, but like
everything else the package does not understand they are kept when the file is saved. So
opening and saving a file does not make it safe. `Workbook.HasMacros` tells you whether there
is any, and `Workbook.RemoveMacros()` makes the next save drop the VBA project and turn the
workbook into a plain .xlsx (macro sheets and ActiveX controls cannot be removed and give `Invalid`).

**Formulas from users.** Text written with `Cell.Of` or as a value is always text, even when it
starts with `=`, `+`, `-` or `@`. Only `Cell.FromFormula` and `Worksheet.SetFormula` make
formulas, so never give them text from a user.

**Error messages** name the file path you gave, sheet names and the content of the cell that
was wrong. They are for your log, not for the person who uploaded the file. Show them a
message of your own, chosen by `ErrorKind()`: `NotFound`, `NotXlsx` (an .xls, a
password-protected file, a CSV), `Damaged`, `TooLarge` (past a `KalkylLimits` value) or
`Inaccessible` (another program holds the file). It works on a `Result`, a `Result<T>` and on
the `InvalidDataException` that `XlsxReader` throws for a damaged row.

```csharp
if (opened.IsFailure(out string? error, out _) == true)
{
    logger.LogWarning("Upload rejected: {Error}", error);
    string shown = opened.ErrorKind() switch
    {
        KalkylErrorKind.TooLarge => "The file is too large.",
        KalkylErrorKind.NotXlsx => "Upload an .xlsx file.",
        _ => "The file could not be read.",
    };
}
```

**Concurrency.** A `Workbook` or `XlsxReader` is not safe to change from several threads at
once; separate instances are independent, also with different limits. The limits hold per file,
so a service that takes uploads should also limit how many files it works on at a time, for
example with a `SemaphoreSlim`. `RecalculateAsync` calculates on a thread-pool thread and holds
it until it is done or cancelled, so give long calculations a `CancellationToken`.

## Building

```powershell
dotnet test Avig.Kalkyl.slnx
./publish.ps1          # tests, packs, checks the package shape, copies to the local feed
./release.ps1 -DryRun
dotnet run --project tools/Avig.Kalkyl.Fuzz -c Release -- all --smoke   # fuzzing, see tools/Avig.Kalkyl.Fuzz
tests/Avig.Kalkyl.Compatibility/download.ps1                            # real workbooks, compared with four other libraries
dotnet test tests/Avig.Kalkyl.Compatibility
tools/Avig.Kalkyl.Fuzz/corpus-check.ps1                                 # every fuzz input in its own process (stack, memory, hangs)
dotnet run -c Release --project tools/Avig.Kalkyl.Benchmarks -- --filter *   # baseline in docs/benchmarks/baseline.md
dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check           # files to open in Excel, see docs/excel-checklista.md
```

## Licence

MIT.
