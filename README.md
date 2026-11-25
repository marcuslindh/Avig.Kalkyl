# Avig.Kalkyl

Read, write, edit and calculate Excel `.xlsx` files in .NET. Three ways in, one package:

- **Typed lists** — `[KalkylRow]` classes are written and read as lists; a source generator writes
  the mapping, so there is no reflection.
- **Streaming** — `XlsxReader` and `XlsxWriter` read and write a row at a time with flat memory
  use, however large the file.
- **Document** — `Workbook` opens a file, lets you change any cell, recalculates formulas with a
  built-in engine of 300+ functions, and saves it with everything it does not understand (pivot
  tables, comments, macros, other programs' extensions) left as it was.

Charts, pictures, shapes, conditional formatting, sheet views and hyperlinks can be added with
the streaming writer and the document model, also to files that already have their own, so a
dashboard like Excel's templates can be built from code. Expected failures come back as
[Avig.Result](https://www.nuget.org/packages/Avig.Result) values instead of exceptions. No
dependencies besides Avig.Result. Safe for trimming and Native AOT.

```
dotnet add package Avig.Kalkyl
```

Every example below is a test in `tests/Avig.Kalkyl.Tests/ReadmeExamples.cs` that runs on every
build, and a test checks that the code here is the same as there. They need `using Avig;` and
`using Avig.Kalkyl;`, and follow this repository's style (explicit types, `== true`), which your
code does not have to.

## Which way in?

| You want to | Use | Scenario |
| --- | --- | --- |
| Export a list of objects, or import one | `Xlsx.WriteAsync`, `Xlsx.ReadAsync` | [1](#1-export-and-import-a-list-of-objects) |
| Several typed lists in one file, with your own sheet settings | `SheetWriter.WriteAsync`, `SheetReader.ReadAsync` | [2](#2-typed-lists-in-a-file-you-build-yourself) |
| Write a large report, row by row | `XlsxWriter`, `SheetWriter` | [4](#4-write-a-large-report-row-by-row) |
| Send a file to an HTTP response | `XlsxWriter.CreateAsync(stream)` | [5](#5-write-to-a-stream) |
| Read a large file, row by row | `XlsxReader`, `SheetReader`, `Row` | [6](#6-read-a-large-file-row-by-row) |
| Take files that users upload | `KalkylLimits`, `ErrorKind()` | [7](#7-take-files-that-users-upload) |
| Create a workbook in memory, with formulas | `Workbook.Create`, `Worksheet` | [11](#11-create-a-workbook-in-memory) |
| Change cells in an existing file and save it | `Workbook.OpenAsync`, `Worksheet` | [12](#12-edit-an-existing-workbook), [13](#13-open-from-and-save-to-a-stream) |
| Calculate formulas, or one formula on its own | `RecalculateAsync`, `Evaluate` | [14](#14-calculate-formulas) |
| Add charts, pictures or conditional formatting | `AddChart`, `AddPicture`, `AddConditionalFormat` | [15](#15-charts)–[17](#17-conditional-formatting) |
| Build a dashboard: cards, buttons that link to sheets, SVG icons, styled charts | `AddShape`, `AddGroup`, `Picture.FromSvg`, `Chart` options, `SheetView` | [20](#20-shapes-groups-and-dashboards)–[22](#22-sheet-view-and-cell-links) |
| Find macros, links and connections in a file from outside, and take them out | `Workbook.ActiveContent`, `RemoveActiveContent` | [19](#19-active-content) |

## 0. Results instead of exceptions

Anything that can go wrong for a reason outside your code (a missing file, a damaged upload, a
sheet name Excel refuses) returns a `Result` or `Result<T>`. `IsFailure(out error, out status,
out value)` gives either the error and its `ResultStatus`, or the value, in one call.

<!-- example: result-pattern -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

// One call gives either the error and its status, or the value.
if (opened.IsFailure(out string? error, out ResultStatus status, out Workbook? book) == true)
{
    Console.WriteLine($"{status}: {error}");   // NotFound: 'budget.xlsx' does not exist.

    return;
}

// When only success matters, or only the value:
bool saved = (await book.SaveAsync("copy.xlsx")).IsSuccess(out _, out _);

if (book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == true)
{
    Console.WriteLine($"{sheet.Name}, saved: {saved}");
}
```

## 1. Export and import a list of objects

Mark a `partial` class with `[KalkylRow]`. The source generator writes the code that turns an
object into a row and a row into an object, at compile time.

<!-- example: typed-row-type -->
```csharp
public enum CustomerLevel
{
    Basic,
    Gold,
}

[KalkylRow]
public partial class Customer
{
    [KalkylColumn("Name", Width = 30, Order = 0)]
    public required string Name { get; init; }

    [KalkylColumn(NumberFormat = "# ##0.00")]
    public decimal Balance { get; init; }

    public DateOnly? Since { get; init; }

    public CustomerLevel Level { get; init; }

    [KalkylIgnore]
    public string Notes { get; set; } = "";
}
```

<!-- example: typed-write-read -->
```csharp
List<Customer> customers =
[
    new() { Name = "Ada Lovelace", Balance = 1234.5m, Since = new DateOnly(2024, 3, 1), Level = CustomerLevel.Gold },
    new() { Name = "Alan Turing", Balance = -20m, Level = CustomerLevel.Basic },
];

Result written = await Xlsx.WriteAsync("customers.xlsx", customers, "Customers");

await foreach (Result<Customer> row in Xlsx.ReadAsync<Customer>("customers.xlsx", "Customers"))
{
    if (row.IsFailure(out string? error, out _, out Customer? customer) == true)
    {
        // Cell B3 (row 2, column 1) in 'Customers', column 'Balance', holds 'abc', which is not a number.
        Console.WriteLine(error);

        continue;
    }

    Console.WriteLine($"{customer.Name}: {customer.Balance} ({customer.Level})");
}
```

- The type must be `partial` and have a public parameterless constructor. It can be a class or a
  record.
- Supported property types: `string`, `int`, `long`, `short`, `byte`, `double`, `float`,
  `decimal`, `bool`, `DateTime`, `DateOnly`, `TimeOnly`, `Guid`, enums (by name), `CellValue`,
  and their nullable forms.
- `[KalkylColumn("Header")]` sets the header; `Order`, `NumberFormat` and `Width` set the
  position, the format and the width when writing. `[KalkylIgnore]` leaves a property out.
- A non-nullable property is a required column. A missing required column gives one Invalid
  result and nothing more; a bad cell gives an Invalid result for that row, and reading goes on.
- Headers are matched without regard to order, case or surrounding spaces. Numbers, dates and
  booleans stored as text (`"1,5"`, `"2026-01-02"`, `"ja"`) are understood.
- The generator reports what it cannot map at compile time: KALK001 (type not partial), KALK002
  (property type not supported), KALK003 (no public parameterless constructor), KALK004 (nested
  or generic type).

## 2. Typed lists in a file you build yourself

`WriteAsync` and `ReadAsync` are also extension methods on a `SheetWriter` and a `SheetReader`,
so a typed list can share a file with other sheets, and get frozen panes or a filter. Both take
an `IEnumerable<T>`, and `WriteAsync` also an `IAsyncEnumerable<T>`, which is written as it
arrives.

<!-- example: typed-streaming -->
```csharp
if ((await XlsxWriter.CreateAsync("export.xlsx")).IsSuccess(out XlsxWriter? writer) == true)
{
    await using (writer)
    {
        // A sheet with settings of its own, then the typed rows: a header row and one row per item.
        if ((await writer.AddSheetAsync("Customers")).IsSuccess(out SheetWriter? sheet) == true)
        {
            sheet.FreezePanes(1);
            sheet.EnableAutoFilter();
            await sheet.WriteAsync(LoadCustomers());
        }

        // An IAsyncEnumerable is written as it arrives, for example straight from a database.
        if ((await writer.AddSheetAsync("Stream")).IsSuccess(out SheetWriter? streamed) == true)
        {
            await streamed.WriteAsync(LoadCustomersAsync());
        }
    }
}

if ((await XlsxReader.OpenAsync("export.xlsx")).IsSuccess(out XlsxReader? reader) == true)
{
    await using (reader)
    {
        if (reader.Sheet("Stream").IsSuccess(out SheetReader? sheet) == true)
        {
            await foreach (Result<Customer> row in sheet.ReadAsync<Customer>())
            {
                if (row.IsSuccess(out Customer? customer) == true)
                {
                    Console.WriteLine(customer.Name);
                }
            }
        }
    }
}
```

## 3. Look at a mapping

The generated mapping is `Customer.KalkylMap`: its columns, and the conversion of an object to
cells. (`KalkylCells` and `MappedRow` are used by the generated code and hidden from IntelliSense.)

<!-- example: typed-map -->
```csharp
KalkylMap<Customer> map = Customer.KalkylMap;

foreach (KalkylColumnInfo column in map.Columns)
{
    Console.WriteLine($"{column.Header}: required {column.Required}, format {column.NumberFormat}, width {column.Width}");
}

IReadOnlyList<Cell> cells = map.ToCells(new Customer { Name = "Ada", Balance = 10m });
```

## 4. Write a large report row by row

`XlsxWriter` writes from the top down: add a sheet, write its rows, add the next sheet. Memory
use stays flat however many rows are written. Column widths and frozen panes come before the
first row, since they go at the top of the sheet in the file; the filter, charts, pictures and
rules can be added at any time before the sheet is finished.

<!-- example: stream-write -->
```csharp
Result<XlsxWriter> created = await XlsxWriter.CreateAsync("report.xlsx");

if (created.IsFailure(out string? error, out _, out XlsxWriter? writer) == true)
{
    Console.WriteLine(error);   // the folder is missing, the file is open in Excel ...

    return;
}

await using (writer)
{
    // A name Excel refuses (too long, one of : \ / ? * [ ], already taken) is Invalid or Conflict.
    if ((await writer.AddSheetAsync("Sales")).IsFailure(out string? refused, out _, out SheetWriter? sheet) == true)
    {
        Console.WriteLine(refused);

        return;
    }

    // Widths and frozen panes go at the top of the sheet in the file: set them before the first row.
    sheet.SetColumnWidth(0, 24);
    sheet.SetColumnWidth(2, 14);
    sheet.FreezePanes(1);
    sheet.EnableAutoFilter();

    await sheet.WriteRowAsync([Cell.Of("Product").Bold().Background("#DDEBF7"), Cell.Of("Units").Bold(), Cell.Of("Sold on").Bold(), Cell.Of("Paid").Bold()]);

    for (int index = 0; index < 10_000; index++)
    {
        await sheet.WriteRowAsync(["Product " + index, index % 17, new DateTime(2026, 1, 1).AddDays(index % 365), index % 3 != 0]);
    }

    int last = sheet.RowCount;
    await sheet.WriteRowAsync([Cell.Of("Total").Bold(), Cell.FromFormula($"SUM(B2:B{last})").Bold()]);

    // Adding the next sheet finishes this one; it cannot be written to afterwards.
    if ((await writer.AddSheetAsync("Notes")).IsSuccess(out SheetWriter? notes) == true)
    {
        await notes.WriteRowAsync(["Written by Avig.Kalkyl", DateTime.UtcNow]);
    }

    Console.WriteLine($"{writer.Sheets.Count} sheets, '{sheet.Name}' (index {sheet.Index}) finished: {sheet.IsFinished}");
}
```

## 5. Write to a stream

<!-- example: stream-write-to-stream -->
```csharp
// Any writable stream, also one that cannot seek, such as an HTTP response body.
// leaveOpen: true, so that the response is not closed with the writer.
await using XlsxWriter writer = await XlsxWriter.CreateAsync(response, leaveOpen: true);

if ((await writer.AddSheetAsync("Export")).IsSuccess(out SheetWriter? sheet) == true)
{
    await sheet.WriteRowAsync(["Exported", DateTime.UtcNow]);
}

await writer.CompleteAsync();   // or let `await using` complete it
```

## 6. Read a large file row by row

`XlsxReader` keeps one row in memory at a time. A row is a snapshot: `row[column]` is a `Cell`,
an empty one past the last cell, and the `TryGet…` methods convert as Excel would show the value.
Empty rows in the file are skipped; `Row.Index` is the row's real position.

<!-- example: stream-read -->
```csharp
Result<XlsxReader> opened = await XlsxReader.OpenAsync("big.xlsx");

if (opened.IsFailure(out string? error, out _, out XlsxReader? reader) == true)
{
    Console.WriteLine(error);

    return;
}

await using (reader)
{
    foreach (SheetInfo info in reader.Sheets)
    {
        Console.WriteLine($"{info.Index}: {info.Name} ({info.Visibility})");
    }

    Console.WriteLine($"Dates count from 1904: {reader.Uses1904Dates}");

    if (reader.Sheet("Sales").IsSuccess(out SheetReader? sheet) == true)   // or reader.Sheet(0)
    {
        await foreach (Row row in sheet.ReadRowsAsync())
        {
            if (row.Index == 0 || row.IsEmpty == true)
            {
                continue;   // the header, and rows that are there but hold nothing
            }

            if (row.TryGetText(0, out string? product) == true
                && row.TryGetNumber(1, out double units) == true
                && row.TryGetDate(2, out DateTime soldOn) == true
                && row.TryGetBool(3, out bool paid) == true)
            {
                Console.WriteLine($"{product}: {units} on {soldOn:yyyy-MM-dd}, paid {paid}");
            }

            Cell fifth = row[4];   // beyond the last cell: an empty cell
            Console.WriteLine($"{row.ColumnCount} columns, {row.Cells.Count} cells; E is empty: {fifth.IsEmpty}");
        }
    }
}
```

## 7. Take files that users upload

An .xlsx is a zip file: a few kilobytes can unpack to gigabytes, and a small file can describe
billions of cells. Limits are on by default and fit any workbook people make by hand; lower them
for uploads, and give the work a time limit with the `CancellationToken` every call takes. The
error message is for your log; `ErrorKind()` chooses the message for the user.

<!-- example: upload -->
```csharp
// The default limits fit files people make by hand; lower them for uploads, and limit the time too.
KalkylLimits limits = KalkylLimits.Default with { MaxCells = 1_000_000, MaxUncompressedBytes = 100_000_000 };

using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(10));

Result<XlsxReader> opened = await XlsxReader.OpenAsync(upload, limits, leaveOpen: true, timeout.Token);

if (opened.IsFailure(out string? error, out _, out XlsxReader? reader) == true)
{
    // The error names paths and cell contents: log it, and show the user a message of your own.
    Console.WriteLine($"Upload rejected: {error}");

    return opened.ErrorKind() switch
    {
        KalkylErrorKind.TooLarge => "The file is too large.",
        KalkylErrorKind.NotXlsx => "Upload an .xlsx file.",
        KalkylErrorKind.Damaged => "The file is damaged.",
        _ => "The file could not be read.",
    };
}

await using (reader)
{
    if (reader.Sheet(0).IsSuccess(out SheetReader? sheet) == false)
    {
        return "The file has no sheets.";
    }

    try
    {
        await foreach (Row row in sheet.ReadRowsAsync(timeout.Token))
        {
            Console.WriteLine(row.Cells.Count);
        }
    }
    catch (InvalidDataException damaged)
    {
        // A row damaged half-way through the file is found when it is reached.
        Console.WriteLine($"{damaged.ErrorKind()}: {damaged.Message}");

        return "The file is damaged.";
    }
}

// The document model takes the same limits and keeps them for its formulas: Workbook.Limits.
// A file you made yourself can be opened without them.
if ((await Workbook.OpenAsync("upload.xlsx", KalkylLimits.Unlimited)).IsSuccess(out Workbook? book) == true)
{
    Console.WriteLine(book.Limits.MaxArrayCells);
}

return "ok";
```

| Limit | Default | What it stops |
| --- | --- | --- |
| `MaxUncompressedBytes` | 1 GiB | The parts unpacking to more than this together. This is checked from the zip directory before anything is read, and each part is also held to the size its header gives, so a header cannot lie. A stream that is not seekable is not copied past this either. |
| `MaxCells` | 20 million | The document model holding more cells. `XlsxReader` keeps one row at a time and is not affected. |
| `MaxSharedStrings` | 10 million | More shared strings. They are read into memory when the file is opened, by `XlsxReader` too. |
| `MaxArrayCells` | 10 million | One array in a formula holding more values, for example `A1:A100000*TRANSPOSE(A1:A100000)`. A larger one gives `#NUM!`. Sums, counts and the *IF functions over whole columns read only the cells that are there, so `SUM(A:XFD)` is quick however far away the last cell is. |
| `StrictZip` | on | A zip file that says two things: a local header that differs from the zip directory, entries that overlap, or bytes before, between or after the entries. Such a file can show a virus scanner one workbook and Excel another, and no spreadsheet program writes one. It is damaged. Also on in `Unlimited`. |

`KalkylLimits.Unlimited` turns them off for files you made yourself. More on what is always
checked under [Files from outside](#files-from-outside).

## 8. Values

`CellValue` is one of `Text`, `Number`, `Bool`, `Date`, `Error` or nothing. Excel stores dates as
numbers with a date format; the reader turns them back into `DateTime`, including the made-up
29 February 1900 and workbooks that count from 1904.

<!-- example: values -->
```csharp
CellValue text = "Ada";                          // implicit from string, double, int, long, decimal, bool, DateTime, DateOnly
CellValue number = CellValue.Number(42.5);
CellValue date = new DateOnly(2026, 11, 24);
CellValue error = CellValue.Error("#N/A");
CellValue nothing = CellValue.Empty;

foreach (CellValue value in new[] { text, number, date, error, nothing, CellValue.Bool(true), CellValue.Text("x"), CellValue.Date(DateTime.UtcNow.Date) })
{
    string kind = value.Kind switch
    {
        CellKind.Text => "text",
        CellKind.Number => "number",
        CellKind.Bool => "true or false",
        CellKind.Date => "date",
        CellKind.Error => "error",
        _ => "empty",
    };

    Console.WriteLine($"{value} is {kind}");
}

if (number.TryGetNumber(out double amount) == true && date.TryGetDate(out DateTime day) == true && error.TryGetError(out string? code) == true)
{
    Console.WriteLine($"{amount}, {day:d}, {code}");
}

if (text.TryGetText(out string? name) == true && CellValue.Bool(false).TryGetBool(out bool flag) == true)
{
    Console.WriteLine($"{name} {flag}");
}

Console.WriteLine($"A cell holds at most {CellValue.MaxTextLength} characters; empty: {nothing.IsEmpty}");
```

## 9. Cells and styles

A `Cell` is a value, an optional formula and a `Style`. `Style` is an immutable record compared
by value: equal styles share one entry in the file, so styling every cell on its own never hits
Excel's limit of about 64 000 styles. When reading, font, fill, number format and alignment are
filled in. Text written with `Cell.Of` or as a value is always text, even when it starts with `=`;
only `Cell.FromFormula` and `Worksheet.SetFormula` make formulas, so never give them text from a
user.

<!-- example: styles -->
```csharp
Cell plain = "Total";                                   // implicit, like CellValue
Cell styled = Cell.Of(1234.5)
    .Bold()
    .Italic()
    .Underline()
    .Font("Arial")
    .FontSize(12)
    .FontColor("#C00000")
    .Background("#FFF2CC")
    .NumberFormat("# ##0.00 \"kr\"")
    .Align(HorizontalAlignment.Right)
    .Align(VerticalAlignment.Top)
    .Wrap();

// Style is a record: build it once, vary it with `with`, and give it to any cell.
Style heading = new() { Bold = true, FontSize = 14, Background = "#DDEBF7" };

Cell title = Cell.Of("Report").WithStyle(heading);
Cell subtitle = Cell.Of("2026").WithStyle(heading with { FontSize = 12 });

Cell formula = Cell.FromFormula("=SUM(B2:B9)");                      // calculated by Excel, or by Workbook.RecalculateAsync
Cell cached = Cell.FromFormula("SUM(B2:B9)", CellValue.Number(99));   // with the result that programs that do not calculate show

Console.WriteLine($"{styled.Style.Bold} {styled.Style.NumberFormat} {plain.Style.IsDefault} {Style.Default == plain.Style}");
Console.WriteLine($"{formula.Formula} {cached.Value} {Cell.Empty.IsEmpty} {title.Kind} {subtitle.Style.FontSize}");

if (styled.TryGetNumber(out double value) == true)
{
    Console.WriteLine(value);
}
```

## 10. Addresses and ranges

Everything is 0-based: `row[0]` is column A and `sheet[1, 2]` is the cell Excel calls **C2**.
`sheet["C2"]`, `CellAddress.Parse` and `CellRange.Parse` take Excel's notation. Error messages
give both: "cell C2 (row 1, column 2)".

<!-- example: addresses -->
```csharp
CellAddress c2 = new(1, 2);                            // row 1, column 2: the cell Excel calls C2
Console.WriteLine(c2);                                 // C2

if (CellAddress.Parse("$AA$10").IsSuccess(out CellAddress parsed) == true)
{
    Console.WriteLine($"row {parsed.Row}, column {parsed.Column}");   // row 9, column 26
}

bool valid = CellAddress.TryParse("XFE1", out _);      // false: past the last column
string name = CellAddress.ColumnName(27);              // AB

if (CellAddress.TryParseColumn("AB", out int column) == true)
{
    Console.WriteLine($"{valid} {name} {column} {CellAddress.MaxRows} x {CellAddress.MaxColumns}");   // False AB 27 1048576 x 16384
}

CellRange range = new(1, 1, 9, 3);                     // B2:D10

if (CellRange.TryParse("D10:B2", out CellRange same) == true)   // corners in any order
{
    Console.WriteLine($"{range} {range == same} {range.RowCount}x{range.ColumnCount} {range.Contains(c2)} {range.First}-{range.Last}");
}
```

## 11. Create a workbook in memory

<!-- example: document-create -->
```csharp
Workbook book = Workbook.Create();

if (book.AddSheet("Budget").IsFailure(out string? error, out _, out Worksheet? budget) == true)
{
    Console.WriteLine(error);

    return;
}

budget["A1"] = Cell.Of("Month").Bold();
budget["B1"] = Cell.Of("Amount").Bold();

for (int month = 1; month <= 12; month++)
{
    budget[month, 0] = new DateOnly(2026, month, 1);   // row month, column A
    budget[month, 1] = month * 1000;
}

budget["A14"] = Cell.Of("Total").Bold();
budget.SetFormula(13, 1, "SUM(B2:B13)");
budget.SetColumnWidth(0, 14);
budget.FreezePanes(1);
budget.AutoFilter = true;

if (book.AddSheet("Settings").IsSuccess(out Worksheet? settings) == true)
{
    settings["A1"] = 0.25;
    settings.Visibility = SheetVisibility.Hidden;
}

RecalculationReport report = await book.RecalculateAsync();
Console.WriteLine($"{report.Calculated} formulas, total {budget["B14"].Value}");

Result saved = await book.SaveAsync("budget.xlsx");
```

## 12. Edit an existing workbook

When a file is saved, every part the model does not own is copied byte for byte, worksheets are
copied element by element with only the cell data rewritten, and a cell that was not changed is
written from its original XML. Styles that were in the file keep their index, so borders and
theme colours survive even though `Style` does not describe them. The calculation chain is
dropped and Excel is told to recalculate on open.

<!-- example: document-edit -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true)
{
    Console.WriteLine(error);

    return;
}

if (book.Sheet("Budget").IsFailure(out string? missing, out _, out Worksheet? sheet) == true)
{
    Console.WriteLine(missing);   // NotFound

    return;
}

sheet.SetValue(4, 1, 125_000);                        // B5: the value changes, the cell keeps its style
sheet["C5"] = Cell.FromFormula("B5*1.25").Bold();      // replaces value, formula and style
sheet.Clear(5, 1);                                     // B6: value, formula and style gone

Console.WriteLine($"{sheet.RowCount} rows, {sheet.ColumnCount} columns, A is {sheet.GetColumnWidth(0)} wide, {sheet.FrozenRows} row and {sheet.FrozenColumns} columns frozen, filter {sheet.AutoFilter}");

foreach (Row row in sheet.Rows)
{
    Console.WriteLine($"{row.Index}: {string.Join(" | ", row.Cells.Select(cell => cell.Value.ToString()))}");
}

foreach (Worksheet each in book.Sheets)
{
    Console.WriteLine($"{each.Index} {each.Name} {each.Visibility} in {each.Workbook.Sheets.Count}");
}

if (book.Sheet(0).IsSuccess(out Worksheet? first) == true)
{
    Console.WriteLine(ReferenceEquals(first, sheet));  // True
}

await book.RecalculateAsync();
Result saved = await book.SaveAsync("budget.xlsx");    // saving over the opened file is fine
```

## 13. Open from and save to a stream

<!-- example: document-streams -->
```csharp
// A workbook from a stream (an upload, a blob) and back to one (a response, a blob).
Result<Workbook> opened = await Workbook.OpenAsync(input, KalkylLimits.Default);

if (opened.IsSuccess(out Workbook? book) == true)
{
    book.Sheets[0]["A1"] = "Checked";
    await book.SaveAsync(output);   // the stream is left open
}
```

## 14. Calculate formulas

`RecalculateAsync()` calculates the formulas affected by cells set since the last
recalculation, everything that depends on them, and volatile functions (NOW, TODAY, RAND,
INDIRECT, OFFSET). `RecalculationScope.All` calculates every formula. Circular references get 0,
as in Excel without iterative calculation, and are listed in the report together with unknown
functions (which give `#NAME?`) and formulas that cannot be parsed.

<!-- example: formulas -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
{
    return;
}

sheet["E1"] = Cell.FromFormula("B14/12");
sheet["E2"] = Cell.FromFormula("E1+E2");               // a circular reference
sheet["E3"] = Cell.FromFormula("NOSUCHFUNCTION(1)");

RecalculationReport changed = await book.RecalculateAsync();                     // what changed, and what depends on it
RecalculationReport all = await book.RecalculateAsync(RecalculationScope.All);  // every formula

Console.WriteLine($"{changed.Calculated} then {all.Calculated} calculated, E1 = {sheet["E1"].Value}");
Console.WriteLine("Circular: " + string.Join(", ", all.CircularReferences));    // Budget!E2
Console.WriteLine("Unknown: " + string.Join(", ", all.UnknownFunctions));       // NOSUCHFUNCTION
Console.WriteLine("Broken: " + string.Join(", ", all.InvalidFormulas));

// A formula that is not in a cell. References without a sheet name are on `sheet`.
Result<CellValue> average = book.Evaluate("AVERAGE(B2:B13)", sheet);
Result<CellValue> fromAnywhere = book.Evaluate("SUM(Budget!B2:B13)*1.25");

// Text read as a number or a date follows Workbook.Culture, which is the machine's unless set.
book.Culture = CultureInfo.GetCultureInfo("en-GB");
Result<CellValue> february = book.Evaluate("MONTH(\"3/2/2013\")");             // 2
Console.WriteLine($"Dates count from 1904: {book.Uses1904Dates}");
```

The engine follows Excel's conversions, comparison rules, error propagation, implicit
intersection, array arithmetic (`SUMPRODUCT((A1:A9>0)*B1:B9)`), defined names, shared formulas
and the 1900 and 1904 date systems. Not supported: dynamic-array spilling (an array result shows
its first value), structured table references (`Table1[Column]` gives `#REF!`), R1C1 references
in INDIRECT, and external workbooks. Text that a formula reads as a number or a date
(`"1,5"*2`, `YEAR("3/2/2013")`) is read in `Workbook.Culture`, the machine's unless you set it;
set it to the culture the workbook was made in to get the results its author saw. ISO dates
(`2026-02-01`) are read in every culture, and TEXT, DOLLAR and FIXED are not affected by it.

## 15. Charts

Column, bar, line, area (each also stacked), pie, doughnut and scatter, with a title, axis
titles, a legend position, alternative text and a colour per series. A series refers to cells on
the chart's sheet or another, and Excel reads them when it opens the file, so the chart follows
later changes. The same methods are on `SheetWriter`, where the series can refer to the sheets
added before the chart's, and `AddPicture` returns nothing, since the stream is never Strict OOXML.

<!-- example: charts -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
{
    return;
}

// A series is one row or one column of numbers; the categories are its labels (the x values of a scatter chart).
ChartSeries amount = new("Amount", "B2:B13", "A2:A13", "#4472C4");
ChartSeries target = new("Target", "C2:C13", "A2:A13");

Chart columns = new(ChartType.Column)
{
    Title = "Budget 2026",
    Series = [amount, target],
    CategoryAxisTitle = "Month",
    ValueAxisTitle = "kr",
    Legend = LegendPosition.Bottom,
    Description = "Amount and target per month",    // read by screen readers
};

Result added = sheet.AddChart(columns, "F2:M18");     // covers F2:M18
sheet.AddChart(new Chart(ChartType.Bar) { Stacked = true, Series = [amount, target] }, new CellRange(19, 5, 35, 12));
sheet.AddChart(new Chart(ChartType.Line) { Series = [amount], Legend = LegendPosition.None }, "F38:M54");
sheet.AddChart(new Chart(ChartType.Area) { Stacked = true, Series = [amount, target] }, "F56:M72");
sheet.AddChart(new Chart(ChartType.Pie) { Title = "Share per month", Series = [amount] }, "O2:U18");
sheet.AddChart(new Chart(ChartType.Doughnut) { Series = [amount] }, "O20:U36");
sheet.AddChart(new Chart(ChartType.Scatter) { Series = [new ChartSeries("Amount by target", "B2:B13", "C2:C13")] }, "O38:U54");

// Another sheet's cells: name the sheet, in quotes when the name has spaces.
if (book.AddSheet("Summary").IsFailure(out string? taken, out _, out Worksheet? summary) == true)
{
    Console.WriteLine(taken);

    return;
}

Result other = summary.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries("Amount", "'Budget'!B2:B13", "Budget!A2:A13")] }, "A1:H16");
Result missing = summary.AddChart(new Chart(ChartType.Column) { Series = [new ChartSeries(null, "Nowhere!B2:B13")] }, "A20:H36");

if (missing.IsFailure(out string? reason, out ResultStatus status) == true)
{
    Console.WriteLine($"{status}: {reason}");          // NotFound: there is no sheet Nowhere
}

await book.SaveAsync("budget-charts.xlsx");
```

Colours per point, data labels, gradients, axes, fonts and combination charts are in
[21](#21-chart-styling-and-combination-charts). The charts store no copy of the values: a program
that only previews the file, without calculating, may show them empty. Category labels are written as text references, so dates as
categories show as numbers unless the cells have a date format.

## 16. Pictures

PNG, JPEG and GIF, and SVG with `Picture.FromSvg` (see [20](#20-shapes-groups-and-dashboards)).
The format and the size are read from the file's header, without an image library, and the image
is stored exactly as given. A picture can also be see-through (`Transparency`), cast a `Shadow`
and link somewhere (`Link`), and the same picture put on a sheet twice is stored once.

<!-- example: pictures -->
```csharp
Result<Picture> read = await Picture.FromFileAsync("logo.png");   // or Picture.FromBytes(bytes)

if (read.IsFailure(out string? error, out _, out Picture? logo) == true)
{
    Console.WriteLine(error);   // missing, or not a PNG, JPEG or GIF

    return;
}

Console.WriteLine($"{logo.Format}, {logo.Width}x{logo.Height} pixels");
Picture described = logo with { Description = "Company logo" };   // alternative text

if ((await XlsxWriter.CreateAsync("pictures.xlsx")).IsSuccess(out XlsxWriter? writer) == false)
{
    return;
}

await using (writer)
{
    if ((await writer.AddSheetAsync("Sheet1")).IsSuccess(out SheetWriter? sheet) == true)
    {
        sheet.AddPicture(described, "A1");                         // at its own size, 96 pixels to the inch
        sheet.AddPicture(described, new CellAddress(0, 3), 0.5);   // D1, at half size
        sheet.AddPicture(described, "F1:H6");                      // stretched over F1:H6
    }
}
```

## 17. Conditional formatting

Cell comparisons, between, formulas, text that contains, top and bottom N or N %, above and below
average, duplicates and unique values, two- and three-colour scales, data bars and the 17 icon
sets. A rule's `Style` may set bold, italic, underline, font colour, fill and number format;
Excel lets a rule change nothing else, so anything else is an `ArgumentException`. Formulas are
in English with commas, as Excel stores them, and relative to the range's top-left cell. Rules
added to an opened sheet lose to the ones it had, where both apply.

<!-- example: conditional-formatting -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
{
    return;
}

// A rule's style may set bold, italic, underline, font colour, fill and number format.
Style bad = new() { FontColor = "#9C0006", Background = "#FFC7CE" };

Style good = new() { Bold = true, FontColor = "#006100", Background = "#C6EFCE" };


sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenCell(ComparisonOperator.GreaterThan, 10_000, good));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenCell(ComparisonOperator.LessThan, "$C$2", bad));   // compared with a formula
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenBetween(4000, 6000, new Style { Italic = true }));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenBetween("$C$2/2", "$C$2", new Style { Underline = true }));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.WhenNotBetween(0, 100_000, bad));
sheet.AddConditionalFormat("A2:C13", ConditionalFormat.WhenFormula("$B2>$C2", new Style { Background = "#FFEB9C" }));   // whole rows
sheet.AddConditionalFormat("D2:D13", ConditionalFormat.WhenTextContains("late", bad));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.Top(3, good));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.Bottom(10, bad, percent: true));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.AboveAverage(new Style { Bold = true }));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.BelowAverage(new Style { NumberFormat = "0.0" }));
sheet.AddConditionalFormat("D2:D13", ConditionalFormat.Duplicates(bad));
sheet.AddConditionalFormat("D2:D13", ConditionalFormat.Unique(good));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.ColorScale("#F8696B", "#63BE7B"));
sheet.AddConditionalFormat("C2:C13", ConditionalFormat.ColorScale("#F8696B", "#FFEB84", "#63BE7B"));
sheet.AddConditionalFormat(new CellRange(1, 1, 12, 1), ConditionalFormat.DataBar("#638EC6"));
sheet.AddConditionalFormat("C2:C13", ConditionalFormat.IconSet(IconSetKind.ThreeTrafficLights));
sheet.AddConditionalFormat("B2:B13", ConditionalFormat.IconSet(IconSetKind.FiveRating, reverse: true, showValue: false));

await book.SaveAsync("budget-rules.xlsx");
```

In an opened file, charts and pictures go into the sheet's drawing next to the ones it has, and
rules after its own; what the file had stays as it was. Strict OOXML files get no charts or
pictures: `Worksheet.AddChart` and `AddPicture` return Invalid for them.

## 18. Macros

A VBA project, Excel 4.0 macro sheets and ActiveX controls are never run, but like everything
else the package does not understand they are kept when the file is saved. So opening and saving
a file does not make it safe. `RemoveMacros` drops a VBA project; for macro sheets, ActiveX
controls and the rest, see [19. Active content](#19-active-content).

<!-- example: macros -->
```csharp
if ((await Workbook.OpenAsync("upload.xlsm")).IsSuccess(out Workbook? book) == false)
{
    return;
}

if (book.HasMacros == true)
{
    Result removed = book.RemoveMacros();   // Invalid for macro sheets and ActiveX controls, which cannot be removed

    if (removed.IsFailure(out string? error, out _) == true)
    {
        Console.WriteLine(error);

        return;
    }
}

await book.SaveAsync("clean.xlsx");         // without the VBA project, as an ordinary .xlsx
```

## 19. Active content

Macros are one of the ways a workbook can act on its own. `Workbook.ActiveContent` says which of
them a file has: `Macros` (VBA, Excel 4.0 macro sheets and dialog sheets), `ActiveX`,
`EmbeddedObjects` (OLE objects and packages), `ExternalReferences` (links to other workbooks, DDE
links such as `=cmd|'/c calc'!A0`, and anything else outside the file that Excel fetches by
itself; hyperlinks do not count) and `DataConnections` (connections, query tables and Power
Query). `ExternalReferences` lists the links, with the part each is in; the type,
`OpcExternalRelationship`, is in the namespace `Avig.OpenXml`. `RemoveActiveContent` makes the
next save leave all of it out and write a plain .xlsx that Excel opens without repairing it. A
formula that refers to another workbook keeps the value it last had, a table a query filled
keeps its data, and the button of a form control stays.

<!-- example: active-content -->
```csharp
// A zip file that says two things (KalkylLimits.StrictZip) is refused as damaged.
if ((await Workbook.OpenAsync("upload.xlsx")).IsSuccess(out Workbook? book) == false)
{
    return;
}

if (book.ActiveContent != ActiveContent.None)
{
    Console.WriteLine(book.ActiveContent);   // Macros, ExternalReferences

    foreach (OpcExternalRelationship link in book.ExternalReferences)
    {
        Console.WriteLine($"{link.Kind}: {link.Target}");   // externalLinkPath: file:///\fileserver\share\other.xlsx, ddeLink: cmd|/c calc
    }

    Result removed = book.RemoveActiveContent();   // Invalid for a pivot table or data model fed by a connection

    if (removed.IsFailure(out string? error, out _) == true)
    {
        Console.WriteLine(error);

        return;
    }
}

await book.SaveAsync("clean.xlsx");   // no macro sheet, no links; the linked cells keep their last values
```

`XlsxReader.GetActiveContentAsync` and `GetExternalReferencesAsync` give the same answers
without loading the sheets. `RemoveActiveContent` returns `Invalid`, and changes nothing, when
taking something out would break the workbook: a pivot table or a data model (Power Pivot) that
gets its data from a connection, or a chart drawn from another workbook.

## 20. Shapes, groups and dashboards

Rectangles, rounded rectangles, ellipses, triangles, diamonds, hexagons, block arrows, chevrons,
pentagons, lines and text boxes, each with a fill (none, a colour with transparency, or a linear
gradient), an outline (colour, width, dash), an outer and an inner shadow, soft edges, a turn,
text in runs with their own font, and a link to a cell in the workbook or to a web address.
`ShapeGroup` keeps shapes and pictures together as one button with one link. A drawing goes over
a block of cells, or at a place in points with `DrawingBounds`, where it stays when rows and
columns change size, as a dashboard's cards should. The same methods are on `Worksheet`, where
they return a `Result` (`Invalid` for a Strict OOXML file).

<!-- example: dashboard -->
```csharp
if ((await XlsxWriter.CreateAsync("dashboard.xlsx")).IsSuccess(out XlsxWriter? writer) == false)
{
    return;
}

await using (writer)
{
    // A chart refers to sheets written before it, so the data comes first.
    if ((await writer.AddSheetAsync("Data")).IsFailure(out string? error, out _, out SheetWriter? data) == true)
    {
        Console.WriteLine(error);

        return;
    }

    await data.WriteRowAsync(["Done", 0.72]);
    await data.WriteRowAsync(["Left", 0.28]);

    if ((await writer.AddSheetAsync("Dashboard")).IsSuccess(out SheetWriter? sheet) == false)
    {
        return;
    }

    // No gridlines or headers: the shapes are the page. Set before the first row.
    sheet.SetView(new SheetView { ShowGridLines = false, ShowHeaders = false, Zoom = 90, TabColor = "#2384F9" });

    // A card: a rounded rectangle with a gradient, a shadow, two lines of text and a link.
    Shape card = new(ShapeKind.RoundedRectangle)
    {
        Fill = Fill.Gradient("#2384F9", "#605776", angle: 45),
        Outline = Outline.None,
        Shadow = new Shadow(),                        // Excel's "Offset: Bottom Right"
        CornerRadius = 0.12,
        Text = new ShapeText(
        [
            new TextRun("Projects\n", new TextFont { Size = 11, Color = "#E6ECF5" }),
            new TextRun("24", new TextFont { Size = 26, Bold = true, Color = "#FFFFFF" }),
        ])
        {
            Horizontal = HorizontalAlignment.Left,
            MarginLeft = 16,
        },
        Link = Hyperlink.ToCell("Data", "A1") with { Tooltip = "Open the data" },
        Description = "24 projects",
    };

    sheet.AddShape(card, DrawingBounds.FromPoints(190, 20, 235, 95));   // left, top, width, height
    sheet.AddShape(new Shape(ShapeKind.UpArrow) { Fill = Fill.Solid("#8BFFF7"), Outline = Outline.None }, "J3:J5");

    // An SVG icon. Excel 2016 and later draw the SVG; older versions a transparent pixel.
    Result<Picture> icon = Picture.FromSvg("""<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="3" fill="#8BFFF7"/></svg>"""u8);

    if (icon.IsFailure(out string? invalid, out _, out Picture? svg) == true)
    {
        Console.WriteLine(invalid);   // not an SVG document

        return;
    }

    // A navigation button: shapes and pictures that move, resize and link as one.
    ShapeGroup button = new(
    [
        GroupItem.Of(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Solid("#232D45"), Outline = Outline.None }, DrawingBounds.FromPoints(0, 0, 145, 40)),
        GroupItem.Of(svg with { Description = "Data" }, DrawingBounds.FromPoints(12, 9, 22, 22)),
        GroupItem.Of(new Shape(ShapeKind.TextBox) { Text = new ShapeText("Data", new TextFont { Color = "#FFFFFF" }) { Horizontal = HorizontalAlignment.Left } }, DrawingBounds.FromPoints(42, 0, 100, 40)),
    ])
    {
        Link = Hyperlink.ToCell("Data", "A1"),
    };

    sheet.AddGroup(button, DrawingBounds.FromPoints(20, 20, 145, 40));

    // A gauge: a doughnut with a thick ring, a colour per slice, no border or background.
    Chart gauge = new(ChartType.Doughnut)
    {
        Series = [new ChartSeries("Share", "Data!B1:B2", "Data!A1:A2") { PointColors = ["#00C5AE", "#2E3850"], Outline = Outline.None }],
        HoleSize = 75,
        Legend = LegendPosition.None,
        Background = Fill.None,
        Border = Outline.None,
    };

    Result added = sheet.AddChart(gauge, DrawingBounds.FromPoints(440, 20, 160, 160));

    sheet.AddHyperlink("A1", Hyperlink.ToUrl("https://example.com/help"));   // a cell that links
    await sheet.WriteRowAsync([Cell.Of("Help").Underline().FontColor("#0563C1")]);
}
```

What is left unset looks as a shape just inserted in Excel does: the theme's blue with a darker
outline and white text, or for a text box no fill and no outline. A group's items are placed in
the group's own coordinates and scaled so that the box around them fills the group's place. An
SVG is stored the way Excel 2016 stores it, with a raster picture beside it for older versions:
`FromSvg(svg)` uses a transparent pixel, `FromSvg(svg, fallback)` your PNG. The size comes from
the SVG's `width` and `height` or its `viewBox`. Files Excel saved keep their own shapes, groups,
icons and charts byte for byte; new ones are added after them.

## 21. Chart styling and combination charts

Every option is optional, and a chart without them is written as before. Per chart:
`HoleSize` and `FirstSliceAngle` (doughnut and pie), `GapWidth` and `Overlap` (column and bar),
`VaryColors`, `RoundedCorners`, `Background` and `PlotBackground` (`Fill`), `Border`
(`Outline.None` removes it), `Font`, `TitleFont`, `LegendFont`, `DataLabels`, and a `ChartAxis`
each for `CategoryAxis`, `ValueAxis` and `SecondaryValueAxis` (`Visible`, `Gridlines`,
`GridlineOutline`, `Line`, `Minimum`, `Maximum`, `MajorUnit`, `NumberFormat`, `Font`). Per
series: `Type` (column, line and area mix in one chart), `SecondaryAxis`, `Fill`, `Outline` (the
line of a line series), `Shadow`, `PointColors`, `Smooth`, `Marker`, `MarkerSize` and its own
`DataLabels`.

<!-- example: charts-styled -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
{
    return;
}

// Pick out one month: a colour per point, null for the others.
string?[] highlight = new string?[12];
highlight[7] = "#00C5AE";

Chart chart = new(ChartType.Column)
{
    Title = "Amount and target",
    TitleFont = new TextFont { Size = 12, Bold = true },
    Series =
    [
        new ChartSeries("Amount", "B2:B13", "A2:A13", "#3A4A6B")
        {
            PointColors = highlight,
            DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, NumberFormat = "# ##0", Font = new TextFont { Size = 8, Color = "#FFFFFF" } },
        },

        // A combination chart: this series is a smooth dashed line, on a second axis.
        new ChartSeries("Target", "C2:C13", "A2:A13")
        {
            Type = ChartType.Line,
            Smooth = true,
            Marker = MarkerStyle.Circle,
            MarkerSize = 6,
            Outline = new Outline("#E1541C", 2.25) { Dash = LineDash.Dash },
            SecondaryAxis = true,
        },
    ],
    GapWidth = 60,                                    // narrower gaps between the columns
    Background = Fill.Gradient("#FFFFFF", "#EEF3FA"),
    Border = Outline.None,                            // no frame
    Font = new TextFont { Name = "Segoe UI", Size = 9, Color = "#595959" },
    ValueAxis = new ChartAxis { Minimum = 0, MajorUnit = 2000, NumberFormat = "# ##0", Line = Outline.None, GridlineOutline = new Outline("#D9D9D9", 0.5) },
    SecondaryValueAxis = new ChartAxis { Visible = false, Minimum = 0, Maximum = 14000 },
    Legend = LegendPosition.Top,
};

Result added = sheet.AddChart(chart, "F2:N20");
await book.SaveAsync("budget-styled.xlsx");
```

A funnel is a stacked bar chart whose first series has `Fill = Fill.None`; a gauge is a doughnut
with two points, `PointColors` and a large `HoleSize`. Excel repairs a file whose data labels sit
where the chart type does not allow (`OutsideEnd` on a stacked column, any position on a
doughnut), so `AddChart` throws an `ArgumentException` for those; `DataLabels.Position` lists
what each type takes. A lone series with its own point colours gets a legend entry per
category, as Excel shows it; set `Legend = LegendPosition.None` to leave it out.

## 22. Sheet view and cell links

`SheetView` holds what Excel's View tab sets for a sheet: gridlines, row and column headers,
zoom, the tab's colour, the selected cell and right-to-left. On the streaming writer it is
`SetView` before the first row; in the document model `Worksheet.View`, which is read from an
opened file and rewrites only what changes (a view attribute the model does not know, such as
page layout, is kept). `AddHyperlink` makes a cell a link to a cell in the workbook or to a web
address; the cell keeps its value and style, and an opened file keeps its own links.

<!-- example: sheet-view -->
```csharp
Result<Workbook> opened = await Workbook.OpenAsync("budget.xlsx");

if (opened.IsFailure(out string? error, out _, out Workbook? book) == true || book.Sheet("Budget").IsSuccess(out Worksheet? sheet) == false)
{
    return;
}

Console.WriteLine(sheet.View.Zoom);   // as the file has it: 100
sheet.View = sheet.View with { ShowGridLines = false, Zoom = 120, TabColor = "#00B050", ActiveCell = new CellAddress(1, 1) };

// A cell that links to another sheet, and one that links to the web. The cells keep their values.
sheet["F1"] = Cell.Of("Go to the summary").Underline().FontColor("#0563C1");
Result linked = sheet.AddHyperlink("F1", Hyperlink.ToCell("Summary", "A1"));
sheet.AddHyperlink(new CellAddress(1, 5), Hyperlink.ToUrl("https://example.com/budget") with { Tooltip = "The budget online" });

await book.SaveAsync("budget-view.xlsx");
```

## The API at a glance

| Type | What it is | Scenario |
| --- | --- | --- |
| `Xlsx` | `WriteAsync(path, items, sheetName)` and `ReadAsync<T>(path, sheetName)`: a typed list as a whole file. | 1 |
| `KalkylRowAttribute`, `KalkylColumnAttribute`, `KalkylIgnoreAttribute` | `[KalkylRow]`, `[KalkylColumn("Header", Order, NumberFormat, Width)]`, `[KalkylIgnore]`. | 1 |
| `KalkylMapping` | The extension methods `SheetWriter.WriteAsync(items)` and `SheetReader.ReadAsync<T>()`. | 2 |
| `IKalkylRow<T>`, `KalkylMap<T>`, `KalkylColumnInfo` | Implemented by the generator: `T.KalkylMap` has `Columns`, `ToCells(item)` and `FromRow(row)`. Generic code can ask for `where T : IKalkylRow<T>`. | 3 |
| `KalkylCells`, `MappedRow` | Used by the generated code; hidden from IntelliSense. | — |
| `XlsxWriter`, `SheetWriter` | Streaming writer: `CreateAsync(path)` or `(stream, leaveOpen)`, `AddSheetAsync`, `CompleteAsync` (or `await using`); per sheet `SetColumnWidth`, `FreezePanes`, `SetView`, `EnableAutoFilter`, `WriteRowAsync`, `RowCount`, `IsFinished`, and `AddChart`, `AddPicture`, `AddShape`, `AddGroup`, `AddHyperlink`, `AddConditionalFormat`. | 4, 5, 20 |
| `XlsxReader`, `SheetReader`, `SheetInfo`, `SheetVisibility` | Streaming reader: `OpenAsync(path or stream, limits)`, `Sheets` (`Name`, `Index`, `Visibility`: `Visible`, `Hidden`, `VeryHidden`), `Sheet(name or index)`, `ReadRowsAsync`, `Uses1904Dates`. | 6 |
| `Row` | One row: `Index`, `Cells`, `ColumnCount`, `IsEmpty`, `row[column]`, `TryGetText`, `TryGetNumber`, `TryGetBool`, `TryGetDate`. | 6 |
| `KalkylLimits` | `Default`, `Unlimited` and `with { MaxUncompressedBytes, MaxCells, MaxSharedStrings, MaxArrayCells, StrictZip }`. | 7 |
| `KalkylError`, `KalkylErrorKind` | `ErrorKind()` on a result or exception: `NotFound`, `NotXlsx`, `Damaged`, `TooLarge`, `Inaccessible`, `Unknown`. | 7 |
| `CellValue`, `CellKind` | A value: `Text`, `Number`, `Bool`, `Date`, `Error` or `Empty`, with `TryGet…`. | 8 |
| `Cell`, `Style`, `HorizontalAlignment`, `VerticalAlignment` | A value with a formula and a style. `Style` has `Bold`, `Italic`, `Underline`, `FontColor`, `FontSize`, `FontName`, `Background`, `NumberFormat`, `Horizontal` (`General`, `Left`, `Center`, `Right`, `Justify`), `Vertical` (`Bottom`, `Center`, `Top`) and `WrapText`. | 9 |
| `CellAddress`, `CellRange` | A cell and a block of cells, 0-based, with Excel's notation in `Parse` and `ToString`. | 10 |
| `Workbook`, `Worksheet` | The document model: `Create`, `OpenAsync`, `AddSheet`, `Sheet`, `Sheets`, `SaveAsync`; per sheet `sheet[row, column]`, `sheet["C2"]`, `SetValue`, `SetFormula`, `Clear`, `Rows`, `RowCount`, `ColumnCount`, widths, panes, `View`, `AutoFilter` and `Visibility`, and the same `Add…` methods as `SheetWriter`. | 11–13, 22 |
| `RecalculationScope`, `RecalculationReport` | `RecalculateAsync(Changed or All)` and what it found; `Workbook.Evaluate`, `Culture`, `Uses1904Dates`, `Limits`. | 14 |
| `Chart`, `ChartSeries`, `ChartType`, `LegendPosition` | `AddChart` on both writers, over cells or at `DrawingBounds`. Legend `None`, `Right`, `Bottom`, `Top` or `Left`. | 15 |
| `ChartAxis`, `DataLabels`, `DataLabelPosition`, `MarkerStyle` | Chart styling: axes (`Visible`, `Gridlines`, `Minimum`, `Maximum`, `MajorUnit`, `NumberFormat`, `Font`), labels (`ShowValue`, `ShowPercent`, `ShowCategory`, `ShowSeriesName`, `Position`, `NumberFormat`, `Font`), markers (`None`, `Circle`, `Square`, `Diamond`, `Triangle`, `X`, `Star`, `Dash`, `Dot`, `Plus`). | 21 |
| `Picture`, `PictureFormat` | `FromFileAsync`, `FromBytes`, `FromSvg`, `Format` (`Png`, `Jpeg`, `Gif`, `Svg`), `Width`, `Height`, `Description`, `Link`, `Transparency`, `Shadow`; `AddPicture` on both writers. | 16, 20 |
| `Shape`, `ShapeKind`, `ShapeText`, `TextRun`, `TextFont` | A shape: `Kind` (`Rectangle`, `RoundedRectangle`, `Ellipse`, `Triangle`, `Diamond`, `Hexagon`, `RightArrow`, `LeftArrow`, `UpArrow`, `DownArrow`, `Chevron`, `Pentagon`, `Line`, `TextBox`), `CornerRadius`, `Fill`, `Outline`, `Shadow`, `InnerShadow`, `SoftEdges`, `Text` (runs, `Horizontal`, `Vertical`, `Wrap`, margins), `Rotation`, flips, `Link`, `Name`, `Description`. `TextFont` also styles chart text. | 20 |
| `ShapeGroup`, `GroupItem`, `DrawingBounds` | Shapes and pictures as one (`Items`, `Link`); a place on the sheet in points or EMU (`FromPoints`, `FromEmu`). | 20 |
| `Fill`, `GradientStop`, `Outline`, `LineDash`, `Shadow` | `Fill.None`, `Solid(color, transparency)`, `Gradient(angle, stops)` or `Gradient(from, to, angle)`; `Outline(color, width)` with `Dash` and `Transparency`, `Outline.None`; `Shadow` with `Color`, `Transparency`, `Blur`, `Distance`, `Angle`. For shapes and charts. | 20, 21 |
| `Hyperlink` | `ToCell(sheet, cell)` or `ToUrl(url)`, with `Tooltip`: on a shape, a group, a picture, or a cell with `AddHyperlink`. | 20, 22 |
| `SheetView` | `ShowGridLines`, `ShowHeaders`, `Zoom`, `TabColor`, `ActiveCell`, `RightToLeft`; `SetView` on the streaming writer, `Worksheet.View`. | 22 |
| `ConditionalFormat`, `ComparisonOperator`, `IconSetKind` | `AddConditionalFormat` on both writers. Comparisons `Equal`, `NotEqual`, `GreaterThan`, `GreaterThanOrEqual`, `LessThan`, `LessThanOrEqual`; 17 icon sets from `ThreeArrows` to `FiveQuarters`. | 17 |
| `Workbook.HasMacros`, `RemoveMacros` | Find and drop a VBA project. | 18 |
| `ActiveContent`, `Workbook.ActiveContent`, `ExternalReferences`, `RemoveActiveContent` | What can act on its own (`Macros`, `ActiveX`, `EmbeddedObjects`, `ExternalReferences`, `DataConnections`), the external links as `OpcExternalRelationship` (`SourcePart`, `Kind`, `Target`), and a save without any of it. `XlsxReader.GetActiveContentAsync` and `GetExternalReferencesAsync` on the streaming reader. | 19 |

## What goes wrong, and how you hear about it

| Situation | You get |
| --- | --- |
| File or folder missing, sheet not found, a chart series on a sheet that does not exist | `NotFound` |
| No permission | `Forbidden` |
| File locked by another program, sheet name taken | `Conflict` |
| Not an .xlsx, an old .xls, password-protected, damaged, bad sheet name, bad cell when mapping, a picture that is not PNG, JPEG or GIF | `Invalid` |
| A sheet damaged half-way through a stream, a drawing damaged when something is added to it | `InvalidDataException` |
| Calling something in the wrong order, values Excel cannot store, a rule style Excel does not allow | `InvalidOperationException`, `ArgumentException` |
| A file past one of the limits | `Invalid` |
| A formula past the array limit | `#NUM!` in the cell |
| A cancelled or timed-out call | `OperationCanceledException` |

`ErrorKind()` on a `Result`, a `Result<T>` or an exception gives `NotFound`, `NotXlsx` (an .xls,
a password-protected file, a CSV), `Damaged`, `TooLarge` (past a `KalkylLimits` value),
`Inaccessible` (another program holds the file) or `Unknown`.

## Files from outside

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
- The zip file must say one thing only (`KalkylLimits.StrictZip`, on by default): its local
  headers as the zip directory has them, and nothing hidden before, between or after the entries.

Macros, ActiveX controls, OLE objects, links to other workbooks, DDE links and data connections
are never run, followed or refreshed, but they are kept when a workbook is saved, as everything
else the package does not understand. Check `Workbook.ActiveContent` for a file from outside, and
call `RemoveActiveContent` before you pass it on (see [19. Active content](#19-active-content)).

The package never reaches the network, starts a process, loads code or deserializes objects,
and it reads and writes only the paths you give it. DDE, external workbooks, `WEBSERVICE`,
`CELL("filename")` and `INFO` give `#NAME?` or `#REF!`. A test checks the source for this.

**Error messages** name the file path you gave, sheet names and the content of the cell that
was wrong. They are for your log, not for the person who uploaded the file.

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
tools/Avig.Kalkyl.Fuzz/schedule-night.ps1                               # libFuzzer every night in Task Scheduler
dotnet run -c Release --project tools/Avig.Kalkyl.Benchmarks -- --filter *   # baseline in docs/benchmarks/baseline.md
dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check           # files to open in Excel, see docs/excel-checklista.md
```

## Licence

MIT.
