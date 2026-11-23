# Avig.Kalkyl

Read and write Excel `.xlsx` files. Rows stream in and out one at a time, cell values are a
typed union with a Try pattern, styles are set per cell and deduplicated for you, and dates are
recognised from their number format. Expected failures come back as
[Avig.Result](https://www.nuget.org/packages/Avig.Result) values instead of exceptions.

No dependencies besides Avig.Result. Safe for trimming and Native AOT.

> **Status: in development (0.x).** Streaming read and write are done. Typed mapping, the
> document model and the formula engine are on their way to 1.0.

## Writing

```csharp
using Avig;
using Avig.Kalkyl;

Result<XlsxWriter> created = await XlsxWriter.CreateAsync("customers.xlsx");

if (created.IsFailure(out string? error, out _))
{
    Console.WriteLine(error);   // folder missing, file open in Excel ...
    return;
}

created.IsSuccess(out XlsxWriter? writer);

await using (writer)
{
    Result<SheetWriter> added = await writer.AddSheetAsync("Customers");
    added.IsSuccess(out SheetWriter? sheet);

    sheet.SetColumnWidth(0, 30);
    sheet.FreezePanes(1);
    sheet.EnableAutoFilter();

    await sheet.WriteRowAsync([Cell.Of("Name").Bold(), Cell.Of("Since").Bold(), Cell.Of("Balance").Bold()]);

    foreach (Customer customer in customers)
    {
        await sheet.WriteRowAsync([customer.Name, customer.Since, Cell.Of(customer.Balance).NumberFormat("# ##0.00")]);
    }

    await sheet.WriteRowAsync([Cell.Of("Total").Bold(), Cell.Empty, Cell.FromFormula("SUM(C2:C100)")]);
}
```

Any value converts to a `Cell` implicitly, so a row is just a collection expression. Rows go
from the top down; column widths and frozen panes have to be set before the first row, because
that is where they sit in the file.

## Reading

```csharp
Result<XlsxReader> opened = await XlsxReader.OpenAsync("customers.xlsx");

if (opened.IsFailure(out string? error, out ResultStatus status))
{
    Console.WriteLine(error);   // not found, an old .xls, password-protected ...
    return;
}

opened.IsSuccess(out XlsxReader? reader);

await using (reader)
{
    if (reader.Sheet("Customers").IsSuccess(out SheetReader? sheet))
    {
        await foreach (Row row in sheet.ReadRowsAsync())
        {
            if (row.TryGetText(0, out string? name) && row.TryGetDate(1, out DateTime since))
            {
                Console.WriteLine($"{name} since {since:d}");
            }
        }
    }
}
```

## Addresses are 0-based

`row[0]` is column A and `new CellAddress(1, 2)` is the cell Excel calls **C2**. `CellAddress`
converts both ways: `CellAddress.Parse("C2")` and `address.ToString()`.

## Values

`CellValue` holds one of `Text`, `Number`, `Bool`, `Date`, `Error` or nothing (`Empty`):

```csharp
switch (cell.Value.Kind)
{
    case CellKind.Number: ...
    case CellKind.Date: ...
}

if (cell.TryGetNumber(out double amount)) { ... }
```

`TryGetNumber` does not return a date's serial number: a date is a date. Excel stores dates as
numbers with a date format, and the reader turns them back into `DateTime`, including Excel's
made-up 29 February 1900 and workbooks that count from 1904. A time without a date comes back
on 1899-12-31, Excel's day zero.

## Styles

```csharp
Cell.Of("Total").Bold().Italic().FontColor("#C00000").Background("#FFF2CC").Align(HorizontalAlignment.Right)
```

`Style` is an immutable record compared by value. Two cells with equal styles share one entry in
the file, so styling every cell on its own never hits Excel's limit of about 64 000 styles.
When reading, only the number format is filled in for now.

## What goes wrong, and how you hear about it

| Situation | You get |
| --- | --- |
| File or folder missing | `NotFound` |
| No permission | `Forbidden` |
| File locked by another program, sheet name taken | `Conflict` |
| Not an .xlsx, an old .xls, password-protected, damaged container, bad sheet name | `Invalid` |
| A sheet damaged half-way through | `InvalidDataException` when the damaged row is reached |
| Calling something in the wrong order, values Excel cannot store | `InvalidOperationException`, `ArgumentException` |

## Building

```powershell
dotnet test Avig.Kalkyl.slnx
./publish.ps1          # tests, packs, checks the package shape, copies to the local feed
./release.ps1 -DryRun
```

## Licence

MIT.
