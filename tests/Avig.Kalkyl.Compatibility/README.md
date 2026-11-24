# Avig.Kalkyl.Compatibility

Opens real workbooks from Excel tutorials with Avig.Kalkyl and with four other libraries, and
checks that they agree. The other libraries are
[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader),
[ClosedXML](https://github.com/ClosedXML/ClosedXML),
[Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) and
[MiniExcel](https://github.com/mini-software/MiniExcel). They are test dependencies only. NPOI
and EPPlus are left out because of their licence terms.

## The workbooks

`files.json` lists 38 workbooks:

- Microsoft's Power BI tutorial samples.
- Microsoft's Data Validation Examples.
- The sample files from docs.fileformat.com and elysiatools.com.
- Contextures' practice data.
- Course assignments and class examples from two GitHub repositories of Excel exercises. These
  have most of the formulas: COUNTIFS, SUMIFS, VLOOKUP, INDEX/MATCH, text, statistics, the
  database functions, and 478,000 formulas in one workbook.

The files are **not committed**. The Microsoft samples may only be used for internal reference
(obviEnce notice), and the other pages state no licence. Download them once:

```powershell
./download.ps1          # into files/, skips what is already there; -Force fetches again
dotnet test             # or dotnet test --filter "FullyQualifiedName~Recalculation"
```

Without the files, every test is skipped.

A workbook made in another culture gives that culture in `files.json` (`"culture": "en-IN"`).
Text dates in its formulas, such as `"<=31/12/2018"`, are then read as its author meant them.

## The tests

Each test runs once per workbook.

| Test | What it checks |
|---|---|
| `Reading_agrees_with_the_other_libraries` | Every cell Avig.Kalkyl's `XlsxReader` reads is what the other libraries agree on. A cell counts when at least two of them give the same value and no other value gets as many votes. |
| `The_document_model_reads_what_the_streaming_reader_reads` | `Workbook` and `XlsxReader` read the same values. |
| `A_saved_copy_reads_the_same_in_the_other_libraries` | Opening with `Workbook` and saving unchanged does not change what any other library reads. |
| `A_written_copy_reads_the_same_in_the_other_libraries` | The values written into a new file with `XlsxWriter` read back the same in every other library. |
| `Recalculation_gives_the_values_Excel_saved` | Every formula is recalculated and compared with the value Excel saved in the file. This checks the formula engine against Excel itself. |

When a test fails, every difference is written to `report/<file>.<test>.md`. The copy that the
saving and writing tests made is kept there as well, so that it can be opened. Neither folder
is committed.

### What the comparison allows for

The libraries disagree with each other in a few known ways. Each case below is handled in one
place, with a comment that says why.

- **Dates before 1900-03-01.** These differ by a day between libraries that follow Excel (serial
  1 is 1900-01-01) and libraries that use `DateTime.FromOADate`. Times without a date fall in
  the same range. Such dates are counted apart and are not failures.
- **ClosedXML and line breaks.** ClosedXML reads `\n` in an inline string as `\r\n`. The
  adapter turns it back.
- **Sylvan and dates.** Sylvan gives a date the type `Numeric`, and only the number format
  shows that it is a date. Its `GetDateTime` also refuses serial numbers below 1.
- **Empty text.** An empty string and an empty cell count as the same.
- **Text results that Excel saved.** Excel trims a formula's text result when it saves it
  without `xml:space`, so text from a formula is compared without whitespace at the ends.
- **Volatile functions.** A cell with `RAND`, `RANDBETWEEN`, `NOW` or `TODAY` is set to the
  value Excel saved before recalculation, so the cells that depend on it can be checked.
- **Structured references.** Formulas like `Table1[Column]` are not supported. They are counted
  and not compared.

## What the first run found

The first run found nine bugs. All are fixed and have regression tests in
`tests/Avig.Kalkyl.Tests/CompatibilityFindingsTests.cs`.

| Bug | Effect |
|---|---|
| The used area of a sheet was found by walking every row on every range read. | The 478,000 formulas never finished. Now they take about 6 s. |
| A formula's cached text result was saved without `xml:space`. | Other libraries, and Excel, read `"Age "` as `"Age"`. |
| `XmlWriter` wrote `\n` in text as `Environment.NewLine`. | Files differed between Windows and Linux. |
| A time without a date got a date format. | Excel shows it as `1900-01-00 08:00`, and Sylvan cannot read it. It now gets `hh:mm:ss`. |
| `COUNTIF` with a range of criteria used only the first one. | `SUMPRODUCT(1/COUNTIF(A5:A11,A5:A11))`, the usual way to count distinct values, gave `#DIV/0!`. |
| `CHOOSE` with an array of indexes gave `#REF!`. | `VLOOKUP(x,CHOOSE({1,2},B:B,A:A),2,0)`, the usual way to look to the left, did not work. |
| The database functions were missing. | `DSUM`, `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DMAX`, `DMIN`, `DPRODUCT`, `DGET`, `DSTDEV`, `DSTDEVP`, `DVAR` and `DVARP` gave `#NAME?`. |
| Text dates were always read as US month/day. | `COUNTIFS(...,"<=31/12/2018")` counted 0, and `YEAR("1/13/1966")` gave 1966 where Excel in en-IN gives `#VALUE!`. Text is now read in `Workbook.Culture`, which is the machine's culture by default, as in Excel. |
| `1975-02-23T02:58:41.000Z` was read as a date, and moved into the machine's time zone. | Excel gives `#VALUE!`. |

## Running time

The whole suite takes about 1.5 minutes. Most of it is the workbook with 478,000 formulas:
opening it takes about 11 s and recalculating it about 6 s. Every sheet is cut off at 20,000
rows (`Readers.MaxRows`), so that the other libraries do not hold all of it in memory five times.
