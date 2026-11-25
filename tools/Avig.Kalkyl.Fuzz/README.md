# Avig.Kalkyl.Fuzz

Fuzz harness for Avig.Kalkyl, built on [SharpFuzz](https://github.com/Metalnem/sharpfuzz) and
libFuzzer the way `Medusa.Fuzz` is. It has three modes:

- **smoke**: seeded random input with no coverage feedback, and nothing to install.
- **libFuzzer**: coverage-guided mutation, run through `fuzz.ps1`.
- **repro**: runs one saved input again and prints the whole exception.

The project is not part of the package: `publish.ps1` packs only `src/Avig.Kalkyl`.

## Targets

| Target | Kind | What it checks |
|---|---|---|
| `xlsx-part` | crash and invariant | One part of a valid workbook is replaced by the input. The file is read with the streaming reader and opened in the document model. It is then recalculated, changed and saved. The saved file must open and read through again, with the change and every sheet still there. Then its active content is removed and it is saved again: unless `RemoveActiveContent` refused, that file must open, with no active content and every sheet. |
| `attribute-sweep` | crash and invariant | Every number attribute of the template (style and font indexes, row and column numbers, counts, sizes, sheet ids) gets each of 18 edge values in turn, with the harness of `xlsx-part`. About 2,000 inputs; run with `--iter 3000`. |
| `zip-raw` | crash | The bytes of the valid template workbook are mutated where the zip structure is: fields in local headers and the central directory (sizes, offsets, compression method, name lengths), a truncated file, junk in front, and repeated chunks. Each input is read with `KalkylLimits.StrictZip` on, which refuses most of them at once, and then with it off, so that the checks behind it are reached too. |
| `formula` | crash | The formula engine gets any formula text. **Every** exception is a finding. |
| `formula-sweep` | crash | The same harness as `formula`, but every function is called in turn with every pair of edge values. |
| `formula-shift` | invariant | A formula moved there and back is unchanged. Moving it by nothing changes nothing. If the formula parses, the moved one parses too. |
| `roundtrip` | invariant | A value the package agrees to write reads back as the same value. This holds for the writer and reader, and for the document model. |
| `cell-text` | invariant | `Decode(Encode(s)) == s`, and the encoded text holds only characters XML allows. |
| `excel-date` | invariant | Serial number and date conversions agree in both directions, in both the 1900 and the 1904 system. A later serial number is never an earlier date. |
| `cell-address` | invariant | What `CellAddress` parses, it formats back to the same address. `Parse` and `TryParse` give the same answer. |

### What counts as a finding

`InvalidDataException` is the only exception that is swallowed. It is how the package says a
file is damaged, and the README at the repo root promises exactly that. The package's other
exceptions (`ArgumentException` and `InvalidOperationException`) mean the caller did something
wrong, and the harness never does. So if one of them is thrown, it is a finding. This is stricter
than Medusa, where `SafeRun` also swallows `ArgumentException` and `FormatException`. Here that
would have hidden five of the findings below.

An invariant target that sees a wrong answer throws `InvariantViolationException`. It is its
own type, so it can never be mistaken for an exception the package throws on purpose.

### Why `xlsx-part` does not mutate a whole .xlsx

Nearly every change to a zip file breaks the deflate stream or the CRC. The fuzzer would then
only test .NET's zip reader. Instead, the first input byte picks one of eight parts of a small
template workbook, and the rest of the input replaces that part. The mutations land in the XML
the package actually parses.

### Why `formula-sweep` exists

`REPT("",2147483648)` threw `ArgumentOutOfRangeException` once in the first run of `formula`.
Then it did not come back in 20,000 inputs. A random formula spreads its luck over 333 functions
and about 85 operands, so a bug that needs one function and two particular arguments is a
needle. The sweep goes through every function with every pair of 18 edge values, about 110,000
inputs, and finds such bugs every time. In libFuzzer mode, finding them is the job of coverage
guidance instead.

## Findings

The first round found thirteen bugs. Later rounds found four more: deep formulas, two sheets on one part, and an entry without a name, after the security test plan (`docs/testplan.md`) had shown where to look, and a workbook without root relationships that lost its shared strings in `RemoveActiveContent`. All are fixed and have regression tests in
`tests/Avig.Kalkyl.Tests/FuzzFindingsTests.cs`:

| Target | Bug | Effect |
|---|---|---|
| `xlsx-part` | The saver treated an unknown element at sheet level (`mc:AlternateContent`, another program's extension) as "after everything". It wrote the new `sheetData` before it and then copied the original as well. | The file had two `sheetData` elements and **the edit was lost** when the file was opened again. |
| `xlsx-part` | The same bug in `workbook.xml`. | **Every sheet appeared twice** after saving. |
| `xlsx-part` | The saver rebuilt relationships it does not own, and filled in `Target=""` where the attribute was missing. | A relationship the reader had skipped replaced another one with the same Id, so **a sheet disappeared**. |
| `xlsx-part` | A broken `[Content_Types].xml` was not read when the file was opened, only when it was saved. | `OpenAsync` succeeded and `SaveAsync` threw `XmlException`. |
| `xlsx-part` | Broken XML inside a cell (`<v>`, `<f>`, `<is>`) came out as `XmlException`. | The README promises `InvalidDataException`. |
| `xlsx-part` | `<v>1e309</v>` and `NaN` got through `double.TryParse`. | `ArgumentOutOfRangeException` from `CellValue.Number`. |
| `xlsx-part` | A row number after row 1,048,576 was accepted. | It crashed later, during recalculation. |
| `xlsx-part` | A font size outside 1–409 points in styles.xml. | `ArgumentOutOfRangeException` when the file was opened. |
| `formula` | `REPT` with empty text and a count above `int.MaxValue`. The limit check multiplied, and could also overflow. | `ArgumentOutOfRangeException`. |
| `formula` (added later) | About 820 nested parentheses, 700 terms of `1+1+…+1`, or 730 nested functions overflowed the stack. The parser and the evaluator recursed once per level, and the evaluator also recursed once per term of a left-associative chain. | **The process ended.** A stack overflow cannot be caught, so one formula from a user or a file brought down the program. Nesting is now limited to 255 levels and functions inside functions to Excel's 64; deeper gives Invalid. Chains are evaluated in a loop, and the evaluator checks how much stack is left. |
| `xlsx-part` (added later) | Two `<sheet>` elements with the same relationship id opened as two sheets on one part. | Saving wrote the part twice, so the saved file had a duplicate zip entry that no program can read. Two sheets on one part are now a damaged file. |
| `xlsx-part` and `zip-raw` (2026-11-25) | `RemoveActiveContent` kept only the parts the package root led to, but the reader also finds `xl/workbook.xml` when `_rels/.rels` is missing or empty. | The shared strings were left out, and the saved file was **damaged**. The search now starts from the workbook part too. |
| `zip-raw` | A central directory record with a name length of 0 opened. | Saving failed with `ArgumentException`, since an entry cannot be written without a name. Such a file is now damaged when it is opened. |
| `formula-shift` | The lexer did not accept `Sheet!#REF!`, which Excel itself writes and which the shifter writes when a reference ends up off the sheet. | A shared formula moved off the sheet no longer parsed. |
| `roundtrip` | A lone `\r` in text was written as it was. Every XML reader turns it into `\n`. | The text changed silently. It is now written as `_x000D_`, as Excel does. |
| `roundtrip` | `_xc1e5` before a character that gets encoded (a lone surrogate) became `_xc1e5_xD800_`, where the surrogate's escape closes the escape-lookalike before it. | The text came back as `\uC1E5xD800_`. |
| `roundtrip` and `excel-date` | `TryFromSerial` had a fixed end that was too narrow for the 1900 system and too wide for the 1904 system. | The last millisecond of 9999 came back as a number. In the 1904 system the reader threw instead of returning false. |

Two findings were bugs in the harness, not in the package. `excel-date` complained about the
order around day 60. That is the made-up date 1900-02-29, which reads as the 28th on purpose.
`formula-shift` complained about `==x`, because the shifter takes off one leading `=`. Both cases
are now excluded in the targets, with a comment that says why.

## Checking that a target is alive

A target that cannot find anything reports zero findings just as confidently as a target that
tests bug-free code. The only way to tell them apart is to break the code on purpose and see
whether the fuzzer notices. These are the results with `--seed 11`:

| Target | Sabotage | Result |
|---|---|---|
| `xlsx-part` | remove the row limit in `CellXml.RowIndex` | 5 findings in 456 inputs |
| `xlsx-part` | remove the 32,767-character check in `CellXml.Text` | 5 findings in 1,065 inputs. The first version found nothing: the long-text token rarely landed inside a cell, so the generator now puts long text in the template's own cells. |
| `zip-raw` | remove the empty-name check in `XmlParts.CheckEntriesAsync` | 5 findings in 8,775 inputs |
| `formula` | make `FunctionCall.TryInteger` throw instead of returning `#NUM!` | 4 findings in 20,000 inputs |
| `formula` | raise the parser's nesting limit from 255 to 1,000,000 | the process ends with a stack overflow; run with `--trace` to keep the input |
| `formula-sweep` | remove the empty-text check in `REPT` | 3 findings in 110,000 inputs |
| `formula-shift` | move the limit in `FormulaShifter` by one | 5 findings in 3,134 inputs |
| `roundtrip` | stop encoding `\r` in `CellText` | 5 findings in 35 inputs |
| `cell-text` | use `LooksLikeEscape` instead of `WouldLookLikeEscape` | 5 findings in 690 inputs |
| `excel-date` | round towards zero instead of away from zero | 5 findings in 16 inputs |
| `cell-address` | make `ToString` add a row for columns from 16,000 | 5 findings in 115 inputs |

Two attempts showed why the check is needed.

- **The REPT bug:** removing the REPT fix gave zero findings in `formula`. That is why
  `formula-sweep` exists.
- **Unknown elements:** putting the unknown-element bug back in the saver also gave zero
  findings. The fix that skips an element already written covers the same bug, so the sabotage
  broke nothing.

`formula` is the weakest target. Four findings in 20,000 inputs is enough to show that it is
alive, but not to trust it to find a bug in one particular function. That is the job of
`formula-sweep` and libFuzzer.

## Running

```powershell
# Every target, 20,000 inputs each
dotnet run -c Release -- all --smoke

# One target with a fixed seed, so that a finding can be reproduced
dotnet run -c Release -- roundtrip --smoke --iter 100000 --seed 42

# The whole sweep over the functions (about 110,000 inputs, 40 s)
dotnet run -c Release -- formula-sweep --smoke --iter 110000

# Print the whole exception for a saved input
dotnet run -c Release -- xlsx-part --repro findings/xlsx-part/crash-00000004-000543.bin

# For xlsx-part: also write out the built file and the saved file, to open and compare them
dotnet run -c Release -- xlsx-part --repro <file> --dump C:\temp\dump
```

To run every saved input (`corpus/` and `findings/`) in a process of its own, with a time limit and a 1 GiB heap, use `./corpus-check.ps1`. That catches the stack overflows, out-of-memory failures and hangs that end or stall the smoke runner.

Findings are saved in `findings/<target>/` as `crash-<seed>-<iteration>.bin`. An input that
takes more than two seconds is saved as `slow-…`. A stack overflow ends the process before any
`catch` runs. In that case, run again with the same `--seed` and add `--trace`: every input is
then written to `findings/<target>/last-input.bin` before it runs. `findings/` is not under
version control. A real finding should become a regression test instead.

### Coverage-guided with libFuzzer

Two programs are needed:

1. `sharpfuzz`, which `fuzz.ps1` installs if it is missing: `dotnet tool install -g SharpFuzz.CommandLine`
2. `libfuzzer-dotnet.exe`, from [Metalnem/libfuzzer-dotnet](https://github.com/Metalnem/libfuzzer-dotnet/releases). Put it in `~\bin\` or on `PATH`.

```powershell
./fuzz.ps1 -Target xlsx-part                              # 60 s
./fuzz.ps1 -Target formula -RuntimeSeconds 600 -Workers 4
./fuzz.ps1 -Target roundtrip -RuntimeSeconds 28800         # overnight
./fuzz-night.ps1 -TotalHours 8 -Workers 4                  # every target in turn, then corpus-check
```

The script publishes the harness and instruments only `Avig.Kalkyl.dll`. It then runs
libFuzzer with `corpus/<target>/` as seeds. The inputs libFuzzer finds interesting go to
`%LOCALAPPDATA%\Avig.Kalkyl.Fuzz\corpus-work\<target>\`. That is outside the repo, so the seeds
stay small and OneDrive does not upload a night's thousands of files. The seeds were made with
`--write-corpus 16 --seed 1` and can be made again the same way.

`fuzz-night.ps1` splits the time by weight: the package targets get the most, and
`attribute-sweep` is left out, because under libFuzzer it runs the harness of `xlsx-part`. It
writes a log to `findings/night-<date>.log`, ends with `corpus-check.ps1`, and keeps the
computer awake until it is done.

#### Every night

`schedule-night.ps1` puts `fuzz-night.ps1` in Windows Task Scheduler as the task
`Avig.Kalkyl fuzz night`:

```powershell
./schedule-night.ps1                                      # every night 23:00, 7 h, 4 workers
./schedule-night.ps1 -At 22:30 -TotalHours 8 -Workers 8
./schedule-night.ps1 -Once                                # the coming night only
./schedule-night.ps1 -Show                                # next run, last result, latest log
./schedule-night.ps1 -Remove
```

The task runs as you, and only while you are logged on (a locked screen is fine), so it needs
no stored password and no administrator. It wakes the computer from sleep. A night missed
because the computer was off is not made up in the morning. A last result of 0 means nothing
was found; 1 means a finding, or that corpus-check failed: read the log.

The first libFuzzer runs (2026-11-24, libfuzzer-dotnet v2025.05.02) found two bugs within a
minute that 50,000 smoke inputs had missed:

- **A defect in .NET 10's `System.Xml`.** `XmlTextReaderImpl` throws
  `ArgumentOutOfRangeException` from `Buffer.BlockCopy` for one malformed XML declaration read
  from a deflate stream. .NET 11 gives `XmlException`. `DepthLimitedXmlReader` now reports it as
  broken XML. Note that a file-based app (`dotnet run probe.cs`) runs on the newest runtime
  installed, not on .NET 10, so a probe of this kind must say `#:property TargetFramework=net10.0`.
- **`<sheet>` outside `<sheets>`.** A mutated `<sheetz>` still gave its sheets to the reader,
  and the saver wrote them twice. Sheets and names are now read only from their own sections.

After the fixes, three minutes each with 4 workers (about 610,000 inputs for `xlsx-part` and
2 million for `zip-raw`) gave nothing.

## Adding a target

1. Write a `FuzzTarget` with a harness and a generator. The generator should mostly give valid
   input, with a dense sprinkling of what tends to go wrong. Random bytes never become a
   formula or a sheet, and then the target only tests the branch that rejects the input.
2. Add the target to `Program.Targets` and to the `ValidateSet` in `fuzz.ps1`.
3. Run `--write-corpus 16` for the target.
4. **Break the code on purpose** and check that the target finds it. Add the sabotage to the
   table above.
