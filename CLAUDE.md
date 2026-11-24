# CLAUDE.md

Knowledge-projekt: avigkalkyl
Dokumentation, beslut och agentlogg för projektet finns i Knowledge — se ~/.claude/CLAUDE.md.

Vägledning för Claude Code (claude.ai/code) i det här repot.

## Vad det här är

NuGet-paketet Avig.Kalkyl: läsa och skriva Excel-filer (.xlsx), som ett eget alternativ till
NPOI, ClosedXML och EPPlus. Valen bakom API:t och jämförelsen med förebilderna står i
`docs/analys.md`; planen med milstolparna i `docs/plan.md`.

## Kommandon

```powershell
dotnet build Avig.Kalkyl.slnx
dotnet test Avig.Kalkyl.slnx
./publish.ps1                 # testar, packar, formkontrollerar, lägger i G:\LocalNugetRepo
./release.ps1 -DryRun

# Fuzzning (tools/Avig.Kalkyl.Fuzz, se dess README)
dotnet run --project tools/Avig.Kalkyl.Fuzz -c Release -- all --smoke
dotnet run --project tools/Avig.Kalkyl.Fuzz -c Release -- formula-sweep --smoke --iter 110000
```

`tests/Avig.Kalkyl.Compatibility` läser riktiga arbetsböcker från Excel-handledningar
(hämtas med `download.ps1`, committas inte) med Avig.Kalkyl och ExcelDataReader, ClosedXML,
Sylvan och MiniExcel, och räknar om formlerna mot Excels sparade värden. Kör det efter ändringar
i läsaren, skrivaren, saveren eller formelmotorn. Ett fynd blir ett test i
`CompatibilityFindingsTests.cs`. Huvudtesterna körs med en-US (`TestCulture.cs`), eftersom
`Workbook.Culture` annars är maskinens.

`tools/Avig.Kalkyl.Fuzz` är en SharpFuzz-harness som Medusa.Fuzz, och den ingår inte i paketet.
Kör smoke-läget efter ändringar i läsaren, saveren, `CellText`, `ExcelDate`, lexern eller
funktionerna. Ett fynd blir ett regressionstest i `FuzzFindingsTests.cs`, och ett nytt mål ska
vara verifierat genom sabotage (tabellen i README:n).

## Regler som betyder mest

- **Inga beroenden utom Avig.Result.** Inget zip-, XML- eller bildbibliotek: allt görs med
  `System.IO.Compression` och `XmlReader`/`XmlWriter`. `publish.ps1` failar annars. Testerna
  får använda vad de vill (DocumentFormat.OpenXml används för schemavalidering).
- **0-baserat överallt** i det publika API:t. Excels adresser (C2) bara i `CellAddress.Parse`,
  `ToString` och i felmeddelanden, som alltid ska ange båda: "cell B2 (row 1, column 1)".
- **Förväntade fel är `Result`**, inte undantag: fel format, saknad fil, ogiltigt bladnamn.
  Undantag är för programmeringsfel (fel ordning, värden Excel inte kan lagra) och för en fil
  som visar sig vara skadad mitt i en ström.
- **Bara async** för allt som gör I/O, och `CancellationToken` förs alltid vidare.
- **Säkert för trimning och AOT** (`IsAotCompatible`). Ingen reflektion; den typade
  mappningen i milstolpe 2 blir en källgenerator.
- **En fil utifrån får inte kunna ta slut på minnet eller stacken.** Allt som läser en zip-post går via `XmlParts.OpenEntryAsync`, alla XmlReader skapas med `XmlParts.CreateReader` (djupgräns 100), och allt som skapar en matris går via `ValueArray` (`KalkylLimits.MaxArrayCells`). Ingen rekursion som växer med indata utan gräns: parsern begränsar nästling, och evaluatorn går igenom operatorkedjor i en loop. Aggregat över stora områden går via `FunctionCall.NonBlankValues` eller rutnätets `StoredValues`, inte position för position. Nya attackfiler blir tester i `HostileFilesTests.cs`; testplanen är `docs/testplan.md`.
- **Ett fel om en fil ska få en felkod.** Meddelanden om skadade filer börjar med `The file is damaged`, meddelanden om gränser innehåller `the limit set in KalkylLimits.`, och meddelanden om filer som inte är .xlsx börjar med `The file is not an .xlsx workbook`. Det är så `KalkylError.ErrorKind()` känner igen dem. Ett `InvalidDataException` som blir ett `Result` går via `KalkylError.Describe`. Fuzzern och `HostileFilesTests` kräver en felkod för varje fil som avvisas.
- **Läs XML på lokalt namn**, inte namnrymd, så att Strict OOXML fungerar. Ingen DTD.
- **Det som skrivs ska klara Microsofts schemavalidator.** Testet
  `A_file_with_every_feature_passes_Microsofts_schema_validation` ska täcka varje ny funktion
  i skrivaren.

## Versioner

0.x tills alla fem milstolparna är klara; första publika versionen blir 1.0.0. Ett publicerat
nummer byts aldrig ut (se `publish.ps1` och `release.ps1`).

## Konventioner

- **Explicita typer** — aldrig `var`
- **Filomfattande namnrymder**
- **Klammrar alltid**, även för en rad
- **Explicita bool-jämförelser** — `if (x == true)`, `if (x == false)`
- **Ingen `?:`** (AVIG0024) — skriv en if-sats eller ett switch-uttryck
- CRLF och UTF-8. Filer som skapas med verktyg som skriver LF måste konverteras, annars
  failar bygget på AVIG0006.
- **All XML-dokumentation och alla kommentarer är på engelska**, även på interna och privata
  medlemmar, i projektfiler och i skripten. Det gäller också README och `<Description>`.
  Bara den här filen och `docs/` är på svenska.

Avig.Analyzers kör med alla AVIG-regler på `error`.

## Publiceringen

Som i Avig.Result: `release.ps1` skjuter upp en `vX.Y.Z`-tagg, som kör
`.github/workflows/publish.yml` och publicerar via Trusted Publishing. GitHub-repot och
policyn på nuget.org är inte uppsatta än.
