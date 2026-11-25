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

- **Inga beroenden utom Avig.Result och Avig.OpenXml.** Inget zip-, XML- eller bildbibliotek:
  allt görs med `System.IO.Compression` och `XmlReader`/`XmlWriter`. `publish.ps1` failar annars.
  Testerna får använda vad de vill (DocumentFormat.OpenXml används för schemavalidering).
- **Paketlagret finns i Avig.OpenXml** (`G:\OneDrive\SkunkWorks\Prod\Lib\Avig.OpenXml`), som
  delas med Avig.Ord: zip-kontrollerna, storleksgränserna, XmlReader med djupgräns, relationer
  och sökvägar. `XmlParts` och `Ooxml` är tunna omslag kring det. Ett fel där rättas i
  Avig.OpenXml, och då ska både Kalkyls och Ords tester köras. Paketet är 0.x i
  `G:\LocalNugetRepo` (se `nuget.config`) tills det släpps med Kalkyl 1.0.0; det måste ligga på
  nuget.org innan en tagg pushas, eftersom publiceringsflödet inte når G:\.
- **0-baserat överallt** i det publika API:t. Excels adresser (C2) bara i `CellAddress.Parse`,
  `ToString` och i felmeddelanden, som alltid ska ange båda: "cell B2 (row 1, column 1)".
- **Förväntade fel är `Result`**, inte undantag: fel format, saknad fil, ogiltigt bladnamn.
  Undantag är för programmeringsfel (fel ordning, värden Excel inte kan lagra) och för en fil
  som visar sig vara skadad mitt i en ström.
- **Bara async** för allt som gör I/O, och `CancellationToken` förs alltid vidare.
- **Säkert för trimning och AOT** (`IsAotCompatible`). Ingen reflektion; den typade
  mappningen i milstolpe 2 blir en källgenerator.
- **En fil utifrån får inte kunna ta slut på minnet eller stacken.** Allt som läser en zip-post går via `XmlParts.OpenEntryAsync`, alla XmlReader skapas med `XmlParts.CreateReader` (djupgräns 100), och allt som skapar en matris går via `ValueArray` (`KalkylLimits.MaxArrayCells`). Ingen rekursion som växer med indata utan gräns: parsern begränsar nästling, och evaluatorn går igenom operatorkedjor i en loop. Aggregat över stora områden går via `FunctionCall.NonBlankValues` eller rutnätets `StoredValues`, inte position för position. Nya attackfiler blir tester i `HostileFilesTests.cs`; testplanen är `docs/testplan.md`.
- **Aktivt innehåll och strikt zip.** `KalkylLimits.StrictZip` (på som standard, också i
  `Unlimited`) gör att `XlsxReader.OpenSeekableAsync` kör `OpcZip.CheckStructureAsync` innan
  `ZipArchive` skapas. Alla vägar in (läsaren, dokumentmodellen, den typade läsningen) går dit,
  och en ny väg in måste också göra det. `PackageScan` går en gång igenom innehållstyper,
  relationer, externa länkar (DDE) och customXml (Power Query) och ger `Workbook.ActiveContent`,
  `Workbook.ExternalReferences` och `XlsxReader.GetActiveContentAsync`. Vad som räknas står i
  `ActiveContentRules`; enumvärdena delas med Avig.Ord. `RemoveActiveContent` går via
  `ActiveContentRemoval`: relationerna till aktivt innehåll markeras, de delar som fortfarande går
  att nå från roten behålls (allt annat följer inte med), och kvarvarande delar som pekade på
  något borttaget städas när de kopieras. Det som inte går att ta bort utan att arbetsboken går
  sönder (pivottabell eller datamodell på en anslutning, diagram från en annan arbetsbok) ger
  `Invalid` direkt i `RemoveActiveContent`. Bladlistan skrivs via `SheetList`, som behåller
  diagram-, makro- och dialogblad på sina platser så att `localSheetId` och `activeTab` stämmer.
  Testerna finns i `ActiveContentTests.cs` och `StrictZipTests.cs`; filerna byggs av
  `ActiveWorkbook`, som också ExcelCheck-fil 13 använder. En rensad fil ska öppnas i Excel utan
  reparation.
- **Ett fel om en fil ska få en felkod.** Meddelanden om skadade filer börjar med `The file is damaged`, meddelanden om gränser innehåller `the limit set in KalkylLimits.`, och meddelanden om filer som inte är .xlsx börjar med `The file is not an .xlsx workbook`. Det är så `KalkylError.ErrorKind()` känner igen dem. Ett `InvalidDataException` som blir ett `Result` går via `KalkylError.Describe`. Fuzzern och `HostileFilesTests` kräver en felkod för varje fil som avvisas.
- **Läs XML på lokalt namn**, inte namnrymd, så att Strict OOXML fungerar. Ingen DTD.
- **Det som skrivs ska klara Microsofts schemavalidator.** Testet
  `A_file_with_every_feature_passes_Microsofts_schema_validation` ska täcka varje ny funktion
  i skrivaren.
- **README-exemplen är tester.** Varje kodblock i README.md har en rad `<!-- example: namn -->`
  före sig och ska vara exakt koden i `#region example: namn` i
  `tests/Avig.Kalkyl.Tests/ReadmeExamples.cs`. Ändra koden där, kör testerna, och kopiera
  sedan in den i README. `Every_README_example_is_code_that_runs_here` failar om de skiljer sig,
  och varje csharp-block i README måste vara ett sådant exempel. Ett nytt publikt API ska få ett
  scenario och en rad i README:s "The API at a glance".
- **Diagram, bilder och villkorsstyrd formatering** samlas per blad i `SheetExtras`, som både
  `SheetWriter` och `Worksheet` använder, och skrivs av `DrawingXml`, `DrawingPackage` och
  `ConditionalFormatXml`. Ett blad har högst en ritning: i en öppnad fil läggs nya ankare till i
  den som finns, och nya regler läggs efter de befintliga med högre `priority`. Regelstilarna
  blir `dxf`-poster i `StyleRegistry`. Strict OOXML får inga ritningar. Testerna finns i
  `DrawingTests.cs`; Excel-checkfil 10–12 visar dem i riktiga Excel.
- **Former, grupper, SVG och diagramval.** `DrawingXml` skriver ankare (cellområde, cell med
  storlek, eller `absoluteAnchor` för `DrawingBounds`) och former, bilder och grupper;
  `ChartXml` skriver diagramdelarna. Under skrivningen samlar `DrawingContext` id:n för former,
  diagram, bilder och länkar (samma bild och samma länk delar en relation), och
  `DrawingPackage` skriver delarna efter ritningen, eftersom en zip bara har en post öppen.
  Varje element skrivs i schemats ordning, som Excel kontrollerar; ett diagramval som Excel
  skulle reparera (etikettposition som typen inte tar, typer som inte går att kombinera) ger
  `ArgumentException` i `Chart.Validate`. En länk inuti arbetsboken är relationen
  `#'Blad'!A1` utan `TargetMode`, som Excel skriver den. `SheetView` och cellänkar skrivs av
  `SheetXml`; i en öppnad fil skrivs `sheetViews`, `sheetPr` och `hyperlinks` om med filens
  eget element som utgångspunkt (`patches` i `WritePatchedSheetAsync`), så att det modellen
  inte känner till blir kvar. Testerna finns i `DashboardTests.cs`, dashboarden i
  `DashboardWorkbook.cs` (också ExcelCheck-fil 14–15), och fil 16 har varje val.
  `tools/Avig.Kalkyl.ExcelCheck/open-in-excel.ps1` öppnar filerna i Excel via COM och säger om
  Excel vägrar eller reparerar.

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
