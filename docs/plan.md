# Avig.Kalkyl — plan till 1.0

Alla fem milstolparna ska vara klara innan 1.0.0 släpps (beslut 2026-11-23, se `analys.md`).

| # | Milstolpe | Status |
| --- | --- | --- |
| 1 | Strömmande läsare och skrivare, stilar, datum | Klar 2026-11-23 |
| 2 | Typad mappning med källgenerator (`Read<T>`, `Write(list)`, `Result` per rad) | Klar 2026-11-23 |
| 3 | Dokumentmodell: öppna, ändra, spara, och bevara det paketet inte förstår | Klar 2026-11-23 |
| 4 | Formelmotorns grund: parser, beroendegraf, omräkning, cirkelreferenser, delade formler | Klar 2026-11-23 |
| 5 | Funktionsbiblioteket: 345 funktioner i tio kategorier | Klar 2026-11-23 |

## Arkitektur i korthet

- `Internal/` — datum, talformat, textkodning, sökvägar i paketet, och `CellXml`/`SheetXml` som
  skriver och tolkar celler och bladens omgivande element. Gemensamt för alla tre lägena.
- `Reading/`, `Writing/` — strömläget. `StyleTable` läser styles.xml, `StyleRegistry` skriver den
  (ny fil, eller tillägg till en öppnad fils stilar).
- `Mapping/` — körtidsdelen av den typade mappningen. Generatorn ligger i `Avig.Kalkyl.Generators`
  (netstandard2.0) och packas under `analyzers/dotnet/cs`.
- `Document/` — `Workbook`, `Worksheet`, `DocumentLoader` (behåller varje cells ursprungliga XML)
  och `DocumentSaver` (kopierar allt den inte äger, skriver om `sheetData`).
- `Formulas/` — lexer, parser, `Evaluator`, `FunctionCall`, `Recalculator` och en klass per
  funktionskategori under `Functions/`.

## Funktionsbiblioteket

333 funktioner från början (345 sedan databasfunktionerna lades till 2026-11-24): matematik och trigonometri (79), statistik och fördelningar (106), text (31), information (16), logik (11), datum och tid (25), uppslag och referens (19), finans (27) och teknik (23). Byggt av fem parallella agenter i var sin worktree, en per kategorigrupp, och sammanslaget utan konflikter. 1 490 tester. Kända avvikelser som agenterna själva redovisade:

- AREAS ger alltid 1 och INDEX tar bara område 1, eftersom motorn saknar referenser med flera områden.
- SUBTOTAL 101–111 beter sig som 1–11, eftersom dolda rader inte följs. AGGREGATE 15–19 ger #VALUE!.
- TEXT stöder inte bråkformat (# ?/?) eller villkorsavsnitt ([>100]).
- DOLLAR visar negativa belopp inom parentes, som i Microsofts exempel.
- LARGE/SMALL avrundar ett k med decimaler uppåt (som LibreOffice); det är inte kontrollerat mot Excel.

## Medvetna begränsningar

- Inget dynamiskt "spill": ett matrisresultat visar sitt första värde.
- Strukturerade tabellreferenser (`Tabell[Kolumn]`) ger `#REF!`; R1C1 i INDIRECT stöds inte;
  externa arbetsböcker stöds inte.
- Blad kan läggas till men inte tas bort eller döpas om i dokumentmodellen (formler och
  definierade namn som pekar på bladet skulle behöva skrivas om).
- Ändrar man rubrikcellerna i en Excel-tabell blir tabelldefinitionen inaktuell.
- Skrivaren använder inline-strängar; Excel gör om dem till delade strängar när filen sparas.

## Fuzzning

`tools/Avig.Kalkyl.Fuzz` (2026-11-23) har åtta mål: delar av en arbetsbok, formler, ett svep
över alla funktioner, flyttade formler, rundturer av värden, textkodningen, datumen och
celladresserna. Den första omgången hittade tretton fel, som alla är rättade och har
regressionstester i `FuzzFindingsTests.cs`. Tre av dem var dataförlust vid sparning av en fil
med element som paketet inte känner till eller med en relation utan Target. Varje mål är
verifierat genom sabotage. libFuzzer-läget är inte kört än, eftersom `libfuzzer-dotnet.exe`
saknas på byggmaskinen.

## Kompatibilitet med riktiga filer

`tests/Avig.Kalkyl.Compatibility` (2026-11-24) läser 38 arbetsböcker från Excel-handledningar med Avig.Kalkyl och fyra andra bibliotek, sparar och skriver kopior som de andra läser, och räknar om alla formler mot de värden Excel sparade. Första körningen hittade nio fel, som alla är rättade: bland annat en omräkning som aldrig blev klar (bladets storlek räknades fram vid varje områdesläsning), mellanslag som försvann i sparade textresultat, COUNTIF och CHOOSE med matriser, de saknade databasfunktionerna, och datumtext som alltid lästes amerikanskt. Det sista löstes med `Workbook.Culture`, som är maskinens kultur som standard, som i Excel (beslut 2026-11-24).

## Gränser för filer utifrån

`KalkylLimits` (2026-11-24) är påslagna från början: högst 1 GiB uppackat (kontrolleras mot zip-katalogen innan något läses, och varje del hålls till storleken i sitt huvud), högst 20 miljoner celler i dokumentmodellen och högst 10 miljoner värden i en matris i en formel (större ger #NUM!). Tid begränsas med `CancellationToken`, som nu också `Workbook.Evaluate` tar. Alltid på, som i Excel: högst 8 192 tecken per formel, 255 nivåer nästling och 64 funktioner i varandra (beslut 2026-11-24).

## Testplan

`docs/testplan.md` (2026-11-24) är testplanen för kvalitet och säkerhet. Den har funktionella tester, attackscenarier, prestandatester, fuzzning, beroenden, en riskmatris och acceptanskriterier för 1.0. Alla brister som hittades den dagen är åtgärdade samma dag. Det gäller lång text, XML-djup, dubbla zip-poster och namn utanför paketet, gränsen för delade strängar, aggregat över hela kolumner, makron (`HasMacros` och `RemoveMacros`), content types som saknas, allokering vid öppning, NuGet Audit och låsfiler. Fuzzern har fått målen `zip-raw` och `attribute-sweep` samt `corpus-check.ps1`. Det som återstår kräver Excel, libFuzzer, en isolerad miljö eller en ny lagring av cellerna för snabbare funktioner över hela kolumner (`docs/benchmarks/baseline.md`).

## Att göra före 1.0

- Öppna filer från skrivaren och dokumentmodellen i riktiga Excel och se att de öppnas utan reparation.
- Prova läsaren och dokumentmodellen på riktiga filer från LibreOffice och Google Sheets. Filer från Excel och WPS är provade (kompatibilitetstesterna).
- Jämför formelmotorn med Excel på fler riktiga arbetsböcker; 400 000 formler från handledningar stämmer nu.
- Uppfyll de acceptanskriterier i `docs/testplan.md`, avsnitt 8, som kräver en miljö: Excel, libFuzzer och belastningstester i en container.
- Skapa GitHub-repot och Trusted Publishing-policyn på nuget.org.
- Kör fuzzern täckningsstyrt med libFuzzer (`fuzz.ps1`), minst en natt för `xlsx-part` och `roundtrip`.
