# Avig + Excel — analys av förebilderna och val för API:t

## Valt (2026-11-23)

| Fråga | Val |
| --- | --- |
| Namn | **Avig.Kalkyl** (namespace `Avig.Kalkyl`) |
| Modell i version 1 | **Alla tre**: ström (B), typat (C) och dokument (A) |
| Fel | **Result** från Avig.Result för förväntade fel, och `TryGet…` för cellvärden |
| Index | **0-baserat**: `sheet[1, 2]` är cell C2 i Excel. `"A1"`-adresser stöds också. |
| Format | **Bara xlsx** |
| Cellvärden | **Union + TryGet**: `CellValue` som struct med `Kind` (Text, Number, Bool, Date, Error, Empty) och `TryGetNumber`/`TryGetText`/`TryGetDate` med `[NotNullWhen]` |
| Stilar | **Flytande per cell**, och paketet slår ihop likadana stilar automatiskt |
| Datum | **Tolkas automatiskt** via talformatet, till `DateTime`/`DateOnly`. Skottårsfelet för 1900 och 1904-systemet hanteras. |
| Formler | **Formelmotor**: paketet räknar ut formler självt |
| Beroenden | **net10.0 och bara Avig.Result**. Säkert för trimning och Native AOT, och `publish.ps1` kontrollerar det. |
| Typad mappning | **Källgenerator direkt**: ett attribut på klassen genererar mappningen vid kompilering, och ingen reflektion används |
| Sync/async | **Bara async**, med `CancellationToken` överallt och `IAsyncEnumerable` för läsning |
| Formelmotor: omfång | **Bred täckning, 200+ funktioner**, ungefär som EPPlus, inklusive finansiella och statistiska |
| Formelmotor: okänd funktion | **`#NAME?`** i cellen, som i Excel |
| Formelmotor: när | **När man ber om det**: `await book.RecalculateAsync()` räknar om det som påverkats |
| Formelmotor: paket | **Samma paket** som resten av Avig.Kalkyl |

Följder av valen:

- **Byggordning:** dokumentmodellen byggs ovanpå samma läsare och skrivare som strömläget. Därför
  blir läsaren och skrivaren klara först, även om alla tre ska finnas i 1.0.
- **Redigera befintliga filer:** det är det svåraste med dokumentmodellen. Delar av filen som
  paketet inte förstår (diagram, villkorsstyrd formatering, bilder) måste följa med orörda när
  filen sparas.
- **0-baserat index:** det skiljer sig från radnumren som Excel visar. Felmeddelanden i det typade
  läget ska därför ange både Excels adress och index, till exempel "C2 (rad 1, kolumn 2)", så att
  användaren hittar raden i Excel.
- **Källgeneratorn:** den måste byggas för `netstandard2.0` (det är ett krav för analysatorer,
  precis som i Avig.Analyzers) och skickas med i paketet under `analyzers/dotnet/cs`. Själva
  biblioteket är fortfarande net10.0.
- **Formelmotorn:** den är den största delen av paketet. Den behöver en egen parser för formelspråket,
  en beroendegraf mellan cellerna, ordning för omräkning, cirkelreferenser och ett funktionsbibliotek.
  Excel har över 500 funktioner, så omfånget måste bestämmas.
- **Bara async:** dokumentmodellen öppnas och sparas asynkront (`await Workbook.OpenAsync(...)`).
  Att läsa och sätta enskilda celler i minnet är vanliga egenskaper och metoder, eftersom det inte
  finns någon I/O att vänta på.

*2026-11-23. Underlag för att välja hur ett Avig-paket för att läsa och skriva Excel-filer ska se ut.
Versioner och beroenden hämtade från nuget.org samma dag.*

## Varför ett eget paket

- **NPOI** (2.8.1) har källkoden under Apache 2.0, men den färdiga NuGet-binären levereras nu med
  ett "Open Source Maintenance Fee Agreement" (`OSMFEULA.txt` i paketet). Organisationer med en
  omsättning på minst 10 000 USD som använder den i intäktsbringande verksamhet ska betala en
  månadsavgift, om de inte bygger NPOI själva från källkoden.
- **EPPlus** (8.7.3) har sedan version 5 en egen licens som kräver en kommersiell licens för
  kommersiellt bruk, och koden måste ange licensen innan första användning. KLM-portal använder EPPlus.
- **Beroenden:** NPOI drar in tio paket (BouncyCastle, SkiaSharp med inbyggda bibliotek för Linux,
  MathNet, SharpZipLib med flera), och ClosedXML drar in fem (bland annat DocumentFormat.OpenXml och
  SixLabors.Fonts). Det går att läsa och skriva xlsx med bara det som finns i .NET:
  `System.IO.Compression` och `XmlReader`/`XmlWriter`.
- **Stilen:** alla förebilder signalerar fel med undantag eller `null`. Avig-stilen är `Result` och
  `TryGet` med `[NotNullWhen]`.

## Förebilderna

| Paket | Version | Licens | Modell | Läser | Skriver | Beroenden |
| --- | --- | --- | --- | --- | --- | --- |
| NPOI | 2.8.1 | Apache 2.0 + avgiftsavtal för binären | Objektmodell (port av Javas Apache POI) | xls, xlsx | xls, xlsx | 10 |
| ClosedXML | 0.105.1 | MIT | Objektmodell, flytande API | xlsx | xlsx | 5 |
| EPPlus | 8.7.3 | Egen, kommersiell | Objektmodell + formelmotor | xlsx | xlsx | 6 |
| MiniExcel | 2.0.0 | Apache 2.0 | Strömmande, typad mappning | xlsx, csv | xlsx, csv | 3 (egna delpaket) |
| ExcelDataReader | 3.9.0 | MIT | `IDataReader`, rad för rad | xls, xlsx, csv | — | — |
| Sylvan.Data.Excel | 0.5.8 | MIT | `DbDataReader`, mycket snabb | xls, xlsx, xlsb | xlsx | — |
| LargeXlsx | 2.0.2 | MIT | Strömmande skrivare | — | xlsx | — |
| DocumentFormat.OpenXml | 3.5.1 | MIT | Lågnivå: XML-delarna som klasser | xlsx | xlsx | — |

### NPOI: objektmodell i Java-stil

```csharp
IWorkbook wb = new XSSFWorkbook();                 // HSSFWorkbook för .xls
ISheet sheet = wb.CreateSheet("Kunder");
IRow row = sheet.CreateRow(0);                     // 0-baserat
row.CreateCell(0).SetCellValue("Namn");

ICellStyle bold = wb.CreateCellStyle();            // stilar skapas på arbetsboken
IFont font = wb.CreateFont();
font.IsBold = true;
bold.SetFont(font);
row.GetCell(0).CellStyle = bold;

using var fs = File.Create("kunder.xlsx");
wb.Write(fs);

// Läsa: rader och celler som saknas är null
ICell? cell = new XSSFWorkbook(File.OpenRead("kunder.xlsx")).GetSheetAt(0).GetRow(1)?.GetCell(2);
double belopp = cell?.CellType == CellType.Numeric ? cell.NumericCellValue : 0;
```

- **Bra:** täcker nästan allt (xls, formler, bilder, diagram, kryptering) och är beprövat i 15 år.
  Hierarkin arbetsbok → blad → rad → cell är lätt att förstå.
- **Mindre bra:** pratigt. `GetRow` och `GetCell` ger `null`. Cellens typ måste kontrolleras innan
  värdet läses, annars kommer ett undantag. Stilar måste skapas på arbetsboken och återanvändas för
  hand, eftersom en ny stil per cell slår i taket på cirka 64 000 stilar. Hela filen ligger i minnet;
  för stora filer finns den separata strömmande varianten SXSSF med ett annat API. Både xls-
  och xlsx-klasserna finns kvar under samma gränssnitt.

### ClosedXML: flytande objektmodell

```csharp
using var wb = new XLWorkbook();
var ws = wb.Worksheets.Add("Kunder");
ws.Cell("A1").Value = "Namn";                      // både "A1" och (rad, kolumn), 1-baserat
ws.Cell(1, 1).Style.Font.Bold = true;              // stil per cell, dubbletter slås ihop automatiskt
ws.Cell(2, 1).InsertTable(kunder);                 // lista → Excel-tabell
ws.Columns().AdjustToContents();
wb.SaveAs("kunder.xlsx");

using var wb2 = new XLWorkbook("kunder.xlsx");
double belopp = wb2.Worksheet(1).Cell("C2").GetValue<double>();
foreach (var r in wb2.Worksheet(1).RowsUsed().Skip(1)) { /* ... */ }
```

- **Bra:** det trevligaste API:t att skriva, med intervall (`Range("A1:C3")`), tabeller och
  kolumnbredd efter innehåll. Sedan 0.100 är `XLCellValue` en typad union i stället för `object`.
- **Mindre bra:** hela filen ligger i minnet och det kan bli flera GB för stora filer.
  `GetValue<T>` kastar undantag vid fel typ. API:t är stort och olika delar uppför sig olika.

### EPPlus: objektmodell med formelmotor

```csharp
// En licensinställning krävs innan första användning
using var p = new ExcelPackage();
var ws = p.Workbook.Worksheets.Add("Kunder");
ws.Cells["A1"].LoadFromCollection(kunder, true);   // rubriker från egenskaperna
ws.Cells["A1:C1"].Style.Font.Bold = true;          // stil på ett intervall
p.SaveAs(new FileInfo("kunder.xlsx"));
```

- **Bra:** intervall som första klassens begrepp (`Cells["A1:C1"]`), `LoadFromCollection` och
  `ToCollection`, och en riktig formelmotor som räknar ut formler.
- **Mindre bra:** licensen. Allt annat liknar ClosedXML.

### MiniExcel: strömmande och typat (syntax från 1.x)

```csharp
MiniExcel.SaveAs("kunder.xlsx", kunder);                    // rubriker från egenskaper eller [ExcelColumnName]
IEnumerable<Kund> rader = MiniExcel.Query<Kund>("kunder.xlsx");  // lat: en rad i minnet åt gången
```

- **Bra:** minsta möjliga API för det vanligaste fallet, alltså lista till fil och fil till lista.
  Låg minnesanvändning. Mallar med `{{Namn}}` i en befintlig xlsx.
- **Mindre bra:** det går inte att redigera en befintlig fil cell för cell, och formatering är
  begränsad. Mappningen bygger på reflektion.

### ExcelDataReader och Sylvan: läsare som DataReader

```csharp
using var reader = ExcelReaderFactory.CreateReader(stream);    // ExcelDataReader
while (reader.Read()) { string namn = reader.GetString(0); }

using var edr = Sylvan.Data.Excel.ExcelDataReader.Create("kunder.xlsx");   // Sylvan: en DbDataReader
```

- **Bra:** snabbast och minst minne. Passar direkt in i ADO.NET, till exempel `SqlBulkCopy` och Dapper.
- **Mindre bra:** bara läsning (Sylvan har en enkel skrivare), och ett lågnivå-API.

### LargeXlsx: strömmande skrivare

```csharp
using var xlsx = new XlsxWriter(stream);
xlsx.BeginWorksheet("Kunder").BeginRow().Write("Namn").Write("Belopp");
foreach (var k in kunder) xlsx.BeginRow().Write(k.Namn).Write(k.Belopp);
```

- **Bra:** liten, snabb och konstant minne, med stilar, sammanslagna celler och fryst rubrikrad.
- **Mindre bra:** kan bara skriva framåt. Man kan inte gå tillbaka till en cell som redan är skriven.

## Tre möjliga modeller för Avig

Alla exempel nedan är skisser med preliminära namn.

### A. Dokument (som ClosedXML): öppna, ändra, spara

```csharp
var book = Workbook.Create();
var sheet = book.AddSheet("Kunder");
sheet["A1"].Set("Namn").Bold();
sheet[2, 3].Set(42.5m).NumberFormat("# ##0.00");
sheet.FreezeRows(1);
book.Save("kunder.xlsx");

Result<Workbook> opened = Workbook.Open("kunder.xlsx");
if (opened.IsFailure(out string? error, out _)) { /* inte en xlsx, låst fil ... */ }
if (sheet["C2"].TryGetNumber(out double belopp)) { /* ... */ }
```

Mest flexibel. Svårast att göra rätt: för att redigera en befintlig fil måste allt som paketet inte
förstår (diagram, villkorsstyrd formatering, makron) bevaras orört när filen sparas.

### B. Ström (som LargeXlsx och Sylvan): rad för rad

```csharp
await using var writer = XlsxWriter.Create("kunder.xlsx");
var sheet = writer.AddSheet("Kunder");
sheet.WriteRow(Style.Bold, "Namn", "Ort", "Belopp");
foreach (var k in kunder) sheet.WriteRow(k.Namn, k.Ort, k.Belopp);

await foreach (Row row in XlsxReader.ReadRows("kunder.xlsx", "Kunder"))
{
    if (row.TryGetNumber(2, out double belopp)) { /* ... */ }
}
```

Snabb och konstant minne. Enklast att få korrekt. Täcker export av rapporter och import av tabeller.

### C. Typat (som MiniExcel och LoadFromCollection): lista ↔ fil

```csharp
await Xlsx.Write("kunder.xlsx", kunder);                   // rubriker från egenskaper eller [Column("Namn")]

await foreach (Result<Kund> rad in Xlsx.Read<Kund>("kunder.xlsx"))
{
    if (rad.IsFailure(out string? error, out _)) { /* "rad 17, Belopp: 'abc' är inget tal" */ }
}
```

Minst kod för det vanligaste fallet. Med ett `Result` per rad syns felaktiga rader utan att hela
importen avbryts. Det har ingen av förebilderna.

### Rekommendation

**Version 1 = B som kärna och C ovanpå. A kommer i version 2 och bygger på samma läsare och skrivare.**
Det mesta du gör med Excel är att exportera listor och importera tabeller. B och C täcker det med
en bråkdel av arbetet, och dokumentmodellen kan läggas ovanpå senare utan att API:t bryts.

## Detaljval

| # | Fråga | Alternativ | Rekommendation |
| --- | --- | --- | --- |
| 1 | Format | bara xlsx · xlsx + csv · även gamla xls | **Bara xlsx** i version 1. xls är ett binärformat som kostar mycket. csv kan bli ett eget litet paket. |
| 2 | Index | 1-baserat som Excel (ClosedXML, EPPlus) · 0-baserat som C# (NPOI) | **1-baserat**, plus `"A1"`-adresser. Då betyder `sheet[2, 3]` det som står i Excel, alltså C2. |
| 3 | Cellvärden | `object Value` · typad union (`CellValue`) · `TryGetNumber`/`TryGetText`/`TryGetDate` | **Union + TryGet**: `CellValue` som struct med `Kind`, och `TryGet…` med `[NotNullWhen]` som i Avig.Result |
| 4 | Fel | undantag · `Result` från Avig.Result | **Result** för förväntade fel (fel format, blad saknas, värde går inte att tolka). Undantag bara för programmeringsfel. Paketet beror då på Avig.Result. |
| 5 | Stilar | stilobjekt på arbetsboken som återanvänds (NPOI) · flytande per cell med automatisk sammanslagning (ClosedXML) | **Flytande, och sammanslagningen sköts av paketet**, så att ingen kan slå i taket på cirka 64 000 stilar |
| 6 | Datum | som tal (Excels interna) · tolka automatiskt via talformatet | **Automatiskt**: `DateTime`/`DateOnly` när cellens format är ett datumformat. Det ska också hantera Excels skottårsfel för 1900 och 1904-systemet. |
| 7 | Formler | ingen · skriv formeltext och läs cachat värde · formelmotor | **Skriv formeltext och läs Excels cachade värde.** En formelmotor är ett eget projekt. |
| 8 | Beroenden | fritt · bara .NET | **Bara .NET (och Avig.Result)**, med net10.0 och säkert för trimning och Native AOT |
| 9 | Typad mappning | reflektion · källgenerator | **Reflektion med `[DynamicallyAccessedMembers]`** i version 1, och en källgenerator om det behövs |
| 10 | Sync/async | bara sync · bara async · båda | **Båda**: `IAsyncEnumerable` för läsning, och synkrona varianter för enkla fall |

## Vad version 1 behöver kunna

**Läsa:** delade strängar, inline-strängar, tal, booleska värden, fel (`#N/A`), datum via
talformat, formler (cachat värde), flera blad, tomma celler och glesa rader (`r="C7"` utan
mellanliggande celler), och rubrikrad till egenskapsnamn.

**Skriva:** flera blad, text, tal, datum, booleska värden och formler, fetstil, talformat,
kolumnbredd, fryst rubrikrad, autofilter och en Excel-tabell från en lista.

**Inte i version 1:** redigering av befintliga filer, bilder, diagram, villkorsstyrd formatering,
lösenordsskydd, xls och formelberäkning.

## Namn

Alla nedan var lediga på nuget.org 2026-11-23.

| Namn | För | Emot |
| --- | --- | --- |
| **Avig.Xlsx** | Säger exakt vad det är. Inget varumärke. Namespace `Avig.Xlsx`. | Låter tekniskt, och "xlsx" är svårt att säga |
| **Avig.Sheets** | Vänligt. Rymmer csv senare. | Kan förväxlas med Google Sheets |
| **Avig.Cells** | Kort | Aspose.Cells finns redan som stort kommersiellt paket |
| **Avig.Excel** | Lättast att hitta i NuGet-sökningar | "Excel" är Microsofts varumärke. Många paket använder det (MiniExcel, ExcelDataReader), men risken är inte noll. |
| **Avig.Kalkyl** | Eget och svenskt ("kalkylblad"), sticker ut | Säger inget för den som inte kan svenska |
| **Avig.Ledger** | Ordet för "huvudbok", ger en bild | Låter som bokföring eller blockkedja |
| **Avig.Tabula** | Latin för tavla eller tabell, och lätt att säga | Abstrakt, och finns som PDF-verktyg (Tabula) |

**Min rekommendation:** `Avig.Xlsx` om namnet ska säga vad paketet gör, i samma anda som
Avig.Result och Avig.Analyzers. `Avig.Sheets` om du vill ha ett vänligare namn som rymmer csv.
