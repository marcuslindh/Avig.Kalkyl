# Checklista: öppna filerna i riktiga Excel

Testfall FUN-04 i `docs/testplan.md`. Byggmaskinen har inte Excel, så det här är det enda
testet som en person gör. Det tar cirka 40 minuter. Gör det före en release och efter
ändringar i `DocumentSaver`, `XlsxWriter`, `Internal/CellXml.cs` eller ritningarna och reglerna
(`DrawingXml`, `DrawingPackage`, `ConditionalFormatXml`).

## 1. Skapa filerna

```powershell
dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check
dotnet run --project tools/Avig.Kalkyl.ExcelCheck -- C:\temp\excel-check --xlsm C:\sökväg\fil-med-makron.xlsm
```

Fil 07 och 08 kräver en egen `.xlsm` med makron. Paketet kan inte skapa ett VBA-projekt, så ta
en fil som du vet fungerar. Fil 09 kräver handledningsfilerna i
`tests/Avig.Kalkyl.Compatibility/files`. De laddas ner första gången kompatibilitetstesterna körs.

Generatorn skriver `expected.txt` med de värden som ska synas i varje fil. Alla filer är
kontrollerade mot OOXML-schemat innan du får dem: de var schemagiltiga 2026-11-24, och
rundturen gav inga nya schemafel i de 38 handledningsfilerna. Det som återstår är det schemat
inte ser, alltså hur Excel tolkar filerna.

## 2. Kontrollera varje fil

För **varje** fil gäller:

1. Öppna filen med dubbelklick. Det får **inte** komma en dialogruta med "Vi hittade ett problem
   med en del av innehållet" eller "Excel kunde inte öppna filen". En sådan dialogruta är ett
   fel, även om Excel sedan reparerar filen. Spara då Excels reparationslogg, som ligger i
   `%TEMP%` och nämns i dialogrutan.
2. Jämför med raden för filen i `expected.txt`.
3. För filer med formler: tryck **Ctrl+Alt+F9** (räkna om allt). Inget värde får ändras.
4. Stäng utan att spara.

| Fil | Kontrollera särskilt | Godkänd |
|---|---|---|
| `01-writer.xlsx` | Rad 1 låst och med filterknappar, kolumnbredder, talformat (`#,##0.00`, `0.0%`), datum med och utan klockslag, `#N/A`, rött teckensnitt. Bladnamnet `Blad 2 & 'citat' x` och formeln som pekar på det. | ☐ |
| `02-document.xlsx` | `SUMIF`, `COUNTIF`, `AVERAGEIFS` och `SUM` över hela kolumner. Värdena ska vara lika före och efter Ctrl+Alt+F9. `DATE(2026,2,29)` ger 2026-03-01. | ☐ |
| `03-edited.xlsx` | Samma format som 01 i den ändrade cellen B2. Den nya kolumnen E. Formler som beror på B2 är omräknade. | ☐ |
| `04-no-content-types.xlsx` | Öppnas utan reparation. Källan saknade `[Content_Types].xml`, som Excel kräver. | ☐ |
| `05-formula-like-text.xlsx` | Varje cell i kolumn A är **text**. Inget räknas ut, ingen länk skapas, och Excel varnar inte för externt innehåll eller DDE. | ☐ |
| `06-edge-values.xlsx` | 32 767 tecken i A1, och B1 visar 32767. Det största talet. Datumet 9999-12-31. `[h]:mm:ss` över ett dygn. Radbrytning, emoji. Ctrl+End går till XFD1048576. | ☐ |
| `07-macros-kept.xlsm` | Excel frågar om makron ska aktiveras, och makrona fungerar som i originalet. | ☐ |
| `08-macros-removed.xlsx` | Ingen makrovarning. Utvecklare > Visual Basic visar inget projekt för filen. | ☐ |
| `10-charts-writer.xlsx` | Sju diagram på bladet Data: stapel, liggande stapel (staplad), linje, yta (staplad), cirkel, ring och punkt, med titlar, axeltitlar och förklaringar som i `expected.txt`. Bilden i A10 i sin egen storlek och utsträckt över A16:C22. Reglerna i B, C och D och den gula fyllningen på raderna där D är under 50. Ändra B2 till 500: diagrammen följer med och B2 blir grön. Bladet Rapport har ett diagram över Data. | ☐ |
| `11-charts-document.xlsx` | Samma blad Data som 10, skrivet av dokumentmodellen. | ☐ |
| `12-added\*.xlsx` | Två rader under sista raden på första bladet: tre värden med blå datastaplar, ett stapeldiagram över dem och bilden. Filens egna diagram, bilder och regler finns kvar och ser ut som i originalet. | ☐ |
| `13-active-content-removed.xlsx` | Gjord av en arbetsbok med allt aktivt innehåll (VBA, makroblad, dialogblad, ActiveX, OLE, länk till en annan arbetsbok, DDE, länkad bild, frågetabell, Power Query) efter `RemoveActiveContent`. Öppnas utan reparation, utan makro- eller säkerhetsvarning, utan fråga om att uppdatera länkar och utan fråga om att uppdatera data. Bara bladet Data. B1 = 84 och C1 = `#N/A` som värden, inte formler. A1:A3 är en vanlig tabell. Knappen "Button 2" finns kvar. Namnhanteraren: `Rate` = `#REF!`, inget `Auto_Open`. Data > Frågor och anslutningar är tom. | ☐ |
| `14-dashboard-document.xlsx` | En mörk dashboard utan rutnät och rubriker, 90 % zoom, blå flik. Navigeringsspalten: "PROJECT DASHBOARD", en tunn linje och fyra rundade knappar med SVG-ikon; ett klick går till A1 på bladet, och varje blad har en länk tillbaka i A3. Fyra KPI-kort med toning, skugga och pil; det sista med inre skugga. Stapeldiagram med augusti turkos, etiketter i staplarna och planen som mjuk orange linje. Ringdiagram som mätare med "72 %" i mitten. Tratt med Leads överst. Mjuk linje över en blå tonad yta. Inga diagramramar eller vita bakgrunder. | ☐ |
| `15-dashboard-writer.xlsx` | Samma dashboard från den strömmande skrivaren; den öppnas på bladet Data, som skrivs först. | ☐ |
| `16-drawing-options.xlsx` | Fjorton former, en av varje sort, med tretonade fyllningar, de åtta streckmönstren, skuggor, inre skuggor, mjuka kanter och vridning; ett klick går till det andra bladet. En genomskinlig SVG-prick med länk. Tolv diagram som deras titlar säger, bland annat kombination med sekundär axel, etiketter i varje tillåten position, tre cirklar, mätare och punktdiagram. Bladet "Right to left" öppnas höger till vänster i 150 %. | ☐ |
| `09-roundtrip\*.xlsx` | Jämför med originalet i `tests/Avig.Kalkyl.Compatibility/files`: formatering, diagram, tabeller, pivottabeller, utsnitt, kommentarer, namn och dataverifiering. På första bladet finns en ny fet cell, "Avig.Kalkyl", under sista raden. | ☐ |

För `09-roundtrip` räcker det att öppna alla 38 filer och se att ingen ger en dialogruta. Jämför
sedan fem filer noga med sina original. Välj gärna filer med pivottabeller och diagram:
`class-pivot-slicer-kpi`, `fileformat-pivot-charts`, `financial-sample`, `customer-profitability`
och `data-validation-examples`.

## 3. Skriv ner resultatet

Fyll i tabellen nedan och lägg den i `docs/testplan.md`, under status för FUN-04. Ett fel blir ett
test i `CompatibilityFindingsTests.cs` eller `HostileFilesTests.cs` innan det rättas.

| Datum | Excel-version (Arkiv > Konto > Om Excel) | Godkända | Fel |
|---|---|---|---|
| 2026-11-24 | 16.0, build 20625 (automatiskt via COM) | Steg 1 och 3 för alla 84 filer: ingen reparation, och samma värden efter omräkning. Steg 2 och det visuella i tabellen är inte gjort. Fil 07 och 08 saknades, eftersom ingen `.xlsm` angavs. | Inga |
| 2026-11-25 | 16.0 (automatiskt via COM) | Steg 1 för alla 85 filer, med den nya fil 13: ingen reparation. Omräkning och det visuella är inte gjort. | Inga |
| 2026-11-25 | 16.0 (automatiskt via COM, `open-in-excel.ps1`) | Steg 1 för fil 01–16 (09 och 12 saknades, inga handledningsfiler i arbetskatalogen): ingen reparation. Fil 14–16 renderades via `CopyPicture` och granskades mot `expected.txt`. | Inga |
