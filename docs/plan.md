# Avig.Kalkyl — plan till 1.0

Alla fem milstolparna ska vara klara innan 1.0.0 släpps (beslut 2026-11-23, se `analys.md`).

| # | Milstolpe | Status |
| --- | --- | --- |
| 1 | Strömmande läsare och skrivare, stilar, datum | Klar 2026-11-23 |
| 2 | Typad mappning med källgenerator (`Read<T>`, `Write(list)`, `Result` per rad) | Klar 2026-11-23 |
| 3 | Dokumentmodell: öppna, ändra, spara, och bevara det paketet inte förstår | Klar 2026-11-23 |
| 4 | Formelmotorns grund: parser, beroendegraf, omräkning, cirkelreferenser, delade formler | Klar 2026-11-23 |
| 5 | Funktionsbiblioteket, 300+ funktioner | Pågår |

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

## Medvetna begränsningar

- Inget dynamiskt "spill": ett matrisresultat visar sitt första värde.
- Strukturerade tabellreferenser (`Tabell[Kolumn]`) ger `#REF!`; R1C1 i INDIRECT stöds inte;
  externa arbetsböcker stöds inte.
- Blad kan läggas till men inte tas bort eller döpas om i dokumentmodellen (formler och
  definierade namn som pekar på bladet skulle behöva skrivas om).
- Ändrar man rubrikcellerna i en Excel-tabell blir tabelldefinitionen inaktuell.
- Skrivaren använder inline-strängar; Excel gör om dem till delade strängar när filen sparas.

## Att göra före 1.0

- Öppna filer från skrivaren och dokumentmodellen i riktiga Excel och se att de öppnas utan reparation.
- Prova läsaren och dokumentmodellen på riktiga filer från Excel, LibreOffice och Google Sheets.
- Jämför formelmotorn med Excel på en större uppsättning riktiga arbetsböcker.
- Skapa GitHub-repot och Trusted Publishing-policyn på nuget.org.
