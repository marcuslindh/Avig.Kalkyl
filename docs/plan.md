# Avig.Kalkyl — plan till 1.0

Alla fem milstolparna ska vara klara innan 1.0.0 släpps (beslut 2026-11-23, se `analys.md`).

| # | Milstolpe | Status |
| --- | --- | --- |
| 1 | Strömmande läsare och skrivare, stilar, datum | **Klar 2026-11-23** |
| 2 | Typad mappning med källgenerator (`Read<T>`, `Write(list)`, `Result` per rad) | Inte påbörjad |
| 3 | Dokumentmodell: öppna, ändra, spara, och bevara det paketet inte förstår | Inte påbörjad |
| 4 | Formelmotorns grund: parser, beroendegraf, omräkning, cirkelreferenser, delade formler | Inte påbörjad |
| 5 | Funktionsbiblioteket, 200+ funktioner, i omgångar om cirka 50 | Inte påbörjad |

## Milstolpe 1 — vad som finns

**Skriva** (`XlsxWriter`, `SheetWriter`): flera blad, text (inline), tal, booleska värden,
datum, felvärden, formler med eller utan sparat resultat, stilar per cell (fetstil, kursiv,
understrykning, färg, storlek, typsnitt, fyllning, talformat, justering, radbrytning) med
automatisk sammanslagning, kolumnbredd, frysta rutor, autofilter, kontroll av bladnamn.
Tomma arbetsböcker får ett tomt blad.

**Läsa** (`XlsxReader`, `SheetReader`, `Row`): delade strängar (även formaterad text och
fonetisk text som hoppas över), inline-strängar, `t="str"`, `t="d"`, booleska värden, fel, tal,
datum via inbyggda och egna talformat, 1900- och 1904-systemet, formeltext och cachat värde,
glesa rader och celler utan adress, dolda blad, Strict OOXML, absoluta och relativa sökvägar
i relationer, strömmar som inte kan söka.

**Kontrollerat:** 108 tester, bland annat att en fil med alla funktioner klarar Microsofts
schemavalidator (DocumentFormat.OpenXml). Inte öppnat i Excel: varken Excel eller LibreOffice
fanns på maskinen där milstolpen byggdes.

**Medvetna begränsningar i milstolpe 1:**

- När filen läses fylls bara talformatet i `Cell.Style`. Övrig formatering kommer med
  dokumentmodellen.
- Följande celler i en delad formel får `Formula = null` (värdet läses). Formeltexten räknas
  fram i milstolpe 4.
- Skrivaren använder inline-strängar och inte delade strängar. Det håller minnet konstant men
  ger något större filer när samma text upprepas. Excel gör om dem till delade strängar när
  filen sparas.
- Diagrambladen listas inte i `XlsxReader.Sheets`.

## Att göra före nästa milstolpe

- Öppna en fil från skrivaren i riktiga Excel och se att den öppnas utan reparation.
- Prova läsaren på några riktiga filer från Excel, LibreOffice och Google Sheets.
- Skapa GitHub-repot och Trusted Publishing-policyn på nuget.org när 1.0 närmar sig.
