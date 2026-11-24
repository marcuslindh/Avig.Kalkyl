# Prestandabaslinje

Mätt 2026-11-24 med `dotnet run -c Release --project tools/Avig.Kalkyl.Benchmarks -- --filter * --job short`.
Maskinen var byggmaskinen, Windows 11 och .NET 10, med fuzzern igång samtidigt. Siffrorna är
alltså ungefärliga, och en riktig baslinje ska mätas på en tyst maskin med standardjobbet.
Testplanen (`docs/testplan.md`, PRF-01) säger att en försämring över 20 % ska utredas.

| Benchmark | Storlek | Tid | Allokerat |
|---|---|---:|---:|
| `WriteAsync`: skriva med `XlsxWriter` | 100 000 rader × 10 kolumner | 856 ms | 132 MB |
| `StreamAsync`: läsa med `XlsxReader` | 100 000 × 10 | 691 ms | 917 MB |
| `OpenDocumentAsync`: `Workbook.OpenAsync` | 100 000 × 10 | 4 307 ms | 10 957 MB |
| `RecalculateChainAsync`: kedja av formler | 100 000 formler | 550 ms | 189 MB |
| `RecalculateSumIfAsync`: `SUMIF(B:B,Bn,A:A)` per rad | 10 000 formler över 10 000 rader | 20 415 ms | 1 915 MB |

Två mätningar visar var nästa förbättring finns:

- **`Workbook.OpenAsync`** allokerar cirka 11 kB per cell. Varje cells XML kopieras och läses
  med en egen `XmlReader`, med en buffert på 4 kB. Före rättningen 2026-11-24 var det 67 kB per
  cell, med en asynkron buffert på 64 kB och en `XmlWriter` per cell.
- **Funktioner över hela kolumner** kostar 200–280 ns per rad, både med den glesa och den
  täta vägen, eftersom varje cell ligger i en `SortedDictionary` per rad. 10 000 löpande SUMIF
  blir därför O(n²) och tar 20 sekunder. För att få ner tiden behövs en annan lagring av
  cellerna, till exempel arrayer per rad och kolumn. Det är en arkitekturändring och är inte gjord.
