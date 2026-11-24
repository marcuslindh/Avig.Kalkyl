# Prestandabaslinje

Mätt 2026-11-24 med `dotnet run -c Release --project tools/Avig.Kalkyl.Benchmarks -- --filter * --job short`.
Maskinen var byggmaskinen, Windows 11 och .NET 10, med fuzzern igång samtidigt. Siffrorna är
alltså ungefärliga, och en riktig baslinje ska mätas på en tyst maskin med standardjobbet.
Testplanen (`docs/testplan.md`, PRF-01) säger att en försämring över 20 % ska utredas.

| Benchmark | Storlek | Tid | Allokerat |
|---|---|---:|---:|
| `WriteAsync`: skriva med `XlsxWriter` | 100 000 rader × 10 kolumner | 877 ms | 132 MB |
| `StreamAsync`: läsa med `XlsxReader` | 100 000 × 10 | 697 ms | 917 MB |
| `OpenDocumentAsync`: `Workbook.OpenAsync` | 100 000 × 10 | 4 528 ms | 10 940 MB |
| `RecalculateChainAsync`: kedja av formler | 100 000 formler | 496 ms | 194 MB |
| `RecalculateSumIfAsync`: `SUMIF(B:B,Bn,A:A)` per rad | 10 000 formler över 10 000 rader | 2 744 ms | 1 916 MB |

`RecalculateSumIfAsync` tog 20 415 ms i den första mätningen. Två ändringar gav 7,4 gånger
snabbare omräkning:

- **Cellagringen.** Raderna ligger i en tabell som slås upp på index, och varje rads celler i
  sorterade arrayer (`Document/CellStore.cs`). Tidigare låg de i en `SortedDictionary` per rad.
  En cell slås nu upp på cirka 40 ns i stället för cirka 280 ns.
- **Formelindexet i `Recalculator`.** Det delar upp formlerna både per rad och per kolumn. Att
  hitta formlerna i `B:B` gick förut igenom varje formelrad. För 10 000 formler tog det 9 av
  12 sekunder, innan själva beräkningen började.

Det som är kvar är beräkningen: 10 000 formler läser 10 000 rader vardera. Det växer med
kvadraten på antalet rader, som i Excel. `Workbook.OpenAsync` allokerar fortfarande cirka 11 kB
per cell (se nedan).

Mätningar från 2026-11-24, före cellagringen:

| Benchmark | Tid | Allokerat |
|---|---:|---:|
| `OpenDocumentAsync` | 4 307 ms | 10 957 MB |
| `RecalculateChainAsync` | 550 ms | 189 MB |
| `RecalculateSumIfAsync` | 20 415 ms | 1 915 MB |

- **`Workbook.OpenAsync`** allokerar cirka 11 kB per cell. Varje cells XML kopieras och läses
  med en egen `XmlReader`, med en buffert på 4 kB. Före rättningen 2026-11-24 var det 67 kB per
  cell, med en asynkron buffert på 64 kB och en `XmlWriter` per cell.