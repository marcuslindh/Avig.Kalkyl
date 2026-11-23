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
```

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
- Allt som publiceras (README, `<Description>`, XML-dokumentationen på publika typer) är på
  engelska. Kommentarer i koden, skripten och den här filen är på svenska.

Avig.Analyzers kör med alla AVIG-regler på `error`.

## Publiceringen

Som i Avig.Result: `release.ps1` skjuter upp en `vX.Y.Z`-tagg, som kör
`.github/workflows/publish.yml` och publicerar via Trusted Publishing. GitHub-repot och
policyn på nuget.org är inte uppsatta än.
