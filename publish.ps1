#Requires -Version 7.0
<#
.SYNOPSIS
    Bygger, testar, packar och formkontrollerar Avig.Kalkyl.

.DESCRIPTION
    Skriptet pushar aldrig sjalvt. Publiceringen till nuget.org gors av
    .github/workflows/publish.yml nar en v-tagg skjuts upp, och den taggen satts av
    release.ps1. Det har skriptet stannar vid ett verifierat paket i det lokala flodet.

    Formkontrollen ar sjalva poangen. Ett av skalen till att Avig.Kalkyl finns ar att NPOI
    och ClosedXML drar in fem till tio paket var. Kontrollen failar darfor pa allt utom ett
    paket med lib/net10.0/Avig.Kalkyl.dll, en README och exakt ett beroende: Avig.Result.

.PARAMETER NoBump
    Behall nuvarande version i stallet for att hoja patch-numret. Workflowen kor med den
    har flaggan: versionen ar redan satt och taggad av release.ps1.

.PARAMETER SkipTests
    Hoppa over testkorningen. Anvand nar du redan kort dem.

.PARAMETER SkipVerify
    Hoppa over formkontrollen av nupkg:en. Gor inte det utan ett specifikt skal.

.PARAMETER Feed
    Katalog att kopiera det byggda paketet till for lokal rokprovning.
#>
param(
    [switch]$NoBump,
    [switch]$SkipTests,
    [switch]$SkipVerify,
    [string]$Feed = "G:\LocalNugetRepo"
)

$ErrorActionPreference = "Stop"

$root    = $PSScriptRoot
$project = Join-Path $root "src\Avig.Kalkyl\Avig.Kalkyl.csproj"
$tests   = Join-Path $root "tests\Avig.Kalkyl.Tests\Avig.Kalkyl.Tests.csproj"
$packDir = Join-Path $root "src\Avig.Kalkyl\bin\Release"

Write-Host ""
Write-Host "=== Avig.Kalkyl ===" -ForegroundColor Cyan
Write-Host ""

# --- [1/5] Version ---
Write-Host "[1/5] Version" -ForegroundColor Yellow

# XmlDocument.Load laser filens egen encoding-deklaration. Get-Content utan -Encoding gor
# det inte: i Windows PowerShell 5.1 avkodas UTF-8 som ANSI och Save skriver tillbaka som
# UTF-8, sa varje a-ring, a-umlaut, o-umlaut och tankstreck fordubblas vid varje bump. Sa
# vaxte Medusa.Portal.csproj fran 2 kB till 1,1 MB pa sju versionsbumpar.
$csproj = New-Object System.Xml.XmlDocument
$csproj.PreserveWhitespace = $true
$csproj.Load((Resolve-Path $project))

$currentVersion = [string](($csproj.Project.PropertyGroup | Where-Object { $_.Version }).Version)
if (-not $currentVersion) {
    Write-Host "      FEL: <Version> saknas i $project" -ForegroundColor Red
    exit 1
}

if ($NoBump) {
    $version = $currentVersion
    Write-Host "      $version (ingen bump)" -ForegroundColor Yellow
} else {
    $parts = $currentVersion.Split('.')
    $version = "$($parts[0]).$($parts[1]).$([int]$parts[2] + 1)"
    foreach ($pg in $csproj.Project.PropertyGroup) {
        if ($pg.Version) { $pg.Version = $version }
    }
    $csproj.Save((Resolve-Path $project))
    Write-Host "      $currentVersion -> $version" -ForegroundColor Yellow
}

# Ett redan publicerat paket far aldrig bytas ut under samma versionsnummer: den som
# cachat det behaller sin kopia och kor annan kod an den som hamtar nytt. Tva olika
# Avig.Kalkyl i omlopp utan att nagot syns i bygget. Galler aven efter en bump - ligger den
# bumpade versionen redan i flodet har csproj:en och flodet glidit isar, och att packa
# over ar da precis fel atgard.
$existing = Join-Path $Feed "Avig.Kalkyl.$version.nupkg"
if (Test-Path $existing) {
    Write-Host "      FEL: Avig.Kalkyl $version finns redan i flodet." -ForegroundColor Red
    if ($NoBump) {
        Write-Host "           Hoj <Version> i stallet for att skriva over den." -ForegroundColor Red
    } else {
        Write-Host "           Bumpen landade pa en version som redan ar publicerad." -ForegroundColor Red
        Write-Host "           Satt <Version> i csproj:en till nasta lediga nummer." -ForegroundColor Red
    }
    exit 1
}

# --- [2/5] Tester ---
if ($SkipTests) {
    Write-Host "[2/5] Tester hoppas over" -ForegroundColor DarkYellow
} else {
    Write-Host "[2/5] Testar..." -ForegroundColor Yellow
    dotnet test $tests -c Release --nologo -v quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      Tester misslyckades - inget packat." -ForegroundColor Red
        exit 1
    }
    Write-Host "      Alla tester grona" -ForegroundColor Green
}

# --- [3/5] Pack ---
Write-Host "[3/5] Packar..." -ForegroundColor Yellow

# Gamla nupkg:er i bin\Release far inte ligga kvar och forvaxlas med den nya.
if (Test-Path $packDir) {
    Get-ChildItem -Path $packDir -Filter "*.nupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -Path $packDir -Filter "*.snupkg" -ErrorAction SilentlyContinue | Remove-Item -Force
}

dotnet pack $project -c Release --nologo -v quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "      Pack misslyckades" -ForegroundColor Red
    exit 1
}

$package = Get-ChildItem -Path $packDir -Filter "Avig.Kalkyl.$version.nupkg" | Select-Object -First 1
if (-not $package) {
    Write-Host "      Hittade inget paket for version $version" -ForegroundColor Red
    exit 1
}
Write-Host "      $($package.Name)" -ForegroundColor Green

# --- [4/5] Verifiera paketets form ---
if ($SkipVerify) {
    Write-Host "[4/5] Formkontroll hoppas over" -ForegroundColor DarkYellow
} else {
    Write-Host "[4/5] Verifierar paketets innehall..." -ForegroundColor Yellow

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($package.FullName)
    try {
        $entries = $zip.Entries | ForEach-Object { $_.FullName }
        $nuspecEntry = $zip.Entries | Where-Object { $_.FullName -eq "Avig.Kalkyl.nuspec" }
        if (-not $nuspecEntry) {
            Write-Host "      FEL: nuspec saknas i paketet" -ForegroundColor Red
            exit 1
        }
        $reader = New-Object System.IO.StreamReader($nuspecEntry.Open())
        try { $nuspec = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $zip.Dispose()
    }

    if (-not ($entries | Where-Object { $_ -eq "lib/net10.0/Avig.Kalkyl.dll" })) {
        Write-Host "      FEL: lib/net10.0/Avig.Kalkyl.dll saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # Kallgeneratorn for [KalkylRow] maste folja med, annars kompilerar ingen typad mappning
    # hos den som anvander paketet.
    if (-not ($entries | Where-Object { $_ -eq "analyzers/dotnet/cs/Avig.Kalkyl.Generators.dll" })) {
        Write-Host "      FEL: analyzers/dotnet/cs/Avig.Kalkyl.Generators.dll saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # Avig.Result ar det enda tillatna beroendet. Allt annat - en zip-, XML- eller
    # bildbibliotek - ar precis det som gor NPOI och ClosedXML tunga att dra in.
    $deps = @($nuspec.package.metadata.dependencies.group.dependency) | Where-Object { $_ -ne $null }
    $foreign = @($deps | Where-Object { $_.id -ne "Avig.Result" })
    if ($foreign.Count -gt 0) {
        Write-Host "      FEL: paketet drar in andra beroenden an Avig.Result:" -ForegroundColor Red
        $foreign | ForEach-Object { Write-Host "        $($_.id) $($_.version)" -ForegroundColor Red }
        exit 1
    }

    if (-not ($deps | Where-Object { $_.id -eq "Avig.Result" })) {
        Write-Host "      FEL: beroendet pa Avig.Result saknas i nuspec:en" -ForegroundColor Red
        exit 1
    }

    # README:n ar det enda den som hittar paketet i ett flode har att ga pa.
    if (-not ($entries | Where-Object { $_ -eq "README.md" })) {
        Write-Host "      FEL: README.md saknas i paketet" -ForegroundColor Red
        exit 1
    }

    # Licensen avgor om nagon far anvanda paketet. Utan den ar det oklart.
    if (-not $nuspec.package.metadata.license) {
        Write-Host "      FEL: licens saknas i nuspec:en" -ForegroundColor Red
        exit 1
    }

    Write-Host "      lib/net10.0/Avig.Kalkyl.dll, README.md och licens finns, bara Avig.Result som beroende" -ForegroundColor Green
}

# --- [5/5] Kopiera till det lokala flodet ---
Write-Host "[5/5] Kopierar till flodet..." -ForegroundColor Yellow
if (-not (Test-Path $Feed)) {
    New-Item -Path $Feed -ItemType Directory | Out-Null
}

Get-ChildItem -Path $packDir -Filter "Avig.Kalkyl.$version.*nupkg" | ForEach-Object {
    Copy-Item $_.FullName -Destination $Feed -Force
    Write-Host "      $Feed\$($_.Name)" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Avig.Kalkyl v$version ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Paketet ligger i det lokala flodet. Till nuget.org gar det via en tagg:" -ForegroundColor Gray
Write-Host "    ./release.ps1 -Patch -Push" -ForegroundColor Gray
