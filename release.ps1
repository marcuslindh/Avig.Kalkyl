#Requires -Version 7.0
<#
.SYNOPSIS
    Gor en release: hojer versionen, verifierar paketet, commitar och taggar.

.DESCRIPTION
    Taggen ar det som publicerar. En uppskjuten vX.Y.Z kor .github/workflows/publish.yml,
    som bygger paketet igen pa en runner och pushar det till nuget.org via Trusted
    Publishing.

    Skriptet gor allt fram till dit och stannar, for en tagg som inte lamnat maskinen gar
    fortfarande att ta bort medan en publicerad version aldrig gar att byta ut. Med -Push
    gar det hela vagen i ett kommando.

    Paketet packas och formkontrolleras INNAN taggen skapas, sa en release som inte gar att
    bygga aldrig blir en tagg. Misslyckas kontrollen aterstalls arbetskatalogen.

    Versionerna barar vid tio i stallet for att rakna vidare: 1.9.0 blir 1.10.0, men 1.10.0
    blir 2.0.0, eftersom tio ar det sista varde ett segment antar och inte det forsta det
    vagrar. Patch barar likadant in i minor, som kan bara vidare in i major: 1.10.10 blir
    2.0.0.

.PARAMETER Patch
    Hoj tredje numret. Standard nar ingen bump namns - en release av det har paketet ar
    oftast en rattning, inte ny yta. Typen sitter i botten av andras anropskedjor, och ny
    publik yta dar ar sallan det man vill.

.PARAMETER Minor
    Hoj andra numret och nollstall det tredje. For nya medlemmar pa Result.

.PARAMETER Major
    Hoj forsta numret och nollstall resten. For en andring som bryter befintlig kod.

.PARAMETER Version
    Satt en uttrycklig version i stallet for att hoja, t.ex. 2.0.0 eller 2.0.0-rc.1.

.PARAMETER Push
    Pusha grenen och taggen nar allt gatt igenom, vilket startar publiceringen.

.PARAMETER DryRun
    Rapportera vad som skulle handa och andra ingenting.

.PARAMETER Branch
    Grenen en release far goras fran. Standard main.

.EXAMPLE
    ./release.ps1
    1.0.0 blir 1.0.1, verifierad, commitad och taggad lokalt.

.EXAMPLE
    ./release.ps1 -Minor -Push
    Nya medlemmar gick in: 1.0.1 blir 1.1.0, och publiceringen startar.

.EXAMPLE
    ./release.ps1 -DryRun
    Sager vilken version en vanlig korning skulle ge, och andrar ingenting.
#>
param(
    [switch]$Patch,
    [switch]$Minor,
    [switch]$Major,
    [string]$Version,
    [switch]$Push,
    [switch]$DryRun,
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"

# Tio ar det sista varde ett versionssegment antar, inte det forsta det vagrar. Andra den
# har enda siffran for att andra hela systemet.
$SegmentCeiling = 10

$root    = $PSScriptRoot
$project = Join-Path $root "src\Avig.Kalkyl\Avig.Kalkyl.csproj"
$publish = Join-Path $root "publish.ps1"

function Fail([string]$Message) {
    Write-Host "      $Message" -ForegroundColor Red
    exit 1
}

function Step-Version([int]$MajorPart, [int]$MinorPart, [int]$PatchPart, [string]$Part) {
    switch ($Part) {
        "major" { $MajorPart++; $MinorPart = 0; $PatchPart = 0 }
        "minor" { $MinorPart++; $PatchPart = 0 }
        default { $PatchPart++ }
    }

    # Bars i ordning, sa en patch-bump kan resa hela vagen: 1.10.10 blir 2.0.0.
    if ($PatchPart -gt $SegmentCeiling) {
        $PatchPart = 0
        $MinorPart++
    }

    if ($MinorPart -gt $SegmentCeiling) {
        $MinorPart = 0
        $MajorPart++
    }

    return "$MajorPart.$MinorPart.$PatchPart"
}

function Save-Text([string]$Path, [string]$Text) {
    # UTF-8 utan BOM, och radsluten som redan star i strangen. .gitattributes lar det har
    # repot vid CRLF, sa ingenting har far normalisera till LF.
    [System.IO.File]::WriteAllText($Path, $Text, [System.Text.UTF8Encoding]::new($false))
}

Write-Host ""
Write-Host "=== Avig.Kalkyl release ===" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "    DRY RUN - ingenting skrivs, commitas eller pushas" -ForegroundColor Yellow
}

Write-Host ""

# --- [1/5] Repot maste vara i ett lage som en tagg kan beskriva ---
Write-Host "[1/5] Kontrollerar arbetskatalogen..." -ForegroundColor Yellow

$currentBranch = (git -C $root rev-parse --abbrev-ref HEAD).Trim()
if ($currentBranch -ne $Branch) {
    Fail "Pa grenen '$currentBranch', inte '$Branch'. Slapp fran $Branch, eller ange -Branch."
}

if ((git -C $root status --porcelain)) {
    Fail "Arbetskatalogen har ocommitade andringar. En tagg maste beskriva ett commitat lage."
}

git -C $root fetch origin --tags --quiet
if ($LASTEXITCODE -ne 0) {
    Fail "Nadde inte origin."
}

$behind = [int](git -C $root rev-list --count "$Branch..origin/$Branch")
if ($behind -gt 0) {
    Fail "$Branch ligger $behind commit(s) efter origin/$Branch. Pulla forst."
}

Write-Host "      ren, pa $Branch, i fas med origin" -ForegroundColor Green

# --- [2/5] Rakna ut den nya versionen ---
Write-Host "[2/5] Version" -ForegroundColor Yellow

$csproj = Get-Content $project -Raw
$currentMatch = ([regex]'<Version>([^<]+)</Version>').Match($csproj)
if (-not $currentMatch.Success) {
    Fail "Hittade ingen <Version> i csproj:en."
}

$current = $currentMatch.Groups[1].Value

$named = @($Patch, $Minor, $Major, [bool]$Version) | Where-Object { $_ }
if ($named.Count -gt 1) {
    Fail "Valj en av -Patch, -Minor, -Major eller -Version."
}

if ($Version) {
    if ($Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.\-]+)?$') {
        Fail "'$Version' ar ingen version. Forvantade 1.2.3 eller 1.2.3-rc.1."
    }

    $next = $Version
    $part = "explicit"
} else {
    $core = ($current -split '-')[0]
    $parts = $core.Split('.')
    if ($parts.Count -ne 3) {
        Fail "Kan inte hoja '$current' - forvantade tre tal."
    }

    # Patch som standard: en release av det har paketet ar oftast en rattning.
    $part = if ($Major) { "major" } elseif ($Minor) { "minor" } else { "patch" }
    $next = Step-Version ([int]$parts[0]) ([int]$parts[1]) ([int]$parts[2]) $part
}

if ($next -eq $current) {
    Fail "$next ar redan den nuvarande versionen."
}

$currentCore = [version]($current -split '-')[0]
$nextCore = [version]($next -split '-')[0]
if ($nextCore -lt $currentCore) {
    Fail "$next ar lagre an nuvarande $current. Versioner pa nuget.org gar bara uppat."
}

$tag = "v$next"

if ((git -C $root tag -l $tag)) {
    Fail "Taggen $tag finns redan lokalt. Ta bort den eller valj en annan version."
}

if ((git -C $root ls-remote --tags origin "refs/tags/$tag")) {
    Fail "Taggen $tag finns redan pa origin. En publicerad version gar inte att byta ut."
}

$carried = ""
if ($part -ne "explicit" -and $part -ne "major" -and $nextCore.Major -gt $currentCore.Major) {
    $carried = "   bar over i en ny major"
} elseif ($part -eq "patch" -and $nextCore.Minor -gt $currentCore.Minor) {
    $carried = "   bar over i en ny minor"
}

Write-Host "      $current -> $next   (tagg $tag)$carried" -ForegroundColor Green

# --- [3/5] Bevisa att releasen gar att bygga innan den blir en tagg ---
Write-Host "[3/5] Verifierar paketet..." -ForegroundColor Yellow

if ($DryRun) {
    Write-Host "      hoppas over i en dry run" -ForegroundColor Yellow
} else {
    $updatedCsproj = ([regex]'<Version>[^<]*</Version>').Replace($csproj, "<Version>$next</Version>", 1)
    Save-Text $project $updatedCsproj

    # Ingen pipeline-omslutning: publish.ps1 rapporterar via Write-Host, sa utdatan gar rakt
    # till varden och ingenting skulle komma hit att indentera.
    & $publish -NoBump
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      publish.ps1 misslyckades - aterstaller arbetskatalogen, ingen tagg skapad." -ForegroundColor Red
        git -C $root checkout -- $project
        exit 1
    }

    Write-Host "      tester, pack och formkontroll gick igenom" -ForegroundColor Green
}

# --- [4/5] Commit och tagg ---
Write-Host "[4/5] Commit och tagg" -ForegroundColor Yellow

$message = "Release $tag"

if ($DryRun) {
    Write-Host "      skulle commita '$message' och tagga $tag" -ForegroundColor Yellow
} else {
    git -C $root add -- $project
    git -C $root commit --quiet -m $message
    if ($LASTEXITCODE -ne 0) {
        Fail "Commit misslyckades."
    }

    git -C $root tag -a $tag -m "Avig.Kalkyl $next"
    if ($LASTEXITCODE -ne 0) {
        Fail "Taggningen misslyckades. Release-commiten ligger kvar; ta bort den eller tagga om for hand."
    }

    Write-Host "      commitad och taggad $tag" -ForegroundColor Green
}

# --- [5/5] Pushen ar det som publicerar ---
Write-Host "[5/5] Push" -ForegroundColor Yellow

if ($DryRun) {
    Write-Host "      skulle pusha $Branch och $tag" -ForegroundColor Yellow
} elseif ($Push) {
    git -C $root push origin $Branch
    if ($LASTEXITCODE -ne 0) {
        Fail "Push av $Branch misslyckades. Taggen ar fortfarande lokal; ingenting publicerat."
    }

    git -C $root push origin $tag
    if ($LASTEXITCODE -ne 0) {
        Fail "Push av $tag misslyckades. Ingenting publicerat."
    }

    Write-Host "      pushad - publiceringen startar" -ForegroundColor Green
} else {
    Write-Host "      inte pushad (ange -Push for att gora det harifran)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== $tag forberedd ===" -ForegroundColor Cyan
Write-Host ""

if ($DryRun) {
    Write-Host "Ingenting andrades. Kor om utan -DryRun." -ForegroundColor Cyan
} elseif ($Push) {
    Write-Host "Folj publiceringen:" -ForegroundColor Cyan
    Write-Host "    https://github.com/marcuslindh/Avig.Kalkyl/actions" -ForegroundColor Gray
} else {
    Write-Host "Ingenting har lamnat maskinen an. For att publicera:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    git push origin $Branch && git push origin $tag" -ForegroundColor Gray
    Write-Host ""
    Write-Host "For att angra i stallet:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    git tag -d $tag; git reset --hard HEAD~1" -ForegroundColor Gray
}

Write-Host ""
