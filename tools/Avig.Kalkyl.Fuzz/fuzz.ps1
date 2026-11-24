param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("xlsx-part", "zip-raw", "attribute-sweep", "formula", "formula-sweep", "formula-shift", "roundtrip", "cell-text", "excel-date", "cell-address")]
    [string]$Target,

    [int]$RuntimeSeconds = 60,
    [int]$Workers = 1,
    [int]$TimeoutSeconds = 10
)

# Coverage-guided fuzzing with libFuzzer: publish the harness, instrument Avig.Kalkyl.dll with
# SharpFuzz, and run libfuzzer-dotnet against corpus/<target>. Findings land in findings/<target>.

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$project = Join-Path $scriptDir "Avig.Kalkyl.Fuzz.csproj"
$publishDir = Join-Path $scriptDir "bin\fuzz-publish"
$corpusDir = Join-Path $scriptDir "corpus\$Target"
$findingsDir = Join-Path $scriptDir "findings\$Target"
# libFuzzer writes the inputs it finds interesting to the first corpus folder. That folder is
# outside the repo, so a night's run neither adds thousands of files to the committed seeds in
# corpus/ nor makes a synced folder (the repo lives in OneDrive) upload them as they come.
$workDir = Join-Path $env:LOCALAPPDATA "Avig.Kalkyl.Fuzz\corpus-work\$Target"

Write-Host ""
Write-Host "=== Avig.Kalkyl.Fuzz: $Target ===" -ForegroundColor Cyan
Write-Host ""

# 1. libfuzzer-dotnet.exe: PATH first, then the usual places.
$libfuzzer = $null
$found = Get-Command "libfuzzer-dotnet.exe" -ErrorAction SilentlyContinue
if ($found) { $libfuzzer = $found.Source }
if (-not $libfuzzer) {
    foreach ($candidate in @((Join-Path $HOME "bin\libfuzzer-dotnet.exe"), (Join-Path $env:LOCALAPPDATA "Programs\libfuzzer-dotnet\libfuzzer-dotnet.exe"))) {
        if (Test-Path $candidate) { $libfuzzer = (Resolve-Path $candidate).Path; break }
    }
}
if (-not $libfuzzer) {
    Write-Host "libfuzzer-dotnet.exe not found." -ForegroundColor Red
    Write-Host "Get it from https://github.com/Metalnem/libfuzzer-dotnet/releases and put it in $HOME\bin\ or on PATH." -ForegroundColor Yellow
    Write-Host "Without it, smoke mode still works: dotnet run -c Release -- $Target --smoke" -ForegroundColor Yellow
    exit 1
}
Write-Host "libfuzzer-dotnet: $libfuzzer" -ForegroundColor Gray

# 2. The sharpfuzz command (a dotnet global tool).
$sharpfuzz = $null
$found = Get-Command "sharpfuzz" -ErrorAction SilentlyContinue
if ($found) { $sharpfuzz = $found.Source }
if (-not $sharpfuzz) {
    $toolPath = Join-Path $env:USERPROFILE ".dotnet\tools\sharpfuzz.exe"
    if (Test-Path $toolPath) { $sharpfuzz = $toolPath }
}
if (-not $sharpfuzz) {
    Write-Host "sharpfuzz is missing; installing SharpFuzz.CommandLine..." -ForegroundColor Yellow
    dotnet tool install -g SharpFuzz.CommandLine
    if ($LASTEXITCODE -ne 0) { exit 1 }
    $sharpfuzz = Join-Path $env:USERPROFILE ".dotnet\tools\sharpfuzz.exe"
}
Write-Host "sharpfuzz:        $sharpfuzz" -ForegroundColor Gray

# 3. Publish. Instrumentation rewrites the IL, so the assembly must not be ReadyToRun (see the csproj).
Write-Host "[1/3] Publishing the harness..." -ForegroundColor Yellow
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
dotnet publish $project -c Release -o $publishDir --nologo -v quiet
if ($LASTEXITCODE -ne 0) { exit 1 }

# 4. Instrument the code under test. Only Avig.Kalkyl.dll: coverage of the harness itself or of
#    Avig.Result would steer the mutations towards nothing interesting.
Write-Host "[2/3] Instrumenting Avig.Kalkyl.dll..." -ForegroundColor Yellow
& $sharpfuzz (Join-Path $publishDir "Avig.Kalkyl.dll")
if ($LASTEXITCODE -ne 0) { exit 1 }

# 5. Run.
Write-Host "[3/3] Running libFuzzer for $RuntimeSeconds s with $Workers worker(s)..." -ForegroundColor Yellow
if (-not (Test-Path $findingsDir)) { New-Item -Path $findingsDir -ItemType Directory | Out-Null }
if (-not (Test-Path $corpusDir)) { New-Item -Path $corpusDir -ItemType Directory | Out-Null }
if (-not (Test-Path $workDir)) { New-Item -Path $workDir -ItemType Directory | Out-Null }


$arguments = @(
    "--target_path=$(Join-Path $publishDir 'Avig.Kalkyl.Fuzz.exe')",
    "--target_arg=$Target",
    "-artifact_prefix=$findingsDir\",
    "-print_final_stats=1",
    "-max_total_time=$RuntimeSeconds",
    "-timeout=$TimeoutSeconds"
)
if ($Workers -gt 1) {
    $arguments += "-workers=$Workers"
    $arguments += "-jobs=$Workers"
}

& $libfuzzer @arguments $workDir $corpusDir
$exitCode = $LASTEXITCODE

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "=== Done, nothing found ===" -ForegroundColor Green
}
else {
    Write-Host "=== Finding saved in $findingsDir ===" -ForegroundColor Red
    Write-Host "Reproduce: dotnet run -c Release -- $Target --repro <file>" -ForegroundColor Yellow
}
Write-Host ""

exit $exitCode
