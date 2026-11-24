param(
    [double]$TotalHours = 8,
    [int]$Workers = 1
)

# A night of coverage-guided fuzzing: every target in turn through fuzz.ps1, with the time split
# by weight (the targets that read whole packages get the most), then corpus-check.ps1 over what
# was found, each input in its own process. A log goes to findings/night-<date>.log.
# Needs libfuzzer-dotnet.exe in ~\bin or on PATH, and the sharpfuzz tool (see README.md).
# attribute-sweep is left out: under libFuzzer it runs the same harness as xlsx-part.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$env:PATH = "$HOME\bin;$HOME\.dotnet\tools;$env:PATH"

# Keep the computer from going to sleep until the night is done: a scheduled task wakes it, but
# Windows would otherwise let it sleep again after the idle time. ES_CONTINUOUS together with
# ES_SYSTEM_REQUIRED holds for as long as this process lives.
Add-Type -Namespace Avig -Name Power -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern uint SetThreadExecutionState(uint flags);
"@
[Avig.Power]::SetThreadExecutionState([uint32]2147483649) | Out-Null

$weights = [ordered]@{
    "xlsx-part"     = 3
    "zip-raw"       = 2
    "formula"       = 2
    "formula-shift" = 1
    "formula-sweep" = 1
    "roundtrip"     = 1
    "cell-text"     = 0.5
    "excel-date"    = 0.5
    "cell-address"  = 0.5
}

$total = ($weights.Values | Measure-Object -Sum).Sum
$findings = Join-Path $scriptDir "findings"
if (-not (Test-Path $findings)) { New-Item -ItemType Directory $findings | Out-Null }
$log = Join-Path $findings ("night-{0:yyyy-MM-dd-HHmm}.log" -f (Get-Date))
$failed = @()

foreach ($target in $weights.Keys) {
    $seconds = [int]($TotalHours * 3600 * $weights[$target] / $total)
    "{0:HH:mm} {1}: {2} s" -f (Get-Date), $target, $seconds | Tee-Object -FilePath $log -Append

    & (Join-Path $scriptDir "fuzz.ps1") -Target $target -RuntimeSeconds $seconds -Workers $Workers *>&1 |
        Select-String -NotMatch "^#\d+\s+(NEW|REDUCE|pulse)" | Out-File -FilePath $log -Append

    if ($LASTEXITCODE -ne 0) {
        $failed += $target
        "{0:HH:mm} {1}: FINDING" -f (Get-Date), $target | Tee-Object -FilePath $log -Append
    }
}

"{0:HH:mm} corpus-check over findings/" -f (Get-Date) | Tee-Object -FilePath $log -Append
& (Join-Path $scriptDir "corpus-check.ps1") *>&1 | Out-File -FilePath $log -Append
$checkFailed = $LASTEXITCODE -ne 0

Write-Host ""
if ($failed.Count -eq 0 -and -not $checkFailed) {
    Write-Host "Night done, nothing found. Log: $log" -ForegroundColor Green
    exit 0
}

Write-Host "Findings in: $($failed -join ', '). corpus-check failed: $checkFailed. Log: $log" -ForegroundColor Red
Write-Host "Reproduce one with: dotnet run -c Release -- <target> --repro findings/<target>/<file>" -ForegroundColor Yellow
exit 1
