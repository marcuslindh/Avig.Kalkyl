param(
    [string]$Target = "",
    [int]$TimeoutSeconds = 30,
    [string]$HeapLimit = "0x40000000"
)

# Runs every saved input (corpus/ and findings/) through --repro in a process of its own, with a
# time limit and a managed-heap limit (1 GiB by default). That is the only way to see a stack
# overflow, an out-of-memory or a hang, which end or stall the process instead of being caught.
# Accepted: no exception, or InvalidDataException (a damaged file, as the package promises).
# See docs/testplan.md, 6.3.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
dotnet build (Join-Path $scriptDir "Avig.Kalkyl.Fuzz.csproj") -c Release --nologo -v quiet | Out-Null
if ($LASTEXITCODE -ne 0) { exit 1 }
$harness = Join-Path $scriptDir "bin\Release\net10.0\Avig.Kalkyl.Fuzz.dll"
$failures = 0
$checked = 0

foreach ($folder in @("corpus", "findings")) {
    $root = Join-Path $scriptDir $folder
    if (-not (Test-Path $root)) { continue }

    foreach ($targetDir in Get-ChildItem $root -Directory) {
        if ($Target -and $targetDir.Name -ne $Target) { continue }

        foreach ($file in Get-ChildItem $targetDir.FullName -File | Where-Object { $_.Name -ne "last-input.bin" }) {
            $checked++
            $start = New-Object System.Diagnostics.ProcessStartInfo "dotnet"
            $start.ArgumentList.Add($harness)
            $start.ArgumentList.Add($targetDir.Name)
            $start.ArgumentList.Add("--repro")
            $start.ArgumentList.Add($file.FullName)
            $start.RedirectStandardOutput = $true
            $start.RedirectStandardError = $true
            $start.UseShellExecute = $false
            $start.Environment["DOTNET_GCHeapHardLimit"] = $HeapLimit
            $process = [System.Diagnostics.Process]::Start($start)
            $output = $process.StandardOutput.ReadToEndAsync()
            $errors = $process.StandardError.ReadToEndAsync()

            if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
                $process.Kill($true)
                $failures++
                Write-Host "  HANG      $($targetDir.Name)/$($file.Name) (over $TimeoutSeconds s)" -ForegroundColor Red
                continue
            }

            $text = $output.Result + $errors.Result
            $accepted = $text.Contains("(no exception)") -or ($text -match "^System\.IO\.InvalidDataException" -or $text -match "`nSystem\.IO\.InvalidDataException")

            if ($process.ExitCode -eq -1073741571 -or $process.ExitCode -eq 0xC00000FD) {
                $failures++
                Write-Host "  STACK     $($targetDir.Name)/$($file.Name) (stack overflow)" -ForegroundColor Red
            }
            elseif ($text -match "OutOfMemoryException") {
                $failures++
                Write-Host "  MEMORY    $($targetDir.Name)/$($file.Name) (out of memory under $HeapLimit)" -ForegroundColor Red
            }
            elseif (-not $accepted) {
                $failures++
                $first = ($text -split "`r?`n" | Where-Object { $_ -match "Exception" } | Select-Object -First 1)
                Write-Host "  FAIL      $($targetDir.Name)/$($file.Name): $first" -ForegroundColor Red
            }
        }
    }
}

Write-Host ""
if ($failures -eq 0) {
    Write-Host "OK: $checked inputs, each in its own process, none crashed, ran out of memory or hung." -ForegroundColor Green
    exit 0
}

Write-Host "FAIL: $failures of $checked inputs." -ForegroundColor Red
exit 1
