param(
    [string]$At = "23:00",
    [double]$TotalHours = 7,
    [int]$Workers = 4,
    [switch]$Once,
    [switch]$RunNow,
    [switch]$Remove,
    [switch]$Show
)

# Puts fuzz-night.ps1 in Windows Task Scheduler, so that libFuzzer runs while nobody uses the
# machine. Every night at $At by default; -Once for the coming night only.
#
#   ./schedule-night.ps1                          # every night 23:00, 7 hours, 4 workers
#   ./schedule-night.ps1 -At 22:30 -TotalHours 8 -Workers 8
#   ./schedule-night.ps1 -Once                    # tonight only
#   ./schedule-night.ps1 -Show                    # next run, last run and its result
#   ./schedule-night.ps1 -Remove
#
# The task runs as the current user and only while that user is logged on (a locked screen is
# fine): that needs no stored password and no administrator. It wakes the computer from sleep,
# and fuzz-night.ps1 keeps it awake until the night is done. A run that is missed, because the
# computer was off, is not made up in the morning. The log is findings/night-<date>.log.

$ErrorActionPreference = "Stop"
$taskName = "Avig.Kalkyl fuzz night"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$night = Join-Path $scriptDir "fuzz-night.ps1"

if ($Remove) {
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Removed the task '$taskName'." -ForegroundColor Green
    }
    else {
        Write-Host "There is no task '$taskName'." -ForegroundColor Yellow
    }
    exit 0
}

if ($Show) {
    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    if (-not $task) {
        Write-Host "There is no task '$taskName'." -ForegroundColor Yellow
        exit 1
    }
    $info = Get-ScheduledTaskInfo -TaskName $taskName
    Write-Host "State:       $($task.State)"
    Write-Host "Runs:        $($task.Actions[0].Arguments)"
    Write-Host "Next run:    $($info.NextRunTime)"
    Write-Host "Last run:    $($info.LastRunTime)"
    # 0 = nothing found, 1 = a finding or corpus-check failed, 267009 = running, 267011 = never run.
    Write-Host "Last result: $($info.LastTaskResult)"
    $latest = Get-ChildItem (Join-Path $scriptDir "findings") -Filter "night-*.log" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime | Select-Object -Last 1
    if ($latest) { Write-Host "Latest log:  $($latest.FullName)" }
    exit 0
}

# PowerShell 7 when there is one: Windows PowerShell 5.1 writes the log in UTF-16.
$shell = (Get-Command "pwsh.exe" -ErrorAction SilentlyContinue).Source
if (-not $shell) { $shell = (Get-Command "powershell.exe").Source }

$time = [datetime]::ParseExact($At, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
$arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$night`" " +
    ("-TotalHours {0} -Workers {1}" -f $TotalHours.ToString([System.Globalization.CultureInfo]::InvariantCulture), $Workers)

$action = New-ScheduledTaskAction -Execute $shell -Argument $arguments -WorkingDirectory $scriptDir

if ($Once) {
    $start = (Get-Date).Date.Add($time.TimeOfDay)
    if ($start -lt (Get-Date)) { $start = $start.AddDays(1) }
    $trigger = New-ScheduledTaskTrigger -Once -At $start
}
else {
    $trigger = New-ScheduledTaskTrigger -Daily -At $time
}

# Room for publishing nine times and for corpus-check after the fuzzing itself.
$limit = New-TimeSpan -Minutes ([int]($TotalHours * 60) + 120)
$settings = New-ScheduledTaskSettingsSet -WakeToRun -ExecutionTimeLimit $limit -MultipleInstances IgnoreNew `
    -DontStopOnIdleEnd -Priority 7
$principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description "libFuzzer over Avig.Kalkyl: tools/Avig.Kalkyl.Fuzz/fuzz-night.ps1. Log in findings/night-<date>.log." `
    -Force | Out-Null

$info = Get-ScheduledTaskInfo -TaskName $taskName
Write-Host "Scheduled '$taskName': $TotalHours h with $Workers worker(s), next run $($info.NextRunTime)." -ForegroundColor Green
Write-Host "Check it with ./schedule-night.ps1 -Show, remove it with ./schedule-night.ps1 -Remove." -ForegroundColor Gray

if ($RunNow) {
    Start-ScheduledTask -TaskName $taskName
    Write-Host "Started now." -ForegroundColor Yellow
}
