param(
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)][string[]]$Files
)

# Opens each file in Excel through COM, read-only, with alerts off and macros disabled, and says
# whether Excel accepted it. With alerts off, a file Excel would have to repair makes Open throw
# instead of showing the "We found a problem with some content" prompt (tried 2026-11-25 on
# Excel 365 with a bad style index, a chart without axis ids and broken XML: all FAIL). A
# "[Repaired]" caption or a new repair log (error*.xml) in %TEMP% also counts. Prints one line
# per file and exits with the number of files that failed. Needs Excel on the machine.
#
#   ./open-in-excel.ps1 C:\temp\excel-check\14-dashboard-document.xlsx C:\temp\excel-check\15-dashboard-writer.xlsx
#   ./open-in-excel.ps1 C:\temp\excel-check     (every .xlsx in the folder)

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false
$excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: never run macros
$failures = 0

$paths = foreach ($file in $Files) {
    if (Test-Path $file -PathType Container) {
        Get-ChildItem -Path $file -Filter '*.xlsx' | Sort-Object Name | Select-Object -ExpandProperty FullName
    }
    else {
        $file
    }
}

try {
    foreach ($file in $paths) {
        $path = (Resolve-Path $file).Path
        $before = @(Get-ChildItem -Path $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)

        try {
            $workbook = $excel.Workbooks.Open($path, 0, $true)
            $caption = $excel.ActiveWindow.Caption
            $after = @(Get-ChildItem -Path $env:TEMP -Filter 'error*.xml' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
            $logs = @($after | Where-Object { $before -notcontains $_ })
            $shapes = 0
            $charts = 0

            foreach ($sheet in $workbook.Worksheets) {
                $shapes += $sheet.Shapes.Count
                $charts += $sheet.ChartObjects().Count
            }

            if ($caption -like '*Repaired*' -or $logs.Count -gt 0) {
                $failures++
                "REPAIRED $([IO.Path]::GetFileName($path))  window=$caption"

                foreach ($log in $logs) {
                    Get-Content $log -Raw
                }
            }
            else {
                "OK       $([IO.Path]::GetFileName($path))  sheets=$($workbook.Worksheets.Count) shapes=$shapes charts=$charts"
            }

            $workbook.Close($false)
        }
        catch {
            $failures++
            "FAIL     $([IO.Path]::GetFileName($path))  $($_.Exception.Message)"
        }
    }
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $failures
