param([switch]$Force)

# Downloads the workbooks listed in files.json into files/. They are not committed (see the
# comment in files.json), so run this once before the compatibility tests.

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$target = Join-Path $scriptDir "files"
$manifest = Get-Content (Join-Path $scriptDir "files.json") -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Force $target | Out-Null
$failed = 0

foreach ($file in $manifest.files) {
    $path = Join-Path $target $file.name

    if ((Test-Path $path) -and -not $Force) {
        Write-Host "  have      $($file.name)" -ForegroundColor Gray
        continue
    }

    try {
        Invoke-WebRequest -Uri $file.url -OutFile $path -UseBasicParsing -MaximumRedirection 5

        # Some pages hand out a zip with the workbook inside: take the first entry with the given extension.
        if ($file.extract) {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $zip = [IO.Compression.ZipFile]::OpenRead($path)

            try {
                $entry = $zip.Entries | Where-Object { $_.FullName.EndsWith($file.extract, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
                if (-not $entry) { throw "no $($file.extract) in the zip" }
                $inner = Join-Path $target ($file.name + ".part")
                [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $inner, $true)
            }
            finally {
                $zip.Dispose()
            }

            Move-Item $inner $path -Force
        }

        $bytes = [IO.File]::ReadAllBytes($path)

        # Every .xlsx is a zip file; anything else is an error page.
        if ($bytes.Length -lt 4 -or $bytes[0] -ne 0x50 -or $bytes[1] -ne 0x4B) {
            Remove-Item $path
            throw "not a zip file"
        }

        Write-Host "  fetched   $($file.name) ($([Math]::Round($bytes.Length / 1KB)) KB)" -ForegroundColor Green
    }
    catch {
        $failed++
        Write-Host "  FAILED    $($file.name): $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($failed -gt 0) {
    Write-Host "$failed file(s) could not be downloaded." -ForegroundColor Yellow
}
