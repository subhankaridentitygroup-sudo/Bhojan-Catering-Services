$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$workbook = Join-Path $PSScriptRoot 'assets\BHOJAN POSITIVE.xlsx'
$outputDir = Join-Path $PSScriptRoot 'assets\partners'

if (-not (Test-Path -LiteralPath $workbook)) {
    throw "Workbook not found: $workbook"
}

New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
Get-ChildItem -LiteralPath $outputDir -Filter 'partner-*.jpg' -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $outputDir -Filter 'partner-*.jpeg' -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $outputDir -Filter 'partner-*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force

$zip = [System.IO.Compression.ZipFile]::OpenRead($workbook)
try {
    $media = @($zip.Entries | Where-Object { $_.FullName -like 'xl/media/*' } | Sort-Object { [int]([IO.Path]::GetFileNameWithoutExtension($_.FullName) -replace '\D', '') })
    if ($media.Count -ne 30) {
        throw "Expected 30 embedded media entries, found $($media.Count)."
    }

    $created = @()
    for ($index = 0; $index -lt $media.Count; $index++) {
        $entry = $media[$index]
        $sourceExtension = [IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()
        $extension = if ($sourceExtension -eq '.png') { '.png' } elseif ($sourceExtension -in @('.jpg', '.jpeg')) { '.jpg' } else { $sourceExtension }
        if ([string]::IsNullOrWhiteSpace($extension)) {
            throw "Cannot determine image format for $($entry.FullName)."
        }

        $outputIndex = $index + 1
        if ($sourceExtension -eq '.png' -and $outputIndex -eq 27) {
            $outputIndex = 28
        } elseif ($sourceExtension -ne '.png' -and $outputIndex -eq 28) {
            $outputIndex = 27
        }
        $name = 'partner-{0:D2}{1}' -f $outputIndex, $extension
        $destination = Join-Path $outputDir $name
        $input = $entry.Open()
        try {
            $output = [IO.File]::Create($destination)
            try { $input.CopyTo($output) } finally { $output.Dispose() }
        } finally { $input.Dispose() }
        $created += [PSCustomObject]@{ File = $name; Source = $entry.FullName; Bytes = (Get-Item -LiteralPath $destination).Length }
    }

    Write-Output ('CREATED_COUNT=' + $created.Count)
    $created | ForEach-Object { Write-Output (('{0}`t{1}`t{2}' -f $_.File, $_.Bytes, $_.Source)) }
} finally {
    $zip.Dispose()
}
