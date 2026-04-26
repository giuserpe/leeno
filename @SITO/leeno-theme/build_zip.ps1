$source = 'w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme-v1.2-hybrid'
$dest = 'w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme-v1.2-hybrid.zip'

if (Test-Path $dest) { Remove-Item $dest -Force }

$items = Get-ChildItem -Path $source -Recurse |
Where-Object {
    $_.FullName -notmatch '\\\{assets' -and
    $_.FullName -notmatch '\\\.git' -and
    $_.Name -notmatch '\.zip$' -and
    $_.Name -notmatch 'build_zip\.ps1$'
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($dest, 'Create')

foreach ($item in $items) {
    if (-not $item.PSIsContainer) {
        $relative = $item.FullName.Substring($source.Length + 1)
        $entryName = ('leeno-theme/' + $relative).Replace('\', '/')
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $item.FullName, $entryName, 'Optimal') | Out-Null
        Write-Host ("Aggiunto: " + $entryName)
    }
}

$zip.Dispose()
Write-Host "ZIP creato con successo: $dest"
