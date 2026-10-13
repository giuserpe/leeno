# Pacchettizza il tema in uno ZIP installabile da Aspetto > Temi > Aggiungi nuovo > Carica tema.
# Il sorgente e' la cartella che contiene questo script: il percorso non dipende dalla macchina.
$source = $PSScriptRoot
$dest = Join-Path (Split-Path $source -Parent) 'leeno-theme-v1.2-hybrid.zip'

# File senza i quali il tema pubblicato perderebbe funzioni, spesso in silenzio
# (get_template_part non segnala errori se il file manca).
$richiesti = @('style.css', 'functions.php', 'template-parts\ai-claim.php')
foreach ($r in $richiesti) {
    if (-not (Test-Path (Join-Path $source $r))) {
        Write-Error ("File obbligatorio mancante nel sorgente: " + $r + " (sorgente: " + $source + ")")
        exit 1
    }
}

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

# Verifica finale sul pacchetto appena creato.
$lettura = [System.IO.Compression.ZipFile]::OpenRead($dest)
$presenti = @($lettura.Entries | ForEach-Object { $_.FullName })
$lettura.Dispose()
foreach ($r in $richiesti) {
    $voce = 'leeno-theme/' + $r.Replace('\', '/')
    if ($presenti -notcontains $voce) {
        Write-Error ("Nel pacchetto manca: " + $voce)
        exit 1
    }
}

Write-Host "ZIP creato con successo: $dest"
