# Percorso della cartella contenente i file SVG
$cartella = Get-Location

# Itera attraverso tutti i file SVG nella cartella corrente
Get-ChildItem -Path $cartella -Filter *.svg | ForEach-Object {
    $file = $_.FullName

    # Leggi il contenuto del file SVG
    $contenuto = Get-Content -Path $file -Raw

    # Inverti i colori sostituendo le stringhe
    $contenuto = $contenuto -replace "#000000", "#temp"  # Inverti nero in bianco
    $contenuto = $contenuto -replace "#ffffff", "#000000"  # Inverti nero in bianco
    $contenuto = $contenuto -replace "#temp", "#ffffff"  # Inverti bianco in nero

    # Sovrascrivi il file con i colori invertiti
    $contenuto | Set-Content -Path $file -Force
}

Write-Host "Operazione completata."
