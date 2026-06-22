---
name: leeno-icone-toolbar
description: >
  Genera automaticamente i file di icone in formato .bmp per le toolbar di LeenO 
  a partire dai file vettoriali .svg originali. Da usare quando si aggiungono o 
  modificano le icone in formato SVG.
---

# LeenO – Generazione Icone Toolbar

Questa skill permette di mantenere aggiornati i file `.bmp` (che in realtà mantengono il contenuto vettoriale originale) richiesti dall'interfaccia di LibreOffice per le icone delle toolbar, generandoli in automatico a partire dai file vettoriali `.svg` di origine.

## Quando usarla
- Dopo aver aggiunto una nuova icona in formato SVG alla cartella `src/Ultimus.oxt/icons/svg/`.
- Dopo aver modificato o sostituito un'icona SVG esistente.
- Per rigenerare i file raster prima di pacchettizzare l'estensione se si sospetta la mancanza di qualche file per i bottoni in toolbar.

## Procedura (AI)
Per eseguire la generazione delle icone, utilizza lo strumento `run_command` lanciando il seguente comando in PowerShell dalla radice del repository:

```powershell
Get-ChildItem -Path "src\Ultimus.oxt\icons\svg" -Filter *.svg | ForEach-Object {
    $baseName = $_.BaseName
    $fullPath = $_.FullName
    foreach ($suffix in @('_16.bmp', '_16h.bmp', '_26.bmp', '_26h.bmp')) {
        Copy-Item -Path $fullPath -Destination "src\Ultimus.oxt\icons\$baseName$suffix" -Force
    }
}
```

## Cosa fa la procedura
1. **Analisi**: Legge in automatico tutti i file `.svg` contenuti in `src/Ultimus.oxt/icons/svg/`.
2. **Generazione**: Per ogni file (es. `miaicona.svg`), crea quattro copie nella cartella padre `src/Ultimus.oxt/icons/`:
   - `miaicona_16.bmp`
   - `miaicona_16h.bmp`
   - `miaicona_26.bmp`
   - `miaicona_26h.bmp`
3. **Sovrascrittura**: Forza l'aggiornamento dei file `.bmp` qualora esistessero già, assicurando che contengano le ultime modifiche del vettore.

---
> [!TIP]
> LibreOffice è in grado di leggere il contenuto SVG pur richiedendo storicamente in configurazione estensioni come `.bmp`. Questa skill agisce come un rimpiazzo cross-platform dello script storico `leeno_icons.sh`.
