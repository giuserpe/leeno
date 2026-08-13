# Lezioni apprese – Quirk minori UNO/ODF

Due gotcha isolati, non legati a un modulo specifico.

## Nomi di stile interni LibreOffice

LibreOffice può assegnare a uno stile un nome interno anonimo (es. `uuuuu`) invece del nome leggibile mostrato nell'interfaccia (es. `'Comp TOTALI'`), tipicamente per stili generati o duplicati programmaticamente. Qualunque confronto nel codice del tipo `oCell.CellStyle == "Comp TOTALI"` fallisce silenziosamente in questi casi — nessuna eccezione, solo logica che non scatta mai. Prima di scrivere un confronto su `CellStyle`, verificare il nome interno effettivo dello stile sul documento reale, non assumerlo dal nome visualizzato in LibreOffice.

## Igiene dei template ODS

Se un file `.ods` reale di un progetto viene salvato sopra un template o usato come base per generarne uno nuovo, il percorso del file reale può restare hardcoded in una cella di `content.xml` (tipicamente le celle `F1` dei fogli `COMPUTO` e `CONTABILITA`, usate per riferimenti a percorso). Prima di distribuire o committare un template, verificare sempre queste celle: un template che punta silenziosamente al file di un progetto specifico produce comportamenti anomali difficili da diagnosticare per chi lo usa in un contesto diverso.
