# Lezioni apprese – Quirk minori UNO/ODF

Gotcha isolati, non legati a un modulo specifico.

## Nomi di stile interni LibreOffice

LibreOffice può assegnare a uno stile un nome interno anonimo (es. `uuuuu`) invece del nome leggibile mostrato nell'interfaccia (es. `'Comp TOTALI'`), tipicamente per stili generati o duplicati programmaticamente. Qualunque confronto nel codice del tipo `oCell.CellStyle == "Comp TOTALI"` fallisce silenziosamente in questi casi — nessuna eccezione, solo logica che non scatta mai. Prima di scrivere un confronto su `CellStyle`, verificare il nome interno effettivo dello stile sul documento reale, non assumerlo dal nome visualizzato in LibreOffice.

## Igiene dei template ODS

Se un file `.ods` reale di un progetto viene salvato sopra un template o usato come base per generarne uno nuovo, il percorso del file reale può restare hardcoded in una cella di `content.xml` (tipicamente le celle `F1` dei fogli `COMPUTO` e `CONTABILITA`, usate per riferimenti a percorso). Prima di distribuire o committare un template, verificare sempre queste celle: un template che punta silenziosamente al file di un progetto specifico produce comportamenti anomali difficili da diagnosticare per chi lo usa in un contesto diverso.

## `getDocument()` subito dopo un dialogo modale

`getDocument()` si basa su `desktop.getCurrentComponent()` con fallback alla scansione di `desktop.getComponents()`; nell'istante immediatamente successivo alla chiusura di un dialogo (`oDlg.execute()`), il componente corrente può non essere ancora il foglio Calc atteso, e se nessun componente aperto supera il controllo `is_valid_calc()` la funzione torna `None`. Se una funzione ha già ottenuto un riferimento valido a `oDoc` prima di aprire il dialogo, quel riferimento va passato esplicitamente alle funzioni chiamate dopo la chiusura (es. come parametro opzionale `oDoc=`), invece di richiamare di nuovo `LeenoUtils.getDocument()` — che a quel punto rischia di risolvere `None` e generare `'NoneType' object has no attribute 'CurrentController'` a valle. Vedi il fix in `LeenoToolbars.py` (`Switch`, `On`, `Ordina`, `AllOn`, `AllOff`) e la relativa chiamata da `LeenoConfig.MENU_leeno_conf()`.
