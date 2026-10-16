# Lezioni apprese – Quirk minori UNO/ODF

Gotcha isolati, non legati a un modulo specifico.

## Nomi di stile interni LibreOffice

LibreOffice può assegnare a uno stile un nome interno anonimo (es. `uuuuu`) invece del nome leggibile mostrato nell'interfaccia (es. `'Comp TOTALI'`), tipicamente per stili generati o duplicati programmaticamente. Qualunque confronto nel codice del tipo `oCell.CellStyle == "Comp TOTALI"` fallisce silenziosamente in questi casi — nessuna eccezione, solo logica che non scatta mai. Prima di scrivere un confronto su `CellStyle`, verificare il nome interno effettivo dello stile sul documento reale, non assumerlo dal nome visualizzato in LibreOffice.

## Igiene dei template ODS

Se un file `.ods` reale di un progetto viene salvato sopra un template o usato come base per generarne uno nuovo, il percorso del file reale può restare hardcoded in una cella di `content.xml` (tipicamente le celle `F1` dei fogli `COMPUTO` e `CONTABILITA`, usate per riferimenti a percorso). Prima di distribuire o committare un template, verificare sempre queste celle: un template che punta silenziosamente al file di un progetto specifico produce comportamenti anomali difficili da diagnosticare per chi lo usa in un contesto diverso.

## `getDocument()` subito dopo un dialogo modale

`getDocument()` si basa su `desktop.getCurrentComponent()` con fallback alla scansione di `desktop.getComponents()`; nell'istante immediatamente successivo alla chiusura di un dialogo (`oDlg.execute()`), il componente corrente può non essere ancora il foglio Calc atteso, e se nessun componente aperto supera il controllo `is_valid_calc()` la funzione torna `None`. Se una funzione ha già ottenuto un riferimento valido a `oDoc` prima di aprire il dialogo, quel riferimento va passato esplicitamente alle funzioni chiamate dopo la chiusura (es. come parametro opzionale `oDoc=`), invece di richiamare di nuovo `LeenoUtils.getDocument()` — che a quel punto rischia di risolvere `None` e generare `'NoneType' object has no attribute 'CurrentController'` a valle. Vedi il fix in `LeenoToolbars.py` (`Switch`, `On`, `Ordina`, `AllOn`, `AllOff`) e la relativa chiamata da `LeenoConfig.MENU_leeno_conf()`.

## Freeze e regressioni lontane

Instrumentare la funzione sospetta con `DLG.chi()` a ogni passaggio chiave (l'ultimo checkpoint visto localizza il blocco); per regressioni lontane usare `git bisect` (good = ultimo tag funzionante, bad = `HEAD` di `dev`); diffidare di commit che sembrano toccare solo codice "non collegato": il problema può essere in un file adiacente dello stesso commit, come un file di test.

## Altezza ottimale delle righe

Nel doppio click sul bordo tra due righe, con l'intero foglio selezionato, Calc adatta l'altezza al contenuto in modo quasi istantaneo anche con moltissimi dati. L'azione non viene catturata dal registratore di macro, quindi non esiste un comando `.uno:` registrabile che la riproduca. `Rows.OptimalHeight` impostato via API, anche in un'unica operazione su aree estese, resta lento: non è equivalente al gesto della GUI e non va presentato come tale.

Percorsi proposti ma mai provati, scartati senza test: `.uno:SetOptimalRowHeight` dopo `.uno:SelectAll`; blocco di Undo, controller e ricalcolo automatico durante l'operazione. Ipotesi non verificata: il vantaggio della GUI deriva dal minor lavoro collaterale (Undo, notifiche, ridisegni) del percorso della vista. Se serve gestire l'altezza in LeenO, l'unica direzione ragionevole è ridurre il perimetro (ricalcolo solo sulle righe inserite o modificate, o altezze impostate in modo mirato), da misurare sul caso reale prima di adottarla.
