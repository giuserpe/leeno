# Lezioni apprese – Modulo XPWE (export/import PriMus)

Riferimento di dettaglio per chi modifica `LeenoExport.py` e `LeenoImport_XPWE.py`. Le lezioni derivano dalla correzione di tre bug reali (rilasciati come `leeno-fix-xpwe-idep-vuoto.zip`, verificati end-to-end con documenti reali via LibreOffice headless/UNO).

## Case sensitivity: `diz_ep` vs VLOOKUP

I lookup Python su dizionario sono case-sensitive per natura, ma il VLOOKUP di LibreOffice che gli utenti usano per confrontare gli stessi codici è case-insensitive di default. Normalizzare sempre le chiavi di `diz_ep` e i valori cercati con `.upper()` prima del confronto, altrimenti righe con lo stesso codice scritto in maiuscolo/minuscolo diverso vengono trattate come non trovate.

## Righe IDEP non risolvibili

Mai `except: pass` silenzioso. Un ID prezzo non risolvibile durante l'import va sempre loggato (non solo scartato) e conteggiato in un riepilogo finale mostrato all'utente a fine import: altrimenti un import "riuscito" può in realtà aver perso righe senza alcuna segnalazione.

## Segno invertito sulle righe "vedi voce" con quantità negativa

Il bug bypassava `invertiUnSegno()` e scriveva direttamente `-PRODUCT(...)` nella cella, producendo un doppio segno negativo per le righe generate da `vedi_voce_xpwe()`. Nei percorsi di import il segno va sempre gestito passando dalla funzione dedicata, mai scrivendo l'operatore `-` a mano davanti alla formula.

## `invertiUnSegno()` è pensata per uso interattivo, non per l'import

È un toggle: se chiamata su una riga che ha già lo stile `ROSSO` impostato da `vedi_voce_xpwe()`, inverte il segno una seconda volta riportandolo a quello originale. Nei percorsi di import automatico va bypassata, non riutilizzata così com'è.

## `uFindStringCol()` può restituire `None`

Se la colonna cercata non esiste nel foglio, la funzione torna `None`; un `None` passato senza controllo a `getCellByPosition()` solleva un'eccezione a metà della costruzione dell'elemento XML in corso, lasciando nodi XML parziali/corrotti nel file di export. Va sempre controllato esplicitamente prima dell'uso.

## `numera_voci()` opera sempre su `ActiveSheet`

Mai su un foglio passato esplicitamente. Prima di chiamarla va sempre garantito che il foglio corretto sia quello attivo. Un export XPWE lanciato mentre era attivo il foglio "Elenco Prezzi" invece di "Computo" ha causato corruzione della colonna A su "Elenco Prezzi".
