# Lezioni apprese – Pulizia di codice morto e duplicato (agosto 2026)

Durante una pulizia sistematica di `pythonpath/` con `pyflakes` e analisi AST mirata sono emersi pattern ricorrenti, utili come checklist per le pulizie future.

## Gli script "usa e getta" in `pythonpath/` sono il rischio più grave, non solo disordine

Trovato `_fix_path.py`: a livello di modulo apriva `pyleeno.py` e lo riscriveva su disco sostituendo un range di righe hardcoded — se il modulo fosse mai stato importato dal processo di LibreOffice, avrebbe corrotto silenziosamente `pyleeno.py` con un range ormai disallineato. Stesso discorso per `benchmark.py`: path hardcoded (`W:\...`) e `print()` eseguiti a livello di modulo, quindi un `FileNotFoundError` per chiunque non abbia esattamente quel file. Regola operativa: uno script one-shot va eseguito ed eliminato subito dopo l'uso, mai lasciato in `pythonpath/`, nemmeno "per sicurezza".

## La redefinition nello stesso scope è un indicatore affidabile di codice morto

`python3 -m pyflakes <file>` segnala "redefinition of unused X from line Y" quando una funzione (o import) viene ridefinita nello stesso modulo/classe prima di essere usata: la prima definizione non è mai raggiungibile. Trovati due casi reali in `pyleeno.py` (`count_clipboard_lines`, `struttura_Registro`, entrambe con una versione più vecchia "morta" prima di quella attiva). Va trattato come codice da rimuovere, non come nota di stile.

## I moduli `LeenoImport_Xml*.py` vengono clonati l'uno dall'altro

Spesso portano con sé l'intero header di import del file sorgente, incluso il blocco `from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED)`, quasi mai usato per intero nel nuovo file. Un giro di `pyflakes` sul singolo modulo appena clonato individua questi import morti in pochi secondi — utile farlo subito dopo aver creato un nuovo import regionale, non solo in sede di pulizia generale.

## Prima di rimuovere una variabile "assegnata e mai usata" apparentemente inutile, controllare i moduli fratelli

Se lo stesso pattern (es. un campo estratto dall'XML ma non incluso nel titolo composto) si ripete identico in più moduli `LeenoImport_Xml*.py`, è quasi sempre una scelta di design ricorrente e non un refuso isolato — la rimozione va fatta comunque (il dato resta inutilizzato), ma senza trattarla come "correzione di un bug".

# Lezioni apprese: sweep di novembre 2026 ed eccezioni silenziose

## Sweep con pyflakes: esito e regole

Da 119 avvisi in 33 moduli a 36 (14 in `pyleeno.py`, 5 in `LeenoContab.py`, gli altri 1-2 per modulo). Regole che hanno funzionato:

- I nomi non definiti segnalati da pyflakes restano identici prima e dopo ogni PR.
- Assegnazione con chiamata a metodo a destra (`x = oSheet.metodo(...)`): togliere solo `x =`. Assegnazione con sola lettura (`.String`, `.Value`, `getCount()`, `set(...)`): togliere la riga intera, e il `try` che resta vuoto.
- Un'istruzione isolata dentro un `try` può essere una sonda voluta: `LeenoUtils.is_usable` legge `comp.getSheets()` e `comp.CurrentController` per far fallire il controllo su un documento chiuso. Non toccarla.
- `except Exception as e` con `e` mai usata: togliere solo `as e`.
- Un solo import inutile può generare più avvisi: `from pydoc import doc` (import automatico) causava anche tre "redefinition of unused 'doc'" in `LeenoEtransmit.py`.
- Togliere un import inutile può eliminare un ciclo: `LeenoSheetUtils` e `LeenoAnalysis` si importavano a vicenda.
- Nessuna rimozione era un re-export (controllato per `Modulo.nome`, `PL.nome`, `from Modulo import`, dispatcher e Basic).

## Nomi non definiti storici (8): non sono regressioni

Falsi positivi: `CALL_SET` in `LeenoBasicBridge` (definito in `builtins`), `__global_context__` in `LeenoUtils` (iniettato da LibreOffice), `PL` in `LeenoImport_XmlLiguria` (protetto da `'PL' in globals()`, quindi il ramo non gira mai), `Toolbars` in `LeenoConfig` (import locale in un'altra funzione, errore nascosto da un `try`).

Difetti reali: `pyleeno.MENU_inserisci_nuova_riga_con_descrizione` richiama `inserisci_nuova_riga_con_descrizione_th`, che non esiste, ed è collegata a una voce di menu attiva in `Addons.xcu` (NameError al click). `LeenoExtra.convert_to_pdf` usa `subprocess` con l'import commentato; non è collegata a nessun menu.

## Eccezioni silenziose e registro errori

- `except Exception: pass` nasconde i guasti: se un'operazione a più fasi fallisce a metà, i dati restano incoerenti senza avviso. L'audit ne ha contati 113 in `pyleeno.py` e `LeenoContab.py`.
- `DLG.chi()` non è un log: apre un dialogo modale con `dir()` dell'oggetto, e in un ciclo mostra decine di finestre. Per gli errori a runtime: `LeenoUtils.log_errore(contesto, exc)`.
- `log_errore` scrive su `leeno_errori.log` (cartella di configurazione utente di LibreOffice, altrimenti cartella temporanea), registra al massimo 3 occorrenze per contesto e tipo di eccezione per sessione, ruota il file oltre 512 KB e non solleva mai eccezioni.
- Nel blocco `except` usare `as _exc`, non `as e`: Python cancella la variabile a fine blocco e un `e` già usato nella funzione provoca `NameError` più avanti.
- Operazioni a più fasi (esempio `LeenoContab._annulla_ultimo_sal_core`): registrare ogni errore, raccogliere le fasi fallite in una lista e mostrare un solo avviso a fine operazione. Gli errori di pulizia ricostruibili (CdP rigenerato) si registrano senza avvisare.
- Nei cicli che ripetono una funzione finché lo stato cambia (`MENU_AnnullaTuttiAttiContabili`) verificare che lo stato sia cambiato davvero: se l'eliminazione fallisce prima di rimuovere l'intervallo con nome, `ultimo_sal()` restituisce sempre lo stesso SAL e il ciclo non termina.
- `undo_utils.with_undo_batch` non è usata da nessuna parte e probabilmente non funziona: blocca l'undo con `lock()` e poi apre un contesto vuoto, quindi non raggruppa nulla.

## Parsing XML in Python 3.12+ e parser regionali

- **ElementTree truth value**: In Python 3.12+, valutare direttamente il booleano di un nodo `Element` (`if not elem:`) solleva un `DeprecationWarning` e diventerà un'eccezione nelle versioni successive. Usare sempre controlli espliciti: `if elem is None:` oppure `if len(list(elem)) < ...:`.
- **Parser regionali coperti**: Tutti i parser XML regionali (`Basilicata`, `Liguria`, `Lombardia`, `Sardegna`, `SIX`, `Toscana`, `Veneto`) dispongono ora di test di caratterizzazione dedicati in `tests/test_import_xml_<regione>.py`. In caso di evoluzione dei formati o refactoring dei moduli `LeenoImport_Xml*.py`, verificare sempre che i relativi test continuino a passare senza introdurre dipendenze PyUNO non isolate.
