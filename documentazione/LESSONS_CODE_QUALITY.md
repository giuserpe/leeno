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

Difetti reali risolti:
- `pyleeno.MENU_inserisci_nuova_riga_con_descrizione`: ripristinata l'intestazione della classe thread `inserisci_nuova_riga_con_descrizione_th(threading.Thread)` precedentemente orfana e annidata in `donazioni()`.
- `LeenoExtra.convert_to_pdf`: de-commentato l'import di `subprocess`.

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

# Lezioni apprese: censimento dei duplicati di macro (novembre 2026)

Regola in AGENTS.md: prima di scrivere una macro o funzione nuova, cercarne una equivalente (per nome e per azione) e riusarla o estenderla. Di seguito il censimento di partenza: 16 cluster reali su 1230 funzioni di `python/` e `pythonpath/`.

## Metodo e falsi positivi

- Hash dell'AST con nomi e costanti normalizzati (cloni esatti), poi similarita' testuale (`difflib`, soglia 0.85) sulle funzioni da 8 righe in su.
- Le omonimie tra `LeenoBasicBridge.py` e `pyleeno.py`/moduli (circa 35 nomi: `vai_a_M1`, `invia_voce`, `trova_np`, ...) sono il ponte voluto verso le macro Basic: non sono duplicati.
- Falsi positivi dopo la normalizzazione delle stringhe: `setTabColor`, `setFormatoNumeri`, `txt_Format`, `setAdatta` (stesso schema di dispatch, azioni diverse), `read`/`write` di `LeenoConfig` e `pyleeno`, `MENU_undo`/`MENU_redo`.

## Gia' eliminati (lotti 1 e 2)

`benchmark.py`, `Calendario.pickDate` (ora `Dialogs.pickDate`), `pyleeno.getLastUsedCell` (ora `SheetUtils.getUsedArea`), `_get_ep_last_row` e `_get_analisi_last_row` (ora `SheetUtils.getRigaSentinella`), `_col_letter` in tre copie (ora `SheetUtils.colonnaInLettera`; `LeenoContab._col_letter` resta come delega per `LeenoComputo` e per i test CdP). Il lotto 1 ha anche corretto `LeenoImport.rimuovi_righe_ridondanti`, che chiamava `LeenoUtils.getLastUsedCell`, funzione inesistente: pyflakes non vede gli attributi di modulo, per questo il difetto era rimasto latente.

## Duplicazione voluta: `LeenoDispatcher.py`

`cerca_path_valido`, `dest`, `reloadLeenoModules` e `loVersion` esistono sia nel dispatcher sia in `pythonpath/`. Non vanno unificate: il dispatcher contiene `handle_exception`, l'ultimo gestore di errori, e deve funzionare anche quando un modulo di `pythonpath/` non si importa. Se importasse quei moduli, un errore di import li renderebbe inutilizzabili proprio quando servono. Vale anche per `LeenoDebug.py`, fuori da `pythonpath/`.

## Da unificare con parametri (verificare su documenti reali)

| Cluster | Soluzione |
|---|---|
| `LeenoContab.trova_r_S2`: funzione di modulo e due closure interne | Usare la funzione di modulo. Attenzione: le closure restituiscono a volte l'offset e a volte la riga assoluta |
| `LeenoComputo.circoscriveVoceComputo` e `_circoscrive_voce` (metodo di classe) | Il metodo delega alla funzione (da confrontare riga per riga) |
| `pyleeno.SubSum_SuperCap`, `SubSum_SottoCap`, `SubSum_Cap` | Una funzione con il livello di stile come parametro |
| `LeenoSheetUtils.inserSuperCapitolo` e `inserCapitolo` | Una funzione con livello e testo di default come parametri |
| `LeenoPdf.PdfElencoPrezzi` e `PdfComputoMetrico` | Una funzione con foglio, suffisso e oggetto come parametri (rimuovere anche il `print()` attivo) |
| `LeenoContab._leggi_anticipo_da_S2`, `_leggi_infortuni_da_S2`, `_leggi_iva_da_S2` | Un helper con l'etichetta come parametro |
| `LeenoImport_XPWE.leggiSuperCapitoli`, `leggiCapitoli`, `leggiSottoCapitoli`, `leggiSuperCategorie`, `leggiCategorie`, `leggiSottoCategorie` | Un lettore generico sul tag XML |

## Non toccare

`Dialogs.py` contiene classi widget quasi gemelle (`CheckBox`/`RadioButton`, `ListBox`/`ComboBox`, `getProps`, `dump`, `_fixup` ripetute fino a 13 volte). Il refactor dipende dal comportamento UNO e non si delega a un agente.

## Ordine di intervento

1. Copie esatte (prima tabella): rischio basso, un commit per cluster.
2. Unificazioni con parametri: un commit per cluster, prova su documenti reali.
3. Ogni PR: `scripts/check_pythonpath_safety.py`, pytest, `pyflakes` con nomi non definiti invariati, ricerca del nome rimosso in `Addons.xcu`, `Accelerators.xcu`, `.xba` e `.xdl`.
