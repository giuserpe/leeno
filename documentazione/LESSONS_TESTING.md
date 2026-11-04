# Lezioni apprese – Pipeline di test automatizzato (headless UNO)

Per il test di round-trip XPWE (export → import → confronto riga per riga) è stata realizzata una pipeline che usa istanze reali di LibreOffice in modalità headless, senza mock.

## Gestione del ciclo di vita del processo `soffice`

Include gestione dei file di lock, avvio con `nohup`+`setsid` per staccare il processo dalla sessione del chiamante, e attesa attiva della disponibilità del socket UNO prima di procedere con i comandi di test.

## Catena di import circolare

`pyleeno↔Debug↔Dialogs↔LeenoContab↔LeenoContab`. Va risolta importando sempre `Dialogs` per primo nell'ambiente di test; un ordine diverso reintroduce l'errore di import circolare.

## Monkeypatch per bypassare la registrazione `.oxt`

`LeenO_path()` e `basic_LeenO()` vanno monkeypatchate nei test per bypassare la registrazione dell'estensione come `.oxt`, che altrimenti non è disponibile nell'ambiente di test headless.

## Collocazione dei file di test

Questi file di test, come da regola generale sulla sicurezza di `pythonpath/` (vedi `AGENTS.md`), non vivono in `src/Ultimus.oxt/python/pythonpath/`.

## Testare codice che scrive percorsi relativi al documento

Un documento aperto da template (`loadComponentFromURL` su un file appena copiato) ha `oDoc.getURL()` valorizzato con l'URL di apertura, ma qualunque codice che debba risolvere "la cartella accanto al documento" va testato su un documento effettivamente salvato con `storeToURL()` e poi riaperto: altrimenti il test passa per ragioni accidentali (l'URL è comunque presente) invece di verificare il percorso reale che l'utente avrà (un documento aperto, modificato, mai salvato con un nome proprio). Il modulo di import dell'agenda di cantiere (`LeenoGiornaleImport._estrai_foto`) usa `uno.fileUrlToSystemPath(oDoc.getURL())` per questo scopo e solleva un errore esplicito se l'URL è vuoto, invece di scrivere un percorso privo di senso.

## Inserimento di righe dentro un blocco: ricalcolare sempre per etichetta, mai per offset fisso

Quando una funzione inserisce una riga dentro un blocco esistente (es. una riga con collegamento ipertestuale subito dopo un campo di testo), qualunque codice scritto *dopo* quell'inserimento, nello stesso blocco, deve ricercare le proprie celle per contenuto dell'etichetta (ricalcolando `_righe_giorni`/confini del blocco da zero), mai per un offset di riga calcolato prima dell'inserimento: l'inserimento sposta fisicamente tutto ciò che sta sotto. Un intervallo con nome (es. `giornale_bianco`) che contiene il punto di inserimento si estende da solo per includere le nuove righe, comportamento verificato ma da non dare per scontato senza test: va confermato con un controllo esplicito del intervallo (`oDoc.NamedRanges.getByName(...).Content`) prima e dopo l'inserimento, non assunto.

## Test senza UNO: funzioni pure e parser

- Stanno in `tests/` e girano con `pytest`. La CI (`.github/workflows/test.yml`) esegue `pytest --ignore=tests/test_brogliaccio_export.py tests/` a ogni push e PR su `dev`, e il test del brogliaccio (Playwright) in un job separato.
- Due modi per caricare il codice. (1) Importare il modulo dentro `unittest.mock.patch.dict(sys.modules, {...})` con `MagicMock` per `uno`, `unohelper`, `com.*`, `pyleeno`, `Dialogs`, `LeenoDialogs`, rimuovendolo da `sys.modules` nel `tearDown` (esempi: `tests/test_import_xml_*.py`, `tests/test_import_xpwe.py`). (2) Estrarre la sola funzione con `ast` ed eseguirla in un namespace minimo (esempi: `tests/test_log_errore.py`, `tests/test_utils_pure.py`).
- Test di caratterizzazione: fotografano il comportamento attuale. Un difetto trovato si documenta con un commento `BUG NOTO:` (input e output osservato) e si corregge in un commit a parte, aggiornando il test.
- Fixture sempre sintetiche: descrizioni, importi e nomi inventati. Per l'XPWE c'è un test di andata e ritorno: `dcf_parser.generate_xpwe` genera il file e le funzioni `leggi*` di `LeenoImport_XPWE` lo rileggono.
- Un test asserisce valori restituiti, non solo l'assenza di eccezioni. I mock servono a caricare il modulo, mai come valore atteso.
- Restano senza copertura automatica le funzioni che richiedono UNO (`compilaComputo`, `XPWE_import`, il rollback del SAL su documento reale): prova manuale in Calc.

## Difetti trovati dai test di caratterizzazione

Corretti:
- Numeri in lettere: `LeenoContab._intero_in_lettere` non elideva "cento" davanti a "ottanta" (180 diventava "centoottanta" invece di "centottanta", 900 divergenze su 0-10000 rispetto a `LeenoUtils.int_to_italian`), e la funzione genera l'importo in lettere del CdP. Un test confronta le due funzioni su 0-20000.
- `LeenoImport._readXmlFile`: i file senza BOM e non in UTF-8 perdevano gli accenti, perché la decodifica usava `errors='ignore'`. Ora ripiega sulla codifica dichiarata nell'intestazione XML e poi su `cp1252`.

Aperti, documentati nei test:
- `paradox_to_xpwe._read_number` restituisce positivi anche i numeri negativi, e `_read_long` tratta i negativi come senza segno. Gli archivi Paradox non sono una priorità.
- "ventitre" e simili sono scritti senza accento nei numeri in lettere, in entrambe le funzioni (decisione da prendere).
- `leggiElencoPrezzi` lascia `prezzo2`-`prezzo5` come stringhe, ma nessuno li usa.
- `Debug._format_time` non ha le ore. `LeenoImport.MENU_ValdAosta` è dichiarata non funzionante nella docstring.
