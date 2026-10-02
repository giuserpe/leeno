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
