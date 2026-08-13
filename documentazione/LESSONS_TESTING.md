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
