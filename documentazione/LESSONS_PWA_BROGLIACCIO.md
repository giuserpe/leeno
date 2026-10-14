# Lezioni apprese – Agenda di cantiere (Brogliaccio, PWA in `tools/appunti-cantiere/`)

PWA statica (HTML/CSS/JS puro, nessun framework, nessuna dipendenza esterna in produzione) per raccogliere appunti di cantiere su smartphone e consolidarli in LeenO tramite `LeenoGiornaleImport.py`. Pubblicata su GitHub Pages all'indirizzo `https://giuserpe.github.io/leeno/brogliaccio/` dal workflow `genera_versioni.yml`, non sotto `leeno.org`: l'origine è schema più host, non il percorso, quindi un percorso condivide `localStorage`/IndexedDB con tutto il resto dell'host. L'origine `giuserpe.github.io` è condivisa con ogni altra pagina GitHub Pages dell'account: su quell'host non devono girare pagine con script di terzi né pagine che toccano lo storage.

## Non rinominare mai una chiave localStorage/IndexedDB senza un recupero

Una chiave di storage (`appunti-cantiere-v1`, poi brevemente `-v2`, poi di nuovo `-v1` con forma interna diversa) è stata cambiata durante lo sviluppo partendo dal presupposto "nessuno la sta ancora usando". Quando l'app è arrivata nelle mani dell'utente reale, un cambio di versione successivo ha fatto sparire una giornata già compilata: l'app cercava i dati sotto il nome/forma corrente e non trovava nulla sotto quello vecchio.

Regola operativa, da applicare anche durante lo sviluppo pre-rilascio: qualunque cambio al nome o alla forma di una chiave di storage persistente va accompagnato, nello stesso cambiamento, da un controllo di recupero che cerchi i dati nel nome/formato precedente e li migri in automatico (vedi `carica()` in `app.js` per il pattern già in uso: prova il formato corrente, poi la chiave intermedia nota, poi il formato piatto precedente, scrivendo subito il risultato nel formato corrente). Non rimandare la rete di sicurezza a "tanto non l'ha ancora usato nessuno": quel momento passa senza preavviso.

## Web Share API: file pronto prima, condivisione con un tocco fresco

`navigator.share()` può fallire dopo un lavoro asincrono (es. una lettura IndexedDB prima di costruire il file) perché il browser considera scaduta la "transient activation" concessa dal tocco dell'utente. Brave è più severo di Chrome su questo punto e rifiuta la condivisione dove Chrome l'avrebbe concessa.

Pattern adottato dalla 0.5.4: "Esporta per LeenO" costruisce il file (asincrono) e, se `navigator.canShare({ files })` lo consente, mostra un riquadro con "Invia o condividi" e "Salva sul telefono"; `share()` parte direttamente dal click sul pulsante, quindi con attivazione valida. Senza supporto alla condivisione il file si salva subito via `<a download>`. Qualunque fallimento di `share()` (annullamento incluso) lascia il riquadro aperto: il file non va mai perso e "Salva sul telefono" resta disponibile. `ultimo_export` si registra solo a condivisione riuscita o a file salvato, non alla semplice preparazione.

Correzione della 0.5.5, causa probabile dei rifiuti di `share()` su Brave: Chromium accetta in `share()` solo alcuni tipi di file (testo, immagini, audio, video, PDF; tra le applicazioni solo `application/pdf`, vedi la lista "Permitted File Extensions" di Chromium). `.json` e `.zip` vengono rifiutati con `NotAllowedError` anche quando `canShare()` risponde vero, perché il controllo del tipo avviene nel processo del browser al momento della condivisione. Pattern adottato: il .json si condivide come testo (`.json.txt`, `text/plain`; l'import di LeenO legge il JSON senza guardare l'estensione), lo .zip con foto non si condivide e si salva direttamente. Il messaggio di errore riporta il nome dell'errore per diagnosticare. Verificato solo con uno stub che imita la regola, non su Brave reale.

Perché non più il download automatico per primo (pattern delle versioni fino alla 0.5.3): il file finiva nella cartella Download, difficile da ritrovare per l'utente medio, mentre il menù di condivisione porta il file dove serve (posta a se stessi, cloud).

## Rimozione EXIF e orientamento foto senza libreria dedicata

`createImageBitmap(file, { imageOrientation: 'from-image' })` seguito dal ridisegno su un `<canvas>` e riesportazione (`canvas.toBlob('image/jpeg', qualità)`) applica l'orientamento EXIF corretto e scarta contestualmente tutti i metadati EXIF, inclusa la posizione GPS, senza bisogno di una libreria di parsing EXIF. Utile ogni volta che l'app acquisisce foto da fotocamera/galleria e quei metadati non vanno esportati.

## CSS di stampa: `position: fixed` per un piè di pagina ripetuto

Un elemento con `position: fixed` dentro `@media print` si ripete identico su ogni pagina fisica quando Chromium genera il PDF (comportamento verificato empiricamente, non sempre documentato in modo affidabile). Permette un piè di pagina o un'intestazione ripetuti senza bisogno di CSS Paged Media (`@page` margin boxes, supporto incompleto nei browser).

Verifica: uno screenshot in `emulate_media(media='print')` mostra solo il flusso continuo del contenuto, non l'impaginazione reale, quindi non basta a dimostrare che un elemento si ripeta su più pagine. Serve generare un vero PDF multi-pagina (Playwright `page.pdf()`, richiede Chromium headless) e leggere il testo di ciascuna pagina (es. con `pypdf`) per confermare la presenza dell'elemento su ognuna.

## Non riusare un tag HTML già investito di stile globale

Il blocco di intestazione generato per il contenuto stampabile riusava il tag `<header>`, già definito nel foglio di stile dell'app con sfondo scuro e layout flex per l'intestazione dell'interfaccia. Il risultato: il PDF ereditava quello stile invece di restare testo semplice in nero, un bug silenzioso notato solo rendendo la pagina in modalità stampa. Dare sempre al contenuto generato dinamicamente per un contesto diverso (stampa, export) un proprio tag/classe, mai un tag strutturale già usato altrove nell'app con uno scopo visivo diverso.

## Scrittore ZIP minimo: validare sempre con uno strumento indipendente

Un file ZIP valido (formato "store", senza compressione) si scrive a mano in poche righe quando il contenuto è già compresso (es. foto JPEG): CRC32, intestazioni locali, central directory, end of central directory. Un file del genere può "scaricarsi senza errori" ed essere comunque malformato in modi che solo un lettore ZIP indipendente rileva. Validare sempre con uno strumento esterno al codice che l'ha generato (es. `zipfile` di Python, `ZipFile.testzip()` per i CRC, lettura e confronto byte-per-byte dei contenuti estratti), mai fidarsi del solo fatto che il browser l'abbia prodotto senza eccezioni.

## Playwright: dialoghi multipli in sequenza da una singola azione

Registrare più `page.once('dialog', ...)` in sequenza prima di innescare una singola azione non li accoda un dialogo per listener: tutti i listener attualmente in ascolto ricevono il *primo* dialogo emesso, e il secondo listener che tenta di rispondere a un dialogo già gestito solleva un errore ("dialog already handled"). Per gestire una sequenza di dialoghi di lunghezza o ordine non noti in anticipo, usare un singolo `page.on('dialog', ...)` persistente che decide l'azione in base a `dialog.type` (e, se serve, a un contatore), non più `page.once(...)` accodati.

## Ripristino da .zip: cantiere per nome, foto senza duplicati

Il .zip dell'esportazione porta il nome del cantiere solo in `testata.lavori` del .json: il ripristino cerca il cantiere per nome (senza badare a maiuscole) e, se manca, lo crea. Il pulsante "Ripristina da file" deve restare visibile anche senza cantieri, altrimenti il caso principale (telefono nuovo o dati persi) non è raggiungibile. Le foto non hanno un identificativo stabile fuori dal dispositivo: si riconoscono per giornata, dimensione e CRC32, così ripetere lo stesso ripristino non le duplica. Il progressivo `_NNN` del nome file torna millisecondi di `creato_il`: mantiene l'ordine delle foto dello stesso minuto, e una nuova esportazione riproduce gli stessi nomi. Si scrive prima il testo (localStorage) e poi le foto (IndexedDB): se le foto falliscono a metà, rilanciare lo stesso file completa il lavoro. Verificato con Chromium reale (esporta, ripristina in un browser vuoto, ri-esporta e confronto byte per byte; zip deflate, alterato, troncato, senza json).

## Esportare e ripristinare tutti i cantieri

Il formato di un cantiere solo (json alla radice, foto in `AAAAMMGG/`) è letto da `LeenoGiornaleImport.leggi_pacchetto`, che pretende un solo .json alla radice dello .zip: non va cambiato. L'esportazione di tutti i cantieri usa quindi la stessa struttura dentro una cartella per cantiere, e il ripristino associa ogni foto al .json della propria cartella. Con un solo .json nel file tutte le foto vanno a quel cantiere, ovunque stiano: così funziona anche uno .zip ricompresso dentro un'altra cartella. Due cantieri con lo stesso nome (l'app lo consente) ottengono cartelle `nome` e `nome-2` e, ripristinati, confluiscono in un solo cantiere: il nome è l'unica chiave, per scelta. Un cantiere senza giornate non viene esportato (non ci sarebbe nulla da ripristinare). Il ripristino non scrive nulla finché l'utente non conferma (casella per cantiere, o conferma classica con un cantiere solo). Nei test con Chromium non aprire `brogliaccio-foto` prima dell'app: un `indexedDB.open()` fuori dall'app crea il database vuoto alla versione 1, senza magazzino `foto`, e l'app non lo ricrea più. Verificato su Chromium reale (esporta tutti, ripristina scegliendo, ripetizione senza duplicati, omonimi, zip ricompresso con cartella contenitore), e con il codice vero dell'importatore su uno .zip massivo (errore chiaro) e uno singolo (importato).

## Licenza e file pubblicati

La cartella `tools/appunti-cantiere/` va caricata sul server così com'è: ogni file presente finisce online, anche se non referenziato (è successo con `jspdf.min.js`, MIT, mai usato). Prima di aggiungere una libreria di terzi verificare che sia davvero caricata da `index.html` e, se resta, conservarne intestazione e licenza originali. I file propri portano l'intestazione LGPL 2.1; l'`index.html` la ha come commento HTML, senza testo visibile (nessun impatto su screenshot e manuale).

## Pubblicazione

- `genera_versioni.yml` (copie in `.github/workflows/` e in `tools/version/`) assembla a ogni push su `dev` un solo sito Pages: `versions.html` alla radice e questa cartella in `brogliaccio/`, senza `README.md`, poi lo pubblica con `deploy-pages`. Indirizzo: `https://giuserpe.github.io/leeno/brogliaccio/`. La stessa cartella viene copiata anche su `master`. `versions.html` è incorporata in `leeno.org/versioni-sviluppo/`.
- Serve HTTPS (Service Worker e installazione): GitHub Pages lo fornisce. Il manifest usa percorsi relativi, quindi l'indirizzo deve finire con `/`.
- Non cambiare l'indirizzo dopo l'uso: i dati degli utenti sono legati all'origine. Brogliaccio è stato indicato in passato su altri indirizzi (`brogliaccio.leeno.org`, `www.giuserpe.it/brogliaccio/`): chi ha installato l'app da uno di quelli ha i dati in un archivio diverso e non li vede dall'indirizzo attuale. Per spostarli si esporta dal vecchio indirizzo e si ripristina dal nuovo.
- GitHub Pages ha un solo sito per repository e ogni deploy sostituisce il precedente: un secondo workflow con `deploy-pages` (provato con `brogliaccio-pages.yml`) fa sparire `versions.html`. Non aggiungerne altri.

## Più istanze sulla stessa origine: revisione e merge

Fino alla 0.7.1 `salva()` riscriveva l'intero stato dalla memoria della pagina senza guardare cosa c'era già salvato. Con due istanze aperte sulla stessa origine (scheda del browser e app installata, o una pagina rimasta in pausa) l'istanza più vecchia, al primo salvataggio, cancellava ciò che l'altra aveva ripristinato o scritto. Riprodotto in Chromium: una pagina caricata prima del ripristino, che poi crea un cantiere, fa sparire il cantiere ripristinato.

Dalla 0.7.2 lo stato porta `rev`, incrementato a ogni salvataggio. Prima di scrivere, `salva()` confronta `rev` con quello salvato: se è più alto porta in memoria i dati dell'altra istanza (solo aggiunte e giornate con `modificato_il` più recente, mai cancellazioni; la giornata aperta in modifica non si tocca). Gli eventi `storage` e `visibilitychange` aggiornano la schermata quando non c'è una giornata in modifica. Non togliere questo controllo da `salva()`. La chiave `appunti-cantiere-v1` non cambia: `rev` è un campo in più e `carica()` lo assume 0 se manca.

Per provarlo senza l'aiuto degli eventi, caricare la pagina "vecchia" con un init script che ignora `addEventListener('storage'|'visibilitychange')`, ripristinare da un'altra pagina, poi salvare dalla prima: senza `rev` i dati ripristinati spariscono.

## Diagnostica della perdita dati

Il browser può cancellare lo storage dell'origine (cancellazione all'uscita, spazio scarso, memoria non persistente) e l'app non può accorgersene dopo, perché sparisce anche ciò che avrebbe registrato. Dalla 0.7.2 rende visibile ciò che si può:

- `navigator.storage.persisted()` (e `persist()` se negato) all'avvio: in fondo alla schermata principale compare "Memoria del browser: protetta" oppure "non protetta" (con invito a esportare spesso), più il numero di avvio.
- Chiave `brogliaccio-diag` (nuova, nessuna rinomina): contatore degli avvii e ultimo stato noto (cantieri e giornate), aggiornato a ogni `salva()`. All'avvio, se le giornate sono meno dell'ultimo stato registrato dall'app stessa, i dati sono stati tolti da fuori: compare un avviso fisso ("Ho capito" lo chiude).
- Foto in IndexedDB ma nessun cantiere (testo perso, foto no): avviso analogo.
- Dopo "Ripristina da file" l'app rilegge dalla memoria del browser (non dalle variabili) e scrive "Verificato: ..." oppure un avviso di errore.

Se la cancellazione toglie tutto (testo, diagnostica e foto) non c'è nulla da rilevare: la difesa resta il file esportato, da cui il promemoria "Ultimo export".

## Stampa con foto: attendere la decodifica prima di `window.print()`

Fino alla 0.7.2 `stampaPDF()` inseriva le foto come immagini con data URL in `#stampa` e chiamava subito `window.print()`: le immagini non erano ancora decodificate e il PDF mostrava riquadri vuoti. Riprodotto in Chromium con 8 foto da 1600x1200: al momento di `print()` risultavano pronte 0 su 8. Dalla 0.7.3 `attendiImmagini()` attende `img.decode()` su tutte le immagini (con un tetto di 20 s) e solo dopo stampa. Con foto leggere il difetto può non comparire: provarlo sempre con foto pesanti e controllando `img.complete && img.naturalWidth > 0` dentro un `window.print` sostituito.
