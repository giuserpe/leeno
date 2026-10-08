# Lezioni apprese – Agenda di cantiere (Brogliaccio, PWA in `tools/appunti-cantiere/`)

PWA statica (HTML/CSS/JS puro, nessun framework, nessuna dipendenza esterna in produzione) per raccogliere appunti di cantiere su smartphone e consolidarli in LeenO tramite `LeenoGiornaleImport.py`. Pubblicata via GitHub Pages (`.github/workflows/brogliaccio-pages.yml`) su un sottodominio dedicato, non su un percorso sotto `leeno.org`: un percorso (`leeno.org/brogliaccio`) condivide la stessa origine del sito, quindi lo stesso `localStorage`/IndexedDB di tutto il resto del dominio — l'isolamento esiste solo con un sottodominio vero.

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
