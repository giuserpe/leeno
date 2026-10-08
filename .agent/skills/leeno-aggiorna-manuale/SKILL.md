---
name: leeno-aggiorna-manuale
description: >
  Aggiorna il manuale ufficiale di LeenO (MANUALE_LeenO.fodt) dalle modifiche al codice o
  alla PWA Brogliaccio, senza duplicazioni grazie al file di tracking. Include revisione,
  mappa sezioni, immagini e screenshot (generati o forniti dall'utente), PDF e consegna zip.
---

# LeenO – Aggiornamento Manuale con Tracciamento

Mantiene sincronizzato il manuale utente (`documentazione/MANUALE_LeenO.fodt`) con il codice, senza ripetere il lavoro su commit o funzionalità già documentati.

## Prima di iniziare
1. Leggi `AGENTS.md` (root) e `documentazione/LESSONS_MANUALE.md`: contengono le regole del progetto e le insidie ricorrenti.
2. Lavora sempre e solo sul branch `dev`. Se il repository non è presente: `git clone --depth 1 --branch dev https://github.com/giuserpe/leeno` e, per avere abbastanza storia, `git fetch --depth=300 origin dev`.
3. Tutti i percorsi sotto sono relativi alla radice del repository (sul PC dell'utente: `W:\_dwg\ULTIMUSFREE\_SRC\leeno`).

## File Chiave
- **Manuale**: `documentazione/MANUALE_LeenO.fodt` (XML, oltre 3 MB, fine riga LF)
- **Manuale PDF**: `src/Ultimus.oxt/MANUALE_LeenO.pdf` (generato da `genera_pdf.py`)
- **Registro aggiornamenti**: `documentazione/TRACKING_MANUALE.md`
- **Mappa sezioni**: `.agent/skills/leeno-aggiorna-manuale/MAPPA_SEZIONI.md`
- **Script**: `.agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py`, `aggiorna_indice.py`, `genera_pdf.py`; per gli screenshot di Brogliaccio `scripts/genera_screenshot_brogliaccio.py`

## Sorgenti di Informazione
Non limitarti ai file Python: consulta tutte le sorgenti.

| Sorgente | Percorso | Cosa cercare |
| :--- | :--- | :--- |
| Codice Python | `src/Ultimus.oxt/python/pythonpath/*.py` | Nuove funzioni `MENU_*`, modifiche ai flussi utente |
| Menù | `src/Ultimus.oxt/Addons.xcu` | Voci di menù nuove o rinominate, etichette |
| Scorciatoie | `src/Ultimus.oxt/Accelerators.xcu` | Nuove scorciatoie o modifiche |
| Dialoghi | `src/Ultimus.oxt/dialogs/*.xdl`, `*.properties` | Finestre, controlli, campi rinominati, testi |
| Configurazione | `src/Ultimus.oxt/python/pythonpath/LeenoConfig.py` | Nuove opzioni |
| Brogliaccio (PWA) | `tools/appunti-cantiere/` (`app.js`, `index.html`, `core.js`) | Schermate, etichette, pulsanti, versione (`VERSIONE` in `core.js`) |

> [!IMPORTANT]
> Percorsi di menù, scorciatoie ed etichette nel manuale vanno copiati dalle sorgenti (`Addons.xcu`, `Accelerators.xcu`, `app.js`), mai ricostruiti a memoria.

## Procedura Operativa

### Fase 0: Mappa delle sezioni
Apri `MAPPA_SEZIONI.md`, individua la sezione adatta e annota riga e bookmark del punto di inserimento. Se il manuale è cambiato di recente, rigenera la mappa prima: `python3 .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py`.

### Fase 1: Storico (tracking)
Leggi `documentazione/TRACKING_MANUALE.md` e individua l'ultimo hash o la versione già documentati (per Brogliaccio, l'ultima versione citata, es. "v0.4.7").

### Fase 2: Novità da documentare
1. Elenca le modifiche successive all'ultima voce di tracking: `git log --format='%h %ad %s' --date=short <hash>..HEAD -- <percorsi sorgente>` e, per la PWA, `git diff <hash> HEAD -- tools/appunti-cantiere`.
2. Escludi ciò che è già nel tracking e ciò che non cambia il comportamento visibile all'utente (refactoring, licenze, formattazione).
3. Per ogni novità verifica menù, scorciatoie ed etichette nelle sorgenti della tabella.
4. Se non è chiaro su quale commit o funzione lavorare, chiedi all'utente.

### Fase 3: Modifica del Manuale (FODT)
1. **Niente editor massivi né script complessi.** Usa la mappa per trovare il punto e ispeziona il contesto XML.
2. Sostituzioni stringa-per-stringa in Python, con fine riga preservati e controllo di unicità:
   ```python
   p = 'documentazione/MANUALE_LeenO.fodt'
   t = open(p, encoding='utf-8', newline='').read()
   vecchio = "testo esatto compresi i tag XML"
   nuovo = "nuovo testo"
   assert t.count(vecchio) == 1, t.count(vecchio)
   open(p, 'w', encoding='utf-8', newline='').write(t.replace(vecchio, nuovo))
   ```
3. **Valida sempre l'XML** dopo le modifiche: `python3 -c "import xml.dom.minidom as m; m.parse('documentazione/MANUALE_LeenO.fodt')"`.
4. Riusa gli stili esistenti del contesto (es. `P440` testo, `P437` titolo livello 3, `T838` etichette dell'interfaccia). Non lasciare nel repository script usa-e-getta.
5. Nessuna icona accanto ai nomi dei comandi (decisione del 2026-08-06).
6. Scrivi in **italiano chiaro e formale**, per l'utente finale, senza terminologia informatica (salvo tasti e menù). Terminologia di Brogliaccio: "agenda di cantiere".
7. Evita "automatico/automaticamente" quando descrivono solo un comportamento del software ("viene generato automaticamente" → "viene generato"). Mantienili se fanno parte del significato o del nome di un comando (es. "Pesca codice automatico").
8. Se aggiungi o rinomini un titolo, aggiorna la voce corrispondente dell'indice (testo nel FODT) o rigenera gli indici da LibreOffice.

### Fase 4: Immagini e screenshot (all'occorrenza)
Un'immagine si inserisce quando il testo da solo non basta: dialogo nuovo o modificato, schermata o barra degli strumenti nuova, area del foglio con un aspetto da riconoscere a colpo d'occhio, procedura in più passaggi su finestre diverse. Non serve per comandi semplici, scorciatoie e voci di menù (bastano percorso e tasti). Nessuna icona accanto ai nomi dei comandi.

1. **Decidi per ogni novità** se serve un'immagine. Se la novità cambia un dialogo o una schermata già illustrati, l'immagine esistente va **sostituita**: cerca nel FODT il titolo o il rimando `(Figura N)` della sezione, altrimenti il manuale mostra una finestra che non esiste più.
2. **Origine dell'immagine**:
   - *Brogliaccio*: rigenera con `python3 scripts/genera_screenshot_brogliaccio.py <cartella>` (richiede Playwright, Pillow, `pdftoppm`) e sostituisci i tre PNG del capitolo (frame `BrogliaccioPrincipale`, `BrogliaccioGiornata`, `BrogliaccioStampa`). Schermate da telefono larghe 5 cm, stampa 12 cm, didascalie con stile `PBrogCap`.
   - *Dialoghi, menù e fogli di LeenO*: non si generano in modo affidabile in headless. Usa gli screenshot forniti dall'utente (allegati in chat o in una cartella indicata). Se servono e mancano, elenca all'utente quali (dialogo o foglio, stato, dati fittizi da usare) e attendi: niente segnaposto nel FODT, niente immagini ricostruite a mano, e l'aggiornamento non si dichiara concluso finché le immagini richieste non sono inserite o l'utente non rinuncia.
3. **Screenshot ricevuti dall'utente come input**: aprili con `view` e confronta etichette, titoli e campi con `.xdl`, `.properties` e `Addons.xcu`. Il testo del manuale segue le sorgenti; se lo screenshot differisce, segnalalo (immagine obsoleta o sorgente non ancora distribuita) prima di inserirlo.
4. **Requisiti dell'immagine**: PNG, dati fittizi, nessun dato personale o reale (nomi, committenti, cantieri, importi), ritaglio sull'area utile, tema chiaro. Larghezza coerente con le figure vicine, al massimo 16,8 cm (pagina intera); altezza proporzionale all'originale.
5. **Inserimento nel FODT**: copia il blocco di una figura esistente dello stesso capitolo (tabella a una cella con frame `fr4`, casella di testo, immagine in `office:binary-data` base64, didascalia `Figura <text:sequence text:name="Figure" ...>N</text:sequence>: descrizione`) e adatta immagine, larghezze, altezze e `text:ref-name` (univoco, es. `refFigure48`). Per Brogliaccio restano gli stili `PBrogFig`, `PBrogCap`, `frBrog`.
6. **Rimandi alle figure**: nel testo si scrivono sempre come riferimenti incrociati, mai come numero digitato: `(Figura <text:sequence-ref text:reference-format="value" text:ref-name="refFigureN">N+1</text:sequence-ref>)`, dove `refFigureN` è il `text:ref-name` della didascalia (vale `numero - 1` finché la numerazione è progressiva) e il testo interno è il numero attuale. Così il numero segue la figura se altre ne vengono aggiunte o tolte. Prima di scegliere la figura di destinazione guardane l'immagine: il numero ricordato a memoria può essere obsoleto. Cerca anche eventuali rimandi digitati a mano (`Figura ` seguito da cifre, anche spezzati su più span) e convertili.
7. **Verifica**: valida l'XML (Fase 3) e controlla l'impaginazione nel PDF (Fase 8). Il FODT cresce di molto con le immagini: segnalalo nella conclusione.

### Fase 5: Tracking
Accoda a `TRACKING_MANUALE.md` una riga: `| YYYY-MM-DD | hash o file | cosa è cambiato nel codice | sezione del manuale |`.

### Fase 6: Stato di revisione
Ad **ogni** aggiornamento aggiorna la tabella "Stato di revisione" all'inizio del manuale (bookmark `__RefHeading___Toc12591_579480652`; colonne Numero / Data / Descrizione / Nome).
1. Parti dall'ultima `<table:table-row>` della tabella e copia gli stili delle celle (`Table5.A2` / `Table5.D2`).
2. **Numero**: progressivo coerente (es. `3.26.xx-rev2.12`). **Data**: mese e anno correnti. **Descrizione**: riepilogo di tutte le modifiche della sessione. **Nome**: vuoto, salvo indicazioni.
3. Più aggiornamenti nella stessa sessione: modifica la riga appena aggiunta invece di crearne una per modifica.

### Fase 7: Mappa sezioni
Rigenera sempre: `python3 .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py`.

### Fase 7bis: Indice generale
Dopo ogni modifica al manuale rigenera sempre i numeri di pagina dell'indice, **prima** del PDF: `python3 .agent/skills/leeno-aggiorna-manuale/scripts/aggiorna_indice.py`. Lo script fa aggiornare l'indice a LibreOffice su una copia temporanea e riporta nel FODT originale solo i numeri di pagina, voce per voce, senza riscrivere il resto del file. Se sono stati aggiunti o rinominati titoli si ferma e chiede di aggiornare l'indice da LibreOffice (Strumenti > Aggiorna > Tutti gli indici).

### Fase 8: PDF
Genera sempre il PDF: `python3 .agent/skills/leeno-aggiorna-manuale/scripts/genera_pdf.py` (LibreOffice headless, da `documentazione/MANUALE_LeenO.fodt` a `src/Ultimus.oxt/MANUALE_LeenO.pdf`).
- Verifica con `pdftotext` che compaia una frase appena inserita.
- Se hai aggiunto immagini, controlla l'impaginazione (`pdftoppm -png -r 50 -f N -l M`): niente buchi di pagina.
- La dimensione dipende dalla versione di LibreOffice (osservato 2,0 MB contro 1,8 MB): segnalala all'utente se cambia sensibilmente.

### Fase 9: Consegna
1. **Commit** su `dev`: tipo `docs`, scope `docs`, solo intestazione (max 72 caratteri, italiano, imperativo, senza punto finale); corpo solo se indispensabile e di una riga. Aggiungi le righe di attribuzione richieste dall'ambiente.
2. **Push**: un solo tentativo. Un errore 403 indica permessi mancanti (app GitHub non installata o non collegata): non ripeterlo, riporta la causa.
3. **Zip, sempre**, anche se il push riesce: tutti i file toccati con le cartelle del repository (`documentazione/MANUALE_LeenO.fodt`, `documentazione/TRACKING_MANUALE.md`, `src/Ultimus.oxt/MANUALE_LeenO.pdf`, `.agent/skills/leeno-aggiorna-manuale/MAPPA_SEZIONI.md`, più gli script nuovi o modificati). Verifica il contenuto con `unzip -l`.

### Fase 10: Conclusione
Comunica l'esito mostrando il testo inserito e conferma: tracking, stato di revisione, mappa, immagini inserite o sostituite (e quelle ancora richieste all'utente), PDF (con eventuali anomalie di dimensione), esito del push e zip consegnato.
