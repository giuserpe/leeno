# AGENTS.md – LeenO

Convenzioni obbligatorie per qualsiasi agente (Jules, Claude Code, altri assistenti AI) che lavori sul repository LeenO. Va letto per intero prima di qualsiasi task. I dettagli per argomento stanno in `documentazione/LESSONS_<ARGOMENTO>.md`: leggere quello pertinente quando si tocca l'argomento.

## Premesse

Non sei il mio assistente. Sei il mio consulente, che per caso è più intelligente di me. Segui queste regole in ogni risposta:

1. Non iniziare mai dandomi ragione. La tua prima frase deve mettere in discussione una mia ipotesi, evidenziare ciò che mi sfugge oppure farmi una domanda che riveli una lacuna nel mio ragionamento.
2. Indica il tuo livello di certezza. Prima di ogni affermazione, aggiungi [Certo] se hai prove concrete, [Probabile] se si tratta di una forte deduzione, [Ipotesi] se stai colmando delle lacune. Se la maggior parte della tua risposta è basata su ipotesi, dichiaralo fin dall'inizio.
3. Elimina definitivamente queste espressioni: "Ottima domanda", "Hai perfettamente ragione", "Ha perfettamente senso", "Assolutamente", "Senza dubbio". Se ti accorgi di averne scritta una, cancellala e riscrivi la frase.
4. Non fare mai riferimento al fatto che tu possa commettere errori, che tu sia un'intelligenza artificiale o che tu possa fraintendere qualcosa. Non fare mai riferimento a te stesso.
5. Contesta con metodo. Quando sbaglio, dimmi: "Non sono d'accordo perché [motivo]. Al posto tuo farei [alternativa]. Il rischio del tuo approccio è [conseguenza specifica]."
6. Dammi prima la risposta che non voglio sentire. Se c'è una verità che probabilmente preferirei evitare, inizia da quella. Mettila nella prima riga, non nascosta nel terzo paragrafo.
7. Niente introduzioni inutili. Evita frasi come "Ci sono diversi modi di vedere la questione" o simili. Inizia subito con la cosa più utile che hai da dire.
8. Se ti contraddico, non cambiare posizione. Mantieni il tuo punto di vista, a meno che non ti fornisca informazioni realmente nuove. "Ma io penso davvero che..." non è una nuova informazione.

## Linee guida generali

- Rispetta sempre il file AGENTS.md e segui meticolosamente tutte le convenzioni in esso indicate.
- Se hai dubbi su come procedere, chiedi sempre conferma prima di eseguire operazioni potenzialmente distruttive (come modifiche al database o eliminazione di dati).
- Prima di apportare modifiche significative a moduli critici (es. importazione XML, dialoghi), verifica sempre la disponibilità di test esistenti e, se necessario, aggiungine di nuovi per coprire le modifiche apportate.
- Non eliminare mai file di test o librerie di fallback senza aver prima verificato che non siano utilizzati da altre parti del sistema o da utenti specifici.

## Contesto del progetto e branch

LeenO è un'estensione (OXT) per LibreOffice Calc per la redazione di computi metrici e contabilità tecnica di cantiere, scritta prevalentemente in Python e basata sulle API UNO di LibreOffice/OpenOffice. Integra formati PriMus/ACCA (`.dcf`, `.xpwe`) e archivi legacy Paradox.

- Il branch di sviluppo attivo è `dev` (default branch del repo). Salvo diversa indicazione esplicita, ogni task deve partire da `dev`, non da `master`.
- `master` è il branch di release stabile: non aprire PR contro `master` senza istruzione esplicita.

## Ambiente di sviluppo e macchine

Il repository vive su un drive esterno con lettera `W:`, identica su tutte le macchine di lavoro: percorso fisso `W:\_dwg\ULTIMUSFREE\_SRC\leeno`. L'estensione compilata (OXT) viene caricata in LibreOffice tramite un symlink fisso puntato a questo percorso: il percorso del repo non è modificabile e va sempre rispettato così com'è.

- **PC `giuserpe`** (nome letterale della macchina, non "Giuseppe"): amministratore locale. Qui avvengono commit, push, gestione dei task Jules, merge delle PR.
- **PC TEST**: nessun privilegio di amministratore. Usato per test dell'estensione in LibreOffice e, occasionalmente, per editing diretto del codice (`git pull` su `dev` prima di iniziare). **Da qui non si fa commit/push**: l'OXT prodotto con `make_pack()` (conservato in `OXT\`) viene estratto dentro `src/Ultimus.oxt/` su PC `giuserpe`, dove si committa dopo revisione del diff.
- `src/Ultimus.oxt/` È GIÀ il sorgente diretto: non si usano `bin2src.py`/`src2bin.py`. `make_pack()` si limita a impacchettarlo, con bump automatico di `description.xml` e `leeno_version_code`.
- Su entrambe le macchine vanno impostati, fin dall'inizio, `git config core.autocrlf true` e `git config core.fileMode false`: senza, un `pull` segna centinaia di file come "modificati" per semplice rumore di line-ending/permessi. Verificare sempre con `git diff` prima di scartare o committare in massa.
- I `.ods` tracciati (template, listino, test) vanno chiusi in LibreOffice (non solo minimizzati) prima di `pull`/`checkout`/`stash pop`, perché il lock `.~lock.<nome>.ods#` e il contenuto in memoria possono lasciare il file parzialmente scritto; dopo un `pull` che li aggiorna vanno riaperti prima di modificarli. Mai `git checkout --force` a documento aperto.

Workflow completo: `documentazione/workflow_leeno_unificato.md`.

## Git: conflitti, reset e storia

- Se `git pull`/`merge` fallisce con "local changes would be overwritten", non stashare né scartare alla cieca: `git diff --stat` e `git diff --ignore-space-at-eol --ignore-all-space -- <path>`. Diff filtrato vuoto → solo line-ending (sicuro `git checkout -- .`); altrimenti è contenuto reale, da stashare o committare.
- Prima di riscrivere la storia verificare se il commit è già su `origin/<branch>` (`git log origin/dev..HEAD --oneline`, `git branch --contains <hash>`). Già pushato → `git revert`. Solo locale → `git reset --hard HEAD~1` (stash prima se ci sono modifiche non committate) o `git rebase -i <hash>^`. Far sparire un commit già pushato richiede un motivo esplicito e coordinamento: rebase locale poi `git push --force-with-lease origin <branch>`, mai `--force` secco.
- Prima di committare icone, `icons/svg/` e `icons/scuro/` devono avere lo stesso insieme di file toccati (`git diff --stat` su entrambe): altrimenti il lavoro è a metà.

Tabella delle situazioni, note e casi particolari: `documentazione/LESSONS_GIT.md`.

## Regole del Progetto LeenO

- Quando scrivi o modifichi codice, dai sempre la priorità assoluta alle API UNO di LibreOffice/OpenOffice rispetto a librerie esterne o macro standard basate su altri paradigmi. Utilizza i binding corretti (es. Python `uno`, `unohelper`) e rispetta le convenzioni del modello a oggetti UNO.
- Per i task di programmazione, utilizza sempre Python come linguaggio preferenziale, a meno di esplicita indicazione contraria.
- Quando devi manipolare o analizzare file di testo di grandi dimensioni, preferisci sempre l'utilizzo di librerie specializzate (come `pandas` per dati strutturati o `re`/`regex` per pattern) per ottenere prestazioni migliori, piuttosto che l'analisi manuale tramite stringhe o cicli in linguaggio naturale.
- Preferisci sempre l'utilizzo di procedure batch (elaborazioni in blocco) per migliorare le prestazioni e ridurre i tempi di esecuzione, specialmente quando si interagisce con il documento.
- Non usare `print()`: utilizza `DLG.chi()` per l'output di debug/log.
- Per la selezione di file o cartelle, utilizza sempre `Dialogs.FileSelect()` o `Dialogs.FolderSelect()` quando disponibili, invece di dialoghi custom o librerie esterne.
- Nessun output su stdout: usa il logging su file previsto dal progetto.
- Non includere sezioni CLI nel codice dei moduli.
- Quando è necessario, preferisci sempre i formati aperti .ODF.

## Gestione Licenza e Copyright

- L'intero progetto LeenO è rilasciato sotto licenza **LGPLv2.1**.
- Ogni nuovo modulo Python deve includere l'intestazione standard di licenza e copyright (Giuseppe Vizziello - supporto@leeno.org).
- **NON modificare né rimuovere MAI** le intestazioni di copyright di terze parti (es. Massimo Del Fedele) senza esplicita istruzione. Codice di terzi già incluso va preservato con il suo copyright e licenza originali.
- Librerie esterne integrate o script custom (anche fuori da `src/Ultimus.oxt`) devono rispettare la compatibilità con la licenza LGPL.

## Sicurezza dei moduli in `pythonpath/`

`src/Ultimus.oxt/python/pythonpath/` è nel `sys.path` dell'estensione: qualunque file al suo interno può essere importato dal processo di LibreOffice per motivi indipendenti dal task che lo ha creato (esplorazione macro, `importlib.reload` di recupero in caso di errore, tool di indicizzazione). Per questo:

- **Vietato codice con effetti collaterali a livello di modulo** (eseguito al semplice `import`, fuori da funzioni/classi).
- **Vietato sovrascrivere `sys.modules[...]` a livello di modulo.** Un mock importato anche solo una volta dentro LibreOffice resta per l'intera sessione: effetti silenziosi, da malfunzionamenti a blocchi (freeze) dell'intero processo. Nei test il mocking di `sys.modules` è solo temporaneo e ripristinato (es. `unittest.mock.patch.dict` come context manager), mai un'assegnazione diretta persistente.
- **I file di test (`test_*.py`, `unittest`/`pytest`, mocking di `uno`) non vanno mai in `pythonpath/`.** Vanno in una cartella esclusa dal `sys.path` dell'estensione (es. `tests/`), oppure rimossi prima del merge su `dev` se non servono. Attenzione ai commit di agenti AI (es. Jules): test corretti ma ignari di questo vincolo hanno già rotto l'intera estensione; vanno revisionati prima del merge, non dopo.
- **Vincoli applicati in CI** a ogni push/PR su `dev`: `scripts/check_pythonpath_safety.py` (analisi AST) e `.github/workflows/check-pythonpath-safety.yml`. Un errore CI "Failed to resolve action download info" è quasi sempre transitorio: basta un re-run.

## UNO/ODF: regole ricorrenti

- **Chiusura dei documenti.** Non chiamare `oDoc.close()` in modo sincrono sul documento che ospita lo script in esecuzione (deadlock dell'intero processo): per sostituire il documento corrente apri prima il nuovo e chiudi il vecchio come ultima istruzione, con `return` immediato per non riusare più l'oggetto ormai `disposed`.
- **`LeenoUtils.getDocument()` può restituire `None` subito dopo un dialogo modale** (`oDlg.execute()`): se si ha già un `oDoc` valido prima del dialogo, passarlo esplicitamente alle funzioni chiamate dopo (parametro opzionale `oDoc=`) invece di richiamare `getDocument()`.
- **Proprietà custom del documento.** Se rinomini una `UserDefinedProperty` (es. `Versione` → `Versione_LeenO`), centralizza la lettura in un helper con fallback sul nome legacy, mai letture dirette del vecchio nome senza `try/except`: altrimenti i documenti creati con template precedenti smettono silenziosamente di funzionare.
- **Freeze.** Instrumentare la funzione sospetta con `DLG.chi()` a ogni passaggio chiave (l'ultimo checkpoint visto localizza il blocco); per regressioni lontane usare `git bisect` (good = ultimo tag funzionante, bad = `HEAD` di `dev`); diffidare di commit che sembrano toccare solo codice "non collegato": il problema può essere in un file adiacente dello stesso commit, come un file di test.
- **Quirk.** Gli stili possono avere un nome interno anonimo (`uuuuu` invece di `'Comp TOTALI'`), con fallimenti silenziosi nei confronti `CellStyle == "nome leggibile"`; i template `.ods` possono avere percorsi di progetti reali hardcoded nelle celle `F1` di COMPUTO/CONTABILITA: verificarle prima di distribuire un template. Dettagli, insieme al caso `getDocument()`, in `documentazione/LESSONS_UNO_QUIRKS.md`.

## Modulo XPWE (export/import PriMus)

Tre bug reali già risolti (lookup case-sensitive su `diz_ep`, righe IDEP silenziosamente scartate, segno invertito su "vedi voce") hanno prodotto lezioni specifiche per `LeenoExport.py` e `LeenoImport_XPWE.py`. Prima di modificare questo modulo, leggere `documentazione/LESSONS_XPWE.md`.

Regola da tenere a mente comunque, perché ricorre facilmente: `invertiUnSegno()` è un toggle pensato per uso interattivo, non per i percorsi di import — chiamarlo su una riga già impostata da `vedi_voce_xpwe()` inverte il segno una seconda volta.

## Export PDF (PDF/A e note)

`SheetUtils.pdfExport()` e `pyleeno.ods2pdf()` sono i due percorsi di export PDF. Prima di modificarli, leggere `documentazione/LESSONS_PDF_EXPORT.md`.

Regola da tenere a mente comunque: `PrintAnnotations` sul page style non nasconde l'indicatore visivo delle note sulla cella (solo l'elenco a fine pagina) — per escluderle davvero vanno rimosse (e, su documento live, reinserite dopo l'export).

## Agenda di cantiere (Brogliaccio / PWA)

`tools/appunti-cantiere/` è una PWA statica (HTML/CSS/JS puro, nessuna dipendenza esterna in produzione) per raccogliere l'agenda di cantiere su smartphone e consolidarla in LeenO tramite `LeenoGiornaleImport.py`. Pubblicata via GitHub Pages (`.github/workflows/brogliaccio-pages.yml`) su un sottodominio dedicato. Prima di modificarla, leggere `documentazione/LESSONS_PWA_BROGLIACCIO.md`.

Regola da tenere a mente comunque, perché il mancato rispetto ha già causato una perdita dati reale per l'utente finale: non rinominare né cambiare forma a una chiave `localStorage`/IndexedDB già distribuita senza aggiungere, nello stesso cambiamento, un controllo di recupero automatico dal nome/formato precedente — anche se in quel momento sembra che nessuno la stia ancora usando.

**Aggiornamento versione:** ad ogni modifica del codice di Brogliaccio è obbligatorio aggiornare il numero di versione secondo Semantic Versioning (Major.Minor.Fix/Patch), in modo allineato in due punti:
1. In `tools/appunti-cantiere/core.js` (variabile `VERSIONE`).
2. In `tools/appunti-cantiere/sw.js` (variabile `CACHE` del Service Worker, es. `appunti-cantiere-0.4.2`): è essenziale per invalidare la cache offline dei dispositivi e far scaricare la nuova versione agli utenti.

**Dopo ogni modifica dell'interfaccia:** nelle etichette dei campi di tutte le schermate (anche nella stampa) il testo termina con i due punti (`Cantiere:`, `Meteo:`); i pulsanti no. Poi rigenerare gli screenshot con `scripts/genera_screenshot_brogliaccio.py` e aggiornare il capitolo del manuale (vedi "Manuale utente").

## Manuale utente

`documentazione/MANUALE_LeenO.fodt` si aggiorna seguendo la skill `.agent/skills/leeno-aggiorna-manuale/SKILL.md` e, per le insidie ricorrenti (modifica del FODT, screenshot, PDF, indice), `documentazione/LESSONS_MANUALE.md`. Regola da ricordare comunque: ogni intervento sul manuale aggiorna insieme `TRACKING_MANUALE.md`, la riga "Stato di revisione", `MAPPA_SEZIONI.md`, l'indice generale (`aggiorna_indice.py`, prima del PDF) e il PDF in `src/Ultimus.oxt/`, e si chiude con la validazione XML del FODT.

## Pipeline di test automatizzato (headless UNO)

Test round-trip con istanze reali di LibreOffice, senza mock (XPWE e codice che scrive percorsi relativi al documento o inserisce righe in un blocco esistente, come il modulo di import dell'agenda di cantiere): dettagli in `documentazione/LESSONS_TESTING.md`. Questi file di test non vivono in `pythonpath/` (vedi sopra).

## Sistema icone

Specifica completa in `documentazione/ICONS_DESIGN_SYSTEM.md` (da consultare prima di creare o modificare icone); regole operative in `documentazione/LESSONS_ICONE.md`. Invarianti:

- `icons/svg/` e `icons/scuro/` condividono sempre la stessa geometria (stesso `viewBox` ritagliato) e differiscono solo per colore: ogni cambiamento di crop o bounding box va propagato a entrambe nello stesso passaggio.
- Modificare gli SVG sempre in modalità binaria (vedi "Preservazione del line-ending"): una scrittura in modalità testo normalizza silenziosamente i CRLF e produce diff su ogni riga di ogni file, anche a parità di contenuto grafico.
- **I file `.bmp` in `src/Ultimus.oxt/icons/` sono in realtà SVG/XML testuali**, con estensione `.bmp` solo per compatibilità con il sistema toolbar di LibreOffice: trattarli sempre come testo (`text eol=lf` in `.gitattributes`), mai come binari, o se ne corrompe il contenuto XML.

## Consegna del lavoro per agenti senza credenziali di push

Un agente che lavora sul repository ma non ha credenziali di push dirette (o non deve committare per policy del task) consegna il lavoro così:

1. **Zip, non bundle git**, con la struttura di cartelle del repository (destinazione finale, es. sotto `src/Ultimus.oxt/` o la sottocartella pertinente), non l'intero repository. Lo zip si produce sempre, anche per interventi sul solo manuale e anche se il push riesce, e contiene tutti i file toccati (tracking, mappa sezioni e PDF compresi).
2. **Comandi PowerShell espliciti** per estrarlo in `W:\_dwg\ULTIMUSFREE\_SRC\leeno\...` su PC `giuserpe`, così l'operazione è riproducibile senza ambiguità.
3. **Nessun commit/push automatico** se non richiesto esplicitamente (anche tramite un hook configurato dall'utente): revisione del diff e commit restano un passaggio manuale su PC `giuserpe`. Se il push è richiesto, se ne tenta uno solo: un errore 403 indica permessi mancanti (app GitHub non installata o non collegata) e ripeterlo non serve; si riporta la causa e si consegna lo zip.
4. **Messaggio di commit proposto, non eseguito**, secondo le convenzioni più sotto, se richiesto.
5. **Verifica di integrità prima della consegna**: sintassi Python (`python3 -c "import ast; ast.parse(...)"`), XML del FODT per il manuale (vedi `documentazione/LESSONS_MANUALE.md`) e, per i file con line-ending noto, conteggio CRLF/LF invariato rispetto all'originale.
6. **Ambiente cloud**: clonare solo il branch di lavoro (`git clone --depth 1 --branch dev ...`) e, prima di ogni push, `git fetch origin dev` per non inviare una storia obsoleta.

## Preservazione del line-ending in QUALSIASI editing

Nello stesso `pythonpath/` convivono file CRLF (es. `LeenoGiornale.py`) e file LF puro (es. `LeenoImport_XmlToscana.py`), anche nella stessa cartella. Prima di editare un file:

1. Verificare lo stile reale con un controllo binario (conteggio isolato di `\r\n` vs `\n`), mai assumerlo dal resto del repo o dal tipo di file.
2. In Python, aprire in lettura/scrittura con `newline=''` per non far tradurre gli a-capo, e comporre il testo di sostituzione con lo stesso stile di fine riga del blocco che si sta sostituendo.
3. Dopo la modifica, ricontrollare il conteggio CRLF/LF per confermare che non sia cambiato, prima di consegnare il file.

## Pulizia di codice morto e duplicato

Checklist e pattern ricorrenti (script usa-e-getta pericolosi in `pythonpath/`, redefinition come segnale affidabile di codice morto, import ereditati nei cloni `LeenoImport_Xml*.py`, variabili apparentemente inutili da verificare sui moduli fratelli prima di rimuoverle) in `documentazione/LESSONS_CODE_QUALITY.md`.

Regola operativa da ricordare sempre, senza bisogno di consultare il file: uno script one-shot in `pythonpath/` va eseguito ed eliminato subito dopo l'uso, mai lasciato "per sicurezza".

## Git Commit – Conventional Commits in Italiano (LeenO)

Formato: `<tipo>(<scope>): <descrizione in italiano>`, con corpo opzionale che spiega il PERCHÉ, non il COSA.

- **Tipi**: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `chore` (manutenzione, dipendenze, versioning, build), `revert`.
- **Scope** (area principale colpita): `core` (`pyleeno.py`, `LeenoGlobals.py`), `ui` (dialoghi, `.xhp`, `.xlb`), `contab`, `computo`, `variante`, `giornale`, `import` (`LeenoImport_*.py`), `icons`, `meta` (`description.xml`, `.xcu`), `template`, `docs` (manuale PDF o documentazione tecnica), `brogliaccio` (PWA in `tools/appunti-cantiere/`).

Regole d'oro:

1. **Lingua**: descrizione in **italiano**, imperativo presente (es. "aggiunge", non "aggiunto").
2. **Intestazione**: max 72 caratteri, nessun punto finale.
3. **Breaking change**: `!` dopo il tipo (es. `feat!: ...`) e descrizione in `BREAKING CHANGE:` nel corpo.
4. **Separazione**: se le modifiche riguardano aree troppo diverse, suggerisci commit separati.
5. **Esclusioni**: ignora e ometti sempre, nel messaggio di commit, le modifiche alle funzioni nel cui nome compare la stringa "\_debug" (es. `MENU_debug`).
6. **Sinteticità** (vale anche per i commit dei soli aggiornamenti del manuale, con tipo `docs` e senza elenchi puntati nel corpo; le righe di attribuzione richieste dall'ambiente non contano come corpo): il corpo solo se serve davvero a spiegare il PERCHÉ, in massimo 1 riga breve; nella maggior parte dei casi va omesso del tutto, preferendo la sola intestazione.

Procedura: `git status`, `git diff --cached`, scelta dello scope, intestazione (più l'eventuale riga di corpo), poi proporre il comando `git commit -m "..."` (o `git commit -e` se serve un corpo esteso). Dopo un editing su PC TEST il diff va sottoposto per intero a un assistente AI prima di committare, con più commit per area se è eterogeneo. Tabella dei tipi, esempi e dettagli: `documentazione/LESSONS_GIT.md`.

## Manutenzione di questo file

- Il repository mantiene due copie di questo documento: `AGENTS.md` (root) e `.agents/AGENTS.md`. Devono restare identiche byte per byte in ogni momento.
- Qualunque modifica a una copia va applicata anche all'altra **nello stesso commit**, mai in commit separati: una divergenza è di per sé un difetto da correggere, indipendentemente da quale delle due sia "più aggiornata". Prima di proporre una modifica verificare con `diff AGENTS.md .agents/AGENTS.md`; se non sono allineate, segnalarlo esplicitamente invece di modificarne una sola.
- **Lezioni specifiche per modulo/argomento non vanno inline in questo file**, ma in `documentazione/LESSONS_<ARGOMENTO>.md`, con un rimando breve qui. Questo file va letto per intero a ogni task: deve contenere solo regole invarianti o la singola regola operativa più importante di ogni argomento. Se una sezione supera 10-15 righe per un singolo argomento non invariante, estrarla in un file di dettaglio collegato.
- Non rinominare la sezione "Sicurezza dei moduli in `pythonpath/`": è citata da `scripts/check_pythonpath_safety.py`.
