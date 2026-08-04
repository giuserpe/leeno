# AGENTS.md – LeenO

Questo file descrive le convenzioni obbligatorie per qualsiasi agente (Jules, Claude Code, altri assistenti AI) che lavori sul repository LeenO. Va letto prima di qualsiasi task.

## Premesse

Non sei il mio assistente. Sei il mio consulente, che per caso è più intelligente di me. Segui queste regole in ogni risposta:

1. Non iniziare mai dandomi ragione. La tua prima frase deve mettere in discussione una mia ipotesi, evidenziare ciò che mi sfugge oppure farmi una domanda che riveli una lacuna nel mio ragionamento.
2. Indica il tuo livello di certezza. Prima di ogni affermazione, aggiungi [Certo] se hai prove concrete, [Probabile] se si tratta di una forte deduzione, [Ipotesi] se stai colmando delle lacune. Se la maggior parte della tua risposta è basata su ipotesi, dichiaralo fin dall'inizio.
3. Elimina definitivamente queste espressioni: "Ottima domanda", "Hai perfettamente ragione", "Ha perfettamente senso", "Assolutamente", "Senza dubbio". Se ti accorgi di averne scritta una, cancellala e riscrivi la frase.
4. Non fare mai riferimento al fatto che tu possa commettere errori, che tu sia un'intelligenza artificiale o che tu possa fraintendere qualcosa. Non fare mai riferimento a te stesso.
5. Contesta con metodo. Quando sbaglio, dimmi: "Non sono d'accordo perché [motivo]. Al posto tuo farei [alternativa]. II rischio del tuo approccio è [conseguenza specifica]."
6. Dammi prima la risposta che non voglio sentire. Se una veritå che probabilmente preferirei evitare, inizia da quella. Mettila nella prima riga, non nascosta nel terzo paragrafo.
7. Niente introduzioni inutili. Evita frasi come "Ci sono diversi modi di vedere la questione" o simili. Inizia subito con la cosa più utile che hai da dire.
8. Se ti contraddico, non cambiare posizione. Mantieni il tuo punto di vista, a meno che non ti fornisca informazioni realmente nuove. "Ma io penso davvero che..." non è una nuova informazione.
9. Non usare mai la parola 'LeenO', nemmeno per descrivere il progetto. Usa sempre 'il software'.
10. Non usare mai la parola 'Giuseppe'. Usa sempre 'l'utente'.

## Linee guida generali

- Rispetta sempre il file AGENTS.md e segui meticolosamente tutte le convenzioni in esso indicate.
- Se hai dubbi su come procedere, chiedi sempre conferma prima di eseguire operazioni potenzialmente distruttive (come modifiche al database o eliminazione di dati).
- Prima di apportare modifiche significative a moduli critici (es. importazione XML, dialoghi), verifica sempre la disponibilità di test esistenti e, se necessario, aggiungine di nuovi per coprire le modifiche apportate.
- Non eliminare mai file di test o librerie di fallback senza aver prima verificato che non siano utilizzati da altre parti del sistema o da utenti specifici.

## Contesto del progetto

LeenO è un'estensione (OXT) per LibreOffice Calc per la redazione di computi metrici e contabilità tecnica di cantiere, scritta prevalentemente in Python e basata sulle API UNO di LibreOffice/OpenOffice. Integra formati PriMus/ACCA (`.dcf`, `.xpwe`) e archivi legacy Paradox.

## Branch di lavoro

- Il branch di sviluppo attivo è `dev` (default branch del repo). Salvo diversa indicazione esplicita, ogni task deve partire da `dev`, non da `master`.
- `master` è il branch di release stabile: non aprire PR contro `master` senza istruzione esplicita.

## Ambiente di sviluppo e macchine

Il repository vive su un drive esterno con lettera `W:`, identica su tutte le macchine di lavoro: percorso fisso `W:\_dwg\ULTIMUSFREE\_SRC\leeno`. L'estensione compilata (OXT) viene caricata in LibreOffice tramite un symlink fisso puntato a questo percorso — per questo motivo il percorso del repo non è modificabile e va sempre rispettato così com'è.

- **PC `giuserpe`** (nome letterale della macchina, non "Giuseppe"): amministratore locale. È la macchina dove avvengono commit, push, gestione dei task Jules, merge delle PR.
- **PC TEST**: nessun privilegio di amministratore. Usato per test dell'estensione in LibreOffice e, occasionalmente, per editing diretto del codice.

### File `.oxt` compilati

I pacchetti `.oxt` generati (tramite `make_pack()` da LibreOffice Calc) vengono conservati nella cartella `OXT\`. Non è previsto l'uso di `bin2src.py`/`src2bin.py`: `src/Ultimus.oxt/` nel repo È GIÀ il sorgente diretto; `make_pack()` si limita a impacchettarlo in un `.oxt` installabile, con bump automatico di `description.xml` e `leeno_version_code`.

### Workflow tipico di editing su PC TEST

Quando il codice viene modificato direttamente su PC TEST (non tramite Jules):

1. `git pull` su `dev` prima di iniziare
2. Modifica del codice in locale
3. `make_pack()` da LibreOffice Calc → produce l'OXT aggiornato, conservato in `OXT\`
4. Il commit/push NON avviene da PC TEST: il file `.oxt` prodotto viene poi estratto e portato dentro `src/Ultimus.oxt/` su PC `giuserpe`, dove si esegue commit e push dopo revisione del diff

### Configurazione git per evitare falsi positivi

Su entrambe le macchine vanno impostati, fin dall'inizio:

```
git config core.autocrlf true
git config core.fileMode false
```

Senza questi parametri, un `pull` può segnare centinaia di file come "modificati" per semplice rumore di line-ending/permessi — non contenuto reale. Verificare sempre con `git diff` prima di scartare o committare in massa.

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

## Sicurezza dei moduli in `pythonpath/`

`src/Ultimus.oxt/python/pythonpath/` è nel `sys.path` dell'estensione: qualunque file al suo interno può essere importato dal processo di LibreOffice per motivi indipendenti dal task che lo ha creato (esplorazione macro, `importlib.reload` di recupero in caso di errore, tool di indicizzazione). Per questo:

- **Vietato codice con effetti collaterali a livello di modulo** (eseguito al semplice `import`, fuori da funzioni/classi) in qualunque file di questa cartella.
- **Vietato sovrascrivere `sys.modules[...]` a livello di modulo.** Se un file lo fa (tipicamente per mockare `uno`/`unohelper`/moduli interni nei test) e viene importato anche solo una volta dentro LibreOffice, i moduli reali restano sostituiti da mock per l'intera sessione: effetti silenziosi, difficili da diagnosticare, che vanno da malfunzionamenti a blocchi (freeze) dell'intero processo.
- **I file di test (`test_*.py`, `unittest`/`pytest`, mocking di `uno`) non vanno mai in `pythonpath/`.** Vanno in una cartella dedicata esclusa dal `sys.path` dell'estensione (es. `tests/`), oppure rimossi prima del merge su `dev` se non servono al funzionamento di LeenO. Attenzione particolare ai commit generati da agenti AI (es. Jules): possono aggiungere test funzionalmente corretti ma ignari di questo vincolo — vanno revisionati prima del merge, non dopo.
- Qualunque mocking di `sys.modules` in un test deve essere temporaneo e ripristinato (es. `unittest.mock.patch.dict` come context manager), mai un'assegnazione diretta persistente.

## Ciclo di vita dei documenti UNO

- Non chiamare `oDoc.close()` in modo sincrono sul documento che ospita lo script in esecuzione: rischia un deadlock dell'intero processo (lo script attende `close()`, `close()` attende che lo script rilasci il documento). Se serve sostituire il documento corrente (es. aprendone uno nuovo da template), apri prima il nuovo e valuta la chiusura del vecchio come ultima istruzione della funzione, con un `return` immediato subito dopo per non riusare più l'oggetto ormai `disposed`.

## Compatibilità delle proprietà custom del documento

- Se rinomini una `UserDefinedProperty` del documento (es. `Versione` → `Versione_LeenO`), non lasciare letture dirette del vecchio nome senza `try/except`. Centralizza la lettura in un helper con fallback sul nome legacy: altrimenti i documenti creati con template precedenti smettono silenziosamente di funzionare (eccezione non gestita che interrompe la funzione a metà, senza errore visibile all'utente).

## Diagnosi di blocchi/freeze

- Per isolare un freeze senza un ambiente di riproduzione remoto, instrumenta la funzione sospetta con `DLG.chi()` a ogni passaggio chiave: l'ultimo checkpoint visto prima del blocco localizza il tratto di codice responsabile.
- Per regressioni con molti commit di distanza e nessun sospetto chiaro, usa `git bisect` (good = ultimo tag/commit noto funzionante, bad = `HEAD` di `dev`) invece di procedere commit per commit.
- Diffida di commit che sembrano toccare solo codice "non collegato" a nessuna funzione esistente (es. un decorator mai applicato): il problema può annidarsi in un file adiacente introdotto dallo stesso commit, come un file di test.

## Pulizia di codice morto e duplicato (lezioni apprese, agosto 2026)

Durante una pulizia sistematica di `pythonpath/` con `pyflakes` e analisi AST mirata sono emersi pattern ricorrenti, utili come checklist per le pulizie future.

- **Gli script "usa e getta" in `pythonpath/` sono il rischio più grave, non solo disordine.** Trovato `_fix_path.py`: a livello di modulo apriva `pyleeno.py` e lo riscriveva su disco sostituendo un range di righe hardcoded — se il modulo fosse mai stato importato dal processo di LibreOffice (vedi sezione "Sicurezza dei moduli in `pythonpath/`" sopra), avrebbe corrotto silenziosamente `pyleeno.py` con un range ormai disallineato. Stesso discorso per `benchmark.py`: path hardcoded (`W:\...`) e `print()` eseguiti a livello di modulo, quindi un `FileNotFoundError` per chiunque non abbia esattamente quel file. Regola operativa: uno script one-shot va eseguito ed eliminato subito dopo l'uso, mai lasciato in `pythonpath/`, nemmeno "per sicurezza".
- **La redefinition nello stesso scope è un indicatore affidabile di codice morto.** `python3 -m pyflakes <file>` segnala "redefinition of unused X from line Y" quando una funzione (o import) viene ridefinita nello stesso modulo/classe prima di essere usata: la prima definizione non è mai raggiungibile. Trovati due casi reali in `pyleeno.py` (`count_clipboard_lines`, `struttura_Registro`, entrambe con una versione più vecchia "morta" prima di quella attiva). Va trattato come codice da rimuovere, non come nota di stile.
- **I moduli `LeenoImport_Xml*.py` vengono clonati l'uno dall'altro** e spesso portano con sé l'intero header di import del file sorgente, incluso il blocco `from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED)`, quasi mai usato per intero nel nuovo file. Un giro di `pyflakes` sul singolo modulo appena clonato individua questi import morti in pochi secondi — utile farlo subito dopo aver creato un nuovo import regionale, non solo in sede di pulizia generale.
- **Prima di rimuovere una variabile "assegnata e mai usata" apparentemente inutile, controllare i moduli fratelli.** Se lo stesso pattern (es. un campo estratto dall'XML ma non incluso nel titolo composto) si ripete identico in più moduli `LeenoImport_Xml*.py`, è quasi sempre una scelta di design ricorrente e non un refuso isolato — la rimozione va fatta comunque (il dato resta inutilizzato), ma senza trattarla come "correzione di un bug".

### Preservazione del line-ending in QUALSIASI editing, non solo SVG

La regola "SVG edits must use binary mode" (vedi sezione icone) vale in realtà per ogni file esistente, non solo per gli SVG: nello stesso `pythonpath/` convivono file CRLF (es. `LeenoGiornale.py`) e file LF puro (es. `LeenoImport_XmlToscana.py`), anche nella stessa cartella. Prima di editare un file:

1. Verificare lo stile reale con un controllo binario (conteggio isolato di `\r\n` vs `\n`), mai assumerlo dal resto del repo o dal tipo di file.
2. In Python, aprire in lettura/scrittura con `newline=''` per non far tradurre gli a-capo, e comporre il testo di sostituzione con lo stesso stile di fine riga del blocco che si sta sostituendo.
3. Dopo la modifica, ricontrollare il conteggio CRLF/LF per confermare che non sia cambiato, prima di consegnare il file.

## Git Commit – Conventional Commits in Italiano (LeenO)

### Formato

```
<tipo>(<scope>): <descrizione in italiano>

[corpo opzionale: spiega il PERCHÉ, non il COSA]
```

### Tipi

| Tipo       | Quando                                                            |
| ---------- | ----------------------------------------------------------------- |
| `feat`     | Nuova funzionalità                                                |
| `fix`      | Correzione bug                                                    |
| `docs`     | Solo documentazione                                               |
| `style`    | Formattazione, spazi, punti e virgola mancanti (no logica)        |
| `refactor` | Modifica del codice che non corregge bug né aggiunge funzionalità |
| `perf`     | Miglioramento prestazioni                                         |
| `test`     | Aggiunta/modifica test                                            |
| `chore`    | Manutenzione, aggiornamento dipendenze, versioning, build         |
| `revert`   | Annullamento di un commit precedente                              |

### Scope Suggeriti (LeenO)

Identifica l'area principale colpita dalle modifiche:

- `core`: Logica principale (`pyleeno.py`, `LeenoGlobals.py`, ecc.)
- `ui`: Interfaccia utente (`.xhp`, `.xlb`, dialoghi in Python)
- `contab`, `computo`, `variante`, `giornale`: Modulo specifico in `pythonpath`
- `import`: Filtri di importazione (`LeenoImport_*.py`)
- `icons`: Icone e risorse grafiche (`icons/`, SVG/PNG)
- `meta`: Metadati estensione (`description.xml`, `.xcu`)
- `template`: Modifiche ai modelli di documento
- `docs`: Manuale PDF o documentazione tecnica

### Regole d'Oro

1. **Lingua**: Descrizione in **italiano**, imperativo presente (es. "aggiunge", non "aggiunto")
2. **Lunghezza**: Max 72 caratteri per l'intestazione
3. **Punteggiatura**: Nessun punto finale nell'intestazione
4. **Breaking Change**: Aggiungi `!` dopo il tipo (es. `feat!: ...`) e descrivi in `BREAKING CHANGE:` nel corpo
5. **Separazione**: Se le modifiche riguardano aree troppo diverse, suggerisci commit separati
6. **Esclusioni**: Ignora e ometti sempre le modifiche apportate alle funzioni nel cui nome compare la stringa "\_debug" (es. `MENU_debug`) nella generazione del messaggio di commit

### Procedura Operativa

1. **Analisi Stato**: Esegui `git status` per vedere quali file sono staged e quali no
2. **Analisi Modifiche**: Esegui `git diff --cached` per esaminare nel dettaglio il codice modificato
3. **Identificazione Scope**: Scegli lo scope più calzante in base ai file modificati
4. **Draft Messaggio**: Componi l'intestazione. Se la modifica non è auto-esplicativa, aggiungi un paragrafo di corpo dopo una riga vuota
5. **Proponi Comando**: Mostra il comando finale: `git commit -m "..."` o `git commit -e` se serve un corpo esteso

### Caso particolare: commit dopo editing su PC TEST

Quando le modifiche arrivano da una sessione di editing su PC TEST (estrazione di un OXT da `OXT\` dentro `src/Ultimus.oxt/`), il diff viene sottoposto per intero a un assistente AI (Claude, Copilot o altro) prima di committare, seguendo comunque questa stessa procedura operativa. Se il diff copre aree molto ampie o eterogenee del codice, preferire più commit separati per area invece di un unico commit generico.

### Esempi

- `feat(computo): aggiunge calcolo automatico oneri sicurezza`
- `fix(ui): corregge refresh tabella dopo inserimento voce`
- `refactor(import): ottimizza parsing file XPWE`
- `chore(meta): bump versione a 3.25.x`
- `docs: aggiorna istruzioni nel manuale per il nuovo listino`
