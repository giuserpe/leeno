# AGENTS.md – LeenO

Questo file descrive le convenzioni obbligatorie per qualsiasi agente (Jules, Claude Code, altri assistenti AI) che lavori sul repository LeenO. Va letto prima di qualsiasi task.

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
- Per la selezione di file o cartelle, utilizza sempre `Dialogs.FileSelect()` invece di dialoghi custom o librerie esterne.
- Nessun output su stdout: usa il logging su file previsto dal progetto.
- Non includere sezioni CLI nel codice dei moduli.

## Git Commit – Conventional Commits in Italiano (LeenO)

### Formato
```
<tipo>(<scope>): <descrizione in italiano>

[corpo opzionale: spiega il PERCHÉ, non il COSA]
```

### Tipi

| Tipo | Quando |
|------|--------|
| `feat` | Nuova funzionalità |
| `fix` | Correzione bug |
| `docs` | Solo documentazione |
| `style` | Formattazione, spazi, punti e virgola mancanti (no logica) |
| `refactor` | Modifica del codice che non corregge bug né aggiunge funzionalità |
| `perf` | Miglioramento prestazioni |
| `test` | Aggiunta/modifica test |
| `chore` | Manutenzione, aggiornamento dipendenze, versioning, build |
| `revert` | Annullamento di un commit precedente |

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
6. **Esclusioni**: Ignora e ometti sempre le modifiche apportate alla funzione `MENU_debug` nella generazione del messaggio di commit

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
