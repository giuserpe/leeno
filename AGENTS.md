# AGENTS.md – LeenO

Questo file descrive le convenzioni obbligatorie per qualsiasi agente (Jules, Claude Code, altri assistenti AI) che lavori sul repository LeenO. Va letto prima di qualsiasi task.

## Contesto del progetto

LeenO è un'estensione (OXT) per LibreOffice Calc per la redazione di computi metrici e contabilità tecnica di cantiere, scritta prevalentemente in Python e basata sulle API UNO di LibreOffice/OpenOffice. Integra formati PriMus/ACCA (`.dcf`, `.xpwe`) e archivi legacy Paradox.

## Branch di lavoro

- Il branch di sviluppo attivo è `dev`. Salvo diversa indicazione esplicita, ogni task deve partire da `dev`, non da `master`.
- `master` è il branch di release stabile: non aprire PR contro `master` senza istruzione esplicita.

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
6. **Esclusioni**: Ignora e ometti sempre le modifiche apportate alle funzioni nel cui nome compare la stringa "\_debug" nella generazione del messaggio di commit

### Procedura Operativa

1. **Analisi Stato**: Esegui `git status` per vedere quali file sono staged e quali no
2. **Analisi Modifiche**: Esegui `git diff --cached` per esaminare nel dettaglio il codice modificato
3. **Identificazione Scope**: Scegli lo scope più calzante in base ai file modificati
4. **Draft Messaggio**: Componi l'intestazione. Se la modifica non è auto-esplicativa, aggiungi un paragrafo di corpo dopo una riga vuota
5. **Proponi Comando**: Mostra il comando finale: `git commit -m "..."` o `git commit -e` se serve un corpo esteso

### Esempi

- `feat(computo): aggiunge calcolo automatico oneri sicurezza`
- `fix(ui): corregge refresh tabella dopo inserimento voce`
- `refactor(import): ottimizza parsing file XPWE`
- `chore(meta): bump versione a 3.25.x`
- `docs: aggiorna istruzioni nel manuale per il nuovo listino`
