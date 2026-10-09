# Lezioni apprese – Git (conflitti, storia, commit)

Dettagli estratti da `AGENTS.md`, che ne conserva le regole invarianti. Il testo qui sotto è quello originale, non riassunto.

## Rumore vs contenuto reale prima di stash/scarto

Quando `git pull`/`git merge` fallisce con "local changes would be overwritten", non stashare né scartare alla cieca, anche con `core.autocrlf`/`core.fileMode` già corretti:

```
git diff --stat
git diff --ignore-space-at-eol --ignore-all-space -- <path>
```

Se il diff filtrato è vuoto → solo normalizzazione riga, sicuro `git checkout -- .`. Se resta invariato rispetto al diff non filtrato → è contenuto reale (es. geometria SVG modificata), non va scartato: stash o commit a seconda che il lavoro sia finito o meno.

Passare più path a `git diff -- ` uno per riga o quotati singolarmente: una riga tagliata o concatenata dal terminale (es. autocomplete PSReadLine) produce un `fatal: bad revision` fuorviante, non un errore di git.

## Reset, rebase, revert, force-push: quale usare

La scelta dipende da un solo fattore: il commit da modificare è già su `origin/<branch>` o solo locale? Verificare sempre prima di agire:

```
git log origin/dev..HEAD --oneline   # commit locali non ancora pushati
git branch --contains <hash>          # su quali branch/remoti compare <hash>
git show --stat <hash>                # conferma che sia il commit giusto
```

| Situazione | Comando | Nota |
| --- | --- | --- |
| Ultimo commit, mai pushato, da eliminare del tutto | `git reset --hard HEAD~1` | Nessuna rete di sicurezza per modifiche non committate: stash prima se serve salvare qualcosa |
| Commit sepolto sotto altri, mai pushato | `git rebase -i <hash>^` → `drop` sulla riga | Può fermarsi in conflitto se commit successivi toccano le stesse righe |
| Commit già su `origin/<branch>`, effetto da annullare | `git revert <hash>` | Non riscrive la storia condivisa: sicuro anche se altre macchine hanno già fatto pull |
| Commit già su `origin/<branch>`, da far sparire dalla storia (motivo esplicito, es. dati sensibili) | rebase interattivo locale poi `git push --force-with-lease origin <branch>` | Mai `--force` secco. Richiede coordinamento esplicito con chi altro ha già pullato: al prossimo pull troverà storia divergente e dovrà riallinearsi con un reset manuale |

Un `reset --hard` a un commit molto indietro rispetto a `origin/<branch>` può far ripresentare conflitti già risolti a monte: prima di lanciarlo, controllare `git log <hash>..origin/<branch> --oneline` per capire quanta storia si sta per riattraversare.

Nota su GitHub: un force-push che rimuove un merge di PR dalla storia di un branch non ritira lo stato "Merged" della PR nell'interfaccia web — restano due cose distinte.

## Parità `icons/svg/` e `icons/scuro/` prima di ogni commit sulle icone

Prima di committare modifiche alle icone, verificare che entrambe le cartelle abbiano lo stesso insieme di file toccati:

```
git diff --stat -- src/Ultimus.oxt/icons/svg/
git diff --stat -- src/Ultimus.oxt/icons/scuro/
```

Se una delle due risulta modificata e l'altra no, il lavoro è a metà (vedi "Sistema icone" più sotto): non committare finché il crop non è propagato a entrambe.

## Blocco dei file `.ods` durante operazioni git

- Se un file `.ods` del repository (template, listino, foglio di test) è aperto in LibreOffice Calc mentre si esegue `git pull`, `git checkout` o `git stash pop`, l'operazione può fallire o lasciare un file parzialmente scritto: LibreOffice mantiene un lock (file nascosto `.~lock.<nome>.ods#`) e un handle sul contenuto in memoria che non coincide più con quello su disco dopo l'operazione git.
- Prima di eseguire operazioni git che toccano `.ods` tracciati, chiudere il documento in LibreOffice (non solo minimizzarlo) oppure verificare l'assenza del file di lock (`.~lock.*.ods#`) nella cartella del repository.
- Se un'operazione git segnala errori di permesso o file "in uso" su un `.ods`, non forzare con `git checkout --force` a documento ancora aperto: chiudere prima il documento, poi ripetere l'operazione.
- Dopo un `pull` che ha aggiornato un `.ods` già aperto, riaprire il documento (chiudi e riapri) prima di modificarlo: LibreOffice non rileva automaticamente la sostituzione del file su disco e un salvataggio successivo rischia di sovrascrivere la versione aggiornata con quella in memoria, obsoleta.

## Conventional Commits in italiano: dettagli

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
- `brogliaccio`: PWA in `tools/appunti-cantiere/`

### Procedura Operativa

1. **Analisi Stato**: Esegui `git status` per vedere quali file sono staged e quali no
2. **Analisi Modifiche**: Esegui `git diff --cached` per esaminare nel dettaglio il codice modificato
3. **Identificazione Scope**: Scegli lo scope più calzante in base ai file modificati
4. **Draft Messaggio**: Componi l'intestazione. Aggiungi un corpo di 1 riga breve solo se il PERCHÉ non è già ovvio dall'intestazione stessa; altrimenti ometti il corpo
5. **Proponi Comando**: Mostra il comando finale: `git commit -m "..."` o `git commit -e` se serve un corpo esteso

### Caso particolare: commit dopo editing su PC TEST

Quando le modifiche arrivano da una sessione di editing su PC TEST (estrazione di un OXT da `OXT\` dentro `src/Ultimus.oxt/`), il diff viene sottoposto per intero a un assistente AI (Claude, Copilot o altro) prima di committare, seguendo comunque questa stessa procedura operativa. Se il diff copre aree molto ampie o eterogenee del codice, preferire più commit separati per area invece di un unico commit generico.

### Esempi

- `feat(computo): aggiunge calcolo automatico oneri sicurezza`
- `fix(ui): corregge refresh tabella dopo inserimento voce`
- `refactor(import): ottimizza parsing file XPWE`
- `chore(meta): bump versione a 3.25.x`
- `docs: aggiorna istruzioni nel manuale per il nuovo listino`
