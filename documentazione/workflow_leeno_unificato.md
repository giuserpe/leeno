# LeenO — Workflow Sheet

Percorso repo: `W:\_dwg\ULTIMUSFREE\_SRC\leeno` (drive esterno, stessa lettera `W:` su tutte le macchine — vincolo fisso, il symlink dell'estensione in LibreOffice punta qui)
Branch di lavoro: `dev` (default branch su GitHub)
Cartella pacchetti compilati: `OXT\` (contiene i file `.oxt` prodotti da `make_pack()`)

## Macchine coinvolte

- **PC `giuserpe`** (nome letterale della macchina): amministratore locale. Qui avvengono gli editing diretti con relativo commit/push (Ciclo B), oltre alla gestione ordinaria del repo.
- **PC TEST**: nessun privilegio di amministratore. Usato per test dell'estensione in LibreOffice, lancio task Jules, merge PR da GitHub (operazione web, non richiede privilegi locali) e, occasionalmente, editing diretto del codice — ma senza mai fare `git commit`/`git push` locali da qui (vedi Ciclo B).

Non si usano `bin2src.py`/`src2bin.py`: `src/Ultimus.oxt/` nel repo è già il sorgente diretto. `make_pack()` (lanciato da LibreOffice Calc) impacchetta il sorgente in un `.oxt` installabile, con bump automatico di `description.xml` e `leeno_version_code`.

## Configurazione git (una tantum, IDENTICA su entrambe le macchine)

```cmd
git config user.name "Il tuo nome"
git config user.email "tua-email@esempio.com"
git config core.pager cat
git config core.autocrlf true
git config core.fileMode false
```

> `autocrlf`/`fileMode` vanno impostati **subito, su entrambe le macchine**, non solo se emergono problemi — sono la causa più comune delle lunghe liste di file "modified" fantasma dopo un pull. Verifica con `git config --get core.autocrlf` prima di dare per scontato che siano già a posto: è già capitato che una macchina li avesse e l'altra no.

## Regola d'oro: sincronizzati prima di lavorare

Prima di lanciare un nuovo task Jules o iniziare qualunque attività, su qualunque macchina:

```cmd
cd /d W:\_dwg\ULTIMUSFREE\_SRC\leeno
git status
```

Deve risultare pulito e allineato a `origin/dev`. Se non lo è:

```cmd
leeno_pull.bat
```

## Ciclo A — Task Jules (lanciabile da qualunque macchina)

Il merge avviene sempre da GitHub (operazione web), quindi questo ciclo non richiede privilegi admin né una macchina specifica — può partire e concludersi sia da PC `giuserpe` sia da PC TEST.

1. **Lancia un task su Jules** (jules.google.com)
   - Base branch: `dev`
   - Prompt preciso, con riferimento ad `AGENTS.md`

2. **Jules apre una PR** verso `dev` — annota il nome branch (`jules-XXXXXXXXXX-XXXXXXXX`) dalla pagina della PR

3. **Testa il branch prima del merge**
   ```cmd
   leeno_test_jules.bat
   ```
   (fetch → checkout branch → richiede il nome branch)
   LibreOffice può restare aperto durante pull/checkout ("cofano aperto"): va chiuso e riaperto solo se le modifiche toccano UI/dialoghi/toolbar/icone. Per modifiche a codice Python puro, LibreOffice lo vede senza riavvio.

4. **Decidi l'esito**

   OK → merge su GitHub (*Merge pull request*, valutare *Squash and merge*) →
   ```cmd
   leeno_restore_dev.bat
   ```

   NON OK → chiudi la PR su GitHub (*Close pull request*, opzionale *Delete branch*) →
   ```cmd
   leeno_restore_dev.bat
   ```

5. **Opzionale — genera il pacchetto distribuibile**: una volta che `dev` è aggiornato col merge, se serve un OXT pronto per distribuzione/installazione:
   ```
   make_pack() da LibreOffice Calc → produce l'OXT, conservato in OXT\
   ```
   Questo passaggio è indipendente dal merge stesso — il codice è già entrato in `dev` tramite la PR, quindi non serve alcuna estrazione manuale del pacchetto (a differenza del Ciclo B).

## Ciclo B — Editing diretto su PC TEST

**Su PC TEST:**
1. `leeno_pull.bat` (parti allineato a `dev`)
2. Modifica il codice in locale
3. `make_pack()` da LibreOffice Calc → produce l'OXT aggiornato, conservato in `OXT\`

**Su PC `giuserpe`** (il commit avviene SEMPRE qui, mai da PC TEST):
1. `leeno_pull.bat` (allineati prima di procedere)
2. Estrai il contenuto dell'OXT dentro `src/Ultimus.oxt/`, sovrascrivendo:
   ```cmd
   mkdir _tmp_extract
   tar -xf "OXT\LeenO-x.y.z.oxt" -C _tmp_extract
   robocopy _tmp_extract src\Ultimus.oxt /MIR /XD .git /L
   ```
   (il `/L` è solo simulazione — controlla l'elenco, poi rilancia senza `/L` per l'esecuzione reale)
3. Rimuovi eventuali artefatti indesiderati che non dovrebbero essere tracciati (es. `staged_diff.txt`, `.vscode/settings.json`, `leeno_version_code_test`, `log`, se presenti)
4. Rivedi il diff:
   ```cmd
   git status
   git add -A
   git diff --cached > diff_da_revisionare.txt
   ```
5. **Fai leggere il diff a un assistente AI** (Claude, Copilot o altro) perché componga il messaggio di commit seguendo le convenzioni in `AGENTS.md`. Se il diff copre aree molto eterogenee, preferire più commit separati
6. Commit e push:
   ```cmd
   git commit -m "..."
   git push
   ```
7. Pulisci `_tmp_extract`

## Se qualcosa non ti convince dopo un merge/commit già fatto

```cmd
git add <file-o-cartella-da-ripristinare>
git commit -m "revert(<scope>): ripristina versione precedente"
git push
```

## Comandi di emergenza

| Situazione | Comando |
|---|---|
| Merge bloccato, voglio annullare | `git merge --abort` |
| `merge --abort` fallisce ("not up to date") | `git reset --hard HEAD` poi ripeti `git merge --abort` |
| Working tree sporco, voglio scartare tutto | `git checkout -- .` seguito da `git clean -fd` |
| Voglio solo mettere da parte senza perdere nulla | `git add -A` poi `git stash push -m "..."` |
| Recupero da stash | `git stash pop` |
| Voglio azzerare tutto e allinearmi esattamente a GitHub | `git fetch origin` poi `git reset --hard origin/dev` |
| Verificare commit locali non voluti prima di un reset | `git log origin/dev..dev --oneline` |

## File locali non versionati (in `.gitignore`, presenti su entrambe le macchine)

```
leeno_test_jules.bat
leeno_restore_dev.bat
leeno_pull.bat
```

## Riferimento rapido nomi branch Jules

Sempre nel formato `jules-<numero>-<hash>` — copialo per intero dalla pagina della PR su GitHub, non abbreviare.

## Nota sul drive esterno

Se sposti/ricolleghi il drive tra PC, verifica sempre che la lettera resti `W:`. Se cambia, fissarla da **Gestione Disco** (`diskmgmt.msc`).

## Nota su Google Drive for Desktop

Su PC `giuserpe` è attivo Google Drive for Desktop, che sincronizza automaticamente una copia del workspace su Drive. Questo NON è un componente vitale del workflow: serve solo a dare a Claude (o altri assistenti) un modo di consultare lo stato del codice quando serve, senza interrompere il ciclo di lavoro se non è attivo o aggiornato. Il repo su `W:\` via git resta l'unica fonte di verità. Non è attivabile su PC TEST (richiede diritti di amministratore per l'installazione) e non serve che lo sia.
