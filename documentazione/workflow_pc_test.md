# LeenO — Workflow Sheet: PC TEST (no admin)

Percorso repo: `W:\_dwg\ULTIMUSFREE\_SRC\leeno` (stesso drive esterno, stessa lettera)
Branch di lavoro: `dev`
Ruolo macchina: **passiva** — solo pull e test in LibreOffice, MAI commit locali intenzionali

## Configurazione git (una tantum — già impostata)

```cmd
git config core.autocrlf true
git config core.fileMode false
git config core.pager cat
```

> Questi due parametri evitano che ogni `pull` segni centinaia di file come "modificati" per rumore di line-ending/permessi (problema riscontrato e risolto in precedenza).

## Ciclo di lavoro standard

1. **Verifica che il working tree sia pulito prima di qualunque operazione**
   ```cmd
   git status
   ```
   Deve dire `nothing to commit, working tree clean`. Se non lo è, NON procedere: prima capire cosa c'è (vedi sezione emergenza).

2. **Aggiorna a `dev`** (dopo che PC Giuseppe ha già mergiato/pushato)
   ```cmd
   leeno_pull.bat
   ```
   (checkout `dev` + pull)

3. **Testa in LibreOffice** — il symlink dell'estensione punta sempre a questa cartella, quindi carica automaticamente ciò che è checkoutato

4. **Se serve testare un branch specifico di Jules** (raro su questa macchina, di solito il test avviene già su PC Giuseppe prima del merge):
   ```cmd
   leeno_test_jules.bat
   ```
   Al termine, tornare su `dev`:
   ```cmd
   leeno_restore_dev.bat
   ```

## Regola d'oro per questa macchina

Non fare mai `git add` / `git commit` qui, nemmeno per errore (es. LeenO che tocca `description.xml` o `leeno_version_code` durante l'uso normale). Se capita per sbaglio:
```cmd
git checkout -- .
```
scarta tutto e basta, senza pensarci — non c'è mai nulla di prezioso da salvare su questa macchina.

## Comandi di emergenza (usare solo se necessario)

| Situazione | Comando |
|---|---|
| Merge bloccato / conflitti dopo un pull fallito | `git merge --abort` |
| `merge --abort` fallisce ("not up to date") | `git reset --hard HEAD` poi ripeti `git merge --abort` |
| Voglio azzerare tutto e allinearmi esattamente a GitHub | `git fetch origin` poi `git reset --hard origin/dev` |
| Verificare se ho commit locali non voluti prima di un reset | `git log origin/dev..dev --oneline` |

## Nota sul drive esterno

Se sposti/ricolleghi il drive tra PC, verifica sempre che la lettera resti `W:`. Se cambia, fissarla da **Gestione Disco** (`diskmgmt.msc`).

## File locali non versionati (in `.gitignore`, presenti anche qui)

```
leeno_test_jules.bat
leeno_restore_dev.bat
leeno_pull.bat
```
