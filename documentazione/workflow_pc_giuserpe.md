# LeenO — Workflow Sheet: PC GIUSERPE (admin)

Percorso repo: `W:\_dwg\ULTIMUSFREE\_SRC\leeno` (drive esterno, stessa lettera su tutti i PC)
Branch di lavoro: `dev` (default branch su GitHub)
Ruolo macchina: **attiva** — qui si gestiscono task Jules, review, merge, push

## Configurazione git (una tantum)

```cmd
git config user.name "Il tuo nome"
git config user.email "tua-email@esempio.com"
git config core.pager cat
```

> Nota: `core.autocrlf`/`core.fileMode` NON servono qui se non editi mai file manualmente — servono su PC TEST dove i pull generano rumore.

## Ciclo di lavoro standard

1. **Lancia un task su Jules** (jules.google.com)
   - Base branch: `dev` (verificare sia selezionato, non `master`)
   - Prompt preciso, con riferimento ad `AGENTS.md` per le convenzioni
   - Task continua in cloud anche a browser chiuso

2. **Jules apre una PR** verso `dev`
   - Vai su `github.com/giuserpe/leeno/pulls`
   - Annota il nome branch (es. `jules-XXXXXXXXXX-XXXXXXXX`)

3. **Testa il branch prima del merge**
   ```cmd
   leeno_test_jules.bat
   ```
   (chiude LibreOffice → fetch → checkout branch → richiede il nome branch)
   Riapri LibreOffice, testa la funzionalità.

4. **Decidi l'esito**

   Se OK → merge su GitHub (pulsante *Merge pull request*, considerare *Squash and merge* per history pulita) → poi:
   ```cmd
   leeno_restore_dev.bat
   ```
   (checkout `dev` + pull, porta dentro il merge)

   Se NON OK → chiudi la PR su GitHub (*Close pull request*, opzionale *Delete branch*) →
   ```cmd
   leeno_restore_dev.bat
   ```
   (torna su `dev` pulito, pull non necessario ma innocuo)

5. **Ripulisci branch locali obsoleti** (periodico, non ad ogni ciclo)
   ```cmd
   git branch -a
   git branch -D nome-branch-jules-vecchio
   ```

## Se qualcosa non ti convince dopo un merge già fatto

Ripristino rapido (hai già il vecchio materiale a disposizione):
```cmd
git add <file-o-cartella-da-ripristinare>
git commit -m "revert(<scope>): ripristina versione precedente"
git push
```

## Comandi di emergenza (usare solo se necessario)

| Situazione | Comando |
|---|---|
| Merge bloccato, voglio annullare | `git merge --abort` |
| Working tree sporco, voglio scartare tutto | `git checkout -- .` seguito da `git clean -fd` |
| Voglio solo mettere da parte senza perdere nulla | `git add -A` poi `git stash push -m "..."` |
| Recupero da stash | `git stash pop` |

## File locali non versionati (in `.gitignore`)

```
leeno_test_jules.bat
leeno_restore_dev.bat
leeno_pull.bat
```

## Riferimento rapido nomi branch Jules

Sempre nel formato `jules-<numero>-<hash>` — copialo per intero dalla pagina della PR su GitHub, non abbreviare.
