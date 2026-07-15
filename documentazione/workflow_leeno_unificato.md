# LeenO — Workflow Sheet (PC Giuserpe \+ PC TEST, entrambi attivi)

Percorso repo: `W:\_dwg\ULTIMUSFREE\_SRC\leeno` (drive esterno, stessa lettera `W:` su tutte le macchine) Branch di lavoro: `dev` (default branch su GitHub) Ruolo macchine: **entrambe attive** — task Jules, test, merge e push possono partire da qualunque PC

## Configurazione git (una tantum, IDENTICA su entrambe le macchine)

git config user.name "Il tuo nome"

git config user.email "tua-email@esempio.com"

git config core.pager cat

git config core.autocrlf true

git config core.fileMode false

`autocrlf`/`fileMode` vanno impostati **subito, su entrambe le macchine**, non solo se emergono problemi — sono la causa più comune delle lunghe liste di file "modified" fantasma dopo un pull. Verifica con `git config --get core.autocrlf` prima di dare per scontato che siano già a posto: è già capitato che una macchina li avesse e l'altra no.

## Regola d'oro: sincronizzati prima di lavorare

Prima di lanciare un nuovo task Jules o iniziare qualunque attività, su qualunque macchina:

cd /d W:\\\_dwg\\ULTIMUSFREE\\\_SRC\\leeno

git status

Deve risultare pulito e allineato a `origin/dev`. Se non lo è:

leeno\_pull.bat

## Ciclo di lavoro standard (uguale su entrambe le macchine)

1. **Lancia un task su Jules** (jules.google.com)  
     
   - Base branch: `dev`  
   - Prompt preciso, con riferimento ad `AGENTS.md`

   

2. **Jules apre una PR** verso `dev` — annota il nome branch (`jules-XXXXXXXXXX-XXXXXXXX`) dalla pagina della PR  
     
3. **Testa il branch prima del merge**  
     
   leeno\_test\_jules.bat  
     
   (chiude LibreOffice → fetch → checkout branch → richiede il nome branch) Riapri LibreOffice, testa.  
     
4. **Decidi l'esito**  
     
   OK → merge su GitHub (*Merge pull request*, valutare *Squash and merge*) →  
     
   leeno\_restore\_dev.bat  
     
   NON OK → chiudi la PR su GitHub (*Close pull request*, opzionale *Delete branch*) →  
     
   leeno\_restore\_dev.bat  
     
5. **Prima di passare all'altra macchina**, assicurati di aver pushato tutto e che `git status` sia pulito — evita di lasciare lavoro "a metà" solo in locale

## Coordinamento tra le due macchine (nuovo, importante ora che entrambe sono attive)

- **Evita task Jules paralleli sulla stessa area di codice** da macchine diverse nello stesso momento — rischio di PR che confliggono al merge  
- **Tieni traccia mentale/informale** di quale macchina ha l'ultimo lavoro non ancora pushato, specialmente se lavori su entrambe nella stessa giornata  
- Prima di spegnere/lasciare una macchina, `git push` se hai commit locali

## Se qualcosa non ti convince dopo un merge già fatto

git add \<file-o-cartella-da-ripristinare\>

git commit \-m "revert(\<scope\>): ripristina versione precedente"

git push

## Comandi di emergenza

| Situazione | Comando |
| :---- | :---- |
| Merge bloccato, voglio annullare | `git merge --abort` |
| `merge --abort` fallisce ("not up to date") | `git reset --hard HEAD` poi ripeti `git merge --abort` |
| Working tree sporco, voglio scartare tutto | `git checkout -- .` seguito da `git clean -fd` |
| Voglio solo mettere da parte senza perdere nulla | `git add -A` poi `git stash push -m "..."` |
| Recupero da stash | `git stash pop` |
| Voglio azzerare tutto e allinearmi esattamente a GitHub | `git fetch origin` poi `git reset --hard origin/dev` |
| Verificare commit locali non voluti prima di un reset | `git log origin/dev..dev --oneline` |

## File locali non versionati (in `.gitignore`, presenti su entrambe le macchine)

leeno\_test\_jules.bat

leeno\_restore\_dev.bat

leeno\_pull.bat

## Riferimento rapido nomi branch Jules

Sempre nel formato `jules-<numero>-<hash>` — copialo per intero dalla pagina della PR su GitHub, non abbreviare.

## Nota sul drive esterno

Se sposti/ricolleghi il drive tra PC, verifica sempre che la lettera resti `W:`. Se cambia, fissarla da **Gestione Disco** (`diskmgmt.msc`).  
