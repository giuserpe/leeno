---
name: wordpress-theme-build
description: >
  Skill per pacchettizzare il tema WordPress di LeenO ("leeno-theme") in un archivio
  ZIP pronto per l'installazione tramite la bacheca di WordPress.
---

# LeenO – Costruzione Tema WordPress (Packaging)

Questa skill automatizza la creazione dell'archivio ZIP del tema WordPress `leeno-theme`, utilizzabile per l'installazione o l'aggiornamento del tema sul sito. La skill fa leva su uno script PowerShell apposito che esclude file di sviluppo non necessari (es. file di git, script di build, vecchi zip).

## Posizione dei File
- **Sorgente del Tema:** `w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme`
- **Script di Build:** `w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme\build_zip.ps1`
- **Destinazione ZIP:** `w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme-v1.2-hybrid.zip`

---

## Procedura Operativa

### Fase 1: Esecuzione del Build
1. Per pacchettizzare il tema, esegui il comando PowerShell `build_zip.ps1`. Puoi usare il tool `run_command` per lanciare lo script:
   ```powershell
   cd w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme
   .\build_zip.ps1
   ```
2. Attendi la fine dell'esecuzione dello script. L'output mostrerà la lista dei file aggiunti e, al termine, confermerà il successo dell'operazione.

### Fase 2: Verifica
1. Dopo l'esecuzione, verifica che il file ZIP sia stato effettivamente generato o sovrascritto nella directory padre (`W:\_dwg\ULTIMUSFREE\@SITO\`).
2. Puoi farlo controllando la data di ultima modifica usando un comando analogo a `Get-Item w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme-v1.2-hybrid.zip` in PowerShell.

### Fase 3: Conclusione
Comunica all'utente che l'archivio ZIP è stato creato con successo e che si trova in `w:\_dwg\ULTIMUSFREE\@SITO\leeno-theme-v1.2-hybrid.zip`, pronto per essere caricato sul sito WordPress.
