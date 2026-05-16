---
name: leeno-aggiorna-manuale
description: >
  Skill per aggiornare sistematicamente il manuale ufficiale di LeenO (MANUALE_LeenO.fodt)
  partendo dalle modifiche apportate al codice, evitando duplicazioni grazie a un file di tracking.
---

# LeenO – Aggiornamento Manuale con Tracciamento

Questa skill ha l'obiettivo di mantenere sincronizzato il manuale utente di LeenO (`documentazione/MANUALE_LeenO.fodt`) con le ultime modifiche al codice, assicurandosi di non ripetere il lavoro su commit o funzionalità già documentate.

## File Chiave
- **Manuale**: `w:\_dwg\ULTIMUSFREE\_SRC\leeno\documentazione\MANUALE_LeenO.fodt` (formato XML)
- **Registro Aggiornamenti**: `w:\_dwg\ULTIMUSFREE\_SRC\leeno\documentazione\TRACKING_MANUALE.md`
- **Mappa Sezioni**: `MAPPA_SEZIONI.md` (nella directory di questa skill)

## Risorse della Skill
- `scripts/genera_mappa.py` — Script Python per rigenerare la mappa delle sezioni del manuale.

---

## Sorgenti di Informazione

Per individuare le novità da documentare, non limitarti ai soli file Python. Consulta **tutte** le seguenti sorgenti:

| Sorgente | Percorso | Cosa cercare |
| :--- | :--- | :--- |
| Codice Python | `src/Ultimus.oxt/python/pythonpath/*.py` | Nuove funzioni `MENU_*`, modifiche a flussi utente |
| Menù e sotto-menù | `src/Ultimus.oxt/Addons.xcu` | Nuove voci di menù, riorganizzazione menù, etichette |
| Scorciatoie da tastiera | `src/Ultimus.oxt/Accelerators.xcu` | Nuove scorciatoie, modifiche a scorciatoie esistenti |
| Finestre di dialogo | `src/Ultimus.oxt/dialogs/*.xdl` | Nuove finestre, nuovi controlli, campi rinominati |
| File di proprietà | `src/Ultimus.oxt/dialogs/*.properties` | Etichette e testi dei dialoghi |
| Configurazione | `src/Ultimus.oxt/python/pythonpath/LeenoConfig.py` | Nuove opzioni di configurazione |

> [!IMPORTANT]
> Quando documenti una funzionalità nel manuale, verifica **sempre** il percorso di menù reale in `Addons.xcu` e le scorciatoie in `Accelerators.xcu`, in modo che le istruzioni nel manuale riflettano fedelmente l'interfaccia utente.

---

## Procedura Operativa

### Fase 0: Consultare la Mappa delle Sezioni
1. Apri `MAPPA_SEZIONI.md` (nella directory di questa skill) con `view_file`.
2. Individua la sezione più adatta in cui inserire la nuova documentazione.
3. Annota il numero di riga e il nome del bookmark del punto di inserimento.

> [!TIP]
> Se il manuale è stato modificato di recente e la mappa potrebbe non essere aggiornata, rigenera la mappa:
> ```
> python .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py
> ```

### Fase 1: Verifica dello Storico (Tracking)
1. Apri e leggi il file `documentazione/TRACKING_MANUALE.md` utilizzando il tool `view_file`.
2. Controlla la tabella per capire quali sono le ultime funzionalità o gli ultimi commit (`Commit Hash`) già documentati. In questo modo saprai esattamente da quale punto temporale o da quale modifica ripartire.

### Fase 2: Individuazione delle Novità da Documentare
1. Chiedi all'utente su quale commit/file sta lavorando o usa `git log` per vedere gli ultimi commit sul codice sorgente.
2. Confronta queste novità con il file di tracking.
3. Seleziona le funzionalità o le modifiche che non sono ancora presenti nella tabella di `TRACKING_MANUALE.md` e che necessitano di una spiegazione nel manuale utente.
4. Per ciascuna novità, consulta anche:
   - `Addons.xcu` per il percorso di menù esatto e le etichette visibili all'utente.
   - `Accelerators.xcu` per le scorciatoie da tastiera associate.
   - I file `.xdl` e `.properties` dei dialoghi per i nomi dei campi e i testi dei pulsanti.

### Fase 3: Modifica del Manuale (FODT)
Essendo il file `MANUALE_LeenO.fodt` un file XML di grandi dimensioni (oltre 3 MB e 37.000 righe):
1. **NON usare** script complessi o editor massivi che potrebbero corrompere la struttura XML.
2. Usa la **Mappa Sezioni** (`MAPPA_SEZIONI.md`) per trovare la posizione di inserimento senza dover scansionare l'intero file.
3. Usa `view_file` con il numero di riga dalla mappa per ispezionare il contesto XML del punto di inserimento.
4. Usa Python per effettuare un semplice replace testuale stringa-per-stringa sulla riga interessata. Esempio di snippet sicuro:
   ```python
   import sys
   p = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/documentazione/MANUALE_LeenO.fodt'
   with open(p, 'r', encoding='utf-8') as f:
       text = f.read()
   target = "vecchio testo esatto compresi i tag XML"
   replacement = "nuovo testo con le nuove istruzioni"
   if target in text:
       text = text.replace(target, replacement)
       with open(p, 'w', encoding='utf-8') as f:
           f.write(text)
       print("Sostituzione completata")
   else:
       print("Testo target non trovato")
   ```
5. Scrivi il testo in **italiano chiaro e formale**, orientato all'utente finale (senza terminologia informatica, a meno che non si parli di tasti o menu).

### Fase 4: Aggiornamento del Tracking
Una volta che la modifica al manuale è andata a buon fine, devi aggiornare il registro in modo che nessuno ripeta questo lavoro in futuro.
1. Edita `documentazione/TRACKING_MANUALE.md` accodando una nuova riga alla tabella in fondo al file.
2. La riga dovrà contenere: `| YYYY-MM-DD | Hash del commit o Nome del file modificato | Breve descrizione della modifica nel codice | Sezione del manuale in cui hai scritto |`

### Fase 5: Rigenera la Mappa delle Sezioni
Dopo aver modificato il manuale, rigenera sempre la mappa per mantenerla aggiornata:
```
python .agent/skills/leeno-aggiorna-manuale/scripts/genera_mappa.py
```

### Fase 6: Conclusione
Comunica all'utente l'avvenuto aggiornamento del manuale, mostrando la porzione di testo inserita, e conferma l'avvenuta registrazione in `TRACKING_MANUALE.md`.
