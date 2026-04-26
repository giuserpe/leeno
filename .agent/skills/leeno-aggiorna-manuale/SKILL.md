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

---

## Procedura Operativa

### Fase 1: Verifica dello Storico (Tracking)
1. Apri e leggi il file `documentazione/TRACKING_MANUALE.md` utilizzando il tool `view_file`.
2. Controlla la tabella per capire quali sono le ultime funzionalità o gli ultimi commit (`Commit Hash`) già documentati. In questo modo saprai esattamente da quale punto temporale o da quale modifica ripartire.

### Fase 2: Individuazione delle Novità da Documentare
1. Chiedi all'utente su quale commit/file sta lavorando o usa `git log` per vedere gli ultimi commit sul codice sorgente (es. in `src/Ultimus.oxt/python/pythonpath/`).
2. Confronta queste novità con il file di tracking.
3. Seleziona le funzionalità o le modifiche che non sono ancora presenti nella tabella di `TRACKING_MANUALE.md` e che necessitano di una spiegazione nel manuale utente.

### Fase 3: Modifica del Manuale (FODT)
Essendo il file `MANUALE_LeenO.fodt` un file XML di grandi dimensioni (oltre 3 MB e 35.000 righe):
1. **NON usare** script complessi o editor massivi che potrebbero corrompere la struttura XML.
2. Usa script Python con `xml.etree.ElementTree` o `grep_search` per trovare la posizione esatta in cui inserire il testo (cerca parole chiave del paragrafo esistente).
3. Usa Python per effettuare un semplice replace testuale stringa-per-stringa sulla riga interessata. Esempio di snippet sicuro:
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
4. Scrivi il testo in **italiano chiaro e formale**, orientato all'utente finale (senza terminologia informatica, a meno che non si parli di tasti o menu).

### Fase 4: Aggiornamento del Tracking
Una volta che la modifica al manuale è andata a buon fine, devi aggiornare il registro in modo che nessuno ripeta questo lavoro in futuro.
1. Edita `documentazione/TRACKING_MANUALE.md` accodando una nuova riga alla tabella in fondo al file.
2. La riga dovrà contenere: `| YYYY-MM-DD | Hash del commit o Nome del file modificato | Breve descrizione della modifica nel codice | Sezione del manuale in cui hai scritto |`

### Fase 5: Conclusione
Comunica all'utente l'avvenuto aggiornamento del manuale, mostrando la porzione di testo inserita, e conferma l'avvenuta registrazione in `TRACKING_MANUALE.md`.
