---
name: leeno-sync-scorciatoie
description: >
  Sincronizza le scorciatoie da tastiera definite in Accelerators.xcu con il 
  foglio 'Scorciatoie' nel template Computo_LeenO.ods. Da usare dopo ogni 
  modifica dei tasti rapidi per mantenere aggiornata la documentazione interna.
---

# LeenO – Sincronizzazione Scorciatoie

Questa skill permette di mantenere allineata la documentazione delle scorciatoie da tastiera contenuta nel template principale di LeenO con l'effettiva configurazione XML dell'estensione.

## Quando usarla
- Dopo aver aggiunto, rimosso o modificato una scorciatoia in `Accelerators.xcu`.
- Dopo aver cambiato il titolo o la descrizione di un comando in `Addons.xcu`.
- Dopo aver aggiunto o modificato un uso di `GetModifiers()` in un file Python del progetto (vedi sotto: non solo `Accelerators.xcu`).
- Prima di ogni rilascio ufficiale, per garantire che l'utente veda informazioni corrette.

## Procedura manuale (AI)
Per eseguire la sincronizzazione, invoca lo script Python contenuto nella cartella della skill:

```powershell
python .agent/skills/leeno-sync-scorciatoie/scripts/sync_shortcuts.py
```

## Cosa fa lo script
1. **Analisi**: Legge `Accelerators.xcu` per identificare le macro associate ai tasti.
2. **Mapping**: Cerca i nomi leggibili dei comandi in `Addons.xcu`.
3. **Aggiornamento Template**:
   - Decomprime temporaneamente `Computo_LeenO.ods`.
   - Modifica `content.xml` per rigenerare la tabella dei tasti nel foglio "Scorciatoie".
   - Riorganizza le righe per categoria (CTRL, SHIFT, CTRL+SHIFT, ALT, Combinazioni toolbar + tastiera).
   - Ricomprime il pacchetto ODS.
4. **Backup**: Crea automaticamente una copia `.bak` del template prima di sovrascriverlo.

## Copertura di `GetModifiers()` oltre `Accelerators.xcu`

`Accelerators.xcu` documenta solo le scorciatoie da tastiera vere e proprie. Diverse funzioni del progetto chiamano `GetModifiers()` (o `PL.GetModifiers()`) per rilevare CTRL/SHIFT al click su un **pulsante toolbar**, senza che questo compaia in `Accelerators.xcu`: sono comportamenti alternativi (es. Ctrl+Click su un'icona esegue un'azione diversa dal click semplice) che lo script non individua automaticamente e che vanno mappati a mano nella lista `toolbar_combos` dentro `sync_shortcuts.py`.

Prima di eseguire la sincronizzazione, verificare se sono comparsi nuovi usi di `GetModifiers()` non ancora coperti:

```bash
grep -rn "GetModifiers()" src/Ultimus.oxt/python/pythonpath/
```

Per ogni occorrenza non ancora presente in `toolbar_combos`:
1. Risalire alla funzione che la contiene e al comando/URL corrispondente in `Addons.xcu` (per il titolo leggibile del pulsante toolbar).
2. Leggere il corpo della funzione per capire cosa cambia quando `is_ctrl`/`is_shift` è vero (es. azione inversa, target alternativo, comportamento distruttivo).
3. Aggiungere una tupla `('Ctrl + Click', "Toolbar '<Titolo>': <descrizione del comportamento alternativo>", '<nome funzione>')` alla lista `toolbar_combos`, coerente con lo stile delle voci esistenti.
4. Se `is_shift` non è mai usato nella funzione, non serve una voce separata per Shift+Click.

---
> [!IMPORTANT]
> Lo script si aspetta che la struttura delle cartelle sia quella standard del repository LeenO. Non spostare lo script al di fuori della cartella della skill se non per refactoring pianificati.

> [!TIP]
> Dopo l'esecuzione, è consigliabile aprire il template aggiornato con LibreOffice Calc per verificare visivamente che il layout del foglio "Scorciatoie" sia corretto e leggibile.
