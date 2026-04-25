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
   - Riorganizza le righe per categoria (CTRL, SHIFT, CTRL+SHIFT, ALT).
   - Ricomprime il pacchetto ODS.
4. **Backup**: Crea automaticamente una copia `.bak` del template prima di sovrascriverlo.

---
> [!IMPORTANT]
> Lo script si aspetta che la struttura delle cartelle sia quella standard del repository LeenO. Non spostare lo script al di fuori della cartella della skill se non per refactoring pianificati.

> [!TIP]
> Dopo l'esecuzione, è consigliabile aprire il template aggiornato con LibreOffice Calc per verificare visivamente che il layout del foglio "Scorciatoie" sia corretto e leggibile.
