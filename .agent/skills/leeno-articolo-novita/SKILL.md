---
name: leeno-articolo-novita
description: >
  Genera un articolo di pubblicazione in italiano basato sulle novità più
  significative introdotte a LeenO rispetto all'ultimo rilascio stabile.
  Linguaggio orientato all'utente finale (tecnici, professionisti).
---

# LeenO – Generazione Articolo Novità (User-Oriented)

Questa skill trasforma il registro tecnico dei commit in un racconto di valore per l'utente finale (geometri, architetti, ingegneri, contabili).

## Procedura

1. **Individuazione Punto di Partenza**:
   Trova l'ultimo tag di rilascio stabile (es. `v3.25.0`):
   ```bash
   git describe --tags --abbrev=0 --match "v*"
   ```

2. **Estrazione Log delle Modifiche**:
   Ottieni l'elenco dei commit dal tag ad oggi:
   ```bash
   git log <ultimo_tag>..HEAD --pretty=format:"%s" --reverse
   ```

3. **Traduzione in Valore per l'Utente**:
   Non elencare i commit così come sono. Usa questa tabella per "tradurre" i cambiamenti:

   | Termine Tecnico | Traduzione per l'Utente (Beneficio) |
   |:---|:---|
   | **Fix IndexError / Bug** | Maggiore stabilità; Risolto errore nel calcolo/importazione |
   | **Refactor / Optimization** | Software più veloce e leggero; Operazioni più fluide |
   | **Feature / New Parser** | Nuova funzione per...; Gestione più semplice dei file X |
   | **Update Dependencies** | Migliore compatibilità con le ultime versioni di LibreOffice |
   | **UI/Style Improvement** | Interfaccia più chiara; Migliore visualizzazione dei dati |

4. **Redazione dell'Articolo**:
   - **Tono**: Entusiasta ma concreto. Parla di tempo risparmiato e precisione.
   - **No Informatichese**: Evita termini come "commit", "merge", "debug", "array", "parser".
   - **Focus**: "Cosa posso fare oggi che ieri non potevo fare (o facevo peggio)?"

5. **Struttura Suggerita**:
   - **Titolo**: Orientato al risultato (es: "LeenO: Importazioni più sicure e calcoli più veloci").
   - **Cosa cambia per te**: Riepilogo dei vantaggi pratici.
   - **Le Novità nel dettaglio**: Spiegate in modo operativo.
   - **Supporto e Community**: Dove scaricare e come chiedere aiuto.

## Esempio di Output Ideale

```markdown
# LeenO: Nuovi miglioramenti per la tua contabilità quotidiana

Continua l'evoluzione di LeenO per offrirti uno strumento sempre più affidabile per i tuoi computi metrici e la contabilità lavori. Ecco le principali novità introdotte negli ultimi giorni:

### 🚀 Importazioni più facili e sicure
- **Listini Esterni**: Abbiamo reso l'importazione dei file XPWE molto più robusta. Se prima alcuni file con strutture particolari potevano creare problemi, ora il sistema li gestisce in modo intelligente, garantendo che tutti i tuoi dati arrivino correttamente a destinazione.
- **Velocità Operativa**: Chi usa listini molto grandi noterà una maggiore fluidità nel caricamento e nella gestione delle impostazioni generali.

### 🛠️ Stabilità e Precisione
- Abbiamo affinato i calcoli interni per garantire la massima precisione, specialmente quando si lavora con molte categorie e sottocategorie.
- Migliorata la compatibilità con le versioni più recenti di LibreOffice per evitare rallentamenti grafici.

### 📦 Prova subito le novità
Puoi scaricare l'ultima versione di prova dall'area [Versioni di Sviluppo](https://leeno.org/versioni-sviluppo/).
```
