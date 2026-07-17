# Specifica del Design System delle Icone di LeenO

**Versione 2.0 (Nuova Generazione)**
**Autore:** Senior Product Designer, Icon Designer e Design System Architect

---

## 1. Filosofia di Design

Il sistema di icone di nuova generazione per **LeenO** si basa sui principi di **chiarezza semantica, coerenza visiva e riconoscibilità funzionale**. È progettato specificamente per le esigenze professionali di architetti, ingegneri, geometri e professionisti della pubblica amministrazione che lavorano con computi metrici estimativi, analisi dei prezzi e contabilità dei lavori.

Il sistema passa da un insieme di icone arcaico, frammentato e con metafore multiple a una **famiglia di icone outline pulita, moderna e minimalista**.

### Principi Cardine:

- **Minimalista e Professionale:** Stile outline con tratti da 2px, angoli arrotondati e geometrie semplici.
- **Coerenza Semantica:** Le icone appartenenti alla stessa categoria (es. categorie, voci di lavoro, file) utilizzano le stesse "primitive" geometriche di base e una logica visiva condivisa.
- **Versatilità del Tema:** Progettato innanzitutto in SVG per apparire nitido e leggibile sia in ambienti chiari che scuri, con una tavolozza di colori estremamente limitata.
- **Alta leggibilità a 16×16 px:** Limiti dei pixel ottimizzati per prevenire l'anti-aliasing e la sfocatura.

---

## 2. Primitive Riutilizzabili e Anatomia delle Icone

Per fare in modo che ogni icona sembri appartenere alla stessa famiglia, stabiliamo un insieme fondamentale di primitive geometriche (forme di base).

### A. Forme di Base (Primitive Documento e Cartella)

1.  **Forma Documento:** Un rettangolo verticale con l'angolo in alto a destra piegato.
    - _Dimensioni (su griglia 24×24):_ larghezza `14px` × altezza `18px`.
    - _Dimensione piega:_ piega dell'angolo di `4px`.
2.  **Forma Cartella:** Una classica scheda di cartella.
    - _Dimensioni (su griglia 24×24):_ larghezza `20px` × altezza `16px`.
    - _Posizione linguetta:_ allineata a sinistra, larghezza `8px`, altezza `3px`.
3.  **Forma WBS / Elenco:** Linee parallele orizzontali, strutturate in modo gerarchico o sequenziale.

### B. Sovrapposizioni e Badge

Le sovrapposizioni sono simboli d'azione standard posizionati nel **quadrante in basso a destra** (o in alto a destra per azioni specifiche) delle forme di base per indicare le operazioni.

- **Badge Più (Aggiunta):** Semplici linee intersecanti lunghe `5px`, tratto `2px`, posizionate in basso a destra.
- **Badge Meno/Elimina (Rimozione/Sottrazione):** Tratto standard `-` o `×` diagonale.
- **Badge Cerca/Trova:** Una piccola icona a forma di lente d'ingrandimento sovrapposta alla forma di base.
- **Badge Avviso/Errore:** Un triangolo di avviso `▲` o un cerchio con punto esclamativo `!`.
- **Badge Successo/Spunta:** Una spunta pulita `✓`.

### C. Frecce

- **Frecce Direzionali:** Semplice chevron `>` o freccia a linea `→` con un tratto di `2px`, giunti a `90°` e terminazioni arrotondate.

---

## 3. Griglia, Proporzioni e Regole di Spaziatura

Per garantire una resa nitida alle risoluzioni native del desktop, le icone sono sviluppate su una griglia vettoriale.

```
       Griglia Master 24x24 px
  +--------------------------+
  |  . . . . . . . . . . . . |  <-- 2px di padding dell'area di sicurezza (nessun tracciato vitale)
  |  . +------------------+ . |
  |  . |                  | . |
  |  . |                  | . |  Area di progettazione attiva: 20x20 px
  |  . |                  | . |
  |  . |                  | . |
  |  . +------------------+ . |
  |  . . . . . . . . . . . . |
  +--------------------------+
```

- **Griglia Master:** `24×24 pixel` (scalabile a `16×16`, `32×32`, `48×48`).
- **Spessori del Tratto:**
  - Bordi e contorni principali: `2px` (esattamente sulle linee della griglia dei pixel).
  - Dettagli interni o accenti secondari: `1.5px` o `1px`.
- **Raggio dell'Angolo:**
  - Angoli esterni: raggio di `2px`.
  - Dettagli interni / giunti di piega: `1px` o netti a `0px` a seconda del contesto.
- **Padding / Area di Sicurezza:**
  - Margine di `2px` su tutti i lati della tela `24×24`.
  - Nessun punto di ancoraggio strutturale o elemento chiave deve trovarsi nell'area di sicurezza, a meno che non trabocchi intenzionalmente per l'equilibrio visivo (es. punte di freccia sottili).
- **Allineamento Ottico:** Centrato visivamente. Gli elementi orizzontali devono essere allineati lungo la linea centrale orizzontale della griglia; gli elementi verticali lungo la linea centrale verticale.

---

## 4. Tavolozza dei Colori e Uso del Tema

La tavolozza dei colori è strettamente limitata a 8 colori funzionali per garantire la massima coerenza e un contrasto elevato sia nei temi Chiari che in quelli Scuri.

### A. La Tavolozza dei Colori di Nuova Generazione

| Nome Colore          | Codice Hex | Significato Semantico / Utilizzo                                               |
| :------------------- | :--------- | :----------------------------------------------------------------------------- |
| **Verde Primario**   | `#5D7400`  | Branding principale, strutture primarie, successo, stati affermativi           |
| **Lime di Accento**  | `#AAD400`  | Evidenziazioni, dettagli ausiliari, accento distintivo del brand               |
| **Arancione Azione** | `#FF4D2E`  | Eliminazioni, sottrazioni, operazioni distruttive, avvertimenti                |
| **Blu Info**         | `#3B82F6`  | Viste, documenti, collegamenti esterni, badge di info/aiuto                    |
| **Giallo Avviso**    | `#F4B400`  | Avvisi di stato, stati temporanei, evidenziazioni di ricerca, utility          |
| **Scuro**            | `#1A2010`  | Colore contorno predefinito per il Tema Chiaro, testo, griglie                 |
| **Sfondo**           | `#F0F4E0`  | Riempimenti interni (semi-opachi), sfondi dei contenitori                      |
| **Grigio**           | `#808080`  | Stati disabilitati, linee della griglia, guide strutturali, elementi secondari |

### B. Adattabilità del Tema (Chiaro vs. Scuro)

- **Tema Chiaro (`icons/svg/`):** I tratti primari utilizzano lo `Scuro` (`#1A2010`) o il `Verde Primario` (`#5D7400`). I riempimenti interni (se presenti) sono trasparenti o presentano riempimenti chiari del contenitore (`#F0F4E0`).
- **Tema Scuro (`icons/scuro/`):** I tratti primari si invertono automaticamente in colori chiari (`#FFFFFF` o `#F0F4E0`). I colori semantici come `Arancione Azione` e `Blu Info` rimangono identici ma vengono leggermente regolati per la luminanza.

---

## 5. Fase 1: Analisi Approfondita e Critica dell'Inventario delle Icone

L'attuale libreria di icone di LeenO soffre di diversi colli di bottiglia nel design:

1.  **Concetti Duplicati:** Molte icone utilizzano gli stessi simboli generici (es. `image15`, `vintage`) per azioni completamente diverse, rendendo le barre degli strumenti ripetitive e confuse.
2.  **Nomi Generici/Numerati:** I nomi dei file come `image14`, `image15`, `image18`, `image37`, `image93`, `image100`, `image444` mancano di significato semantico. Ciò ostacola la manutenibilità del codice e l'inserimento di nuovi designer.
3.  **Metafore Obsolete e Superate:**
    - `Caschetto` (elmetto da cantiere) utilizzato per duplicare una voce di lavoro in una voce di sicurezza è estremamente letterale e visivamente pesante.
    - `falegname` per importare un file DAT personalizzato è molto specifico e manca di una chiara traduzione dell'utilità software.
    - `sfera_gialla` per le importazioni di stile non ha alcun collegamento logico con fogli di stile o modelli.
4.  **Simboli Ambigui:** `sf_Ver` (sfera/pulsante verde) è utilizzato per "Numeri in lettere". Questo ha zero metafore tipografiche o numeriche.

---

## 6. Fase 2: Famiglie Semantiche

Organizziamo tutte le icone di LeenO in 9 chiare famiglie semantiche per stabilire pattern visivi funzionali.

### Categoria 1: Principale e Navigazione

Operazioni principali, punti di ingresso e collegamenti alla documentazione.

- `leeno`: Menu Principale dell'Estensione / Dashboard
- `manuale`: Manuale di Istruzioni in PDF
- `teleg`: Gruppo di Supporto della Community su Telegram

### Categoria 2: Struttura di Scomposizione del Lavoro (WBS)

Definizione della gerarchia visiva del progetto di costruzione.

- `supcat`: SuperCategoria (Livello 1)
- `cat`: Categoria (Livello 2)
- `subcat`: SottoCategoria (Livello 3)
- `image8` (`struttura_on`): Organizza / Attiva vista struttura
- `image9` (`struttura_off`): Pulisci / Disattiva vista struttura
- `rinumCap`: Rinumera Voci di Lavoro e Categorie

### Categoria 3: Voci di Lavoro (Voci)

Operazioni riguardanti singole voci dell'elenco, misurazioni e descrizioni.

- `image93` (`nuova_voce`): Inserisce una nuova voce di lavoro vuota
- `Corta` (`voce_breve`): Alterna la descrizione completa / vista codice breve
- `vedivoce`: Alterna la vista della voce di riferimento precedente
- `pesca`: Cattura/Eredita il codice dalla selezione attiva
- `invia_voce_ep`: Invia le voci selezionate all'Elenco Prezzi principale (DP)
- `compo` (`aggiungi_misura`): Aggiunge una nuova riga di misura (rigo di misura)
- `image37` (`sposta_voce`): Sposta la voce selezionata verticalmente

### Categoria 4: Elenchi Prezzi e Analisi dei Costi

Operazioni all'interno dell'Elenco Prezzi Regionale e dell'Analisi dei Prezzi (Analisi).

- `2ep` (`analisi_a_prezzo`): Crea una nuova voce di prezzo dai dettagli dell'analisi
- `perc` (`utili_maggiorazioni`): Configura le percentuali di utile / spese generali (%)
- `image21B` (`elimina_doppioni`): Rimuove codici di voce identici (deduplica)
- `riordina`: Ordina le voci alfabeticamente

### Categoria 5: Quantità e Contabilità

Formule, subtotali e tenuta della contabilità in cantiere.

- `parz`: Inserisce un subtotale parziale (parziale)
- `invert` (`inverti_segno`): Alterna quantità di lavoro positive/negative (+/-)
- `azzera`: Imposta a zero (0) le quantità delle voci selezionate
- `part_agg` (`partita_provvisoria_piu`): Inserisce una registrazione contabile provvisoria positiva
- `part_det` (`partita_provvisoria_meno`): Detrae una registrazione contabile provvisoria negativa
- `strutt_voci_zero`: Nasconde le voci di lavoro con quantità pari a zero
- `elimina_azzerate`: Elimina dall'elenco le voci di lavoro con quantità pari a zero
- `elimina_vuote`: Pulisce le righe del foglio di calcolo completamente vuote

### Categoria 6: Layout, Fogli e Viste

Controlli visivi, griglie e strutture di visualizzazione.

- `image18` (`scelta_viste`): Seleziona le viste del foglio di lavoro (Computo / Stampa / Computo & Stampa)
- `adattaH`: Adatta automaticamente l'altezza delle righe alla lunghezza del testo
- `griglia3` (`mostra_griglia`): Alterna le griglie del foglio di calcolo
- `vintage` (`copertine`): Gestisce/Visualizza le copertine di progetto
- `colore_tematico`: Personalizzatore del colore del tema

### Categoria 7: Reporting, Stampa ed Esportazione

Pubblicazione di report di progetto, elenchi e stime.

- `riepilogo`: Firme e totali riassuntivi del progetto
- `riepilogo_quantita`: Report riassuntivo quantitativo dei materiali
- `riepilogo_a2`: Riepilogo WBS complessivo dei costi
- `print_ok` (`anteprima_stampa`): Configurazione visiva dell'anteprima di stampa
- `image100` (`riga_rossa`): Inserisce una barra orizzontale rossa spessa di chiusura (fine computo)

### Categoria 8: Utility e Configurazioni

Strumenti di sistema, convertitori e configurazioni.

- `config`: Preferenze generali di sistema
- `image16` (`stringhe_numeri`): Converte le rappresentazioni di stringhe in numeri
- `image17` (`sproteggi_tutto`): Sblocca/Sprotegge tutti i fogli in Calc
- `sfera_gialla` (`importa_stili`): Importa stili tipografici e di layout da un modello esterno
- `sf_Ver` (`numeri_lettere`): Converte i valori numerici in lettere (es. 100 -> "cento")

### Categoria 9: Sviluppatore e Importazioni Legacy

Strumenti amministrativi, di diagnostica di sistema e convertitori legacy.

- `py` (`python_debug`): Apre la console del debugger della shell Python
- `refresh`: Ricarica a caldo il file `Addons.xcu` e le strutture del menu
- `falegname` (`importa_dat`): Convertitore speciale di importazione legacy per file DAT (falegnameria e settori affini)

---

## 7. Fase 4: Revisione Esaustiva delle Icone e Specifiche di Riprogettazione

Di seguito è riportato il progetto completo di riprogettazione per ogni icona nella libreria di LeenO.

| Nome File Icona                   | Metafora / Simbolo Attuale                                                 | Proposta di Metafora Moderna                                                         | Dettagli di Design e Motivazione                                                                                                                                  | Priorità  |
| :-------------------------------- | :------------------------------------------------------------------------- | :----------------------------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------- | :-------- |
| **leeno**                         | Icona quadrata con lettere L, O e sfumatura verde                          | Marchio piatto e coeso. Lettere 'L' e 'O' intrecciate in stile outline vettoriale.   | Unifica l'identità visiva del brand. Contorno ad alto contrasto della lettera 'L' (in grassetto, #5D7400) annidata all'interno della lettera 'O' (#AAD400).       | **Bassa** |
| **manuale**                       | Raccoglitore giallo con logo LibreOffice                                   | Un'icona documento piatta con un angolo piegato e un simbolo di informazioni (`i`)   | Alta visibilità a 16px. Combina la primitiva del documento con una linea verticale centrale pulita della lettera minuscola 'i'.                                   | **Media** |
| **teleg**                         | Vecchio aeroplano di carta blu circolare di Telegram                       | Profilo minimalista di aeroplano di carta con tratto da 2px                          | Aeroplano di carta moderno in stile Tabler. Scalato perfettamente per adattarsi all'area di progettazione 20x20.                                                  | **Bassa** |
| **supcat**                        | Schede gialle gerarchiche con frecce su/giù                                | Una primitiva cartella contenente il numero romano "I"                               | La "SuperCategoria" rappresenta il Livello 1 della gerarchia di progetto. L'uso di un contenitore cartella con il numero romano "I" rende la gerarchia intuitiva. | **Alta**  |
| **cat**                           | Scheda divisoria orizzontale con linguetta rossa                           | Una primitiva cartella contenente il numero romano "II"                              | Segue la gerarchia principale. Stabilisce la cartella di Livello 2 contenente il numero romano "II" all'interno dell'area di sicurezza.                           | **Alta**  |
| **subcat**                        | Scheda cartella blu annidata con piccola gerarchia                         | Una primitiva cartella contenente il numero romano "III"                             | Completa la gerarchia delle cartelle. Stabilisce la cartella di Livello 3 contenente il numero romano "III". Eccellente coerenza familiare.                       | **Alta**  |
| **image8** (`struttura_on`)       | Pulsanti grigi di espansione ad albero                                     | Un elenco strutturato con un indicatore di espansione (`+` o rientro gerarchico)     | Sostituisce il nome generico con `struttura_on`. Mostra tre punti dell'elenco strutturato con guide di rientro.                                                   | **Alta**  |
| **image9** (`struttura_off`)      | Pulsanti grigi di compressione ad albero                                   | Un elenco strutturato con un indicatore di riduzione o una linea diagonale           | Sostituisce il nome generico con `struttura_off`. Indicatore visivo chiaro per comprimere i dettagli.                                                             | **Alta**  |
| **rinumCap**                      | Freccia di ricaricamento verde circolare con righe di elenco               | Primitiva elenco standard (`≡`) con un simbolo cancelletto adiacente (`#`)           | Altamente leggibile. La combinazione di righe di elenco e un indicatore di numeri/cancelletto comunica immediatamente la "rinumerazione".                         | **Media** |
| **image93** (`nuova_voce`)        | Documento bianco con badge più verde                                       | Primitiva documento con un badge Più Verde (`+`) pulito in basso a destra            | Sostituisce il nome generico con `nuova_voce`. Segue perfettamente le regole della grammatica visiva.                                                             | **Alta**  |
| **Corta**                         | Forbici diagonali che tagliano un foglio di documento                      | Icona piatta di forbici combinata con una linea tratteggiata orizzontale             | Forbici outline modernizzate. Indica il taglio o l'accorciamento delle descrizioni sullo schermo.                                                                 | **Media** |
| **vedivoce**                      | Freccia circolare blu sinistra/destra                                      | Un occhio aperto che guarda una primitiva documento                                  | Molto più descrittivo. Permette di rivedere una voce di documento precedentemente referenziata.                                                                   | **Media** |
| **pesca**                         | Un frutto di pesca (gioco di parole in italiano: pesca = frutto / cattura) | Un amo da pesca sagomato o una freccia che estrae il codice da una cella             | Sebbene il gioco di parole sia divertente, un amo che afferra il contorno di una cella/codice è più professionale. Un amo semplificato è altamente riconoscibile. | **Media** |
| **invia_voce_ep**                 | Freccia curva blu che salta sopra una riga verticale                       | Primitiva documento con una freccia uscente (`→`) verso destra                       | Metafora universalmente compresa per esportare o inviare una voce selezionata a un altro elenco.                                                                  | **Alta**  |
| **compo**                         | Pila di righe di misura verdi/grigie                                       | Una primitiva foglio con una linea orizzontale e un badge più (`+`)                  | Indica l'inserimento di una riga figlia di dettaglio di calcolo o dimensione.                                                                                     | **Alta**  |
| **image37** (`sposta_voce`)       | Due frecce verdi verticali opposte                                         | Due frecce verticali pulite che puntano in alto e in basso in parallelo              | Sostituisce il nome generico con `sposta_voce`. Indica lo spostamento delle righe selezionate su/giù.                                                             | **Alta**  |
| **2ep**                           | Frecce di copia cartella rosse/blu                                         | Due contorni di documento sovrapposti con un percorso di frecce                      | Modernizza la copia cartella-a-cartella legacy. Indica la generazione di una nuova voce di prezzo dai dettagli dell'analisi.                                      | **Alta**  |
| **perc**                          | Segno di percentuale blu all'interno di un cerchio giallo                  | Segno di percentuale pulito `%` in Verde Primario con tratto da 2px                  | Rimuove la sfera gialla non necessaria. Leggibile a 16px con un allineamento netto dei pixel.                                                                     | **Media** |
| **image21B** (`elimina_doppioni`) | Cartelle rosse/verdi sovrapposte con una croce                             | Due primitive foglio sovrapposte con una sovrapposizione di sottrazione/cestino      | Sostituisce il nome generico con `elimina_doppioni`. Indica intuitivamente la rimozione dei codici duplicati dal database.                                        | **Alta**  |
| **riordina**                      | Frecce di ordinamento elenco A-Z                                           | Freccia verticale adiacente alle lettere 'A' e 'Z' impilate                          | Classica metafora di ordinamento universalmente compresa. Molto facile da leggere a 16px.                                                                         | **Media** |
| **parz**                          | Icona parentesi subtotale                                                  | Un segno di sommatoria matematica (`∑`) all'interno di parentesi                     | Indica la sommatoria parziale. Molto più professionale rispetto a un semplice contorno di parentesi.                                                              | **Alta**  |
| **invert**                        | Segni più e meno in pulsanti circolari grigi                               | Segni `+` e `-` puliti affiancati con una freccia di commutazione orizzontale        | Chiara indicazione dell'inversione dei segni matematici da positivi a negativi.                                                                                   | **Media** |
| **azzera**                        | Un cerchio grigio con una linea diagonale rossa e uno zero                 | Una grande cifra `0` in Arancione Azione con un tratto da 2px                        | Forte e chiaro. Imposta a zero le metriche di selezione attive.                                                                                                   | **Alta**  |
| **part_agg**                      | Pila di celle arancioni del foglio con badge più                           | Pila di schede contabili con un badge Più Verde (`+`)                                | Un libro mastro contabile con indicatore di aggiunta. Rappresenta aggiunte provvisorie.                                                                           | **Alta**  |
| **part_det**                      | Pila di celle arancioni del foglio con badge meno                          | Pila di schede contabili con un badge Meno Arancione (`-`)                           | Un libro mastro contabile con indicatore di sottrazione. Rappresenta detrazioni provvisorie.                                                                      | **Alta**  |
| **strutt_voci_zero**              | Albero di espansione con zero su un libro mastro                           | Primitiva struttura ad albero con uno zero sbarrato (`Ø`)                            | Indica il filtraggio o il nascondere gli elementi con valore zero dalla vista.                                                                                    | **Media** |
| **elimina_azzerate**              | Foglio di mastro con croce e zero                                          | Primitiva documento con uno zero (`0`) e un chiaro badge di eliminazione (`×`)       | Indicatore pulito di eliminazione per le righe con valore zero.                                                                                                   | **Alta**  |
| **elimina_vuote**                 | Libro mastro pulito con riga eliminata barrata                             | Primitiva elenco multi-riga con righe vuote evidenziate e un badge di eliminazione   | Indica la depurazione del foglio di calcolo dalle righe inutilizzate e vuote.                                                                                     | **Alta**  |
| **image18** (`scelta_viste`)      | Tre fogli multicolore impilati                                             | Uno schermo di monitor diviso verticalmente in diversi layout di visualizzazione     | Sostituisce il nome generico con `scelta_viste`. Rappresentazione software moderna e chiara delle viste dello schermo.                                            | **Alta**  |
| **adattaH**                       | Due frecce verticali che espandono righe orizzontali                       | Linea orizzontale pulita delimitata da frecce esterne su/giù                         | Indicatore di adattamento automatico dello spazio verticale universalmente compreso.                                                                              | **Media** |
| **griglia3**                      | Una griglia di linee in un foglio                                          | Un profilo di griglia `3×3` pulito in Scuro con angoli esterni arrotondati           | Attivazione visiva della griglia. Semplice, leggibile e strutturalmente equilibrato.                                                                              | **Bassa** |
| **vintage**                       | Vecchio cassetto di schedario con file                                     | Un contorno visivo di cartella ad anelli che mostra i segnaposto della copertina     | Sostituisce la metafora legacy di un cassetto fisico di un archivio cartaceo. Rappresenta le copertine dei progetti.                                              | **Alta**  |
| **colore_tematico**               | Tavolozza di colori rossa e blu                                            | Un contorno di secchio di vernice che versa una goccia di colore Lime Accent         | Classico personalizzatore del colore del tema del design system. Altamente intuitivo.                                                                             | **Media** |
| **riepilogo**                     | Foglio di raccoglitore arancione con firme                                 | Primitiva documento contenente linee e una penna stilografica in miniatura per firma | Rappresentazione visiva dei totali di progetto finalizzati e dei blocchi di firma esecutivi.                                                                      | **Media** |
| **riepilogo_quantita**            | Documento con tre barre orizzontali multicolore                            | Primitiva documento contenente il contorno di un mini grafico a barre                | Rappresenta i report quantitativi di distribuzione dei materiali e dei pesi.                                                                                      | **Media** |
| **riepilogo_a2**                  | Documento con metriche di griglia verdi/arancioni                          | Primitiva documento contenente una griglia di matrice dei costi                      | Indica calcoli complessi di ripartizione dei costi tra varianti.                                                                                                  | **Media** |
| **print_ok**                      | Foglio di documento che entra in una stampante                             | Un profilo di stampante piatta elegante e moderno con carta in uscita                | Icona di configurazione della stampa e del layout leggibile e ad alto contrasto.                                                                                  | **Bassa** |
| **image100** (`riga_rossa`)       | Barra rettangolare rossa spessa                                            | Un evidenziatore rosso che punta a una linea di chiusura orizzontale                 | Sostituisce il nome generico con `riga_rossa`. Indica chiaramente il blocco di chiusura del progetto.                                                             | **Alta**  |
| **config**                        | Ingranaggio grigio e chiave inglese incrociati                             | Due ingranaggi annidati di dimensioni diverse con denti arrotondati                  | Metafora universale dell'ingranaggio per le impostazioni di configurazione.                                                                                       | **Bassa** |
| **image16** (`stringhe_numeri`)   | Testo 'abc' con una freccia che punta ai numeri '123'                      | Profilo del testo `A` che punta tramite una freccia destra (`→`) a un numero `1`     | Sostituisce il nome generico con `stringhe_numeri`. Indicatore pulito per la conversione da testo a numero.                                                       | **Alta**  |
| **image17** (`sproteggi_tutto`)   | Lucchetto dorato aperto                                                    | Un contorno di lucchetto aperto in Giallo Avviso con un tratto da 2px                | Sostituisce il nome generico con `sproteggi_tutto`. Chiara metafora di sblocco dei fogli.                                                                         | **Alta**  |
| **sfera_gialla**                  | Semplice sfera gialla tridimensionale                                      | Un pennello in stile moderno sovrapposto a una scheda di foglio di calcolo           | Decisamente superiore. Rappresenta l'importazione di modelli di stile (colori, caratteri, bordi).                                                                 | **Alta**  |
| **sf_Ver** (`numeri_lettere`)     | Semplice sfera verde                                                       | Le lettere `123` con una freccia a fumetto che punta alla parola `abc`               | Rappresenta la conversione di cifre numeriche nel corrispondente testo in lettere.                                                                                | **Alta**  |
| **py** (`python_debug`)           | Due serpenti Python                                                        | Il logo Python (profilo vettoriale semplificato di due serpenti)                     | Icona del debugger Python leggibile. Si adatta alla tavolozza dei colori limitata.                                                                                | **Bassa** |
| **refresh**                       | Frecce di ricaricamento circolari                                          | Due frecce circolari che formano un ciclo continuo                                   | Simbolo dell'azione di ricarica/aggiornamento. Nitido, simmetrico e chiaro.                                                                                       | **Bassa** |
| **falegname**                     | Strumento letterale da falegname/carpentiere                               | Una primitiva di parentesi di codice (`<>`) con una freccia di importazione          | Sostituisce la metafora letterale del falegname. Indica l'importazione di file database DAT standard.                                                             | **Alta**  |

---

## 8. Fase 5: Icone Mancanti per un Flusso di Lavoro Ottimale

Per completare l'esperienza utente di LeenO, specifichiamo 5 nuove icone personalizzate per colmare le lacune funzionali esistenti.

### A. Nome Icona: `importa_xml`

- **Esigenza:** LeenO contiene importatori di parsing XML personalizzati (es. Listini regionali), ma attualmente non ha un'icona dedicata nei menu/barre degli strumenti.
- **Metafora Visiva:** Primitiva documento con la scritta `XML` stampata sopra, abbinata a una freccia in entrata in basso a sinistra (`↓`).
- **Posizionamento:** Sottomenu Principale Importazione File.

### B. Nome Icona: `esporta_gantt`

- **Esigenza:** Converte le quantità e le durate del progetto in formato CSV per GanttProject. Questa è una potente funzionalità attualmente nascosta nei menu senza alcuna iconografia.
- **Metafora Visiva:** Un profilo di un piccolo diagramma di Gantt (barre di attività orizzontali sfalsate) con una freccia di esportazione rivolta a destra (`→`).
- **Posizionamento:** Sottomenu Importa/Esporta.

### C. Nome Icona: `documento_bollo`

- **Esigenza:** Formatta le relazioni tecniche in documenti legali (documento uso bollo) con strutture a margini.
- **Metafora Visiva:** Un foglio di documento bordato contenente il profilo di un timbro a cera rotondo in Arancione Azione.
- **Posizionamento:** Sottomenu Nuovo Documento.

### D. Nome Icona: `unisci_fogli`

- **Esigenza:** Unisce tutti i fogli di lavoro aperti in un unico file di progetto consolidato.
- **Metafora Visiva:** Due singole schede di foglio che si uniscono in un unico foglio contenitore in primo piano.
- **Posizionamento:** Sottomenu Utility Fogli.

### E. Nome Icona: `somma_colore`

- **Esigenza:** Utility speciale che calcola i totali dei costi in base ai colori di evidenziazione del foglio di calcolo Calc.
- **Metafora Visiva:** Un segno di sigma (`∑`) adiacente al profilo di un evidenziatore colorato.
- **Posizionamento:** Sottomenu Utility di Calcolo.

---

## 9. Specifiche Colore e Monocromatiche

Il sistema di icone di nuova generazione opera in due modalità operative principali per supportare diversi motori di rendering dei client.

### A. Modalità a Colori (Predefinita)

- Utilizza la tavolozza funzionale limitata a 8 colori.
- Le linee sono principalmente `#1A2010` (Scuro) o `#5D7400` (Verde Primario) su sfondi chiari.
- Le evidenziazioni e gli accenti sfruttano il `#AAD400` (Lime) e il `#3B82F6` (Blu).
- Gli indicatori di azione e stato utilizzano `#FF4D2E` (Arancione) e `#F4B400` (Giallo).
- I riempimenti interni devono rimanere vuoti (trasparenti) o utilizzare il colore ad alto contrasto `#F0F4E0` (Sfondo) a un livello semi-opaco (vettore di tracciato piatto o `rgba`).

### B. Modalità Monocromatica (Alta Accessibilità / Temi a basso contrasto)

- Tutti i tracciati colorati sono convertiti in nero piatto (`#000000`) per i temi chiari o bianco piatto (`#FFFFFF`) per i temi scuri.
- Gli spessori dei tratti sono impostati uniformemente su `2px`.
- Le sovrapposizioni (es. `+`, `-`, `×`) sono separate dalla forma genitore utilizzando un intervallo di contorno trasparente di `1.5px` (maschera di confine dello spazio negativo) per garantire una chiara leggibilità anche senza variazioni di colore.

---

## 10. Raccomandazioni di SVG e Implementazione Tecnica

Per garantire un'implementazione impeccabile all'interno di LibreOffice Calc:

1.  **Standard Vettoriali Rigorosi:** Evitare l'esportazione con anteprime bitmap incorporate (i metadati `sodipodi` o `inkscape` devono essere eliminati utilizzando `scour` o `svgo` prima della distribuzione).
2.  **ViewBox e Confini:** Tutti i file sorgente devono essere centrati esattamente all'interno di `viewBox="0 0 24 24"`.
3.  **Nessuna Trasformazione:** Comprimere tutti i livelli annidati e applicare le trasformazioni direttamente ai tracciati.
4.  **Nessuno Stile HTML:** Utilizzare gli attributi di presentazione SVG in linea (`stroke`, `fill`, `stroke-width`, `stroke-linecap="round"`, `stroke-linejoin="round"`) invece dei blocchi di stile CSS, evitando che i motori di layout di Calc ignorino gli stili.
5.  **Nomi dei Codici Puliti:** Assicurarsi che i nomi dei file corrispondano alla loro descrizione semantica della famiglia anziché ai tag di layout, utilizzando il formato minuscolo snake_case (es. `nuova_voce.svg` invece di `image93.svg`).

# Addendum v2.1 — Correzioni Anti-Sovrapposizione e Anti-Confusione

**Da leggere insieme a "Specifica del Design System delle Icone di LeenO v2.0".**
Questo addendum **sostituisce** la sezione 2 (Primitive) e la sezione 9 (Colore) dove in conflitto, e **aggiunge** una nuova sezione 2.5 e una sezione 11 (Checklist di validazione).

Motivo: le icone generate nel primo passaggio (vedi allegato immagine) mostrano tratti troppo spessi che si sovrappongono e forme piene (non outline) che si fondono visivamente tra loro, rendendo il significato illeggibile anche a dimensione piena. Questo addendum introduce vincoli geometrici *misurabili* invece di indicazioni stilistiche generiche, in modo che Jules non possa interpretarle in modo troppo libero.

---

## 2.5 Regole Rigide Anti-Sovrapposizione (OBBLIGATORIE)

Queste regole hanno **priorità assoluta** su qualunque altra indicazione estetica del documento. Se una regola qui sotto è in conflitto con la sezione 5-8 (proposte di metafora), vince questa sezione.

### A. Le primitive "a linea" sono TRATTI, non blocchi pieni
- La "Forma WBS / Elenco" (linee parallele) e qualunque "barra" (es. Gantt, grafico) devono essere disegnate come **`<line>` o `<rect>` sottili con altezza massima 3px**, mai come pillole/capsule spesse riempite che occupano più del 12% dell'altezza della griglia.
- **Vietato**: rettangoli arrotondati (`rx` grande) usati come "barre piene" di colore saturo che assomigliano a pulsanti. Se serve un blocco pieno, l'altezza massima è 3px su griglia 24×24.

### B. Distanza minima tra elementi distinti
- Tra due primitive/badge/frecce diverse deve esserci **uno spazio vuoto (nessun tracciato) di almeno 1px** su griglia 24×24, salvo il caso C (alone/knockout).
- Nessuna forma può toccare o intersecare un'altra forma senza passare dalla regola C.

### C. Regola dell'Alone (Knockout) per ogni intersezione — [IMPLEMENTAZIONE AGGIORNATA, vedi sezione 12]
Quando un badge, una freccia o un accento **deve** sovrapporsi a una forma di base (es. freccia che attraversa una barra, badge in basso a destra su un documento), il principio resta: *nessun elemento può intersecare un altro senza una separazione visiva netta*. Il **meccanismo** con cui ottenerla, però, non è più uno stroke colore-sfondo (approccio fragile), ma il layering di forme piene descritto nella sezione 12: si disegna prima una forma piena "cuscinetto" colore Sfondo/Bianco leggermente più grande dell'elemento in primo piano, poi sopra l'elemento in primo piano con i suoi colori normali. Il risultato visivo è lo stesso (un piccolo margine netto attorno all'elemento che si sovrappone), ma la costruzione è coerente con il resto dell'icona (solo fill, mai stroke) e non produce artefatti di bordo sfocato.

Senza questo cuscinetto i tratti si fondono e l'icona diventa illeggibile — è esattamente il difetto visto nella prima generazione.

### D. Un solo elemento "principale" per icona
- Ogni icona ha **una sola forma primitiva dominante** (documento, cartella, lista, ecc.) + **al massimo un badge/accento** + **al massimo una freccia**.
- Se un'icona richiede più di 3 elementi grafici distinti, va semplificata concettualmente prima, non compressa graficamente. Esempio: `esporta_gantt` non deve avere più di 3 barre orizzontali, mai 4+.

### E. [SUPERATA — vedi sezione 12] Costruzione a strati di fill, non stroke
Questa regola è stata rivista dopo l'analisi delle icone reali di riferimento nella cartella Drive `1FKWfOENRWhc6CEeA8pIjt6FB2UjCibVc` (icone LibreOffice/Colibre): quel set **non usa mai l'attributo `stroke`**. Ogni "contorno", "anello" o "bordo" è ottenuto impilando forme piene (`fill`) di dimensione decrescente, mai tracciando una linea aperta che possa sovrapporsi ad un'altra. Vedi la sezione 12 per l'algoritmo esatto: sostituisce sia questa regola E sia il meccanismo dell'alone descritto in C, che va ora implementato come layering di fill invece che come stroke con margine colore-sfondo.

---

## 9. Tavolozza dei Colori — Regola di Limite (sostituisce 9A/9B)

- **Massimo 2 colori per icona**, oltre allo Scuro/Bianco del tratto principale:
  1. Colore del tratto principale: `Scuro` (#1A2010) o `Verde Primario` (#5D7400).
  2. **Un solo** colore d'accento semantico (es. Arancione per eliminazioni, Blu per info, Lime per evidenziazioni) usato *solo* sul badge o sulla freccia, mai su entrambi contemporaneamente con colori diversi.
- **Vietato** usare 3+ colori saturi diversi nella stessa icona (es. lime + arancione + blu insieme): è la causa diretta della confusione osservata.
- Gli sfondi chiari (`#F0F4E0`) sono ammessi solo come alone (vedi 2.5.C) o riempimento contenitore, mai come colore "decorativo" a sé stante.

---

## 11. Checklist di Validazione Pre-Consegna (Jules deve auto-verificare ogni icona)

Prima di considerare un'icona completa, Jules deve rispondere SÌ a tutte le domande seguenti; se anche una sola risposta è NO, l'icona va rifatta:

1. C'è **una sola** forma dominante e al massimo un badge + una freccia? (regola D)
2. Ogni tratto ha `stroke-width` costante = 2px (o 1.5/1px solo per dettagli secondari dichiarati)?
3. Nessun tratto satura più del 12% dell'area come blocco pieno colorato? (regola A/E)
4. Tra ogni coppia di elementi distinti c'è almeno 1px di spazio vuoto, oppure un alone esplicito dove si intersecano? (regola B/C)
5. L'icona usa al massimo 2 colori oltre allo Scuro? (sezione 9)
6. Renderizzando l'icona a 16×16px, ogni forma resta distinguibile senza fondersi con le altre? (test visivo esplicito, non assunto)
7. Rimuovendo il colore (scala di grigi), l'icona è ancora comprensibile solo dai contorni?

Se un'icona fallisce il punto 6, il problema quasi sempre è che due tratti si toccano senza alone (torna al punto 2.5.C) o che uno stroke è troppo spesso rispetto allo spazio disponibile.

---

## 12. Regole Tecniche di Costruzione (ricavate dalle icone di riferimento LibreOffice)

Analizzando i file SVG reali nella cartella `1FKWfOENRWhc6CEeA8pIjt6FB2UjCibVc` (es. `print.svg`, `runbasic.svg`, `radiomono2.svg`, `sc_aligntop.svg`, `odb_32_8.svg`) emergono regole tecniche precise, molto più rigorose di quelle stilistiche descritte finora. Jules deve seguire **esattamente** questo metodo di costruzione, non un'interpretazione libera.

### 12.1 Nessuno stroke, solo fill impilati
Nessuno dei file di riferimento usa `stroke` per il rendering (l'attributo `stroke-width` a volte presente è un residuo dell'editor, non produce un tratto visibile perché manca `stroke` con un colore). Ogni elemento è un `<path>` o `<rect>` con `fill` pieno. I "contorni" a 1px non sono linee: sono **forme piene sottratte/sovrapposte**. Da qui in avanti, ogni icona LeenO va costruita così: solo `fill`, mai `stroke` per il rendering finale.

### 12.2 Griglia e dimensioni canoniche
Il set di riferimento usa griglie diverse a seconda del contesto d'uso, non un'unica dimensione:
- **32×32**: icone principali di documento/toolbar grande (es. `print.svg`, `runbasic.svg`, `odb_32_8.svg`).
- **16×16**: icone di toolbar compatta (es. `sc_aligntop.svg`).
- **13×13** (con un piccolo offset di traslazione, es. `translate(.5 -.5)`): glifi piccolissimi come indicatori radio/stato.

Per LeenO: mantenere **24×24** come griglia di lavoro primaria (coerente con il resto del documento e con la UI di LibreOffice Calc), ma generare anche una variante **16×16** ottimizzata a parte per i casi di toolbar compatta, non un semplice ridimensionamento della 24×24 — a 16px i dettagli vanno semplificati (regola 12.4).

### 12.3 Spessori "a linea" reali osservati
Le barre sottili (es. le righe nell'icona `odb_32_8.svg`, i due rettangoli in `sc_aligntop.svg`) hanno altezza **1 unità su una griglia a 32** (~3% dell'altezza totale) o **1 unità su una griglia a 16** (~6% dell'altezza). Proporzionalmente su una griglia a 24, questo equivale a uno spessore di circa **0.75-1 unità**, non ai 2-3px "spessi" prodotti nel primo batch. Gli angoli di questi rettangoli sottili hanno un raggio (`rx`/`ry`) minimo, tra 0.4 e 0.5 unità su griglia 32 — un arrotondamento appena percettibile, non una pillola.

**Regola pratica per LeenO**: ridurre lo spessore massimo delle barre/linee da 3px (regola 2.5.A) a **1.5px su griglia 24×24**, con `rx` massimo 0.4px — le barre devono restare nettamente sottili anche affiancate.

### 12.4 Tecnica dell'anello (ring) per badge e cerchi
Osservata in `radiomono2.svg` e `runbasic.svg`: un cerchio con "bordo" è in realtà **due o tre cerchi pieni concentrici sovrapposti**, ciascuno più piccolo del precedente di circa 0.5-1 unità di raggio:
1. Cerchio esterno pieno, colore scuro/principale (raggio R).
2. Cerchio pieno colore chiaro/sfondo sopra, raggio ≈ R − 1 unità (crea l'anello).
3. Se serve un contenuto interno (icona play, pallino), cerchio o forma pieni colore accento, raggio ≈ R − 2/3 unità.

Questa è la tecnica da usare per **qualunque badge circolare o "sigillo"** in LeenO (es. `documento_bollo`), al posto di un cerchio pieno unico con motivo interno che genera confusione: outer scuro → inner chiaro (anello) → piccola forma accento centrale, senza croce/reticolo interno.

### 12.5 Forme con "buco" (fill-rule evenodd)
Quando una forma ha una parte cava (es. una cornice, una lettera con contro-forma), i file di riferimento usano `fill-rule="evenodd"` su un unico path con il contorno esterno e quello interno, invece di sovrapporre più shape separate. Jules può scegliere l'uno o l'altro metodo (path unico con evenodd, oppure due path sovrapposti di colore alternato) ma **non deve mai lasciare il bordo interno "aperto"** (cioè uno stroke che delimita senza un fill che lo chiude): è un'altra causa comune di forme che sembrano "sfrangiate".

### 12.6 Palette per ruolo, non per decorazione
Nei file reali ogni colore ha un ruolo fisso indipendentemente dall'icona: scuro quasi-nero per struttura/linee (`#3a3a38`, non nero puro), chiaro quasi-bianco per i cutout (`#fafafa`, non bianco puro), un solo accento semantico per icona (blu per azioni informative, verde per esecuzione/successo, viola per dati/database, grigio per elementi secondari/disabilitati). Per LeenO questo conferma ed estende la regola 9 già approvata: mantenere lo scuro e il chiaro leggermente "sporcati" (non nero/bianco puro) per evitare contrasto eccessivo che a icone piccole appare come un bordo troppo duro.

### 12.7 Checklist tecnica aggiuntiva (da unire alla sezione 11)
8. L'SVG finale contiene attributi `stroke` con un colore impostato? Se sì, l'icona **non è conforme** — va ricostruita solo con `fill`.
9. Ogni forma "sottile" (barra, linea) ha spessore ≤ 1.5px su griglia 24×24, non 2-3px?
10. Ogni badge/cerchio con bordo è costruito con almeno 2 cerchi pieni concentrici (anello), non un singolo cerchio con stroke?

---

## Nota sulle 4 icone del primo batch (pilota, non da correggere singolarmente)

Le 4 icone generate nel primo passaggio erano un **pilota tecnico**, utile per individuare i difetti, ma **non vanno corrette una per una**: vanno scartate e rigenerate da zero insieme a tutto il resto del set, seguendo la sezione 13 qui sotto, in modo che ogni icona (comprese queste 4) nasca fin dall'inizio con le regole 2.5/9/12 già applicate, invece di essere "aggiustata" sopra una base sbagliata.

- **`documento_bollo`**: il "timbro a cera" era un mirino/reticolo dentro un cerchio pieno arancione troppo grande. Da rifare con la tecnica dell'anello (12.4): scuro → chiaro → piccola forma accento, senza croce interna.
- **`esporta_gantt`**: barre troppo spesse (pillole) che si toccano con la freccia senza cuscinetto. Da rifare con barre ≤1.5px (12.3) e layering di fill al posto dello stroke.
- **`importa_xml`**: la X di "XML" si sovrappone alla freccia senza spazio. Da rifare con almeno 1px di distanza o freccia spostata fuori dal testo.
- **Icona con nastro/cerchio blu incrociato** (`somma_colore` o `unisci_fogli`): troppi elementi intrecciati (viola regola D, un solo elemento dominante). Da semplificare radicalmente.

---

## 13. Piano di Produzione Completa del Set — Da Zero

Lo scopo di questa sezione è che Jules produca **l'intero inventario di icone di LeenO** (non solo le 4 del pilota), applicando fin dal primo tracciato le regole geometriche (2.5), cromatiche (9) e tecniche (12) già stabilite. Nessuna icona del vecchio set (`image8.svg`, `image93.svg`, `sfera_gialla.svg`, ecc.) va riutilizzata, ricolorata o adattata: **si riparte da un tracciato vuoto per ciascuna**.

### 13.1 Regola di partenza pulita
- Ignorare completamente gli asset legacy (bitmap, nomi `imageNN`, sfere colorate, metafore letterali come `falegname` o `caschetto`) come *riferimento visivo*: contano solo come mappa nome-vecchio → nome-nuovo → significato (sezioni 6-8 della specifica), non come punto di partenza grafico.
- Ogni file nuovo usa il nome semantico snake_case definitivo (sezione 10.5), mai il nome legacy (`nuova_voce.svg`, non `image93.svg`).

### 13.2 Costruire prima le primitive condivise, poi le famiglie
Per evitare che 50 icone generate in sequenza divergano leggermente l'una dall'altra (il rischio principale generando un set così ampio), Jules deve procedere in quest'ordine:

1. **Definire una volta sole le 3 primitive di base** (Documento, Cartella, Elenco/WBS — sezione 2.A) con le regole tecniche di sezione 12 (solo fill, spessori 1.5px, raggi coerenti), come "componenti" di riferimento.
2. **Definire una volta sole le sovrapposizioni/badge standard** (Più, Meno/Elimina, Cerca, Avviso, Spunta — sezione 2.B) e le frecce (sezione 2.C), inclusa la tecnica dell'anello (12.4) per i badge circolari.
3. Solo dopo aver fissato questi "mattoncini", generare ogni icona della sezione 13.3 **componendo** questi stessi mattoncini, senza ridisegnarli ogni volta leggermente diversi.

Questo è l'equivalente grafico di riusare le stesse funzioni invece di riscrivere codice simile 50 volte: garantisce coerenza reale, non solo dichiarata.

### 13.3 Manifest completo delle icone da produrre

Produrre **tutte** le icone seguenti, raggruppate per famiglia semantica (sezione 6), più le 5 icone nuove (sezione 8). Per ciascuna: nome file definitivo, primitiva di base da riusare, badge/accento se presente.

**Categoria 1 — Principale e Navigazione**
- `leeno` — marchio piatto L+O intrecciate (Verde/Lime)
- `manuale` — Documento + lettera "i"
- `teleg` — aeroplano di carta, outline autonomo (nessuna primitiva documento)

**Categoria 2 — WBS**
- `supcat` — Cartella + numero romano "I"
- `cat` — Cartella + numero romano "II"
- `subcat` — Cartella + numero romano "III"
- `struttura_on` — Elenco/WBS + indicatore espansione
- `struttura_off` — Elenco/WBS + indicatore compressione
- `rinumCap` — Elenco/WBS + simbolo "#"

**Categoria 3 — Voci di Lavoro**
- `nuova_voce` — Documento + badge Più
- `voce_breve` — Documento + linea tratteggiata (forbici)
- `vedivoce` — Documento + occhio
- `pesca` — amo semplificato (autonomo)
- `invia_voce_ep` — Documento + freccia uscente destra
- `aggiungi_misura` — Documento + linea orizzontale + badge Più
- `sposta_voce` — due frecce verticali opposte (autonomo)

**Categoria 4 — Elenchi Prezzi**
- `analisi_a_prezzo` — due Documenti sovrapposti + percorso freccia
- `utili_maggiorazioni` — simbolo "%" (autonomo)
- `elimina_doppioni` — due Documenti sovrapposti + badge Elimina
- `riordina` — freccia verticale + lettere A/Z

**Categoria 5 — Quantità e Contabilità**
- `parz` — simbolo "∑" + parentesi
- `inverti_segno` — "+"/"-" affiancati + freccia commutazione
- `azzera` — cifra "0" grande (Arancione Azione)
- `partita_provvisoria_piu` — pila schede + badge Più
- `partita_provvisoria_meno` — pila schede + badge Meno
- `strutt_voci_zero` — Elenco/WBS + "Ø"
- `elimina_azzerate` — Documento + "0" + badge Elimina
- `elimina_vuote` — Elenco/WBS multi-riga + badge Elimina

**Categoria 6 — Layout, Fogli e Viste**
- `scelta_viste` — monitor diviso verticalmente (autonomo)
- `adattaH` — linea orizzontale + frecce esterne su/giù
- `mostra_griglia` — griglia 3×3 (autonomo)
- `copertine` — cartella ad anelli (autonomo)
- `colore_tematico` — secchio di vernice + goccia Lime

**Categoria 7 — Reporting e Stampa**
- `riepilogo` — Documento + linee + penna stilografica
- `riepilogo_quantita` — Documento + mini grafico a barre
- `riepilogo_a2` — Documento + griglia matrice
- `anteprima_stampa` — profilo stampante (autonomo)
- `riga_rossa` — evidenziatore + linea di chiusura orizzontale

**Categoria 8 — Utility e Configurazioni**
- `config` — due ingranaggi annidati (autonomo)
- `stringhe_numeri` — "A" + freccia + "1"
- `sproteggi_tutto` — lucchetto aperto (Giallo Avviso)
- `importa_stili` — pennello + scheda foglio
- `numeri_lettere` — "123" + fumetto + "abc"

**Categoria 9 — Sviluppatore e Import Legacy**
- `python_debug` — logo Python semplificato (autonomo)
- `refresh` — due frecce circolari a ciclo
- `importa_dat` — parentesi codice "&lt;&gt;" + freccia import

**Icone nuove (sezione 8 della specifica)**
- `importa_xml` — Documento + testo "XML" + freccia entrante (badge/freccia con cuscinetto, regola 12.4/2.5.C)
- `esporta_gantt` — barre Gantt (≤1.5px, regola 12.3) + freccia uscente destra
- `documento_bollo` — Documento + timbro con tecnica ad anello (12.4), senza croce interna
- `unisci_fogli` — due schede foglio che confluiscono in una (max 3 elementi, regola D)
- `somma_colore` — "∑" + profilo evidenziatore, un solo elemento dominante

Totale: **46 icone rinominate/riprogettate + 5 icone nuove = 51 icone**.

### 13.4 Struttura di consegna
- Generare ogni icona in **entrambi i temi**: `icons/svg/<nome>.svg` (chiaro) e `icons/scuro/<nome>.svg` (scuro), secondo la sezione 4.B.
- Generare anche la variante compatta **16×16** per le icone destinate a toolbar (priorità Alta nella tabella di sezione 7), come file separato ottimizzato (non uno scale-down automatico), secondo la regola 12.2.
- Nessun file con nome legacy (`imageNN.svg`, `sfera_gialla.svg`, ecc.) deve comparire nella consegna finale.

### 13.5 Validazione di insieme (oltre alla checklist per singola icona, sezioni 11+12.7)
Prima della consegna, Jules verifica l'intero set (non icona per icona) rispondendo SÌ a tutte queste domande:

1. Tutte le icone che usano la primitiva Documento hanno **esattamente** le stesse proporzioni/piega d'angolo (sezione 2.A.1), non varianti leggermente diverse tra loro?
2. Tutti i badge Più/Meno/Elimina hanno la stessa geometria in ogni icona in cui compaiono?
3. Nessuna icona usa `stroke` con colore impostato (regola 12.1) in tutto il set?
4. Il conteggio totale dei file consegnati corrisponde a 51 icone × 2 temi (+ variante 16×16 dove richiesta), senza mancanze né duplicati?
5. Scansionando l'intero set in sequenza (come una sprite sheet), le icone appartenenti alla stessa famiglia semantica sono immediatamente riconoscibili come "parenti" tra loro?

Se la risposta 1 o 2 è NO, il problema è quasi sempre che le primitive sono state ridisegnate a mano ogni volta invece di essere riusate come da regola 13.2 — tornare a quel passaggio prima di procedere oltre.
