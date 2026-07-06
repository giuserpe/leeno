# Regole del Progetto LeenO

- Quando scrivi o modifichi codice, dai sempre la priorità assoluta alle API UNO di LibreOffice/OpenOffice rispetto a librerie esterne o macro standard basate su altri paradigmi. Utilizza i binding corretti (es. Python uno, unohelper) e rispetta le convenzioni del modello a oggetti UNO.

- Per i task di programmazione, utilizza sempre Python come linguaggio preferenziale, a meno di esplicita indicazione contraria.

- Quando devi manipolare o analizzare file di testo di grandi dimensioni, preferisci sempre l'utilizzo di librerie specializzate (come `pandas` per dati strutturati o `re`/`regex` per pattern) per ottenere prestazioni migliori, piuttosto che l'analisi manuale tramite stringhe o cicli in linguaggio naturale.

- Preferisci sempre l'utilizzo di procedure batch (elaborazioni in blocco) per migliorare le prestazioni e ridurre i tempi di esecuzione, specialmente quando si interagisce con il documento.
