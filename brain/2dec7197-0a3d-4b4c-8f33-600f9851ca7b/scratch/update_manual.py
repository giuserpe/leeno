import sys

p = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/documentazione/MANUALE_LeenO.fodt'
with open(p, 'r', encoding='utf-8') as f:
    text = f.read()

# We look for the paragraph about "Trova voci ricorrenti" to insert our new section after it
partial = 'Mostra una lista di codici di voce dalla quale'
pos = text.find(partial)

if pos != -1:
    # Find the full tag containing this text
    start = text.rfind('<text:p', 0, pos)
    end = text.find('</text:p>', pos) + 9
    target = text[start:end]
    
    new_content = (
        '\n   <text:h text:style-name="OOoHeading_20_4" text:outline-level="4">Somma per Colore nella Colonna</text:h>'
        '\n   <text:p text:style-name="P364">Questa funzionalità permette di generare rapidamente un totale basato sulla colorazione delle celle, facilitando il computo di voci raggruppate visivamente.</text:p>'
        '\n   <text:p text:style-name="P364">Per utilizzare il comando: 1. Selezionare una cella nella colonna d’interesse che abbia il colore di sfondo desiderato (ad esempio, una cella evidenziata in giallo); 2. Attivare la funzione: LeenO eseguirà una scansione automatica dell’intera colonna, individuando tutte le celle che presentano la medesima colorazione e che contengono valori numerici diversi da zero; 3. Il cursore si trasformerà in modalità selezione: è sufficiente fare clic sulla cella di destinazione dove si desidera visualizzare il risultato. LeenO inserirà automaticamente nella cella scelta una formula di somma contenente i riferimenti a tutte le celle colorate individuate e applicherà ad essa lo stesso colore di sfondo.</text:p>'
    )
    
    text = text.replace(target, target + new_content)
    
    with open(p, 'w', encoding='utf-8') as f:
        f.write(text)
    print("Manuale aggiornato con successo.")
else:
    print("Errore: Posizione per l'inserimento non trovata.")
