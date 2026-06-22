import sys

filepath = 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/documentazione/MANUALE_LeenO.fodt'

with open(filepath, 'r', encoding='utf-8') as f:
    text = f.read()

target = """   <text:p text:style-name="P364"><text:soft-page-break/>Per utilizzare il comando: 1. Selezionare una cella nella colonna d’interesse che abbia il colore di sfondo desiderato (ad esempio, una cella evidenziata in giallo); 2. Attivare la funzione: LeenO eseguirà una scansione automatica dell’intera colonna, individuando tutte le celle che presentano la medesima colorazione e che contengono valori numerici diversi da zero; 3. Il cursore si trasformerà in modalità selezione: è sufficiente fare clic sulla cella di destinazione dove si desidera visualizzare il risultato. LeenO inserirà automaticamente nella cella scelta una formula di somma contenente i riferimenti a tutte le celle colorate individuate e applicherà ad essa lo stesso colore di sfondo.</text:p>"""

replacement = target + """
   <text:h text:style-name="OOoHeading_20_4" text:outline-level="4"><text:bookmark-start text:name="__RefHeading___Riepilogo_importi"/>Riepilogo importi A2 cartella<text:bookmark-end text:name="__RefHeading___Riepilogo_importi"/></text:h>
   <text:p text:style-name="P364">Questa funzionalità, accessibile dal menù <text:span text:style-name="T897">LeenO &gt; UTILITY...</text:span>, permette di ottenere un riepilogo rapido degli importi di tutti i file di LeenO presenti in una determinata cartella.</text:p>
   <text:p text:style-name="P364">Selezionando la cartella desiderata tramite l'apposita finestra di dialogo, LeenO analizzerà tutti i file ODS contenuti al suo interno. Verrà quindi mostrata una finestra riepilogativa contenente i valori presenti nella cella A2 dei fogli COMPUTO, VARIANTE e CONTABILITA per ogni file. Cliccando su una delle voci nell'elenco, sarà possibile aprire direttamente il file corrispondente per la consultazione.</text:p>"""

if target in text:
    text = text.replace(target, replacement)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(text)
    print("Sostituzione completata con successo nel MANUALE_LeenO.fodt")
else:
    print("ERRORE: Testo target non trovato!")
