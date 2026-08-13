# Lezioni apprese – Pulizia di codice morto e duplicato (agosto 2026)

Durante una pulizia sistematica di `pythonpath/` con `pyflakes` e analisi AST mirata sono emersi pattern ricorrenti, utili come checklist per le pulizie future.

## Gli script "usa e getta" in `pythonpath/` sono il rischio più grave, non solo disordine

Trovato `_fix_path.py`: a livello di modulo apriva `pyleeno.py` e lo riscriveva su disco sostituendo un range di righe hardcoded — se il modulo fosse mai stato importato dal processo di LibreOffice, avrebbe corrotto silenziosamente `pyleeno.py` con un range ormai disallineato. Stesso discorso per `benchmark.py`: path hardcoded (`W:\...`) e `print()` eseguiti a livello di modulo, quindi un `FileNotFoundError` per chiunque non abbia esattamente quel file. Regola operativa: uno script one-shot va eseguito ed eliminato subito dopo l'uso, mai lasciato in `pythonpath/`, nemmeno "per sicurezza".

## La redefinition nello stesso scope è un indicatore affidabile di codice morto

`python3 -m pyflakes <file>` segnala "redefinition of unused X from line Y" quando una funzione (o import) viene ridefinita nello stesso modulo/classe prima di essere usata: la prima definizione non è mai raggiungibile. Trovati due casi reali in `pyleeno.py` (`count_clipboard_lines`, `struttura_Registro`, entrambe con una versione più vecchia "morta" prima di quella attiva). Va trattato come codice da rimuovere, non come nota di stile.

## I moduli `LeenoImport_Xml*.py` vengono clonati l'uno dall'altro

Spesso portano con sé l'intero header di import del file sorgente, incluso il blocco `from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED)`, quasi mai usato per intero nel nuovo file. Un giro di `pyflakes` sul singolo modulo appena clonato individua questi import morti in pochi secondi — utile farlo subito dopo aver creato un nuovo import regionale, non solo in sede di pulizia generale.

## Prima di rimuovere una variabile "assegnata e mai usata" apparentemente inutile, controllare i moduli fratelli

Se lo stesso pattern (es. un campo estratto dall'XML ma non incluso nel titolo composto) si ripete identico in più moduli `LeenoImport_Xml*.py`, è quasi sempre una scelta di design ricorrente e non un refuso isolato — la rimozione va fatta comunque (il dato resta inutilizzato), ma senza trattarla come "correzione di un bug".
