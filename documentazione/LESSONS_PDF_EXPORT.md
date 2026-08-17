# Lezioni apprese – Export PDF

Riguarda `SheetUtils.pdfExport()` e `pyleeno.ods2pdf()`, i due percorsi di export PDF di LeenO.

## FilterData annidato va tipizzato esplicitamente

`FilterData` (es. per `SelectPdfVersion=2`, PDF/A-2b) è una `PropertyValue` il cui `Value` è a sua volta una sequenza di `PropertyValue`. Un tuple Python passato così com'è viene marshallato da PyUNO come `sequence<any>` generico, non `sequence<PropertyValue>`: il filtro `calc_pdf_Export` lo ignora in silenzio, senza eccezioni, e produce comunque un PDF ma non PDF/A. Va tipizzato esplicitamente con `uno.Any('[]com.sun.star.beans.PropertyValue', (...))` (o `LeenoUtils.dictToProperties(..., unoAny=True)`).

## `PrintAnnotations` non nasconde l'indicatore della nota

La proprietà del page style `PrintAnnotations` controlla solo l'elenco testuale delle note in coda al documento stampato. Non ha alcun effetto sull'icona/indicatore visivo che Calc disegna sulla cella per qualunque nota presente (anche non impostata come "sempre visibile"): quell'icona viene comunque inclusa nel rendering di stampa/export PDF. Per escluderla davvero:

- su un documento temporaneo usa-e-getta (come `nDoc` in `pdfExport()`): rimuovere le annotazioni con `Annotations.removeByIndex()` prima dell'export, nessun ripristino necessario.
- su un documento live (come `oDoc` in `ods2pdf()`): salvare `(nome foglio, Position, Text.getString())` di ogni annotazione, rimuoverle, esportare, poi reinserirle identiche in un blocco `finally`. Il reinserimento non preserva formattazione ricca del testo né autore/timestamp originali — accettabile per note editoriali in testo semplice come quelle di `vedi_voce_xpwe()`.
