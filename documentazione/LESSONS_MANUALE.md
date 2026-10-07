# Lezioni apprese – Manuale utente (`documentazione/MANUALE_LeenO.fodt`)

Il manuale è un unico file Flat ODT (XML, oltre 3 MB, fine riga LF). La procedura operativa completa vive nella skill `.agent/skills/leeno-aggiorna-manuale/SKILL.md`; qui restano le lezioni che ricorrono.

## Modificare il FODT senza corromperlo

- Sostituzioni stringa-per-stringa in Python, con `open(..., encoding='utf-8', newline='')` per non tradurre i fine riga, e `assert testo.count(vecchio) == 1` prima di ogni `replace`: un target ambiguo o assente deve fermare lo script, non passare in silenzio.
- Dopo **ogni** serie di modifiche validare l'XML: `python3 -c "import xml.dom.minidom as m; m.parse('documentazione/MANUALE_LeenO.fodt')"`. Un tag non chiuso si scopre solo quando LibreOffice rifiuta il file.
- Stili ricorrenti del capitolo Brogliaccio, da riusare invece di inventarne di nuovi: `P440` (paragrafo di testo), `P437` (titolo livello 3, con bookmark `__RefHeading__<nome>`), `T838` (grassetto-corsivo per le etichette dell'interfaccia).
- Non lasciare nel repository script usa-e-getta per modificare il manuale (con percorsi `w:/...` fissi): vanno eseguiti ed eliminati, come per `pythonpath/`.

## Contenuto

- Le etichette dell'interfaccia vanno riportate esattamente come compaiono (stessa grafia, stessi due punti finali nelle etichette di Brogliaccio). Percorsi di menù da `Addons.xcu`, scorciatoie da `Accelerators.xcu`, non dalla memoria.
- Nessuna icona accanto ai nomi dei comandi (decisione del 2026-08-06, soggette ad aggiornamenti frequenti). Restano le icone "Attenzione" e le immagini illustrative.
- Terminologia di Brogliaccio: "agenda di cantiere" (non più "appunti"), "il Giornale dei Lavori".

## Screenshot e immagini

- Il manuale contiene circa 60 immagini incorporate (PNG in base64, `office:binary-data`) in un frame dentro una tabella a una cella, con didascalia a campo sequenza `Figure`. Per una nuova immagine si copia il blocco di una figura vicina, non se ne inventa la struttura.
- Solo le schermate di Brogliaccio si generano da script (`scripts/genera_screenshot_brogliaccio.py`, dati fittizi, versione visibile nell'intestazione) e vanno rigenerate a ogni modifica dell'interfaccia. Stili `PBrogFig`, `PBrogCap`, `frBrog`; larghezza 5 cm per le schermate da telefono, 12 cm per la stampa: a 6 cm l'immagine non entrava in pagina e lasciava circa il 40% di pagina vuota.
- Per dialoghi e fogli di LeenO gli screenshot li fornisce l'utente: l'agente li chiede con l'elenco preciso (dialogo, stato, dati fittizi) e non inserisce segnaposto né immagini ricostruite.
- Etichette sempre dalle sorgenti (`.xdl`, `.properties`, `Addons.xcu`): se lo screenshot differisce, si segnala la discrepanza.
- I rimandi alle figure sono riferimenti incrociati (`text:sequence-ref` con `text:reference-format="value"` e `text:ref-name` della didascalia): seguono la numerazione da soli. Un rimando digitato a mano può essere già obsoleto (trovato un "Figura 23" che doveva essere 26) o spezzato su più span ("Figura 2" + "3"): verificare sempre l'immagine di destinazione.
- Controllo obbligatorio dell'impaginazione: `pdftoppm -png -r 50 -f N -l M ...` sulle pagine interessate e verifica che figure e didascalie non lascino buchi.

## PDF

- `genera_pdf.py` usa LibreOffice headless: la dimensione del PDF varia con la versione di LibreOffice (osservato: 2,0 MB sul PC di sviluppo, 1,8 MB in ambiente cloud a parità di contenuto). Prima di distribuire un PDF generato altrove controllare impaginazione e immagini, o rigenerarlo sul PC `giuserpe`.
- Verifica minima: `pdftotext` sul PDF e ricerca di una frase appena inserita.

## Indice

L'indice del manuale è testo nel FODT. Se si aggiunge o rinomina un titolo verificare che la voce corrispondente nell'indice (bookmark `__RefHeading__...`) sia presente e aggiornata, oppure rigenerare gli indici da LibreOffice (Strumenti > Aggiorna > Tutti gli indici).
