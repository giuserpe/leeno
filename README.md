Ultimus
=======

Ultimus alias LEENO - Computo metrico assistito con Libreoffice


Cos’è UltimusFree

Fare architettura è divertente, la gestione della parte finanziaria molto meno!

Ma avere il controllo del budget, sia a monte che durante la gestione di un progetto, può rendere la cosa piacevole.

In commercio ci sono diversi programmi orientati allo scopo, ma li trovo poco flessibili e – in genere – specifici per un solo sistema operativo.

Per questo ho messo a punto UlltimusFree, un applicativo per LibreOffice (e/o OpenOffice.org) per la stesura e la gestione dei Computi Metrici Estimativi e della Contabilità Lavori.

UltimusFree può “girare” su piattaforme diverse (GNU/Linux, MS Windows e Mac) e i documenti contabili prodotti (normali tabelle di calcolo) possono essere aperti, manipolati e stampati anche con programmi diversi da LibreOffice (o da OpenOffice.org).

Ho cercato il massimo della flessibilità e della potenza, anche a scapito della semplicità… Il risultato sono una serie di tabelle collegate fra loro, manipolabili anche a mano attraverso l’interfaccia standard e senza utilizzare le macro.


Installazione
=============

Per poter permettere un versionamento del codice sono stati aggiunti due script
in python: bin2src.py e bin2src.py. Il primo permette di estrarre i sorgenti
in modo da poterli versionare ed il secondo di archiviare i file sorgente
in un nuovo ed aggiornato file di estensione di LibreOffice (.oxt) su cui poter
lavorare.

Una volta scaricato il sorgente è sufficiente lanciare dalla cartella radice
della repository il seguente comando per iniziare:
    $ src2bin.py