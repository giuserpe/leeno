LeenO - Computo metrico assistito con LibreOffice
=====

Cos’è LeenO
===========

Fare architettura è divertente, gestire la parte finanziaria molto meno!

Ma avere il controllo del budget, sia a monte che durante la gestione di un progetto, può rendere la cosa piacevole.

In commercio ci sono diversi programmi orientati allo scopo, ma li trovo poco flessibili e – in genere – specifici per un solo sistema operativo.

Per questo ho messo a punto UltimusFree (LeenO), un applicativo per LibreOffice per la stesura e la gestione dei Computi Metrici Estimativi e della Contabilità Lavori.

LeenO può “girare” su piattaforme diverse (GNU/Linux, MS Windows e Mac) e i documenti contabili prodotti (normali tabelle di calcolo) possono essere aperti, manipolati e stampati anche con programmi diversi da LibreOffice.

Ho cercato il massimo della flessibilità e della potenza, talvolta a scapito della semplicità... Il risultato è una serie di tabelle collegate fra loro, manipolabili anche a mano attraverso l’interfaccia standard di Calc senza utilizzare le macro.

LibreOffice genera e gestisce file in formato OpenDocument (ODF) ISO/IEC 26300:2006 che è l’unico formato standard aperto riconosciuto. LeenO, essendo un add-on per LibreOffice, aderisce perfettamente allo stesso standard.
OpenDocument, in quanto standard ISO, garantisce interoperabilità senza barriere tecniche e legali anche tra sistemi operativi diversi, assicurando la scambio dei dati corretto e sicuro oltre che l’accesso agli stessi a lungo termine. Le sue specifiche tecniche sono di pubblico dominio, per cui favorisce la concorrenza impedendo che dette specifiche siano detenute da un singolo produttore di software.

Installazione
=============

Per poter permettere un versionamento del codice sono disponibili due script
in python: bin2src.py e src2bin.py. Il primo permette di estrarre i sorgenti
in modo da poterli versionare ed il secondo di archiviare i file sorgente
in un nuovo ed aggiornato file di estensione di LibreOffice (.oxt) su cui poter
lavorare.

Una volta scaricato il sorgente è sufficiente lanciare dalla cartella radice
della repository il seguente comando per iniziare:
    $ src2bin.py
