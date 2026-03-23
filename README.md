LeenO - Computo metrico assistito con LibreOffice
=====
Estensione per LibreOffice basata su UltimusFree di Bartolomeo Aimar e
distribuita con licenza LGPL.

Cos’è LeenO
===========

LeenO è l’estensione per LibreOffice Calc specifica per la redazione di computi metrici e contabilità tecnica di cantiere. È un fork derivato da Ultimus che fu ideato e scritto da Bartolomeo Aimar, è coperto da Licenza LGPL  perciò se ne può fare un uso personale e professionale libero. E’ distribuito con nome LeenO.oxt.
LeenO lavora con LibreOffice, quindi sotto tutte le piattaforme su cui è possibile installare questa office suite.

LeenO, che eredita da LibreOffice tutte le sue funzionalità, ne completa l’offerta, proponendo all’ufficio tecnico un sistema integrato per la gestione delle progettazioni e degli appalti, a partire dalla organizzazione della documentazione a base di gara, fino alla gestione dell’appalto in fase di esecuzione dell’opera ed alla sua conclusione. Lavorando in ambiente Calc, LeenO consente un’ampia manovrabilità e personalizzazione degli elaborati.

LibreOffice genera e gestisce file in formato OpenDocument (ODF) che è lo Standard aperto ISO/IEC 26300:2015. LeenO, essendo un add-on per LibreOffice, aderisce perfettamente allo stesso standard. OpenDocument garantisce:

    interoperabilità senza barriere tecniche e legali anche tra sistemi operativi diversi,
    scambio dei dati corretto e sicuro,
    accesso agli stessi a lungo termine.

Le sue specifiche tecniche sono di pubblico dominio, per cui favorisce la concorrenza impedendo che queste siano detenute da un singolo produttore di software.

Per queste peculiarità LeenO è perfettamente in linea con le prescrizioni dettate dal comma 1. dell’art. 15 (Strumenti elettronici di contabilità e contabilità semplificata) del DM 7 marzo 2018 n.49 – Regolamento recante: «Approvazione delle linee guida sulle modalità di svolgimento delle funzioni del direttore dei lavori e del direttore dell’esecuzione».

LeenO è destinato ad utenti di LibreOffice Calc che vogliono guadagnare velocità nella compilazione dei documenti contabili tecnici senza perdere margini di controllo degli elaborati preferendo, proprio per questo, l’utilizzo di fogli di calcolo.

Installazione
=============

Per poter permettere un versionamento del codice sono disponibili due
script in python: bin2src.py e src2bin.py. Il primo permette di
estrarre i sorgenti in modo da poterli versionare ed il secondo di
archiviare i file sorgente in un nuovo ed aggiornato file di
estensione di LibreOffice (.oxt) su cui poter lavorare.

Una volta scaricato il sorgente è sufficiente lanciare dalla cartella
radice della repository il seguente comando per iniziare:

  $ src2bin.py
