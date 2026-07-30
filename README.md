LeenO - Computo metrico assistito con LibreOffice
=================================================

[![Sito Ufficiale](https://img.shields.io/badge/Sito-LeenO.org-blue?style=flat-square)](https://leeno.org/)
[![Licenza](https://img.shields.io/badge/Licenza-LGPL%20v3-green?style=flat-square)](https://leeno.org/licenza-2/)
[![Telegram](https://img.shields.io/badge/Telegram-Group-blue?logo=telegram&style=flat-square)](https://t.me/leeno_computometrico)
[![Facebook](https://img.shields.io/badge/Facebook-Group-blue?logo=facebook&style=flat-square)](https://m.facebook.com/groups/433206393972197)
[![Donazioni](https://img.shields.io/badge/Sostienici-Dona-red?style=flat-square)](https://leeno.org/donazioni/)

**LeenO** è l’estensione open-source per LibreOffice Calc specifica per la redazione di computi metrici estimativi e contabilità tecnica di cantiere.

È un fork derivato da *Ultimus*, ideato e scritto da Bartolomeo Aimar, ed è coperto da Licenza LGPL v3. Può quindi essere utilizzato liberamente sia per scopi personali che professionali.

Visita il sito ufficiale: [https://leeno.org/](https://leeno.org/)

---

## Caratteristiche Principali

LeenO completa l’offerta di LibreOffice Calc proponendo all’ufficio tecnico un sistema integrato per la gestione delle progettazioni e degli appalti:

*   **Automazione Completa:** Genera computi metrici, contabilità di cantiere, varianti e libretti delle misure automaticamente. Ogni calcolo e operazione è gestito da macro integrate e altamente ottimizzate.
*   **Prezzari Regionali:** Accesso ai prezzari regionali aggiornati annualmente (es. Piemonte, Sardegna, Friuli Venezia Giulia, Umbria, Emilia Romagna, RFI, ecc.). Consente l'importazione rapida di elenchi prezzi di tutte le regioni ed enti.
*   **Conforme al DM 49/2018:** Rispetta perfettamente le prescrizioni dell'art. 15 (Strumenti elettronici di contabilità e contabilità semplificata) del DM 7 marzo 2018 n.49.
*   **Formato Aperto ed Interoperabile:** Lavorando in ambiente LibreOffice Calc, LeenO adotta lo standard aperto ISO/IEC 26300:2015 (OpenDocument Format - ODF), garantendo la massima interoperabilità tra diversi sistemi operativi, scambio corretto e sicuro dei dati e accessibilità a lungo termine.

---

## Sostieni il Progetto (Donazioni)

LeenO è gratuito per chi lo utilizza, ma richiede tempo, risorse e investimenti continui per essere mantenuto (server per i download, hosting, infrastruttura).

Se LeenO ti ha permesso di risparmiare ore di lavoro o costose licenze software annuali, considera di effettuare una donazione per assicurarne il futuro e lo sviluppo continuo:

*   **Dona ora:** [Sostieni LeenO con una donazione](https://leeno.org/donazioni/)

---

## Link Utili e Community

Rimani in contatto con il progetto LeenO e unisciti alla nostra community di professionisti:

*   **Sito Ufficiale:** [leeno.org](https://leeno.org/)
*   **Documentazione (Manuale, guide, tutorial):** [Documentazione LeenO](https://leeno.org/category/documentazione/)
*   **Forum di Supporto:** [Forum LeenO](https://leeno.org/forums/)
*   **Canale Telegram:** [Telegram LeenO](https://t.me/leeno_computometrico)
*   **Gruppo Facebook:** [Facebook LeenO](https://m.facebook.com/groups/433206393972197)
*   **Pagina LibreOffice Extension:** [Extensions LibreOffice - LeenO](https://extensions.libreoffice.org/extensions/leeno-2)

### Sviluppo & Codice Sorgente
*   **GitHub:** [github.com/giuserpe/leeno](https://github.com/giuserpe/leeno)
*   **GitLab:** [gitlab.com/giuserpe/leeno](https://gitlab.com/giuserpe/leeno)

---

## Installazione e Sviluppo

Per gli sviluppatori e coloro che vogliono compilare l'estensione a partire dai file sorgente:

È disponibile lo script `src2bin.py` per archiviare i file sorgente in un pacchetto di estensione LibreOffice (`.oxt`) aggiornato su cui poter lavorare.

Dalla cartella radice del repository, esegui il comando:

```bash
python3 src2bin.py
```

L'estensione generata verrà salvata all'interno della cartella `bin/`.
