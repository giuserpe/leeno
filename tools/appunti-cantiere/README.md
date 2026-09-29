# Brogliaccio per LeenO

PWA per raccogliere sul telefono gli appunti del giornale dei lavori e portarli in LeenO
(menu Importa/Esporta, "Importa appunti di cantiere nel Giornale Lavori...").
Non è un registro ufficiale: il giornale si consolida solo in LeenO.

- Formato di scambio: `documentazione/schemi/giornale_appunti.schema.json` (v1).
- I dati restano nel browser del dispositivo: esportare spesso.
- Più cantieri: ogni cantiere ha il proprio elenco di giornate, isolato dagli altri. "Esporta per LeenO" esporta solo il cantiere aperto, in un file con il suo nome nel nome del file.
- Tema: il pulsante in alto sceglie tra automatico (segue il sistema), chiaro e scuro. La scelta resta salvata sul dispositivo.
- Se una giornata ha il campo "Evento infortunistico" compilato, l'esportazione mostra un doppio avviso (informazione, poi conferma) sul rischio di condividere dati sulla salute con terzi.
- Stampa o PDF: genera una pagina A4 con le giornate compilate del cantiere aperto (solo i campi non vuoti), pensata per essere letta o salvata come PDF dal comando di stampa del browser. Lo stesso doppio avviso privacy vale anche qui. Non riproduce l'impaginazione ufficiale di LeenO: niente numerazione, niente firme.
- Solo file statici, nessun server. Per pubblicare basta servire questa cartella via HTTPS (es. GitHub Pages).
- Con una nuova versione: aggiornare `CACHE` in `sw.js` e `VERSIONE` in `core.js`.
- Le etichette dei campi in `core.js` (`CAMPI`) devono restare allineate a `LeenoGiornaleImport.ETICHETTE`.
- In fondo a ogni schermata c'è un link a https://leeno.org/donazioni/ (stesso indirizzo e testo "Dona!" del footer del sito).
- Grafica: colori e logo vengono dal tema del sito (`@SITO/leeno-theme`, variabili in `assets/css/main.css`, logo in `assets/images/logo-leeno.png`). Se la palette del sito cambia, aggiornare `:root` in `index.html` e rigenerare `logo.png` e le icone.
