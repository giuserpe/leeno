# Appunti di cantiere per LeenO

PWA per raccogliere sul telefono gli appunti del giornale dei lavori e portarli in LeenO
(menu Importa/Esporta, "Importa appunti di cantiere nel Giornale Lavori...").
Non è un registro ufficiale: il giornale si consolida solo in LeenO.

- Formato di scambio: `documentazione/schemi/giornale_appunti.schema.json` (v1).
- I dati restano nel browser del dispositivo: esportare spesso.
- Solo file statici, nessun server. Per pubblicare basta servire questa cartella via HTTPS (es. GitHub Pages).
- Con una nuova versione: aggiornare `CACHE` in `sw.js` e `VERSIONE` in `core.js`.
- Le etichette dei campi in `core.js` (`CAMPI`) devono restare allineate a `LeenoGiornaleImport.ETICHETTE`.
