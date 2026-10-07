/* Brogliaccio per LeenO: logica pura (schema JSON v1), senza DOM. */
(function (root) {
  'use strict';
  var VERSIONE = '0.5.2';
  var AVVISO = 'Brogliaccio da consolidare in LeenO. Non costituisce registro ufficiale.';
  // chiave JSON, etichetta a video: stesso ordine del foglio GIORNALE
  var CAMPI = [
    ['meteo', 'Meteo'],
    ['presenti', 'Presenti/intervenuti'],
    ['annotazioni', 'Annotazioni, attività svolte'],
    ['operai', 'Qualifica e n. operai'],
    ['attrezzature', 'Attrezzature impiegate'],
    ['provviste', 'Provviste'],
    ['rifiuti', 'Rifiuto di materiali e/o manufatti'],
    ['disposizioni', 'Disposizioni e ordini di servizio del R.U.P. e del D.L.'],
    ['relazione_rup', 'Relazione indirizzata al R.U.P.'],
    ['verbali', 'Verbali di accertamento e prove'],
    ['contestazioni', 'Contestazioni, sospensioni e riprese lavori'],
    ['varianti', 'Varianti disposte, modifiche e/o aggiunte prezzi'],
    ['infortuni', 'Evento infortunistico'],
    ['osservazioni', 'Osservazioni, prescrizioni, avvertenze della D.L.']
  ];
  var METEO = ['Sereno.', 'Poco nuvoloso.', 'Nuvoloso.', 'Pioggia leggera.', 'Pioggia',
    'Pioggia intensa.', 'Temporale.', 'Grandine.', 'Nevicata leggera.', 'Nevicata.',
    'Nevicata intensa.', 'Nebbia.'];

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function oggiISO(d) { d = d || new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function dataValida(s) {
    if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
    var d = new Date(s + 'T00:00:00Z');
    return !isNaN(d) && d.toISOString().slice(0, 10) === s;
  }

  // giornate: { 'AAAA-MM-GG': { campi: {chiave: testo}, modificato_il: ISO } }
  function buildExport(giornate, adesso) {
    var out = { schema_version: '1', origine: 'appunti-mobile', avviso: AVVISO,
      generato_il: adesso.toISOString(), app_versione: VERSIONE, giornate: [] };
    Object.keys(giornate).sort().forEach(function (d) {
      var g = giornate[d], campi = {}, voce;
      CAMPI.forEach(function (c) {
        var t = g.campi && g.campi[c[0]];
        if (typeof t === 'string' && t.trim()) campi[c[0]] = t;
      });
      voce = { data: d, campi: campi };
      if (g.modificato_il) voce.modificato_il = g.modificato_il;
      out.giornate.push(voce);
    });
    return out;
  }

  // Valida un file esportato e restituisce { 'AAAA-MM-GG': {campi, modificato_il} }
  function parseImport(testo) {
    var dati, res = {}, noti = {};
    CAMPI.forEach(function (c) { noti[c[0]] = true; });
    try { dati = JSON.parse(testo); } catch (e) { throw new Error('File non leggibile.'); }
    if (!dati || dati.schema_version !== '1' || dati.origine !== 'appunti-mobile' || !Array.isArray(dati.giornate)) {
      throw new Error('Il file non è un export di questa app (schema v1).');
    }
    dati.giornate.forEach(function (g, i) {
      var campi = {};
      if (!g || !dataValida(g.data) || !g.campi || typeof g.campi !== 'object') {
        throw new Error('Giornata n. ' + (i + 1) + ': dati non validi.');
      }
      if (res[g.data]) throw new Error('Data duplicata: ' + g.data + '.');
      Object.keys(g.campi).forEach(function (k) {
        if (!noti[k]) return;
        if (typeof g.campi[k] !== 'string') throw new Error('Giornata ' + g.data + ': campo ' + k + ' non è testo.');
        if (g.campi[k].trim()) campi[k] = g.campi[k];
      });
      res[g.data] = { campi: campi, modificato_il: typeof g.modificato_il === 'string' ? g.modificato_il : new Date().toISOString() };
    });
    return res;
  }

  root.Core = { VERSIONE: VERSIONE, AVVISO: AVVISO, CAMPI: CAMPI, METEO: METEO,
    oggiISO: oggiISO, dataValida: dataValida, buildExport: buildExport, parseImport: parseImport };
  if (typeof module !== 'undefined') module.exports = root.Core;
})(typeof window !== 'undefined' ? window : globalThis);
