/* ########################################################################
 * LeenO - Computo Metrico
 * Copyright (C) Giuseppe Vizziello - supporto@leeno.org
 * Licenza LGPL http://www.gnu.org/licenses/lgpl.html
 * ######################################################################## */
/* Archivio locale delle foto (IndexedDB) e preparazione immagini.
   Le foto restano solo su questo dispositivo, come il resto dei dati del Brogliaccio.
   La preparazione ridisegna l'immagine su un canvas: questo applica l'orientamento
   corretto e scarta ogni altro dato EXIF (inclusa la posizione GPS), che altrimenti
   verrebbe esportato insieme alla foto. */
(function (root) {
  'use strict';
  var DB_NOME = 'brogliaccio-foto', DB_VERSIONE = 1, STORE = 'foto';
  var LATO_MASSIMO = 1600, QUALITA_JPEG = 0.82;
  var dbPromise = null;

  function apriDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      if (!('indexedDB' in window)) { reject(new Error('Le foto non sono supportate su questo browser.')); return; }
      var richiesta = indexedDB.open(DB_NOME, DB_VERSIONE);
      richiesta.onupgradeneeded = function () {
        var db = richiesta.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: 'id' }).createIndex('cantiereId', 'cantiereId', { unique: false });
        }
      };
      richiesta.onsuccess = function () { resolve(richiesta.result); };
      richiesta.onerror = function () { reject(richiesta.error || new Error('Apertura archivio foto non riuscita.')); };
    });
    return dbPromise;
  }

  function transazione(modo) {
    return apriDb().then(function (db) { return db.transaction(STORE, modo).objectStore(STORE); });
  }

  function nuovoId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

  function aggiungi(cantiereId, giorno, blob) {
    var riga = { id: nuovoId(), cantiereId: cantiereId, giorno: giorno, blob: blob, creato_il: new Date().toISOString() };
    return transazione('readwrite').then(function (store) {
      return new Promise(function (resolve, reject) {
        var r = store.add(riga);
        r.onsuccess = function () { resolve(riga.id); };
        r.onerror = function () { reject(r.error); };
      });
    });
  }

  function elencaPerGiorno(cantiereId, giorno) {
    return transazione('readonly').then(function (store) {
      return new Promise(function (resolve, reject) {
        var risultati = [], cursore = store.index('cantiereId').openCursor(IDBKeyRange.only(cantiereId));
        cursore.onsuccess = function (e) {
          var c = e.target.result;
          if (!c) { resolve(risultati); return; }
          if (c.value.giorno === giorno) risultati.push(c.value);
          c.continue();
        };
        cursore.onerror = function () { reject(cursore.error); };
      });
    });
  }

  function elencaPerCantiere(cantiereId) {
    return transazione('readonly').then(function (store) {
      return new Promise(function (resolve, reject) {
        var risultati = [], cursore = store.index('cantiereId').openCursor(IDBKeyRange.only(cantiereId));
        cursore.onsuccess = function (e) {
          var c = e.target.result;
          if (!c) { resolve(risultati); return; }
          risultati.push(c.value); c.continue();
        };
        cursore.onerror = function () { reject(cursore.error); };
      });
    });
  }

  function elimina(id) {
    return transazione('readwrite').then(function (store) {
      return new Promise(function (resolve, reject) {
        var r = store.delete(id);
        r.onsuccess = function () { resolve(); };
        r.onerror = function () { reject(r.error); };
      });
    });
  }

  function eliminaDoveCursore(cantiereId, filtro) {
    return transazione('readwrite').then(function (store) {
      return new Promise(function (resolve, reject) {
        var cursore = store.index('cantiereId').openCursor(IDBKeyRange.only(cantiereId));
        cursore.onsuccess = function (e) {
          var c = e.target.result;
          if (!c) { resolve(); return; }
          if (filtro(c.value)) c.delete();
          c.continue();
        };
        cursore.onerror = function () { reject(cursore.error); };
      });
    });
  }

  function eliminaPerCantiere(cantiereId) { return eliminaDoveCursore(cantiereId, function () { return true; }); }
  function eliminaPerGiorno(cantiereId, giorno) {
    return eliminaDoveCursore(cantiereId, function (r) { return r.giorno === giorno; });
  }

  // Ridisegna il file su un canvas: corregge l'orientamento e rimuove ogni dato EXIF
  // (inclusa la posizione GPS), oltre a limitare il lato massimo per non appesantire l'archivio.
  function preparaImmagine(file) {
    return createImageBitmap(file, { imageOrientation: 'from-image' }).then(function (bitmap) {
      var scala = Math.min(1, LATO_MASSIMO / Math.max(bitmap.width, bitmap.height));
      var w = Math.max(1, Math.round(bitmap.width * scala)), h = Math.max(1, Math.round(bitmap.height * scala));
      var canvas = document.createElement('canvas');
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(bitmap, 0, 0, w, h);
      if (bitmap.close) bitmap.close();
      return new Promise(function (resolve, reject) {
        canvas.toBlob(function (b) { b ? resolve(b) : reject(new Error('Conversione della foto non riuscita.')); }, 'image/jpeg', QUALITA_JPEG);
      });
    });
  }

  root.FotoStore = { aggiungi: aggiungi, elencaPerGiorno: elencaPerGiorno, elencaPerCantiere: elencaPerCantiere,
    elimina: elimina, eliminaPerCantiere: eliminaPerCantiere, eliminaPerGiorno: eliminaPerGiorno, preparaImmagine: preparaImmagine };
})(typeof window !== 'undefined' ? window : globalThis);
