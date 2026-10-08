/* Scrittore ZIP minimo, senza compressione (le foto JPEG sono già compresse), e lettore
   per il ripristino. Nessuna libreria esterna, nessuna dipendenza di rete. */
(function (root) {
  'use strict';

  function tabellaCrc() {
    var tabella = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      tabella[n] = c >>> 0;
    }
    return tabella;
  }
  var TABELLA_CRC = tabellaCrc();
  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = TABELLA_CRC[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function dataOraDos(d) {
    return {
      ora: ((d.getHours() & 0x1F) << 11) | ((d.getMinutes() & 0x3F) << 5) | ((d.getSeconds() >> 1) & 0x1F),
      data: (((Math.max(0, d.getFullYear() - 1980)) & 0x7F) << 9) | (((d.getMonth() + 1) & 0xF) << 5) | (d.getDate() & 0x1F)
    };
  }

  // voci: [{ percorso: 'cartella/file.jpg', dati: Uint8Array }] -> Blob ZIP
  function creaZip(voci) {
    var quando = dataOraDos(new Date());
    var parti = [], centrale = [], offset = 0;

    voci.forEach(function (voce) {
      var nome = new TextEncoder().encode(voce.percorso.replace(/\\/g, '/'));
      var dati = voce.dati, crc = crc32(dati);

      var header = new Uint8Array(30 + nome.length);
      var dv = new DataView(header.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0, true);
      dv.setUint16(8, 0, true);
      dv.setUint16(10, quando.ora, true);
      dv.setUint16(12, quando.data, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, dati.length, true);
      dv.setUint32(22, dati.length, true);
      dv.setUint16(26, nome.length, true);
      dv.setUint16(28, 0, true);
      header.set(nome, 30);
      parti.push(header, dati);

      var voceCentrale = new Uint8Array(46 + nome.length);
      var dvc = new DataView(voceCentrale.buffer);
      dvc.setUint32(0, 0x02014b50, true);
      dvc.setUint16(4, 20, true);
      dvc.setUint16(6, 20, true);
      dvc.setUint16(8, 0, true);
      dvc.setUint16(10, 0, true);
      dvc.setUint16(12, quando.ora, true);
      dvc.setUint16(14, quando.data, true);
      dvc.setUint32(16, crc, true);
      dvc.setUint32(20, dati.length, true);
      dvc.setUint32(24, dati.length, true);
      dvc.setUint16(28, nome.length, true);
      dvc.setUint16(30, 0, true);
      dvc.setUint16(32, 0, true);
      dvc.setUint16(34, 0, true);
      dvc.setUint16(36, 0, true);
      dvc.setUint32(38, 0, true);
      dvc.setUint32(42, offset, true);
      voceCentrale.set(nome, 46);
      centrale.push(voceCentrale);

      offset += header.length + dati.length;
    });

    var dimensioneCentrale = centrale.reduce(function (s, b) { return s + b.length; }, 0);
    var fine = new Uint8Array(22);
    var dvf = new DataView(fine.buffer);
    dvf.setUint32(0, 0x06054b50, true);
    dvf.setUint16(4, 0, true);
    dvf.setUint16(6, 0, true);
    dvf.setUint16(8, voci.length, true);
    dvf.setUint16(10, voci.length, true);
    dvf.setUint32(12, dimensioneCentrale, true);
    dvf.setUint32(16, offset, true);
    dvf.setUint16(20, 0, true);

    return new Blob(parti.concat(centrale, [fine]), { type: 'application/zip' });
  }

  // -------------------- lettura (per il ripristino da .zip) --------------------
  var MAX_VOCE = 64 * 1024 * 1024; // difesa da archivi malformati o enormi

  // Voce compressa "deflate" (metodo 8): si decomprime con il browser, senza librerie.
  function decomprimi(dati) {
    if (typeof DecompressionStream === 'undefined') {
      return Promise.reject(new Error('Questo browser non sa aprire file .zip compressi. Usa il file originale creato dall\'esportazione, senza ricomprimerlo.'));
    }
    var flusso = new Blob([dati]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Response(flusso).arrayBuffer().then(function (b) { return new Uint8Array(b); });
  }

  // buffer: ArrayBuffer di un .zip -> Promise di [{ percorso, dati: Uint8Array, crc }] (solo file, niente cartelle).
  // Legge la central directory, quindi accetta anche archivi creati da altri programmi.
  function leggiZip(buffer) {
    var buf = new Uint8Array(buffer), dv = new DataView(buffer), invalido = new Error('Il file non è un archivio .zip valido.');
    var fine = -1, i;
    for (i = buf.length - 22; i >= 0 && i >= buf.length - 22 - 65535; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { fine = i; break; }
    }
    if (fine < 0) return Promise.reject(invalido);
    var n = dv.getUint16(fine + 10, true), pos = dv.getUint32(fine + 16, true), elenco = [], k;
    for (k = 0; k < n; k++) {
      if (pos + 46 > buf.length || dv.getUint32(pos, true) !== 0x02014b50) return Promise.reject(invalido);
      var lNome = dv.getUint16(pos + 28, true), voce = {
        flag: dv.getUint16(pos + 8, true), metodo: dv.getUint16(pos + 10, true), crc: dv.getUint32(pos + 16, true),
        dimC: dv.getUint32(pos + 20, true), dimU: dv.getUint32(pos + 24, true), offset: dv.getUint32(pos + 42, true),
        percorso: new TextDecoder().decode(buf.subarray(pos + 46, pos + 46 + lNome)).replace(/\\/g, '/')
      };
      pos += 46 + lNome + dv.getUint16(pos + 30, true) + dv.getUint16(pos + 32, true);
      if (/\/$/.test(voce.percorso)) continue; // cartella
      if (voce.flag & 1) return Promise.reject(new Error('Il file .zip è protetto da password: non si può aprire.'));
      if (voce.dimC === 0xFFFFFFFF || voce.dimU === 0xFFFFFFFF || voce.offset === 0xFFFFFFFF || voce.dimU > MAX_VOCE) {
        return Promise.reject(new Error('Il file .zip contiene voci troppo grandi per essere aperte qui.'));
      }
      elenco.push(voce);
    }
    var risultati = [];
    return elenco.reduce(function (prec, voce) {
      return prec.then(function () {
        var o = voce.offset;
        if (o + 30 > buf.length || dv.getUint32(o, true) !== 0x04034b50) throw invalido;
        var inizio = o + 30 + dv.getUint16(o + 26, true) + dv.getUint16(o + 28, true);
        if (inizio + voce.dimC > buf.length) throw invalido;
        var grezzo = buf.subarray(inizio, inizio + voce.dimC);
        if (voce.metodo === 0) return grezzo;
        if (voce.metodo === 8) return decomprimi(grezzo);
        throw new Error('Il file .zip usa un metodo di compressione non supportato (' + voce.metodo + ').');
      }).then(function (dati) {
        if (dati.length !== voce.dimU || crc32(dati) !== voce.crc) {
          throw new Error('Il file .zip è danneggiato (voce "' + voce.percorso + '" non integra).');
        }
        risultati.push({ percorso: voce.percorso, dati: dati, crc: voce.crc });
      });
    }, Promise.resolve()).then(function () { return risultati; });
  }

  root.ZipStore = { creaZip: creaZip, leggiZip: leggiZip, crc32: crc32 };
})(typeof window !== 'undefined' ? window : globalThis);
