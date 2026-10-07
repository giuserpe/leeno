/* Scrittore ZIP minimo, senza compressione (le foto JPEG sono già compresse).
   Usato solo per l'esportazione: nessuna libreria esterna, nessuna dipendenza di rete. */
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

  root.ZipStore = { creaZip: creaZip };
})(typeof window !== 'undefined' ? window : globalThis);
