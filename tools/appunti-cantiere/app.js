/* Appunti di cantiere per LeenO: interfaccia. Dati solo su questo dispositivo. */
(function () {
  'use strict';
  var C = window.Core, KEY = 'appunti-cantiere-v1';
  var app = document.getElementById('app'), msg = document.getElementById('msg');
  var stato = carica();

  function carica() {
    try { var s = JSON.parse(localStorage.getItem(KEY)); if (s && s.giornate) return s; } catch (e) { /* stato vuoto */ }
    return { giornate: {}, ultimo_export: null };
  }
  function salva() {
    try { localStorage.setItem(KEY, JSON.stringify(stato)); return true; }
    catch (e) { avviso('Salvataggio non riuscito: memoria piena o bloccata. Esporta subito i dati.'); return false; }
  }
  function giornate(n) { return n + (n === 1 ? ' giornata' : ' giornate'); }
  function avviso(t) { msg.textContent = t; msg.hidden = !t; }
  function h(tag, attr, figli) {
    var e = document.createElement(tag), k;
    for (k in (attr || {})) { if (k === 'on') { Object.keys(attr.on).forEach(function (ev) { e.addEventListener(ev, attr.on[ev]); }); } else if (k === 'text') { e.textContent = attr.text; } else { e.setAttribute(k, attr[k]); } }
    (figli || []).forEach(function (f) { if (f) e.appendChild(f); });
    return e;
  }
  function btn(testo, fn, cls) { return h('button', { type: 'button', 'class': cls || '', text: testo, on: { click: fn } }); }
  function dataEstesa(iso) {
    return new Date(iso + 'T12:00:00').toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  }
  function svuota() { while (app.firstChild) app.removeChild(app.firstChild); avviso(''); window.scrollTo(0, 0); }

  function elenco() {
    svuota();
    var date = Object.keys(stato.giornate).sort().reverse();
    var oggi = h('input', { type: 'date', id: 'nuova', value: C.oggiISO() });
    var modificate = date.filter(function (d) { return !stato.ultimo_export || stato.giornate[d].modificato_il > stato.ultimo_export; }).length;
    app.appendChild(h('section', {}, [
      h('label', { 'for': 'nuova', text: 'Giornata' }), oggi,
      btn('Apri o crea giornata', function () { if (C.dataValida(oggi.value)) modifica(oggi.value); }, 'primary')
    ]));
    app.appendChild(h('ul', { 'class': 'lista' }, date.map(function (d) {
      var a = stato.giornate[d].campi.annotazioni || stato.giornate[d].campi.meteo || '';
      return h('li', {}, [h('button', { type: 'button', on: { click: function () { modifica(d); } } }, [
        h('strong', { text: dataEstesa(d) }), h('span', { text: a.slice(0, 80) })])]);
    })));
    if (!date.length) app.appendChild(h('p', { 'class': 'nota', text: 'Nessuna giornata salvata.' }));
    app.appendChild(h('section', {}, [
      h('p', { 'class': modificate ? 'nota alert' : 'nota', text: !date.length ? '' :
        (giornate(modificate) + (modificate === 1 ? ' non ancora esportata' : ' non ancora esportate') + '. I dati esistono solo su questo dispositivo.') }),
      btn('Esporta per LeenO', esporta, 'primary'),
      btn('Ripristina da file', function () { document.getElementById('file').click(); })
    ]));
  }

  function modifica(iso) {
    svuota();
    var g = stato.giornate[iso] || { campi: {} };
    app.appendChild(h('div', { 'class': 'barra' }, [btn('Indietro', elenco), h('h2', { text: dataEstesa(iso) })]));
    C.CAMPI.forEach(function (c) {
      var id = 'c_' + c[0], el;
      if (c[0] === 'meteo') {
        el = h('input', { id: id, type: 'text', list: 'meteo', autocomplete: 'off' });
      } else {
        el = h('textarea', { id: id, rows: c[0] === 'annotazioni' ? '6' : '3' });
      }
      el.value = g.campi[c[0]] || '';
      el.addEventListener('input', function () {
        g.campi[c[0]] = el.value; g.modificato_il = new Date().toISOString();
        stato.giornate[iso] = g; salva();
      });
      app.appendChild(h('div', { 'class': 'campo' }, [h('label', { 'for': id, text: c[1] }), el]));
    });
    app.appendChild(btn('Elimina questa giornata', function () {
      if (confirm('Eliminare la giornata ' + dataEstesa(iso) + '?')) { delete stato.giornate[iso]; salva(); elenco(); }
    }, 'danger'));
  }

  function esporta() {
    var dati = C.buildExport(stato.giornate, new Date());
    if (!dati.giornate.length) { avviso('Nessuna giornata da esportare.'); return; }
    var nome = 'appunti-cantiere-' + C.oggiISO().replace(/-/g, '') + '.json';
    var file = new File([JSON.stringify(dati, null, 2)], nome, { type: 'application/json' });
    function fatto() { stato.ultimo_export = new Date().toISOString(); salva(); elenco(); avviso('Esportate ' + giornate(dati.giornate.length) + ': ' + nome); }
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      navigator.share({ files: [file], title: 'Appunti di cantiere' }).then(fatto, function (e) { if (e.name !== 'AbortError') avviso('Condivisione non riuscita.'); });
    } else {
      var a = h('a', { href: URL.createObjectURL(file), download: nome }); document.body.appendChild(a); a.click(); a.remove(); fatto();
    }
  }

  document.getElementById('file').addEventListener('change', function (ev) {
    var f = ev.target.files[0]; ev.target.value = '';
    if (!f) return;
    f.text().then(function (t) {
      var nuove = C.parseImport(t), n = Object.keys(nuove), doppie = n.filter(function (d) { return stato.giornate[d]; }).length;
      if (!confirm('Ripristinare ' + giornate(n.length) + '? ' + doppie + ' già presenti su questo dispositivo verranno sostituite.')) return;
      n.forEach(function (d) { stato.giornate[d] = nuove[d]; });
      salva(); elenco();
    }).catch(function (e) { avviso(e.message); });
  });

  var lista = document.getElementById('meteo');
  C.METEO.forEach(function (m) { lista.appendChild(h('option', { value: m })); });
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
  elenco();
})();
