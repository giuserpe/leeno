/* Copyright (C) Giuseppe Vizziello - supporto@leeno.org. Licenza LGPL 2.1 (vedi LICENSE nella radice del repository). */
/* Brogliaccio per LeenO: interfaccia. Dati solo su questo dispositivo. */
(function () {
  'use strict';
  var C = window.Core, KEY = 'appunti-cantiere-v1', TEMA_KEY = 'appunti-cantiere-tema';
  var cantiereApertoId = null, dataAperta = null, urlFotoAttive = [];
  var cantiereApertoRif = null, giornoApertoRif = null;
  // Stato di salvaguardia: DIAG_KEY e' una chiave nuova (non rinomina nulla di distribuito).
  var DIAG_KEY = 'brogliaccio-diag', diag = null, allarme = null, memoriaProtetta, inModifica = false;
  var app = document.getElementById('app'), msg = document.getElementById('msg'), pannello = document.getElementById('pannello-export');
  var stato = carica();

  // -------------------- tema: automatico, chiaro o scuro, a scelta --------------------
  var ORDINE_TEMI = ['auto', 'chiaro', 'scuro'];
  var ETICHETTE_TEMA = { auto: 'Tema: automatico', chiaro: 'Tema: chiaro', scuro: 'Tema: scuro' };
  var bottoneTema = document.getElementById('tema');
  function temaSalvato() {
    try { var t = localStorage.getItem(TEMA_KEY); return ORDINE_TEMI.indexOf(t) >= 0 ? t : 'auto'; }
    catch (e) { return 'auto'; }
  }
  function applicaTema(t) {
    if (t === 'auto') document.documentElement.removeAttribute('data-tema');
    else document.documentElement.setAttribute('data-tema', t);
    bottoneTema.textContent = ETICHETTE_TEMA[t];
    bottoneTema.setAttribute('aria-label', ETICHETTE_TEMA[t] + '. Tocca per cambiare.');
  }
  bottoneTema.addEventListener('click', function () {
    var attuale = temaSalvato(), successivo = ORDINE_TEMI[(ORDINE_TEMI.indexOf(attuale) + 1) % ORDINE_TEMI.length];
    try { localStorage.setItem(TEMA_KEY, successivo); } catch (e) { /* il tema resta solo per questa sessione */ }
    applicaTema(successivo);
  });
  applicaTema(temaSalvato());

  function nuovoId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  function carica() {
    try {
      var s = JSON.parse(localStorage.getItem(KEY));
      if (s && s.cantieri) { if (typeof s.rev !== 'number') s.rev = 0; return s; }
    } catch (e) { /* prosegue con i tentativi di recupero sotto */ }

    // Recupero: durante lo sviluppo il nome del "cassetto" di memoria e' cambiato
    // due volte. Se qui dentro c'e' ancora qualcosa di riconoscibile, lo si
    // riporta al formato attuale invece di ripartire da zero in silenzio.
    var recuperato = null;
    try {
      var v2 = JSON.parse(localStorage.getItem('appunti-cantiere-v2'));
      if (v2 && v2.cantieri) recuperato = v2; // formato "a piu' cantieri" salvato sotto il nome intermedio
    } catch (e) { /* niente da recuperare da qui */ }
    if (!recuperato) {
      try {
        var v1piatto = JSON.parse(localStorage.getItem(KEY));
        if (v1piatto && v1piatto.giornate && Object.keys(v1piatto.giornate).length) {
          var id = nuovoId();
          recuperato = { cantieri: {}, attivo: id };
          recuperato.cantieri[id] = { nome: 'Cantiere recuperato', giornate: v1piatto.giornate,
            ultimo_export: v1piatto.ultimo_export || null }; // formato a un solo cantiere, precedente a questo
        }
      } catch (e) { /* niente da recuperare da qui */ }
    }
    if (recuperato) {
      try { localStorage.setItem(KEY, JSON.stringify(recuperato)); } catch (e) { /* si tenta comunque di usarlo in questa sessione */ }
      return recuperato;
    }
    return { cantieri: {}, attivo: null, rev: 0 };
  }

  // Revisione: ogni salvataggio incrementa stato.rev. Se nella memoria c'e' una revisione piu' alta della propria,
  // un'altra istanza dell'app (scheda del browser, app installata, pagina rimasta in pausa) ha scritto dopo il
  // nostro ultimo caricamento: prima di scrivere si portano in memoria i suoi dati, per non cancellarli.
  function leggiSalvato() {
    try { var s = JSON.parse(localStorage.getItem(KEY)); return (s && s.cantieri) ? s : null; }
    catch (e) { return null; }
  }
  // Solo aggiunte e giornate piu' recenti, mai cancellazioni. Restituisce true se qualcosa e' cambiato.
  function unisciSalvato() {
    var s = leggiSalvato();
    if (!s || !(s.rev > (stato.rev || 0))) return false;
    var cambiato = false;
    Object.keys(s.cantieri).forEach(function (id) {
      var sc = s.cantieri[id], c = stato.cantieri[id];
      if (!c) { stato.cantieri[id] = sc; cambiato = true; return; }
      Object.keys(sc.giornate || {}).forEach(function (d) {
        var sg = sc.giornate[d], g = c.giornate[d];
        var aperta = inModifica && id === cantiereApertoId && d === dataAperta; // la giornata in modifica resta com'e'
        if (!g || (!aperta && (sg.modificato_il || '') > (g.modificato_il || ''))) { c.giornate[d] = sg; cambiato = true; }
      });
      if ((sc.ultimo_export || '') > (c.ultimo_export || '')) { c.ultimo_export = sc.ultimo_export; cambiato = true; }
    });
    stato.rev = s.rev;
    return cambiato;
  }
  function contaDati() {
    var g = 0, ids = Object.keys(stato.cantieri);
    ids.forEach(function (id) { g += Object.keys(stato.cantieri[id].giornate || {}).length; });
    return { cantieri: ids.length, giornate: g };
  }
  function leggiDiag() {
    try { var d = JSON.parse(localStorage.getItem(DIAG_KEY)); return (d && typeof d === 'object') ? d : null; }
    catch (e) { return null; }
  }
  function scriviDiag() { try { if (diag) localStorage.setItem(DIAG_KEY, JSON.stringify(diag)); } catch (e) { /* diagnostica facoltativa */ } }
  function salva() {
    try {
      unisciSalvato();
      stato.rev = (stato.rev || 0) + 1;
      localStorage.setItem(KEY, JSON.stringify(stato));
      if (diag) { diag.ultimo_stato = contaDati(); scriviDiag(); } // l'ultimo stato noto permette di accorgersi di cancellazioni esterne
      return true;
    } catch (e) { avviso('Salvataggio non riuscito: memoria piena o bloccata. Esporta subito i dati.'); return false; }
  }
  // Un'altra istanza ha scritto: si aggiorna lo stato e, se non c'e' una giornata in modifica, la schermata.
  function aggiornaDaAltrove() {
    if (!unisciSalvato()) return;
    if (inModifica) avviso('I dati sono stati aggiornati da un\'altra finestra dell\'app. Torna indietro per vederli.');
    else principale();
  }
  function verificaPersistenza() {
    var s = navigator.storage;
    if (!s || !s.persisted) return Promise.resolve(null);
    return s.persisted().then(function (p) { return p ? true : (s.persist ? s.persist() : false); })
      .catch(function () { return null; });
  }
  function testoMemoria() {
    var avvio = diag ? ' Avvio n. ' + diag.avvii + '.' : '';
    if (memoriaProtetta === undefined) return 'Memoria del browser: verifica in corso.' + avvio;
    if (memoriaProtetta === true) return 'Memoria del browser: protetta.' + avvio;
    if (memoriaProtetta === false) return 'Memoria del browser: non protetta, il browser può cancellare questi dati. Esporta spesso.' + avvio;
    return 'Memoria del browser: protezione non verificabile. Esporta spesso.' + avvio;
  }
  function aggiornaStatoMemoria() {
    var el = document.getElementById('stato-memoria');
    if (!el) return;
    el.textContent = testoMemoria();
    el.className = memoriaProtetta === false || memoriaProtetta === null ? 'nota alert' : 'nota';
  }
  function statoMemoria() { return h('p', { id: 'stato-memoria', 'class': memoriaProtetta === false || memoriaProtetta === null ? 'nota alert' : 'nota', text: testoMemoria() }); }
  // Giornate mai esportate o modificate (testo o foto) dopo l'ultimo export del cantiere.
  function giornateDaEsportare(c) {
    return Object.keys(c.giornate).filter(function (d) {
      return !c.ultimo_export || c.giornate[d].modificato_il > c.ultimo_export;
    }).length;
  }
  function ultimoExportTesto(c) {
    var t = c.ultimo_export ? new Date(c.ultimo_export) : null;
    if (!t || isNaN(t.getTime())) return 'Mai esportato.';
    var ora = new Date();
    var n = Math.round((new Date(ora.getFullYear(), ora.getMonth(), ora.getDate())
      - new Date(t.getFullYear(), t.getMonth(), t.getDate())) / 86400000);
    if (n <= 0) return 'Ultimo export: oggi.';
    return 'Ultimo export: ' + (n === 1 ? 'ieri' : n + ' giorni fa') + '.';
  }
  function avviso(t) { msg.textContent = t; msg.hidden = !t; }
  function giornate(n) { return n + (n === 1 ? ' giornata' : ' giornate'); }
  function cantieri(n) { return n + (n === 1 ? ' cantiere' : ' cantieri'); }
  function h(tag, attr, figli) {
    var e = document.createElement(tag), k;
    for (k in (attr || {})) { if (k === 'on') { Object.keys(attr.on).forEach(function (ev) { e.addEventListener(ev, attr.on[ev]); }); } else if (k === 'text') { e.textContent = attr.text; } else { e.setAttribute(k, attr[k]); } }
    (figli || []).forEach(function (f) { if (f) e.appendChild(f); });
    return e;
  }
  function btn(testo, fn, cls) { return h('button', { type: 'button', 'class': cls || '', text: testo, on: { click: fn } }); }
  // AAAAMMGGhhmm dall'istante di scatto/salvataggio di una foto (creato_il), ora locale.
  function marcaTemporale(iso) {
    var d = new Date(iso);
    function due(n) { return (n < 10 ? '0' : '') + n; }
    return '' + d.getFullYear() + due(d.getMonth() + 1) + due(d.getDate()) + due(d.getHours()) + due(d.getMinutes());
  }

  function dataEstesa(iso) {
    return new Date(iso + 'T12:00:00').toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  }
  function slug(nome) {
    return nome.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase()
      .replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'cantiere';
  }
  function svuota() {
    urlFotoAttive.forEach(function (u) { URL.revokeObjectURL(u); }); urlFotoAttive = [];
    while (app.firstChild) app.removeChild(app.firstChild); avviso(''); window.scrollTo(0, 0);
  }
  function cantiereCorrente() { return stato.attivo ? stato.cantieri[stato.attivo] : null; }

  var MESI = ['GEN', 'FEB', 'MAR', 'APR', 'MAG', 'GIU', 'LUG', 'AGO', 'SET', 'OTT', 'NOV', 'DIC'];
  var GIORNI = ['Dom', 'Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab'];

  // Testo di anteprima di una giornata per l'elenco: meteo piu' il primo altro campo
  // compilato, oppure un avviso esplicito se non c'e' ancora nulla (giornata con sole foto compresa).
  function anteprimaGiorno(g) {
    var pezzi = [];
    if (g.campi.meteo) pezzi.push(g.campi.meteo);
    for (var i = 0; i < C.CAMPI.length; i++) {
      var chiave = C.CAMPI[i][0];
      if (chiave !== 'meteo' && g.campi[chiave]) { pezzi.push(g.campi[chiave]); break; }
    }
    var testo = pezzi.join(' · ');
    if (!testo) return 'Nessuna lavorazione scritta';
    return testo.length > 70 ? testo.slice(0, 70) + '…' : testo;
  }

  // -------------------- schermata unica: cantiere (scelta/creazione) + sue giornate --------------------
  function principale() {
    chiudiPannello();
    svuota();
    inModifica = false;
    if (allarme) {
      app.appendChild(h('section', {}, [h('p', { 'class': 'nota alert', text: allarme }),
        btn('Ho capito', function () { allarme = null; principale(); })]));
    }
    var ids = Object.keys(stato.cantieri).sort(function (a, b) {
      return stato.cantieri[a].nome.localeCompare(stato.cantieri[b].nome, 'it');
    });
    if (stato.attivo && !stato.cantieri[stato.attivo]) stato.attivo = null;
    if (!stato.attivo && ids.length) stato.attivo = ids[0];

    var campi = [];
    if (ids.length) {
      var selettore = h('select', { id: 'sel-cantiere' });
      ids.forEach(function (id) {
        var opz = h('option', { value: id, text: stato.cantieri[id].nome });
        if (id === stato.attivo) opz.setAttribute('selected', 'selected');
        selettore.appendChild(opz);
      });
      selettore.addEventListener('change', function () { stato.attivo = selettore.value; salva(); principale(); });
      campi.push(h('div', { 'class': 'campo' }, [h('label', { 'for': 'sel-cantiere', text: 'Cantiere:' }), selettore]));
    }

    var nuovoNome = h('input', { type: 'text', id: 'nuovo-cantiere', placeholder: 'Scrivi il nome del nuovo cantiere', autocomplete: 'off' });
    function creaDaCampo() {
      var nome = nuovoNome.value.trim();
      if (!nome) return;
      var id = nuovoId();
      stato.cantieri[id] = { nome: nome, giornate: {}, ultimo_export: null };
      stato.attivo = id; salva(); principale();
    }
    nuovoNome.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') creaDaCampo(); });
    campi.push(h('div', { 'class': 'campo' }, [
      h('label', { 'for': 'nuovo-cantiere', text: ids.length ? 'Aggiungi nuovo cantiere:' : 'Nome del cantiere:' }), nuovoNome
    ]));
    campi.push(btn('Crea cantiere', creaDaCampo, ids.length ? '' : 'primary'));
    app.appendChild(h('section', {}, campi));

    var c = cantiereCorrente();
    if (!c) {
      app.appendChild(h('p', { 'class': 'nota', text: 'Scrivi un nome e crea il primo cantiere per iniziare, oppure ripristina un file esportato in precedenza.' }));
      app.appendChild(btn('Ripristina da file', function () { document.getElementById('file').click(); }));
      app.appendChild(statoMemoria());
      return;
    }

    app.appendChild(h('div', { 'class': 'barra' }, [
      btn('Rinomina', rinominaCantiere), btn('Elimina cantiere', eliminaCantiere, 'danger')
    ]));
    app.appendChild(btn('+ Giornata di oggi', function () { modifica(C.oggiISO()); }, 'primary cta-grande'));
    var dataAltra = h('input', { type: 'date', id: 'altra-data', value: C.oggiISO(), 'aria-label': 'Un\'altra data' });
    app.appendChild(h('div', { 'class': 'campo-data-secondario' }, [
      dataAltra,
      btn('Aggiungi per un\'altra data', function () { if (C.dataValida(dataAltra.value)) modifica(dataAltra.value); })
    ]));

    var date = Object.keys(c.giornate).sort().reverse();
    var modificate = giornateDaEsportare(c);

    app.appendChild(h('div', { 'class': 'barra' }, [h('h2', { text: 'Giornate:' })]));
    app.appendChild(h('ul', { 'class': 'lista' }, date.map(function (d) {
      var dt = new Date(d + 'T12:00:00');
      return h('li', {}, [h('button', { type: 'button', on: { click: function () { modifica(d); } } }, [
        h('span', { 'class': 'giorno-numero' }, [
          h('strong', { text: GIORNI[dt.getDay()] + ' ' + dt.getDate() }),
          h('span', { 'class': 'giorno-mese', text: MESI[dt.getMonth()] + " '" + String(dt.getFullYear()).slice(-2) })
        ]),
        h('span', { 'class': 'giorno-dettagli' }, [
          h('strong', { text: c.nome }), h('span', { text: anteprimaGiorno(c.giornate[d]) })
        ])
      ])]);
    })));
    if (!date.length) app.appendChild(h('p', { 'class': 'nota', text: 'Nessuna giornata salvata per questo cantiere.' }));
    app.appendChild(h('section', {}, [
      h('p', { 'class': modificate ? 'nota alert' : 'nota', text: !date.length ? '' :
        (giornate(modificate) + (modificate === 1 ? ' non ancora esportata' : ' non ancora esportate') + '. ' + ultimoExportTesto(c) + ' I dati esistono solo su questo dispositivo.') }),
      btn('Esporta per LeenO', esporta, 'cta'),
      btn('Stampa o PDF', stampaPDF),
      btn('Ripristina da file', function () { document.getElementById('file').click(); }),
      statoMemoria()
    ]));
  }

  function rinominaCantiere() {
    var c = cantiereCorrente();
    var nome = (prompt('Nuovo nome del cantiere:', c.nome) || '').trim();
    if (!nome || nome === c.nome) return;
    c.nome = nome; salva(); principale();
  }

  function eliminaCantiere() {
    var c = cantiereCorrente(), n = Object.keys(c.giornate).length;
    var daSalvare = giornateDaEsportare(c);
    if (daSalvare && confirm('Questo cantiere ha ' + giornate(daSalvare) + (daSalvare === 1 ? ' non ancora esportata' : ' non ancora esportate')
      + '.\n\nOK: salva prima una copia di sicurezza (poi potrai eliminare il cantiere).\nAnnulla: prosegui senza copia.')) {
      esportaCantiere(); return;
    }
    var testo = 'Eliminare il cantiere "' + c.nome + '"' + (n ? ', con ' + giornate(n) + '?' : '?');
    if (n && !c.ultimo_export) testo += '\n\nAttenzione: non è mai stato esportato per LeenO.';
    if (!confirm(testo)) return;
    var idEliminato = stato.attivo;
    delete stato.cantieri[idEliminato]; stato.attivo = null; salva(); principale();
    FotoStore.eliminaPerCantiere(idEliminato).catch(function () { /* pulizia foto: nessun blocco per l'utente */ });
  }

  function modifica(iso) {
    svuota();
    var c = cantiereCorrente(), g = c.giornate[iso] || { campi: {} };
    cantiereApertoId = stato.attivo; dataAperta = iso; inModifica = true;
    cantiereApertoRif = c; giornoApertoRif = g;
    app.appendChild(h('div', { 'class': 'barra' }, [btn('Indietro', principale), h('h2', { text: dataEstesa(iso) })]));
    app.appendChild(h('p', { 'class': 'cantiere-corrente' }, [
      h('span', { 'class': 'etichetta', text: 'Cantiere:' }), h('span', { text: c.nome })
    ]));
    app.appendChild(h('p', { 'class': 'suggerimento', text: 'Per scrivere a voce, tocca il microfono sulla tastiera del telefono.' }));
    C.CAMPI.forEach(function (campo) {
      var id = 'c_' + campo[0], el;
      if (campo[0] === 'meteo') {
        el = h('input', { id: id, type: 'text', list: 'meteo', autocomplete: 'off' });
      } else {
        el = h('textarea', { id: id, rows: campo[0] === 'annotazioni' ? '6' : '3' });
      }
      el.value = g.campi[campo[0]] || '';
      el.addEventListener('input', function () {
        g.campi[campo[0]] = el.value; g.modificato_il = new Date().toISOString();
        c.giornate[iso] = g; salva();
      });
      app.appendChild(h('div', { 'class': 'campo' }, [h('label', { 'for': id, text: etichettaCampo(campo) }), el]));
    });
    app.appendChild(h('div', { 'class': 'campo' }, [
      h('label', { text: 'Foto:' }),
      h('div', { id: 'foto-lista', 'class': 'foto-lista' }),
      h('div', { 'class': 'azioni-destra' }, [btn('Aggiungi foto', function () { document.getElementById('file-foto').click(); })])
    ]));
    aggiornaFotoLista();
    app.appendChild(btn('Indietro', principale, 'cta-grande')); // in fondo alla scheda, lontano da Elimina
    app.appendChild(h('div', { 'class': 'zona-pericolo' }, [btn('Elimina questa giornata', function () {
      if (!confirm('Eliminare la giornata ' + dataEstesa(iso) + '?')) return;
      delete c.giornate[iso]; salva(); principale();
      FotoStore.eliminaPerGiorno(stato.attivo, iso).catch(function () { /* pulizia foto: nessun blocco per l'utente */ });
    }, 'danger')]));
  }

  // Etichetta di un campo degli appunti: trattino iniziale (rende leggibile la stampa) e due punti finali.
  function etichettaCampo(campo) { return '- ' + campo[1] + ':'; }

  function aggiornaFotoLista() {
    var el = document.getElementById('foto-lista');
    if (!el) return;
    while (el.firstChild) el.removeChild(el.firstChild);
    if (!FotoStore) return;
    FotoStore.elencaPerGiorno(cantiereApertoId, dataAperta).then(function (righe) {
      if (el !== document.getElementById('foto-lista')) return; // schermata cambiata nel frattempo
      righe.forEach(function (r) {
        var url = URL.createObjectURL(r.blob); urlFotoAttive.push(url);
        el.appendChild(h('figure', { 'class': 'foto' }, [
          h('img', { src: url, alt: 'Foto del ' + dataEstesa(dataAperta) }),
          btn('Elimina', function () {
            if (!confirm('Eliminare questa foto?')) return;
            FotoStore.elimina(r.id).then(function () {
              if (cantiereApertoRif && cantiereApertoRif.giornate[dataAperta]) {
                giornoApertoRif.modificato_il = new Date().toISOString(); salva();
              }
              aggiornaFotoLista();
            });
          }, 'danger')
        ]));
      });
    }).catch(function (e) { avviso('Impossibile caricare le foto: ' + e.message); });
  }

  function contieneDatiSensibili(dati) {
    if (!dati) return false;
    if (Array.isArray(dati)) {
      return dati.some(function (item) {
        var d = item.dati || item;
        return d && d.giornate && d.giornate.some(function (g) { return g && g.campi && g.campi.infortuni; });
      });
    }
    return dati.giornate ? dati.giornate.some(function (g) { return g && g.campi && g.campi.infortuni; }) : false;
  }

  function contaGiornate(dati) {
    if (!dati) return 0;
    if (Array.isArray(dati)) {
      var tot = 0;
      dati.forEach(function (item) {
        var d = item.dati || item;
        if (d && d.giornate) tot += d.giornate.length;
      });
      return tot;
    }
    return dati.giornate ? dati.giornate.length : 0;
  }

  // azione: 'esportare il file' | 'creare il PDF', per adattare il testo dell'avviso.
  // Restituisce true se si può procedere (nessun dato sensibile, o confermato).
  function confermaSeSensibile(dati, azione) {
    if (!contieneDatiSensibili(dati)) return true;
    alert('Questo contiene il campo "Evento infortunistico" di una o più giornate, che può riguardare '
      + 'dati sulla salute di una persona nominata.\n\n'
      + 'È corretto usarlo per portarlo in LeenO. Se invece hai intenzione di condividerlo con altre persone '
      + '(oltre a te stesso), verifica prima con chi segue la privacy in azienda: condividere dati sulla salute '
      + 'richiede attenzioni che questa app non gestisce.');
    return confirm('Vuoi comunque ' + azione + '?');
  }

  function chiudiPannello() {
    while (pannello.firstChild) pannello.removeChild(pannello.firstChild);
    pannello.hidden = true;
  }
  function registraExport(target) {
    var adesso = new Date().toISOString();
    if (Array.isArray(target)) {
      target.forEach(function (item) {
        var c = item.c || item;
        if (c && typeof c === 'object') c.ultimo_export = adesso;
      });
    } else if (target && typeof target === 'object') {
      target.ultimo_export = adesso;
    }
    salva(); principale();
  }
  function salvaSulTelefono(file, dati, c, senzaGiornate) {
    // Download diretto via <a download>: non richiede alcuna attivazione utente.
    var a = h('a', { href: URL.createObjectURL(file), download: file.name });
    document.body.appendChild(a); a.click(); a.remove();
    registraExport(c);
    var sgText = senzaGiornate ? ' ' + cantieri(senzaGiornate) + ' senza giornate non ' + (senzaGiornate === 1 ? 'e\' incluso' : 'sono inclusi') + '.' : '';
    avviso('Esportate ' + giornate(contaGiornate(dati)) + '. File salvato sul telefono: ' + file.name
      + '. Di solito si trova nella cartella Download (app File).' + sgText);
  }
  function condividiFile(file, dati, c, senzaGiornate) {
    if (!confermaSeSensibile(dati, 'condividere il file')) { avviso('Condivisione annullata.'); return; }
    var titolo = 'Brogliaccio';
    if (c && !Array.isArray(c) && c.nome) titolo += ': ' + c.nome;
    try {
      var p = navigator.share({ files: [file], title: titolo });
      if (p && p.then) {
        p.then(
          function () { registraExport(c); avviso('File condiviso: ' + file.name); },
          function (e) {
            if (e && e.name === 'AbortError') { avviso('Condivisione annullata. Puoi riprovare oppure salvare il file sul telefono.'); return; }
            avviso('Condivisione non riuscita (' + (e && e.name ? e.name : 'errore') + '). Il file viene salvato sul telefono.');
            salvaSulTelefono(file, dati, c, senzaGiornate);
          }
        );
      } else {
        salvaSulTelefono(file, dati, c, senzaGiornate);
      }
    } catch (e) {
      if (e && e.name === 'AbortError') { avviso('Condivisione annullata.'); return; }
      avviso('La condivisione diretta non è supportata da questo browser. Il file viene salvato sul telefono.');
      salvaSulTelefono(file, dati, c, senzaGiornate);
    }
  }
  // Il file e' gia' pronto (costruirlo e' asincrono, e dopo un lavoro asincrono il browser puo'
  // rifiutare share(): vedi LESSONS_PWA_BROGLIACCIO.md). Se il dispositivo sa condividere file,
  // un riquadro offre pulsanti da toccare: il tocco fresco rende valida la condivisione.
  // L'ultimo export si registra solo a condivisione riuscita o a file salvato.
  // Il .json si condivide come testo (.json.txt) per massima compatibilita' con i selettori MIME di Chromium,
  // mentre i file .zip (singolo cantiere con foto o tutti i cantieri) si condividono direttamente come archivio.
  function fileDaCondividere(file) {
    if (/\.json$/i.test(file.name)) return new File([file], file.name + '.txt', { type: 'text/plain' });
    if (/\.zip$/i.test(file.name)) return file;
    return file;
  }
  function condividiOScarica(file, dati, c, senzaGiornate) {
    chiudiPannello();
    var daCondividere = fileDaCondividere(file);
    var supportaShare = false;
    try {
      supportaShare = !!(daCondividere && navigator.canShare && navigator.canShare({ files: [daCondividere] }));
    } catch (e) {
      supportaShare = false;
    }
    pannello.hidden = false;
    pannello.appendChild(h('p', { 'class': 'pannello-titolo', text: 'File pronto: ' + file.name }));
    var numG = contaGiornate(dati);
    var notaTesto = giornate(numG) + '. ' + (supportaShare
      ? 'Per portarle in LeenO invia il file a te stesso (posta, messaggi, cloud) oppure salvalo sul telefono.'
      : 'Il tuo browser non supporta la condivisione diretta. Salva il file sul telefono per portarlo in LeenO.');
    pannello.appendChild(h('p', { 'class': 'nota', text: notaTesto }));
    if (supportaShare) {
      pannello.appendChild(btn('Invia o condividi', function () { condividiFile(daCondividere, dati, c, senzaGiornate); }, 'cta cta-grande'));
    }
    pannello.appendChild(btn('Salva sul telefono', function () { salvaSulTelefono(file, dati, c, senzaGiornate); }, supportaShare ? '' : 'cta cta-grande'));
    pannello.appendChild(btn('Chiudi', chiudiPannello, ''));
    avviso('');
    pannello.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // Prepara i file di un cantiere. prefisso: '' per un cantiere esportato da solo (file alla radice dello .zip),
  // 'cartella/' quando si esportano tutti i cantieri (una cartella per cantiere, con la stessa struttura).
  function preparaCantiere(id, prefisso) {
    var c = stato.cantieri[id];
    var dati = C.buildExport(c.giornate, new Date());
    dati.testata = { lavori: c.nome };
    var base = 'agenda-' + slug(c.nome) + '-' + C.oggiISO().replace(/-/g, '');
    return FotoStore.elencaPerCantiere(id).catch(function () { return []; }).then(function (foto) {
      var perGiorno = {};
      foto.forEach(function (f) { (perGiorno[f.giorno] = perGiorno[f.giorno] || []).push(f); });
      var voci = [{ percorso: prefisso + base + '.json', promessa: Promise.resolve(new TextEncoder().encode(JSON.stringify(dati, null, 2))) }];
      Object.keys(perGiorno).sort().forEach(function (giorno) {
        var cartella = giorno.replace(/-/g, '');
        perGiorno[giorno]
          .slice().sort(function (a, b) { return a.creato_il < b.creato_il ? -1 : 1; })
          .forEach(function (f, i) {
            var nome = marcaTemporale(f.creato_il) + '_' + String(i + 1).padStart(3, '0') + '.jpg';
            voci.push({
              percorso: prefisso + cartella + '/' + nome,
              promessa: f.blob.arrayBuffer().then(function (buf) { return new Uint8Array(buf); })
            });
          });
      });
      return { c: c, dati: dati, base: base, foto: foto.length, voci: voci };
    });
  }
  function vociPronte(voci) {
    return Promise.all(voci.map(function (v) { return v.promessa.then(function (d) { return { percorso: v.percorso, dati: d }; }); }));
  }

  // Esporta il solo cantiere aperto: .json semplice, oppure .zip se ha foto.
  function esportaCantiere() {
    var c = cantiereCorrente();
    var dati = C.buildExport(c.giornate, new Date());
    if (!dati.giornate.length) { avviso('Nessuna giornata da esportare.'); return; }
    // Nessun blocco qui: il salvataggio in locale resta sul dispositivo, non e' un
    // rischio di condivisione. L'avviso privacy scatta sulla condivisione vera e
    // propria, dentro condividiOScarica().
    preparaCantiere(stato.attivo, '').then(function (p) {
      if (!p.foto) {
        condividiOScarica(new File([JSON.stringify(p.dati, null, 2)], p.base + '.json', { type: 'application/json' }), p.dati, c);
        return;
      }
      return vociPronte(p.voci).then(function (pronte) {
        condividiOScarica(new File([ZipStore.creaZip(pronte)], p.base + '.zip', { type: 'application/zip' }), p.dati, c);
      }).catch(function (e) { avviso('Creazione del file compresso non riuscita: ' + e.message); });
    });
  }

  // Esporta tutti i cantieri con giornate in un unico .zip: una cartella per cantiere, ciascuna con il proprio
  // .json e le proprie foto.
  function esportaTutti() {
    var ids = Object.keys(stato.cantieri).sort(function (a, b) {
      return stato.cantieri[a].nome.localeCompare(stato.cantieri[b].nome, 'it') || (a < b ? -1 : 1);
    });
    var conGiornate = ids.filter(function (id) { return Object.keys(stato.cantieri[id].giornate).length; });
    if (!conGiornate.length) { avviso('Nessuna giornata da esportare.'); return; }
    var usate = {};
    avviso('Preparazione del file in corso...');
    Promise.all(conGiornate.map(function (id) {
      var cartella = slug(stato.cantieri[id].nome), scelta = cartella, k = 2;
      while (usate[scelta]) scelta = cartella + '-' + (k++); // due cantieri con lo stesso nome: cartelle distinte
      usate[scelta] = true;
      return preparaCantiere(id, scelta + '/');
    })).then(function (pr) {
      var voci = [];
      pr.forEach(function (p) { voci = voci.concat(p.voci); });
      return vociPronte(voci).then(function (pronte) {
        var file = new File([ZipStore.creaZip(pronte)], 'agenda-tutti-i-cantieri-' + C.oggiISO().replace(/-/g, '') + '.zip', { type: 'application/zip' });
        condividiOScarica(file, pr, pr, ids.length - conGiornate.length);
      });
    }).catch(function (e) { avviso('Creazione del file compresso non riuscita: ' + e.message); });
  }

  // Con piu' cantieri sul telefono si sceglie cosa esportare; con uno solo si esporta subito.
  function esporta() {
    var ids = Object.keys(stato.cantieri);
    if (ids.length < 2) { esportaCantiere(); return; }
    var c = cantiereCorrente(), n = ids.filter(function (id) { return Object.keys(stato.cantieri[id].giornate).length; }).length;
    chiudiPannello();
    pannello.hidden = false;
    pannello.appendChild(h('p', { 'class': 'pannello-titolo', text: 'Cosa vuoi esportare?' }));
    pannello.appendChild(h('p', { 'class': 'nota', text: 'Su questo telefono ci sono ' + cantieri(ids.length)
      + '. Tutti i cantieri finiscono in un unico file .zip, con una cartella per ciascuno.' }));
    pannello.appendChild(btn('Solo "' + c.nome + '"', function () { chiudiPannello(); esportaCantiere(); }, 'cta cta-grande'));
    pannello.appendChild(btn('Tutti i cantieri (' + n + ')', function () { chiudiPannello(); esportaTutti(); }, 'cta cta-grande'));
    pannello.appendChild(btn('Chiudi', chiudiPannello, ''));
    avviso('');
    pannello.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // Risolve quando tutte le immagini del contenitore sono decodificate (o dopo 20 s, per non bloccare la stampa).
  function attendiImmagini(contenitore) {
    var imgs = Array.prototype.slice.call(contenitore.querySelectorAll('img'));
    var tutte = Promise.all(imgs.map(function (img) {
      if (img.decode) return img.decode().catch(function () { /* non decodificabile: si stampa comunque */ });
      return img.complete ? Promise.resolve() : new Promise(function (r) { img.onload = img.onerror = r; });
    }));
    return Promise.race([tutte, new Promise(function (r) { setTimeout(r, 20000); })]);
  }

  function costruisciStampa(nomeCantiere, dati, fotoPerGiorno, colonne) {
    fotoPerGiorno = fotoPerGiorno || {};
    colonne = colonne || 4;
    var el = document.getElementById('stampa');
    while (el.firstChild) el.removeChild(el.firstChild);
    el.appendChild(h('div', { 'class': 's-intestazione' }, [
      h('h1', { text: 'Brogliaccio ' + C.VERSIONE }),
      h('p', { 'class': 's-cantiere', text: 'Cantiere: ' + nomeCantiere }),
      h('p', { 'class': 's-avviso', text: 'Agenda da consolidare in LeenO. Non è un registro ufficiale. '
        + 'Generato il ' + new Date().toLocaleDateString('it-IT') + '.' })
    ]));
    dati.giornate.forEach(function (g) {
      var dl = h('dl');
      C.CAMPI.forEach(function (campo) {
        var testo = g.campi[campo[0]];
        if (!testo) return;
        dl.appendChild(h('dt', { text: etichettaCampo(campo) }));
        dl.appendChild(h('dd', { text: testo }));
      });
      var sezione = h('section', { 'class': 's-giorno' }, [h('h2', { text: dataEstesa(g.data) }), dl]);
      var foto = (fotoPerGiorno[g.data] || []).filter(function (f) { return f.dataUrl; });
      if (foto.length) {
        sezione.appendChild(h('div', { 'class': 's-foto-elenco', 'data-colonne': String(colonne) },
          foto.map(function (f) { return h('img', { src: f.dataUrl, alt: 'Foto del ' + dataEstesa(g.data) }); })));
      }
      el.appendChild(sezione);
    });
    el.appendChild(h('div', { 'class': 's-piede' }, [h('span', { text: 'realizzato con LeenO.org' })]));
  }

  function eseguiStampaPDF(colonne) {
    var c = cantiereCorrente();
    var dati = C.buildExport(c.giornate, new Date());
    if (!dati.giornate.length) { avviso('Nessuna giornata da stampare.'); return; }
    if (!confermaSeSensibile(dati, 'creare il PDF')) { avviso('Operazione annullata.'); return; }
    var idCantiere = stato.attivo;
    FotoStore.elencaPerCantiere(idCantiere).catch(function () { return []; }).then(function (foto) {
      var perGiorno = {};
      foto.forEach(function (f) { (perGiorno[f.giorno] = perGiorno[f.giorno] || []).push(f); });
      Object.keys(perGiorno).forEach(function (g) {
        perGiorno[g].sort(function (a, b) { return a.creato_il < b.creato_il ? -1 : 1; });
      });
      var tutte = [];
      Object.keys(perGiorno).forEach(function (g) { perGiorno[g].forEach(function (f) { tutte.push(f); }); });
      Promise.all(tutte.map(function (f) {
        return new Promise(function (resolve) {
          var lettore = new FileReader();
          lettore.onload = function () { resolve(lettore.result); };
          lettore.onerror = function () { resolve(null); };
          lettore.readAsDataURL(f.blob);
        });
      })).then(function (urlDati) {
        tutte.forEach(function (f, i) { f.dataUrl = urlDati[i]; });
        costruisciStampa(c.nome, dati, perGiorno, colonne);
        // Con window.print() subito dopo l'inserimento le foto non sono ancora decodificate e finiscono nel PDF come riquadri vuoti.
        attendiImmagini(document.getElementById('stampa')).then(function () { window.print(); });
      });
    });
  }

  function stampaPDF() {
    var c = cantiereCorrente();
    var dati = C.buildExport(c.giornate, new Date());
    if (!dati.giornate.length) { avviso('Nessuna giornata da stampare.'); return; }
    chiudiPannello();
    pannello.hidden = false;
    pannello.appendChild(h('p', { 'class': 'pannello-titolo', text: 'Stampa o PDF: ' + c.nome }));
    pannello.appendChild(h('p', { 'class': 'nota', text: 'Scegli quante colonne di immagini inserire nella stampa (da 1 a 4):' }));

    var divColonne = h('div', { 'class': 'scelta-colonne' });
    [1, 2, 3, 4].forEach(function (n) {
      var rad = h('input', { type: 'radio', name: 'colonne-foto', value: String(n), id: 'col-' + n });
      if (n === 4) rad.checked = true;
      divColonne.appendChild(h('label', { 'for': 'col-' + n }, [rad, h('span', { text: n + (n === 1 ? ' colonna' : ' colonne') })]));
    });
    pannello.appendChild(divColonne);

    pannello.appendChild(btn('Genera stampa/PDF', function () {
      var sel = pannello.querySelector('input[name="colonne-foto"]:checked');
      var numColonne = sel ? parseInt(sel.value, 10) : 4;
      chiudiPannello();
      eseguiStampaPDF(numColonne);
    }, 'cta cta-grande'));
    pannello.appendChild(btn('Chiudi', chiudiPannello, ''));
    avviso('');
    pannello.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // Ripristino da un .json (solo testo): le giornate vanno nel cantiere aperto.
  function ripristinaDaJson(f) {
    var c = cantiereCorrente();
    if (!c) { avviso('Apri o crea prima un cantiere.'); return Promise.resolve(); }
    return f.text().then(function (t) {
      var nuove = C.parseImport(t), n = Object.keys(nuove), doppie = n.filter(function (d) { return c.giornate[d]; }).length;
      if (!confirm('Ripristinare ' + giornate(n.length) + ' nel cantiere "' + c.nome + '"? ' +
        doppie + ' già presenti su questo dispositivo verranno sostituite.')) return;
      n.forEach(function (d) { c.giornate[d] = nuove[d]; });
      salva(); principale();
    });
  }

  // Istante di una foto esportata, dal nome AAAAMMGGhhmm_NNN.jpg (ora locale). Il progressivo diventa
  // millisecondi, cosi' le foto dello stesso minuto restano nell'ordine originale. Nome diverso: mezzogiorno del giorno.
  function creatoIlDaNome(giorno, nome, ordine) {
    var m = /^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(?:_(\d+))?/.exec(nome), ms = m && m[6] ? Math.min(+m[6], 999) : ordine % 1000;
    var d = m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], 0, ms)
      : new Date(+giorno.slice(0, 4), +giorno.slice(5, 7) - 1, +giorno.slice(8, 10), 12, 0, 0, ms);
    return isNaN(d.getTime()) ? new Date().toISOString() : d.toISOString();
  }

  // Cantieri contenuti in uno .zip di esportazione. Un file di un solo cantiere ha il .json alla radice e le foto
  // in cartelle AAAAMMGG/; l'esportazione di tutti i cantieri ha una cartella per cantiere con la stessa struttura:
  // le foto appartengono al .json che sta nella cartella che le contiene.
  function leggiPacchetti(voci) {
    var json = voci.filter(function (v) { return /\.json$/i.test(v.percorso); });
    if (!json.length) throw new Error('Nel file .zip manca il file dell\'agenda (.json).');
    var pacchetti = [], primoErrore = null, ignorati = 0;
    json.forEach(function (v) {
      var testo = new TextDecoder().decode(v.dati), nuove, intest;
      try { nuove = C.parseImport(testo); intest = JSON.parse(testo); }
      catch (e) { primoErrore = primoErrore || e; ignorati++; return; }
      var generato = new Date(intest.generato_il);
      pacchetti.push({
        dir: v.percorso.slice(0, v.percorso.lastIndexOf('/') + 1), nuove: nuove, n: Object.keys(nuove), foto: [],
        nome: (intest.testata && typeof intest.testata.lavori === 'string' ? intest.testata.lavori : '').trim(),
        marca: isNaN(generato.getTime()) ? new Date().toISOString() : generato.toISOString()
      });
    });
    if (!pacchetti.length) throw primoErrore;
    voci.forEach(function (v) {
      if (/\.json$/i.test(v.percorso)) return;
      var m = /^(.*\/)?(\d{4})(\d{2})(\d{2})\/([^\/]+)\.jpe?g$/i.exec(v.percorso), giorno = m && (m[2] + '-' + m[3] + '-' + m[4]);
      if (!m || !C.dataValida(giorno) || v.dati.length < 4 || v.dati[0] !== 0xFF || v.dati[1] !== 0xD8) { ignorati++; return; }
      var dir = m[1] || '';
      // un solo cantiere nel file: tutte le foto sono sue; con piu' cantieri conta la cartella
      var p = pacchetti.length === 1 ? pacchetti[0] : pacchetti.filter(function (x) { return x.dir === dir; })[0];
      if (!p) { ignorati++; return; }
      p.foto.push({ giorno: giorno, dati: v.dati, crc: v.crc, creato_il: creatoIlDaNome(giorno, m[5], p.foto.length) });
    });
    pacchetti.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'it') || (a.dir < b.dir ? -1 : 1); });
    return { pacchetti: pacchetti, ignorati: ignorati };
  }

  // Cantiere con lo stesso nome (senza badare a maiuscole) gia' presente sul telefono, oppure null.
  function cercaCantiere(nome) {
    var trovato = null;
    Object.keys(stato.cantieri).some(function (k) {
      if (stato.cantieri[k].nome.trim().toLowerCase() === nome.toLowerCase()) { trovato = k; return true; }
      return false;
    });
    return trovato;
  }

  // Decide dove va il cantiere del file e quali foto sono davvero da aggiungere (stessa giornata e stesso
  // contenuto di una foto gia' presente: non si duplica).
  function analizzaPacchetto(p, unico) {
    var id = null;
    if (p.nome) id = cercaCantiere(p.nome);
    else if (unico) { id = cantiereCorrente() ? stato.attivo : null; p.nome = id ? stato.cantieri[id].nome : 'Cantiere ripristinato'; } // file senza nome: va nel cantiere aperto
    else p.nome = p.dir.replace(/\/$/, '').split('/').pop() || 'Cantiere ripristinato';
    p.id = id; p.esistente = id !== null;
    p.doppie = p.esistente ? p.n.filter(function (d) { return stato.cantieri[id].giornate[d]; }).length : 0;
    var dimensioni = {}, presenti = {};
    p.foto.forEach(function (x) { dimensioni[x.giorno + '|' + x.dati.length] = true; });
    var esaminate = p.esistente ? FotoStore.elencaPerCantiere(id) : Promise.resolve([]);
    return esaminate.then(function (righe) {
      return Promise.all(righe.filter(function (r) { return dimensioni[r.giorno + '|' + r.blob.size]; }).map(function (r) {
        return r.blob.arrayBuffer().then(function (b) { presenti[r.giorno + '|' + r.blob.size + '|' + ZipStore.crc32(new Uint8Array(b))] = true; });
      }));
    }).then(function () {
      p.daAggiungere = p.foto.filter(function (x) {
        var k = x.giorno + '|' + x.dati.length + '|' + x.crc;
        if (presenti[k]) return false;
        presenti[k] = true; return true;
      });
      p.saltate = p.foto.length - p.daAggiungere.length;
    });
  }

  function dettaglioPacchetto(p) {
    return giornate(p.n.length) + (p.doppie ? ' (' + p.doppie + ' già presenti su questo dispositivo verranno sostituite)' : '')
      + ', ' + p.daAggiungere.length + ' foto da aggiungere' + (p.saltate ? ' (' + p.saltate + ' già presenti, non vengono duplicate)' : '');
  }
  function notaIgnorati(n) {
    return n ? n + (n === 1 ? ' voce del file non riconosciuta, ignorata.' : ' voci del file non riconosciute, ignorate.') : '';
  }

  // File con piu' cantieri: elenco con una casella per cantiere, tutte spuntate. Risolve con i cantieri scelti, o null.
  function scegliPacchetti(pacchetti, ignorati) {
    return new Promise(function (resolve) {
      chiudiPannello();
      pannello.hidden = false;
      var caselle = [];
      pannello.appendChild(h('p', { 'class': 'pannello-titolo', text: 'Ripristino dal file .zip' }));
      pannello.appendChild(h('p', { 'class': 'nota', text: 'Il file contiene ' + cantieri(pacchetti.length) + '. Scegli quali ripristinare. '
        + 'Le giornate già presenti sul telefono vengono sostituite; le foto già presenti non vengono duplicate.' + (ignorati ? ' ' + notaIgnorati(ignorati) : '') }));
      pacchetti.forEach(function (p) {
        var casella = h('input', { type: 'checkbox', checked: 'checked' });
        caselle.push(casella);
        pannello.appendChild(h('label', { 'class': 'scelta-cantiere' }, [casella,
          h('span', { text: p.nome + ': ' + dettaglioPacchetto(p) + '. ' + (p.esistente ? 'Cantiere già presente.' : 'Il cantiere verrà creato.') })]));
      });
      function concludi(risultato) { chiudiPannello(); resolve(risultato); }
      pannello.appendChild(btn('Ripristina selezionati', function () {
        var scelti = pacchetti.filter(function (p, i) { return caselle[i].checked; });
        if (!scelti.length) { avviso('Seleziona almeno un cantiere.'); return; }
        concludi(scelti);
      }, 'cta cta-grande'));
      pannello.appendChild(btn('Annulla', function () { concludi(null); }, ''));
      avviso('');
      pannello.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  }

  // Rilegge dalla memoria del browser (non dalle variabili) e controlla che giornate e foto del file ci siano davvero.
  function verificaRipristino(scelti) {
    var salvato = leggiSalvato();
    return Promise.all(scelti.map(function (p) {
      var cs = salvato && salvato.cantieri[p.idFinale];
      var testoOk = !!cs && p.n.every(function (d) { return cs.giornate && cs.giornate[d]; });
      return FotoStore.elencaPerCantiere(p.idFinale).then(function (righe) { return righe.length >= p.foto.length; },
        function () { return false; }).then(function (fotoOk) { return testoOk && fotoOk; });
    })).then(function (r) { return r.every(Boolean); });
  }

  // Scrive i cantieri scelti: prima il testo (localStorage), poi le foto (IndexedDB). Se le foto si interrompono
  // a meta', ripetere lo stesso ripristino completa il lavoro, perche' quelle gia' presenti vengono saltate.
  function applicaPacchetti(scelti) {
    var totG = 0, totF = 0, totS = 0;
    scelti.forEach(function (p) {
      var id = (p.id && stato.cantieri[p.id]) ? p.id : cercaCantiere(p.nome); // due cantieri omonimi nel file: confluiscono in uno
      if (id === null) { id = nuovoId(); stato.cantieri[id] = { nome: p.nome, giornate: {}, ultimo_export: p.marca }; }
      var c = stato.cantieri[id];
      p.n.forEach(function (d) { c.giornate[d] = p.nuove[d]; });
      p.daAggiungere.forEach(function (x) { if (!c.giornate[x.giorno]) c.giornate[x.giorno] = { campi: {}, modificato_il: p.marca }; });
      p.idFinale = id; totG += p.n.length; totF += p.daAggiungere.length; totS += p.saltate;
    });
    stato.attivo = scelti[0].idFinale;
    if (!salva()) return;
    var saltate = totS ? ' (' + totS + ' foto già presenti, saltate)' : '';
    var verifica = function () { return verificaRipristino(scelti); };
    var riepilogo = scelti.length === 1
      ? 'Ripristinate ' + giornate(totG) + ' e ' + totF + ' foto nel cantiere "' + scelti[0].nome + '"' + saltate + '.'
      : 'Ripristinati ' + cantieri(scelti.length) + ': ' + giornate(totG) + ' e ' + totF + ' foto' + saltate + '.';
    return Promise.all(scelti.map(function (p) {
      return Promise.all(p.daAggiungere.map(function (x) {
        return FotoStore.aggiungi(p.idFinale, x.giorno, new Blob([x.dati], { type: 'image/jpeg' }), x.creato_il);
      }));
    })).then(verifica).then(function (ok) {
      principale();
      avviso(riepilogo + (ok ? ' Verificato: i dati risultano salvati nella memoria del browser.'
        : ' ATTENZIONE: la verifica dei dati salvati non è riuscita. Esporta subito e riprova.'));
    }, function (e) {
      principale();
      avviso('Giornate ripristinate, ma non tutte le foto (' + e.message + '). Riprova con lo stesso file: le foto già presenti vengono saltate.');
    });
  }

  // Ripristino da uno .zip creato da "Esporta per LeenO": uno o piu' cantieri (creati se mancano),
  // contenuto dei .json e foto.
  function ripristinaDaZip(f) {
    avviso('Lettura del file in corso...');
    return f.arrayBuffer().then(ZipStore.leggiZip).then(function (voci) {
      var lettura = leggiPacchetti(voci), pacchetti = lettura.pacchetti, unico = pacchetti.length === 1;
      return Promise.all(pacchetti.map(function (p) { return analizzaPacchetto(p, unico); })).then(function () {
        if (unico) {
          var p = pacchetti[0];
          var domanda = 'Ripristinare dal file .zip?\n\nCantiere "' + p.nome + '": ' + (p.esistente ? 'già presente, le giornate vanno qui.' : 'non presente, verrà creato.')
            + '\n' + giornate(p.n.length) + (p.doppie ? ' (' + p.doppie + ' già presenti su questo dispositivo verranno sostituite)' : '') + '.'
            + '\n' + p.daAggiungere.length + ' foto da aggiungere' + (p.saltate ? ' (' + p.saltate + ' già presenti, non vengono duplicate)' : '') + '.'
            + (lettura.ignorati ? '\n' + notaIgnorati(lettura.ignorati) : '');
          if (!confirm(domanda)) { avviso(''); return; }
          return applicaPacchetti([p]);
        }
        return scegliPacchetti(pacchetti, lettura.ignorati).then(function (scelti) {
          if (!scelti) { avviso(''); return; }
          return applicaPacchetti(scelti);
        });
      });
    });
  }

  document.getElementById('file').addEventListener('change', function (ev) {
    var f = ev.target.files[0]; ev.target.value = '';
    if (!f) return;
    // Si riconosce lo .zip dai primi byte ("PK"), non dall'estensione: i selettori di file dei telefoni cambiano spesso i nomi.
    f.slice(0, 2).arrayBuffer().then(function (b) {
      var t = new Uint8Array(b);
      return (t[0] === 0x50 && t[1] === 0x4B) ? ripristinaDaZip(f) : ripristinaDaJson(f);
    }).catch(function (e) { avviso(e.message); });
  });

  document.getElementById('file-foto').addEventListener('change', function (ev) {
    var file = Array.prototype.slice.call(ev.target.files); ev.target.value = '';
    if (!file.length || !dataAperta) return;
    avviso('Elaborazione foto in corso...');
    Promise.all(file.map(function (f) {
      return FotoStore.preparaImmagine(f).then(function (blob) { return FotoStore.aggiungi(cantiereApertoId, dataAperta, blob); });
    })).then(function () {
      // la prima foto di una giornata mai toccata prima la registra nell'elenco, come farebbe un campo di testo
      if (cantiereApertoRif && cantiereApertoRif.giornate[dataAperta] !== giornoApertoRif) {
        cantiereApertoRif.giornate[dataAperta] = giornoApertoRif;
      }
      giornoApertoRif.modificato_il = new Date().toISOString();
      salva(); avviso(''); aggiornaFotoLista();
    }).catch(function (e) { avviso('Foto non salvata: ' + e.message); });
  });

  var lista = document.getElementById('meteo');
  C.METEO.forEach(function (m) { lista.appendChild(h('option', { value: m })); });
  var elTitolo = document.getElementById('titolo-app');
  if (elTitolo) elTitolo.textContent = 'Brogliaccio ' + C.VERSIONE;
  // Diagnostica di avvio: confronta il numero di giornate con l'ultimo stato registrato dall'app stessa.
  // Se e' diminuito senza un'azione dell'utente, i dati sono stati cancellati fuori dall'app (browser o sistema).
  var precedente = leggiDiag(), adesso = contaDati();
  var statoPrec = precedente && precedente.ultimo_stato;
  diag = { avvii: ((precedente && precedente.avvii) || 0) + 1, primo_avvio: (precedente && precedente.primo_avvio) || new Date().toISOString(),
    ultimo_stato: adesso, versione: C.VERSIONE };
  if (statoPrec && adesso.giornate < statoPrec.giornate) {
    allarme = 'Attenzione: dall\'ultimo utilizzo mancano dati (prima ' + giornate(statoPrec.giornate) + ', ora ' + giornate(adesso.giornate)
      + '). Il browser potrebbe averli cancellati. Ripristina dall\'ultimo file esportato.';
  }
  scriviDiag();
  if (FotoStore && FotoStore.contaTutte) {
    FotoStore.contaTutte().then(function (n) {
      if (n > 0 && adesso.cantieri === 0 && (!precedente || (statoPrec && statoPrec.cantieri > 0)) && !allarme) {
        allarme = 'Sul telefono ci sono ' + n + ' foto ma nessun cantiere: i testi risultano persi. Ripristina dall\'ultimo file esportato.';
        if (!inModifica) principale();
      }
    }).catch(function () { /* nessuna diagnostica sulle foto se l'archivio non si apre */ });
  }
  verificaPersistenza().then(function (p) { memoriaProtetta = p; aggiornaStatoMemoria(); });
  window.addEventListener('storage', function (ev) { if (ev.key === KEY) aggiornaDaAltrove(); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') aggiornaDaAltrove(); });
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
  principale();
})();
