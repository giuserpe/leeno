/* Brogliaccio per LeenO: interfaccia. Dati solo su questo dispositivo. */
(function () {
  'use strict';
  var C = window.Core, KEY = 'appunti-cantiere-v1', TEMA_KEY = 'appunti-cantiere-tema';
  var cantiereApertoId = null, dataAperta = null, urlFotoAttive = [];
  var cantiereApertoRif = null, giornoApertoRif = null;
  var app = document.getElementById('app'), msg = document.getElementById('msg');
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
      if (s && s.cantieri) return s;
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
    return { cantieri: {}, attivo: null };
  }
  function salva() {
    try { localStorage.setItem(KEY, JSON.stringify(stato)); return true; }
    catch (e) { avviso('Salvataggio non riuscito: memoria piena o bloccata. Esporta subito i dati.'); return false; }
  }
  function avviso(t) { msg.textContent = t; msg.hidden = !t; }
  function giornate(n) { return n + (n === 1 ? ' giornata' : ' giornate'); }
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

  // -------------------- schermata: elenco cantieri --------------------
  function cantieri() {
    svuota();
    var ids = Object.keys(stato.cantieri).sort(function (a, b) {
      return stato.cantieri[a].nome.localeCompare(stato.cantieri[b].nome, 'it');
    });
    app.appendChild(h('div', { 'class': 'barra' }, [h('h2', { text: 'Cantieri:' })]));
    app.appendChild(h('ul', { 'class': 'lista' }, ids.map(function (id) {
      var c = stato.cantieri[id], n = Object.keys(c.giornate).length;
      return h('li', {}, [h('button', { type: 'button', on: { click: function () { stato.attivo = id; salva(); elenco(); } } }, [
        h('strong', { text: c.nome }), h('span', { text: giornate(n) + (n === 1 ? ' salvata' : ' salvate') })])]);
    })));
    if (!ids.length) app.appendChild(h('p', { 'class': 'nota', text: 'Nessun cantiere. Creane uno per iniziare.' }));
    app.appendChild(btn('Nuovo cantiere ►', creaCantiere, 'primary'));
  }

  function creaCantiere() {
    var nome = (prompt('Nome del cantiere (es. via del corso, indirizzo o commessa):') || '').trim();
    if (!nome) return;
    var id = nuovoId();
    stato.cantieri[id] = { nome: nome, giornate: {}, ultimo_export: null };
    stato.attivo = id; salva(); elenco();
  }

  function rinominaCantiere() {
    var c = cantiereCorrente();
    var nome = (prompt('Nuovo nome del cantiere:', c.nome) || '').trim();
    if (!nome || nome === c.nome) return;
    c.nome = nome; salva(); elenco();
  }

  function eliminaCantiere() {
    var c = cantiereCorrente(), n = Object.keys(c.giornate).length;
    var testo = 'Eliminare il cantiere "' + c.nome + '"' + (n ? ', con ' + giornate(n) + '?' : '?');
    if (n && !c.ultimo_export) testo += '\n\nAttenzione: non è mai stato esportato per LeenO.';
    if (!confirm(testo)) return;
    var idEliminato = stato.attivo;
    delete stato.cantieri[idEliminato]; stato.attivo = null; salva(); cantieri();
    FotoStore.eliminaPerCantiere(idEliminato).catch(function () { /* pulizia foto: nessun blocco per l'utente */ });
  }

  // -------------------- schermata: elenco giornate di un cantiere --------------------
  function elenco() {
    svuota();
    var c = cantiereCorrente();
    if (!c) { cantieri(); return; }
    var date = Object.keys(c.giornate).sort().reverse();
    var oggi = h('input', { type: 'date', id: 'nuova', value: C.oggiISO() });
    var modificate = date.filter(function (d) { return !c.ultimo_export || c.giornate[d].modificato_il > c.ultimo_export; }).length;
    app.appendChild(h('div', { 'class': 'barra' }, [
      btn('Cantieri ►', cantieri), h('h2', { text: c.nome }), btn('Rinomina', rinominaCantiere)
    ]));
    app.appendChild(h('section', {}, [
      h('label', { 'for': 'nuova', text: 'Giornata' }), oggi,
      btn('Apri o crea giornata', function () { if (C.dataValida(oggi.value)) modifica(oggi.value); }, 'primary')
    ]));
    app.appendChild(h('ul', { 'class': 'lista' }, date.map(function (d) {
      var a = c.giornate[d].campi.annotazioni || c.giornate[d].campi.meteo || '';
      return h('li', {}, [h('button', { type: 'button', on: { click: function () { modifica(d); } } }, [
        h('strong', { text: dataEstesa(d) }), h('span', { text: a.slice(0, 80) })])]);
    })));
    if (!date.length) app.appendChild(h('p', { 'class': 'nota', text: 'Nessuna giornata salvata per questo cantiere.' }));
    app.appendChild(h('section', {}, [
      h('p', { 'class': modificate ? 'nota alert' : 'nota', text: !date.length ? '' :
        (giornate(modificate) + (modificate === 1 ? ' non ancora esportata' : ' non ancora esportate') + '. I dati esistono solo su questo dispositivo.') }),
      btn('Esporta per LeenO ↗', esporta, 'cta'),
      btn('Stampa o PDF ↙', stampaPDF),
      btn('Ripristina da file ↙', function () { document.getElementById('file').click(); }),
      btn('Elimina cantiere ⌫', eliminaCantiere, 'danger')
    ]));
  }

  function modifica(iso) {
    svuota();
    var c = cantiereCorrente(), g = c.giornate[iso] || { campi: {} };
    cantiereApertoId = stato.attivo; dataAperta = iso;
    cantiereApertoRif = c; giornoApertoRif = g;
    app.appendChild(h('div', { 'class': 'barra' }, [btn('◄ Indietro', elenco), h('h2', { text: dataEstesa(iso) })]));
    app.appendChild(h('p', { 'class': 'cantiere-corrente' }, [
      h('span', { 'class': 'etichetta', text: 'Cantiere:' }), h('span', { text: c.nome })
    ]));
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
      app.appendChild(h('div', { 'class': 'campo' }, [h('label', { 'for': id, text: campo[1] }), el]));
    });
    app.appendChild(h('div', { 'class': 'campo' }, [
      h('label', { text: 'Foto' }),
      h('div', { id: 'foto-lista', 'class': 'foto-lista' }),
      btn('Aggiungi foto', function () { document.getElementById('file-foto').click(); })
    ]));
    aggiornaFotoLista();
    app.appendChild(btn('Elimina questa giornata ⌫', function () {
      if (!confirm('Eliminare la giornata ' + dataEstesa(iso) + '?')) return;
      delete c.giornate[iso]; salva(); elenco();
      FotoStore.eliminaPerGiorno(stato.attivo, iso).catch(function () { /* pulizia foto: nessun blocco per l'utente */ });
    }, 'danger'));
  }

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
          btn('Elimina  ⌫', function () {
            if (!confirm('Eliminare questa foto?')) return;
            FotoStore.elimina(r.id).then(aggiornaFotoLista);
          }, 'danger')
        ]));
      });
    }).catch(function (e) { avviso('Impossibile caricare le foto: ' + e.message); });
  }

  function contieneDatiSensibili(dati) {
    return dati.giornate.some(function (g) { return g.campi.infortuni; });
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

  function condividiOScarica(file, dati, c) {
    // Si salva sempre il file in locale per primo: e' affidabile e non dipende da
    // permessi del browser. La condivisione, che fa uscire il file dal dispositivo,
    // viene proposta subito dopo come passo separato, con l'avviso privacy agganciato
    // a quella proposta e non al semplice salvataggio locale.
    var a = h('a', { href: URL.createObjectURL(file), download: file.name });
    document.body.appendChild(a); a.click(); a.remove();
    c.ultimo_export = new Date().toISOString(); salva(); elenco();
    avviso('Esportate ' + giornate(dati.giornate.length) + ': ' + file.name);

    if (!(navigator.canShare && navigator.canShare({ files: [file] }))) return;
    if (!confirm('Vuoi condividere questo file?')) return;
    if (!confermaSeSensibile(dati, 'condividere il file')) { avviso('Condivisione annullata.'); return; }
    navigator.share({ files: [file], title: 'Brogliaccio: ' + c.nome }).then(
      function () { avviso('File condiviso: ' + file.name); },
      function (e) {
        if (e && e.name === 'AbortError') { avviso('Condivisione annullata. Il file resta comunque salvato.'); return; }
        avviso('Condivisione non riuscita. Il file resta comunque salvato: ' + file.name);
      }
    );
  }

  function esporta() {
    var c = cantiereCorrente();
    var dati = C.buildExport(c.giornate, new Date());
    if (!dati.giornate.length) { avviso('Nessuna giornata da esportare.'); return; }
    // Nessun blocco qui: il salvataggio in locale resta sul dispositivo, non e' un
    // rischio di condivisione. L'avviso privacy scatta sulla condivisione vera e
    // propria, dentro condividiOScarica().
    dati.testata = { lavori: c.nome };
    var base = 'agenda-' + slug(c.nome) + '-' + C.oggiISO().replace(/-/g, '');
    var idCantiere = stato.attivo;

    FotoStore.elencaPerCantiere(idCantiere).catch(function () { return []; }).then(function (foto) {
      if (!foto.length) {
        condividiOScarica(new File([JSON.stringify(dati, null, 2)], base + '.json', { type: 'application/json' }), dati, c);
        return;
      }
      var perGiorno = {};
      foto.forEach(function (f) { (perGiorno[f.giorno] = perGiorno[f.giorno] || []).push(f); });
      var voci = [{ percorso: base + '.json', promessa: Promise.resolve(new TextEncoder().encode(JSON.stringify(dati, null, 2))) }];
      Object.keys(perGiorno).sort().forEach(function (giorno) {
        var cartella = giorno.replace(/-/g, '');
        perGiorno[giorno]
          .slice().sort(function (a, b) { return a.creato_il < b.creato_il ? -1 : 1; })
          .forEach(function (f, i) {
            var nome = marcaTemporale(f.creato_il) + '_' + String(i + 1).padStart(3, '0') + '.jpg';
            voci.push({
              percorso: cartella + '/' + nome,
              promessa: f.blob.arrayBuffer().then(function (buf) { return new Uint8Array(buf); })
            });
          });
      });
      Promise.all(voci.map(function (v) { return v.promessa.then(function (d) { return { percorso: v.percorso, dati: d }; }); }))
        .then(function (vociPronte) {
          condividiOScarica(new File([ZipStore.creaZip(vociPronte)], base + '.zip', { type: 'application/zip' }), dati, c);
        })
        .catch(function (e) { avviso('Creazione del file compresso non riuscita: ' + e.message); });
    });
  }

  function stampaPDF() {
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
          lettore.onload = function () {
            var img = new Image();
            img.onload = function() { resolve({ url: lettore.result, w: img.naturalWidth, h: img.naturalHeight }); };
            img.onerror = function() { resolve({ url: lettore.result, w: 400, h: 400 }); };
            img.src = lettore.result;
          };
          lettore.onerror = function () { resolve(null); };
          lettore.readAsDataURL(f.blob);
        });
      })).then(function (datiFoto) {
        tutte.forEach(function (f, i) {
          if (datiFoto[i]) { f.dataUrl = datiFoto[i].url; f.w = datiFoto[i].w; f.h = datiFoto[i].h; }
        });
        generaVeroPDF(c.nome, dati, perGiorno);
      });
    });
  }

  function generaVeroPDF(nomeCantiere, dati, fotoPerGiorno) {
    if (!window.jspdf) { avviso('Libreria PDF non ancora caricata. Riprova tra un attimo.'); return; }
    var doc = new window.jspdf.jsPDF();
    var mar = 20, y = mar;
    var maxW = 210 - mar * 2;
    var riga = 5;
    var nPag = 1;
    
    function addPageNum() {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.text('realizzato con LeenO.org', mar, 285, { align: 'left' });
      doc.text('Pagina ' + nPag, 210 - mar, 285, { align: 'right' });
      nPag++;
    }

    function addText(testo, font, size) {
      doc.setFont('helvetica', font);
      doc.setFontSize(size);
      var linee = doc.splitTextToSize(testo, maxW);
      for (var i = 0; i < linee.length; i++) {
        if (y > 275) {
          addPageNum();
          doc.addPage();
          y = mar;
        }
        doc.text(linee[i], mar, y);
        y += riga + (size > 12 ? 2 : 0); // extra spazio per titoli
      }
      y += 2;
    }
    
    // Intestazione globale
    addText('Brogliaccio ' + C.VERSIONE, 'bold', 16);
    addText('Cantiere: ' + nomeCantiere, 'bold', 14);
    addText('Agenda da consolidare in LeenO. Non è un registro ufficiale.', 'italic', 10);
    addText('Generato il ' + new Date().toLocaleDateString('it-IT'), 'italic', 10);
    y += 8;
    
    dati.giornate.forEach(function (g, index) {
      // Se non è il primo giorno, cambia pagina per ricominciare da 1
      if (index > 0) {
        addPageNum();
        doc.addPage();
        y = mar;
        nPag = 1; // Resetta ad ogni cambio di data
      }
      
      addText(dataEstesa(g.data), 'bold', 12);
      y += 4;
      
      C.CAMPI.forEach(function (campo) {
        var testo = g.campi[campo[0]];
        if (!testo) return;
        addText(campo[1].toUpperCase(), 'bold', 10);
        addText(testo, 'normal', 10);
        y += 2;
      });
      
      var foto = (fotoPerGiorno[g.data] || []).filter(function (f) { return f.dataUrl; });
      if (foto.length) {
        y += 4;
        var maxBox = 50, x = mar;
        var inlineY = y;
        var rigaH = 0;
        for (var i = 0; i < foto.length; i++) {
          var wOrig = foto[i].w || 400, hOrig = foto[i].h || 400;
          var fW, fH;
          if (wOrig > hOrig) {
            fW = maxBox;
            fH = (hOrig / wOrig) * maxBox;
          } else {
            fH = maxBox;
            fW = (wOrig / hOrig) * maxBox;
          }

          if (x + fW > 210 - mar) {
            x = mar;
            inlineY += rigaH + 4;
            rigaH = 0;
          }
          if (inlineY + fH > 275) {
            addPageNum();
            doc.addPage();
            inlineY = mar;
            x = mar;
            rigaH = 0;
          }
          try { doc.addImage(foto[i].dataUrl, x, inlineY, fW, fH); } catch(e) {}
          x += fW + 4;
          if (fH > rigaH) rigaH = fH;
        }
        y = inlineY + rigaH + 6;
      }
    });
    
    if (dati.giornate.length > 0) {
      addPageNum();
    }
    
    var pdfBlob = doc.output('blob');
    var pdfFile = new File([pdfBlob], 'Brogliaccio_' + slug(nomeCantiere) + '.pdf', { type: 'application/pdf' });
    condividiOScarica(pdfFile, dati, cantiereCorrente());
  }

  document.getElementById('file').addEventListener('change', function (ev) {
    var f = ev.target.files[0]; ev.target.value = '';
    if (!f) return;
    var c = cantiereCorrente();
    if (!c) { avviso('Apri o crea prima un cantiere.'); return; }
    f.text().then(function (t) {
      var nuove = C.parseImport(t), n = Object.keys(nuove), doppie = n.filter(function (d) { return c.giornate[d]; }).length;
      if (!confirm('Ripristinare ' + giornate(n.length) + ' nel cantiere "' + c.nome + '"? ' +
        doppie + ' già presenti su questo dispositivo verranno sostituite.')) return;
      n.forEach(function (d) { c.giornate[d] = nuove[d]; });
      salva(); elenco();
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
      salva(); avviso(''); aggiornaFotoLista();
    }).catch(function (e) { avviso('Foto non salvata: ' + e.message); });
  });

  var lista = document.getElementById('meteo');
  C.METEO.forEach(function (m) { lista.appendChild(h('option', { value: m })); });
  var elTitolo = document.getElementById('titolo-app');
  if (elTitolo) elTitolo.textContent = 'Brogliaccio ' + C.VERSIONE;
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js');
  cantieri();
})();
