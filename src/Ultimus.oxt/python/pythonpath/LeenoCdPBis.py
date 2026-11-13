########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################
'''
Aggiornamento delle voci selezionate di COMPUTO, VARIANTE e CONTABILITA a
un nuovo prezzario (file .ods con struttura LeenO).

Il codice di ogni voce viene privato del prefisso (tutto fino all'ultimo
underscore, per qualunque emittente) e cercato nel nuovo prezzario sul solo
suffisso. La descrizione estesa del vecchio e del nuovo articolo viene
confrontata e la voce classificata VERDE, GIALLO o ROSSO. Il prezzo non entra
nel giudizio. L'utente conferma le sostituzioni in un unico dialogo.

Le funzioni pure (taglio del prefisso, normalizzazione, similarita',
classificazione, confronto, report) non usano UNO e sono testate in
tests/test_leeno_cdpbis.py.
'''
import csv
import difflib
import json
import os
import re
import time
from collections import Counter, namedtuple

import Dialogs
import LeenoUtils
import LeenoConfig
import LeenoComputo
import LeenoAnalysis
import SheetUtils
import DocUtils
import uno

FOGLI_VOCI = ('COMPUTO', 'VARIANTE', 'CONTABILITA')

VERDE = 'VERDE'
GIALLO = 'GIALLO'
ROSSO = 'ROSSO'

# Elenco Prezzi: prima riga dati (indice 0), colonne A:J, riga di
# inserimento delle nuove voci e riga modello per le formule L:Z.
_EP_PRIMA_RIGA = 3
_EP_ULTIMA_COL = 9
_EP_RIGA_INSERIMENTO = 4
_EP_RIGA_MODELLO = 3

_EPSILON = 1e-9

Voce = namedtuple('Voce', 'foglio riga codice numero')

Esito = namedtuple(
    'Esito',
    'voce codice_nuovo colore similarita um_vecchia um_nuova '
    'desc_vecchia desc_nuova prezzo_vecchio prezzo_nuovo nota riga_prezzario')

Risultato = namedtuple(
    'Risultato', 'esiti allineate non_trovate ambigue escluse')


########################################################################
# Funzioni pure (senza UNO)
########################################################################

def testo_cella(valore):
    '''Converte un valore di getDataArray in stringa (1.0 diventa "1").'''
    if valore is None:
        return ''
    if isinstance(valore, float):
        if valore.is_integer():
            return str(int(valore))
        return repr(valore)
    return str(valore)


def taglia_prefisso(codice):
    '''
    Restituisce il codice senza prefisso: la parte dopo l'ultimo underscore.
    Senza underscore (o con underscore finale) restituisce il codice intero.
    '''
    testo = testo_cella(codice).strip()
    if '_' not in testo:
        return testo
    suffisso = testo.rsplit('_', 1)[1]
    return suffisso if suffisso else testo


def normalizza_testo(testo):
    '''Maiuscole, punteggiatura trasformata in spazio, spazi compattati.'''
    pulito = re.sub(r'[\W_]+', ' ', testo_cella(testo).upper())
    return ' '.join(pulito.split())


def normalizza_um(um):
    '''
    Unità di misura confrontabile: maiuscole, senza spazi e punti, con
    m2 e m² ridotti a MQ e m3 e m³ a MC.
    '''
    testo = ''.join(testo_cella(um).upper().split()).replace('.', '')
    testo = re.sub(r'(?<![A-Z0-9])M[²2](?![0-9])', 'MQ', testo)
    testo = re.sub(r'(?<![A-Z0-9])M[³3](?![0-9])', 'MC', testo)
    return testo


def similarita(testo_a, testo_b):
    '''Similarità 0..1 tra due descrizioni, dopo normalizzazione.'''
    norm_a = normalizza_testo(testo_a)
    norm_b = normalizza_testo(testo_b)
    if norm_a == norm_b:
        return 1.0
    if not norm_a or not norm_b:
        return 0.0
    return difflib.SequenceMatcher(None, norm_a, norm_b, autojunk=False).ratio()


def classifica(sim, um_uguale, soglia_verde, soglia_giallo):
    '''
    VERDE, GIALLO o ROSSO. Unità di misura diversa: ROSSO a prescindere
    dalla similarità.
    '''
    if not um_uguale:
        return ROSSO
    if sim >= soglia_verde - _EPSILON:
        return VERDE
    if sim >= soglia_giallo - _EPSILON:
        return GIALLO
    return ROSSO


def indice_codici(righe):
    '''Mappa codice (maiuscolo) -> indice della prima riga che lo contiene.'''
    indice = {}
    for i, riga in enumerate(righe):
        codice = testo_cella(riga[0]).strip().upper()
        if codice and codice not in indice:
            indice[codice] = i
    return indice


def indice_suffissi(righe):
    '''
    Mappa suffisso (maiuscolo) -> indice di riga, più l'insieme dei suffissi
    che compaiono su più righe (ambigui, mai usati per sostituire).
    '''
    indice = {}
    ambigui = set()
    for i, riga in enumerate(righe):
        codice = testo_cella(riga[0]).strip()
        if not codice:
            continue
        chiave = taglia_prefisso(codice).upper()
        if chiave in indice:
            ambigui.add(chiave)
        else:
            indice[chiave] = i
    return indice, ambigui


def formatta_prezzo(valore):
    '''Prezzo in formato italiano (1.234,56); stringa vuota se non numerico.'''
    if isinstance(valore, bool) or not isinstance(valore, (int, float)):
        return ''
    testo = f'{valore:,.2f}'
    return testo.replace(',', '#').replace('.', ',').replace('#', '.')


def formatta_similarita(sim):
    '''Similarità con due decimali e virgola.'''
    return f'{sim:.2f}'.replace('.', ',')


def _valore(riga, colonna):
    if riga is None or colonna >= len(riga):
        return ''
    return riga[colonna]


def confronta_voci(voci, righe_doc, righe_prezzario, soglia_verde,
                   soglia_giallo, esclusi=frozenset()):
    '''
    Confronta le voci con il nuovo prezzario.

    voci            : lista di Voce
    righe_doc       : righe A:J di Elenco Prezzi del documento
    righe_prezzario : righe A:J di Elenco Prezzi del nuovo prezzario
    esclusi         : codici (maiuscoli) da non toccare (analisi e costi
                      elementari collegati)

    Restituisce un Risultato. Nessun accesso UNO.
    '''
    idx_doc = indice_codici(righe_doc)
    idx_prez, ambigui = indice_suffissi(righe_prezzario)
    esiti = []
    allineate = []
    non_trovate = []
    ambigue = []
    escluse = []
    for voce in voci:
        codice = voce.codice.strip()
        if codice.upper() in esclusi:
            escluse.append(voce)
            continue
        chiave = taglia_prefisso(codice).upper()
        if chiave in ambigui:
            ambigue.append(voce)
            continue
        i_nuovo = idx_prez.get(chiave)
        if i_nuovo is None:
            non_trovate.append(voce)
            continue
        nuova = righe_prezzario[i_nuovo]
        codice_nuovo = testo_cella(nuova[0]).strip()
        if codice_nuovo.upper() == codice.upper():
            allineate.append(voce)
            continue
        i_vecchio = idx_doc.get(codice.upper())
        vecchia = righe_doc[i_vecchio] if i_vecchio is not None else None
        desc_vecchia = testo_cella(_valore(vecchia, 1))
        desc_nuova = testo_cella(_valore(nuova, 1))
        um_vecchia = testo_cella(_valore(vecchia, 2))
        um_nuova = testo_cella(_valore(nuova, 2))
        nota = ''
        if vecchia is None:
            sim = 0.0
            um_uguale = False
            nota = 'codice assente da Elenco Prezzi del documento'
        else:
            sim = similarita(desc_vecchia, desc_nuova)
            um_uguale = normalizza_um(um_vecchia) == normalizza_um(um_nuova)
            if not um_uguale:
                nota = 'unità di misura diversa'
        colore = classifica(sim, um_uguale, soglia_verde, soglia_giallo)
        esiti.append(Esito(
            voce, codice_nuovo, colore, sim, um_vecchia, um_nuova,
            desc_vecchia, desc_nuova, _valore(vecchia, 4), _valore(nuova, 4),
            nota, i_nuovo))
    return Risultato(esiti, allineate, non_trovate, ambigue, escluse)


def riepilogo_conteggi(risultato):
    '''Conteggi per colore e per categoria, per il dialogo e il messaggio.'''
    per_colore = Counter(e.colore for e in risultato.esiti)
    return {
        VERDE: per_colore[VERDE],
        GIALLO: per_colore[GIALLO],
        ROSSO: per_colore[ROSSO],
        'allineate': len(risultato.allineate),
        'non_trovate': len(risultato.non_trovate),
        'ambigue': len(risultato.ambigue),
        'escluse': len(risultato.escluse),
    }


_PRIORITA_COLORE = {ROSSO: 0, GIALLO: 1, VERDE: 2}


def ordina_per_colore(esiti):
    '''ROSSO, GIALLO, VERDE; a parità resta l'ordine di lettura (stabile).'''
    return sorted(esiti, key=lambda e: _PRIORITA_COLORE[e.colore])


def preselezione(esiti):
    '''Lista di booleani: solo le voci VERDE sono preselezionate.'''
    return [e.colore == VERDE for e in esiti]


def etichetta_voce(voce):
    '''Testo breve per la colonna "Voce" del dialogo.'''
    numero = voce.numero.strip()
    base = f'{voce.foglio} {numero}' if numero else voce.foglio
    return f'{base}  {voce.codice}'


def una_riga(testo, massimo=90):
    '''Testo su una sola riga, troncato, per le celle della tabella.'''
    compatto = ' '.join(testo_cella(testo).split())
    if len(compatto) > massimo:
        return compatto[:massimo - 3].rstrip() + '...'
    return compatto


def righe_tabella(esiti, spuntate):
    '''Righe del dialogo: spunta, voce, colore, similarità, testi, prezzi.'''
    righe = []
    for esito, spunta in zip(esiti, spuntate):
        righe.append((
            '[x]' if spunta else '[ ]',
            etichetta_voce(esito.voce),
            esito.colore,
            formatta_similarita(esito.similarita),
            una_riga(esito.desc_vecchia),
            una_riga(esito.desc_nuova),
            formatta_prezzo(esito.prezzo_vecchio),
            formatta_prezzo(esito.prezzo_nuovo),
        ))
    return righe


def suggerimento_riga(esito):
    '''Testo completo (tooltip) di una riga della tabella.'''
    parti = [f'{esito.voce.codice}  ->  {esito.codice_nuovo}']
    if esito.nota:
        parti.append(f'Nota: {esito.nota}')
    parti.append(f'Vecchia ({esito.um_vecchia}): {esito.desc_vecchia}')
    parti.append(f'Nuova ({esito.um_nuova}): {esito.desc_nuova}')
    return '\n\n'.join(parti)


INTESTAZIONE_REPORT = (
    'Tipo', 'Foglio', 'Riga', 'Voce', 'Codice vecchio', 'Codice nuovo',
    'Colore', 'Similarità', 'UM vecchia', 'UM nuova', 'Descrizione vecchia',
    'Descrizione nuova', 'Prezzo vecchio', 'Prezzo nuovo', 'Sostituita',
    'Nota')


def righe_report(risultato, sostituite):
    '''
    Righe del report delle divergenze: voci non trovate, ambigue e voci
    trovate ma con descrizione o unità di misura diversa. Le voci identiche
    e già allineate non compaiono.

    sostituite : insieme di (foglio, riga) effettivamente sostituite
    '''
    righe = []
    for esito in risultato.esiti:
        identica = esito.similarita >= 1.0 - _EPSILON and not esito.nota
        if identica:
            continue
        voce = esito.voce
        righe.append((
            'DIVERSA', voce.foglio, voce.riga + 1, voce.numero, voce.codice,
            esito.codice_nuovo, esito.colore,
            formatta_similarita(esito.similarita), esito.um_vecchia,
            esito.um_nuova, una_riga(esito.desc_vecchia, 100000),
            una_riga(esito.desc_nuova, 100000),
            formatta_prezzo(esito.prezzo_vecchio),
            formatta_prezzo(esito.prezzo_nuovo),
            'SI' if (voce.foglio, voce.riga) in sostituite else 'NO',
            esito.nota))
    for tipo, elenco, nota in (
            ('NON_TROVATA', risultato.non_trovate,
             'suffisso assente dal nuovo prezzario'),
            ('AMBIGUA', risultato.ambigue,
             'suffisso presente più volte nel nuovo prezzario')):
        for voce in elenco:
            righe.append((
                tipo, voce.foglio, voce.riga + 1, voce.numero, voce.codice,
                '', '', '', '', '', '', '', '', '', 'NO', nota))
    return righe


def codifica_sfondo(stato, colore, trasparente):
    '''Sfondo di una cella come testo: "stato|colore|trasparente".'''
    return f'{stato}|{int(colore)}|{1 if trasparente else 0}'


def decodifica_sfondo(testo):
    '''Inverso di codifica_sfondo: (stato, colore, trasparente) o None.'''
    try:
        stato, colore, trasparente = str(testo).split('|')
        return stato, int(colore), trasparente == '1'
    except ValueError:
        return None


def unisci_registri(esistente, nuovo):
    '''
    Unione ricorsiva di due registri (dizionari annidati) senza
    sovrascrivere quanto già registrato: vale il primo valore salvato.
    '''
    risultato = dict(esistente)
    for chiave, valore in nuovo.items():
        if chiave not in risultato:
            risultato[chiave] = valore
        elif isinstance(risultato[chiave], dict) and isinstance(valore, dict):
            risultato[chiave] = unisci_registri(risultato[chiave], valore)
    return risultato


def nome_report(nome_documento, marca_temporale):
    '''Nome del file di report, con marca temporale per non sovrascrivere.'''
    base = nome_documento or 'documento'
    return f'{base}_CdPBis_{marca_temporale}.csv'


def scrivi_report(percorso, righe):
    '''Scrive il CSV (UTF-8 con BOM, separatore ";") per Excel e Calc.'''
    with open(percorso, 'w', encoding='utf-8-sig', newline='') as f:
        scrittore = csv.writer(f, delimiter=';')
        scrittore.writerow(INTESTAZIONE_REPORT)
        scrittore.writerows(righe)
    return percorso


########################################################################
# Lettura da Calc (UNO)
########################################################################

def _fine_selezione(oDoc, oSheet):
    '''Intervalli di righe selezionati (inizio, fine), limitati all'area usata.'''
    selezione = oDoc.getCurrentSelection()
    try:
        indirizzi = list(selezione.getRangeAddresses())
    except AttributeError:
        try:
            indirizzi = [selezione.getRangeAddress()]
        except AttributeError:
            return []
    ultima = SheetUtils.getLastUsedRow(oSheet)
    intervalli = []
    for ind in indirizzi:
        fine = min(ind.EndRow, ultima)
        if ind.StartRow <= fine:
            intervalli.append((ind.StartRow, fine))
    return intervalli


def leggi_voci_selezionate(oDoc, oSheet):
    '''
    Voci di COMPUTO, VARIANTE o CONTABILITA toccate dalla selezione.
    Restituisce (lista di Voce, insieme dei codici con formula in colonna B).
    Il codice sta in colonna B, nella riga sotto l'attributo iniziale.
    '''
    voci = []
    con_formula = set()
    viste = set()
    for inizio, fine in _fine_selezione(oDoc, oSheet):
        riga = inizio
        while riga <= fine:
            intervallo = LeenoComputo.circoscriveVoceComputo(oSheet, riga)
            if not intervallo:
                riga += 1
                continue
            sr = intervallo.RangeAddress.StartRow
            er = intervallo.RangeAddress.EndRow
            if sr > fine:
                break
            riga = er + 1
            if sr in viste:
                continue
            viste.add(sr)
            r_codice = sr + 1
            cella = oSheet.getCellByPosition(1, r_codice)
            codice = cella.String.strip()
            if not codice:
                continue
            if cella.Formula.startswith('='):
                con_formula.add(codice.upper())
            numero = oSheet.getCellByPosition(0, r_codice).String
            voci.append(Voce(oSheet.Name, r_codice, codice, numero))
    return voci, con_formula


def leggi_elenco_prezzi(oDoc):
    '''
    Elenco Prezzi in lettura a blocchi: (righe A:J, formule colonna B).
    None se il documento non ha il foglio. Una sola lettura per blocco.
    '''
    if not oDoc.getSheets().hasByName('Elenco Prezzi'):
        return None
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    ultima = SheetUtils.getLastUsedRow(oSheet)
    if ultima < _EP_PRIMA_RIGA:
        return [], []
    dati = oSheet.getCellRangeByPosition(
        0, _EP_PRIMA_RIGA, _EP_ULTIMA_COL, ultima).getDataArray()
    formule_b = oSheet.getCellRangeByPosition(
        1, _EP_PRIMA_RIGA, 1, ultima).getFormulaArray()
    return list(dati), [str(r[0]) for r in formule_b]


def _indice_colonna_a(oSheet, prima_riga=2):
    '''Mappa codice (maiuscolo) -> riga, colonna A, una sola lettura.'''
    ultima = SheetUtils.getLastUsedRow(oSheet)
    if ultima < prima_riga:
        return {}
    dati = oSheet.getCellRangeByPosition(0, prima_riga, 0, ultima).getDataArray()
    indice = {}
    for i, riga in enumerate(dati):
        codice = testo_cella(riga[0]).strip().upper()
        if codice and codice not in indice:
            indice[codice] = prima_riga + i
    return indice


def codici_analisi(righe_doc, formule_b):
    '''Codici di Elenco Prezzi con formula in colonna B (voci con analisi).'''
    analisi = set()
    for riga, formula in zip(righe_doc, formule_b):
        if formula.startswith('='):
            codice = testo_cella(riga[0]).strip().upper()
            if codice:
                analisi.add(codice)
    return analisi


def costi_elementari_collegati(oDoc, codici, analisi_ep):
    '''
    Codici dei costi elementari che compaiono nelle analisi dei codici dati,
    seguendo anche le analisi annidate (stessa logica di invia_voce).
    analisi_ep: codici di Elenco Prezzi che sono a loro volta analisi.
    '''
    if not codici or not oDoc.getSheets().hasByName('Analisi di Prezzo'):
        return set()
    oAP = oDoc.getSheets().getByName('Analisi di Prezzo')
    mappa_ap = _indice_colonna_a(oAP)
    trovati = set()
    coda = list(codici)
    visti = set(coda)
    while coda:
        codice = coda.pop(0)
        riga_ap = mappa_ap.get(codice)
        if riga_ap is None:
            continue
        blocco = LeenoAnalysis.circoscriveAnalisi(oAP, riga_ap)
        sr = blocco.RangeAddress.StartRow
        er = blocco.RangeAddress.EndRow
        col_a = oAP.getCellRangeByPosition(0, sr, 0, er).getDataArray()
        for i, riga in enumerate(col_a):
            ce = testo_cella(riga[0]).strip().upper()
            if not ce:
                continue
            if oAP.getCellByPosition(0, sr + i).CellStyle != 'An-lavoraz-Cod-sx':
                continue
            trovati.add(ce)
            if ce not in visti and ce in analisi_ep:
                visti.add(ce)
                coda.append(ce)
    return trovati


########################################################################
# Apertura del prezzario
########################################################################

def _documento_visibile(oDocumento):
    '''True se il documento ha una finestra visibile.'''
    controller = getattr(oDocumento, 'CurrentController', None)
    if controller is None:
        return False
    try:
        return bool(controller.Frame.ContainerWindow.isVisible())
    except Exception as _exc:
        LeenoUtils.log_errore('apri_prezzario: visibilità finestra', _exc)
        return False


def apri_prezzario(percorso):
    '''
    Apre il prezzario in modo nascosto, oppure riusa il documento già aperto.
    Restituisce (documento, da_chiudere). Se è già aperto in una finestra
    visibile non va chiuso; un documento aperto ma nascosto (residuo) sì.
    '''
    aperto = LeenoUtils.findOpenDocument(percorso)
    if aperto is not None:
        return aperto, not _documento_visibile(aperto)
    oPrez = DocUtils.loadDocument(percorso, Hidden=True)
    return oPrez, oPrez is not None


def chiudi_prezzario(oPrez):
    '''Chiude il prezzario nascosto senza salvare.'''
    try:
        oPrez.setModified(False)
        oPrez.close(True)
    except Exception as _exc:
        LeenoUtils.log_errore('chiudi_prezzario: chiusura', _exc)


########################################################################
# Dialogo di conferma
########################################################################

_INTESTAZIONI_COLONNE = (
    ('Sost.', 24, 0), ('Voce', 112, 0), ('Esito', 34, 0), ('Simil.', 28, 0),
    ('Descrizione vecchia', 90, 1), ('Descrizione nuova', 90, 1),
    ('Prezzo vecchio', 46, 0), ('Prezzo nuovo', 44, 0))


def _crea_ascoltatori(oDlg, oGrid, dati, stato, esiti):
    '''
    Collega i pulsanti e il clic sulla prima colonna. Le classi sono definite
    qui (non a livello di modulo) perché dipendono da UNO.
    '''
    import unohelper
    from com.sun.star.awt import XMouseListener, XActionListener

    def imposta(riga, valore):
        stato[riga] = valore
        dati.updateCellData(0, riga, '[x]' if valore else '[ ]')

    class _Clic(unohelper.Base, XMouseListener):
        def mouseReleased(self, ev):
            try:
                riga = oGrid.getRowAtPoint(ev.X, ev.Y)
                colonna = oGrid.getColumnAtPoint(ev.X, ev.Y)
                if riga >= 0 and colonna == 0:
                    imposta(riga, not stato[riga])
            except Exception as _exc:
                LeenoUtils.log_errore('dialogo CdP bis: clic sulla tabella', _exc)

        def mousePressed(self, ev):
            return None

        def mouseEntered(self, ev):
            return None

        def mouseExited(self, ev):
            return None

        def disposing(self, ev):
            return None

    class _Azione(unohelper.Base, XActionListener):
        def actionPerformed(self, ev):
            try:
                comando = ev.ActionCommand
                for i, esito in enumerate(esiti):
                    if comando == 'tutte':
                        valore = True
                    elif comando == 'nessuna':
                        valore = False
                    else:
                        valore = esito.colore == VERDE
                    imposta(i, valore)
            except Exception as _exc:
                LeenoUtils.log_errore('dialogo CdP bis: selezione rapida', _exc)

        def disposing(self, ev):
            return None

    clic = _Clic()
    azione = _Azione()
    oGrid.addMouseListener(clic)
    for nome, comando in (('btnTutte', 'tutte'), ('btnNessuna', 'nessuna'),
                          ('btnVerdi', 'verdi')):
        pulsante = oDlg.getControl(nome)
        pulsante.setActionCommand(comando)
        pulsante.addActionListener(azione)
    return clic, azione


def mostra_dialogo_conferma(esiti, conteggi):
    '''
    Dialogo unico con tabella delle sostituzioni proposte.
    Restituisce la lista di booleani (una per esito, nello stesso ordine)
    oppure None se l'utente preme Annulla.
    '''
    ctx = LeenoUtils.getComponentContext()
    dp = ctx.ServiceManager.createInstance('com.sun.star.awt.DialogProvider')
    oDlg = dp.createDialog(
        'vnd.sun.star.script:UltimusFree2.DlgCdPBis?language=Basic&location=application')
    stato = preselezione(esiti)
    oGrid = _costruisci_tabella(oDlg, esiti, stato)
    dati = oDlg.getModel().getByName('Griglia').GridDataModel
    oDlg.getModel().getByName('lblRiepilogo').Label = (
        f"Voci proposte: {len(esiti)}   VERDE {conteggi[VERDE]}   "
        f"GIALLO {conteggi[GIALLO]}   ROSSO {conteggi[ROSSO]}\n"
        f"Già allineate {conteggi['allineate']}   "
        f"non trovate {conteggi['non_trovate']}   "
        f"ambigue {conteggi['ambigue']}   escluse {conteggi['escluse']}")
    clic, azione = _crea_ascoltatori(oDlg, oGrid, dati, stato, esiti)
    try:
        esito = oDlg.execute()
    finally:
        oDlg.dispose()
    if esito != 1:
        return None
    return list(stato)


def _costruisci_tabella(oDlg, esiti, stato):
    '''Crea il controllo tabella nel dialogo e lo riempie in blocco.'''
    oModello = oDlg.getModel()
    griglia = oModello.createInstance('com.sun.star.awt.grid.UnoControlGridModel')
    griglia.PositionX = 6
    griglia.PositionY = 30
    griglia.Width = 428
    griglia.Height = 176
    griglia.ShowRowHeader = False
    griglia.ShowColumnHeader = True
    griglia.HScroll = True
    griglia.VScroll = True
    colonne = griglia.ColumnModel
    for titolo, larghezza, flessibilita in _INTESTAZIONI_COLONNE:
        colonna = colonne.createColumn()
        colonna.Title = titolo
        colonna.ColumnWidth = larghezza
        colonna.Flexibility = flessibilita
        colonne.addColumn(colonna)
    oModello.insertByName('Griglia', griglia)
    dati = griglia.GridDataModel
    righe = righe_tabella(esiti, stato)
    if righe:
        dati.addRows(tuple('' for _ in righe), tuple(righe))
        for i, esito in enumerate(esiti):
            dati.updateRowToolTip(i, suggerimento_riga(esito))
    return oDlg.getControl('Griglia')


########################################################################
# Applicazione
########################################################################

def _inserisci_voci_prezzario(oDoc, righe_nuove):
    '''
    Aggiunge in Elenco Prezzi le righe (colonne A:H) del nuovo prezzario,
    in un solo inserimento, con stili e formule L:Z della riga modello.
    Restituisce il numero di righe aggiunte.
    '''
    n = len(righe_nuove)
    if n == 0:
        return 0
    oEP = oDoc.getSheets().getByName('Elenco Prezzi')
    oEP.getRows().insertByIndex(_EP_RIGA_INSERIMENTO, n)
    blocco = tuple(tuple(r[:8]) for r in righe_nuove)
    prima = _EP_RIGA_INSERIMENTO
    ultima = _EP_RIGA_INSERIMENTO + n - 1
    oEP.getCellRangeByPosition(0, prima, 7, ultima).setDataArray(blocco)
    origine = oEP.getCellRangeByPosition(
        11, _EP_RIGA_MODELLO, 25, _EP_RIGA_MODELLO).getRangeAddress()
    for riga in range(prima, ultima + 1):
        destinazione = oEP.getCellByPosition(11, riga).getCellAddress()
        oEP.copyRange(destinazione, origine)
    import pyleeno as PL
    PL.stileCelleElencoPrezzi(oEP, prima, ultima)
    return n


def righe_da_aggiungere(scelti, righe_prezzario, righe_doc):
    '''
    Righe del prezzario da aggiungere a Elenco Prezzi per le voci scelte:
    una sola volta per codice nuovo e solo se il codice non esiste già.
    '''
    esistenti = set(indice_codici(righe_doc))
    nuove = []
    for esito in scelti:
        chiave = esito.codice_nuovo.upper()
        if chiave in esistenti:
            continue
        esistenti.add(chiave)
        nuove.append(righe_prezzario[esito.riga_prezzario])
    return nuove


_PROPRIETA_SFONDI = 'LeenoCdPBis.sfondi'
_PREFISSO_STILE_CODICE = 'comp Art-EP'


def _stato_sfondo(cella):
    '''(stato proprietà, colore, trasparente) dello sfondo di una cella.'''
    return (cella.getPropertyState('CellBackColor').value,
            cella.CellBackColor,
            bool(cella.IsCellBackgroundTransparent))


def _colori_cdpbis():
    return (LeenoConfig.COLORE_CDPBIS_VERDE, LeenoConfig.COLORE_CDPBIS_GIALLO,
            LeenoConfig.COLORE_CDPBIS_ROSSO)


def _sfondi_standard_foglio(oSheet):
    '''
    Sfondo più frequente della cella del codice, per stile, tra le celle
    non evidenziate da questo modulo: {stile: sfondo codificato}.
    '''
    nostri = _colori_cdpbis()
    conteggi = {}
    for riga in range(SheetUtils.getLastUsedRow(oSheet) + 1):
        cella = oSheet.getCellByPosition(1, riga)
        stile = cella.CellStyle
        if not stile.startswith(_PREFISSO_STILE_CODICE):
            continue
        if cella.CellBackColor in nostri:
            continue
        conteggi.setdefault(stile, Counter())[codifica_sfondo(*_stato_sfondo(cella))] += 1
    return {stile: c.most_common(1)[0][0] for stile, c in conteggi.items()}


def _leggi_sfondi_registrati(oDoc):
    '''Registro degli sfondi nel documento ({} se assente o illeggibile).'''
    proprieta = oDoc.getDocumentProperties().getUserDefinedProperties()
    if not proprieta.getPropertySetInfo().hasPropertyByName(_PROPRIETA_SFONDI):
        return {}
    try:
        return json.loads(proprieta.getPropertyValue(_PROPRIETA_SFONDI))
    except ValueError as _exc:
        LeenoUtils.log_errore('LeenoCdPBis: sfondi registrati illeggibili', _exc)
        return {}


def _registra_sfondi(oDoc, esiti):
    '''
    Registra nel documento, prima di colorare, lo sfondo originale delle
    celle codice: lo sfondo standard per foglio e stile, più le eccezioni
    cella per cella (celle con uno sfondo diverso dallo standard, con il
    codice presente in quel momento per riconoscerle al ripristino).
    Così la rimozione ripristina lo stato originale anche quando quasi
    tutte le voci sono evidenziate. Il primo valore registrato non viene
    sovrascritto da evidenziazioni successive.
    '''
    nostri = _colori_cdpbis()
    standard = {}
    eccezioni = {}
    for nome in {e.voce.foglio for e in esiti}:
        oSheet = oDoc.getSheets().getByName(nome)
        standard[nome] = _sfondi_standard_foglio(oSheet)
        eccezioni[nome] = {}
    for esito in esiti:
        voce = esito.voce
        cella = oDoc.getSheets().getByName(voce.foglio).getCellByPosition(1, voce.riga)
        if cella.CellBackColor in nostri:
            continue
        sfondo = codifica_sfondo(*_stato_sfondo(cella))
        if sfondo != standard[voce.foglio].get(cella.CellStyle):
            eccezioni[voce.foglio][str(voce.riga)] = [cella.String, sfondo]
    registro = unisci_registri(
        _leggi_sfondi_registrati(oDoc),
        {'standard': standard, 'eccezioni': eccezioni})
    proprieta = oDoc.getDocumentProperties().getUserDefinedProperties()
    testo = json.dumps(registro, sort_keys=True)
    if proprieta.getPropertySetInfo().hasPropertyByName(_PROPRIETA_SFONDI):
        proprieta.setPropertyValue(_PROPRIETA_SFONDI, testo)
    else:
        proprieta.addProperty(_PROPRIETA_SFONDI, 128, testo)


def _cancella_sfondi_registrati(oDoc):
    proprieta = oDoc.getDocumentProperties().getUserDefinedProperties()
    if proprieta.getPropertySetInfo().hasPropertyByName(_PROPRIETA_SFONDI):
        proprieta.removeProperty(_PROPRIETA_SFONDI)


def _evidenzia(oDoc, esiti):
    '''Colora la cella del codice di tutte le voci proposte (anche rifiutate).'''
    colori = {VERDE: LeenoConfig.COLORE_CDPBIS_VERDE,
              GIALLO: LeenoConfig.COLORE_CDPBIS_GIALLO,
              ROSSO: LeenoConfig.COLORE_CDPBIS_ROSSO}
    _registra_sfondi(oDoc, esiti)
    for esito in esiti:
        colore = colori.get(esito.colore)
        if colore is None:
            continue
        voce = esito.voce
        oSheet = oDoc.getSheets().getByName(voce.foglio)
        oSheet.getCellByPosition(1, voce.riga).CellBackColor = colore


def applica_sostituzioni(oDoc, esiti, spuntate, righe_prezzario, righe_doc):
    '''
    Applica le sostituzioni spuntate: aggiunge le righe a Elenco Prezzi e
    cambia il codice nella cella B della voce (descrizione, unità di misura
    e prezzo arrivano da Elenco Prezzi con VLOOKUP). Colora GIALLO e ROSSO.
    Refresh disabilitato durante le scritture e ripristinato in finally.
    Restituisce (numero voci sostituite, righe aggiunte, insieme sostituite).
    '''
    scelti = [e for e, s in zip(esiti, spuntate) if s]
    sostituite = set()
    aggiunte = 0
    LeenoUtils.DocumentRefresh(False, oDoc=oDoc)
    try:
        aggiunte = _inserisci_voci_prezzario(
            oDoc, righe_da_aggiungere(scelti, righe_prezzario, righe_doc))
        for esito in scelti:
            voce = esito.voce
            oSheet = oDoc.getSheets().getByName(voce.foglio)
            oSheet.getCellByPosition(1, voce.riga).String = esito.codice_nuovo
            sostituite.add((voce.foglio, voce.riga))
        _evidenzia(oDoc, esiti)
    finally:
        LeenoUtils.DocumentRefresh(True, oDoc=oDoc)
    _adatta_altezze(oDoc, scelti, aggiunte)
    return len(sostituite), aggiunte, sostituite


def _adatta_altezze(oDoc, scelti, aggiunte):
    '''
    Altezza ottimale delle righe di descrizione cambiate e delle nuove righe
    di Elenco Prezzi. Va fatta a calcolo ripristinato: la descrizione è una
    formula che dipende dal codice appena scritto.
    '''
    oDoc.lockControllers()
    try:
        for esito in scelti:
            voce = esito.voce
            oSheet = oDoc.getSheets().getByName(voce.foglio)
            oSheet.getCellRangeByPosition(0, voce.riga, 0, voce.riga).Rows.OptimalHeight = True
        if aggiunte:
            oEP = oDoc.getSheets().getByName('Elenco Prezzi')
            oEP.getCellRangeByPosition(
                0, _EP_RIGA_INSERIMENTO, 0,
                _EP_RIGA_INSERIMENTO + aggiunte - 1).Rows.OptimalHeight = True
    except Exception as _exc:
        LeenoUtils.log_errore('applica_sostituzioni: altezza righe', _exc)
    finally:
        oDoc.unlockControllers()


########################################################################
# Rimozione delle evidenziazioni
########################################################################

def _ripristina_sfondo(cella, sfondo):
    '''Riporta lo sfondo della cella a (stato, colore, trasparente).'''
    stato, colore, trasparente = sfondo
    if stato == 'DEFAULT_VALUE':
        cella.setPropertyToDefault('CellBackColor')
    elif trasparente:
        cella.CellBackColor = -1
    else:
        cella.CellBackColor = colore


def rimuovi_evidenziazioni(oDoc):
    '''
    Toglie le evidenziazioni di LeenoCdPBis da COMPUTO, VARIANTE e CONTABILITA
    e ripristina lo sfondo originale della cella del codice.

    Lo sfondo di partenza è quello registrato nel documento prima di
    colorare (eccezioni cella per cella, poi standard del foglio). In assenza (documento evidenziato in altro modo) si usa il
    più frequente tra le celle non evidenziate dello stesso foglio e stile;
    se non ce ne sono, lo stile della cella. In CONTABILITA la cella ha uno
    sfondo trasparente diretto diverso da quello dello stile, quindi il
    solo ripristino dello stile cambierebbe l'aspetto del foglio.
    Restituisce il numero di celle ripulite.
    '''
    nostri = _colori_cdpbis()
    registro = _leggi_sfondi_registrati(oDoc)
    ripulite = 0
    for nome in FOGLI_VOCI:
        if not oDoc.getSheets().hasByName(nome):
            continue
        oSheet = oDoc.getSheets().getByName(nome)
        ultima = SheetUtils.getLastUsedRow(oSheet)
        evidenziate = []
        for riga in range(ultima + 1):
            cella = oSheet.getCellByPosition(1, riga)
            if cella.CellBackColor in nostri:
                evidenziate.append((riga, cella.CellStyle))
        if not evidenziate:
            continue
        sfondi = dict(_sfondi_standard_foglio(oSheet))
        sfondi.update(registro.get('standard', {}).get(nome, {}))
        eccezioni = registro.get('eccezioni', {}).get(nome, {})
        for riga, stile in evidenziate:
            cella = oSheet.getCellByPosition(1, riga)
            sfondo = None
            eccezione = eccezioni.get(str(riga))
            if eccezione and eccezione[0] == cella.String:
                sfondo = decodifica_sfondo(eccezione[1])
            if sfondo is None:
                sfondo = decodifica_sfondo(sfondi.get(stile))
            if sfondo is None:
                sfondo = ('DEFAULT_VALUE', -1, False)
            _ripristina_sfondo(cella, sfondo)
            ripulite += 1
    _cancella_sfondi_registrati(oDoc)
    return ripulite


########################################################################
# Punti di ingresso (menu)
########################################################################

def _percorso_report(oDoc, percorso_prezzario):
    '''Cartella e nome del report: accanto al documento, se è salvato.'''
    url = oDoc.getURL()
    marca = time.strftime('%Y%m%d_%H%M%S')
    if url:
        percorso_doc = uno.fileUrlToSystemPath(url)
        cartella = os.path.dirname(percorso_doc)
        nome = os.path.splitext(os.path.basename(percorso_doc))[0]
    else:
        cartella = os.path.dirname(percorso_prezzario)
        nome = ''
    return os.path.join(cartella, nome_report(nome, marca))


def _testo_riepilogo(conteggi, sostituite, aggiunte, percorso_report):
    righe = [
        f'Voci sostituite: {sostituite}',
        f'Righe aggiunte a Elenco Prezzi: {aggiunte}',
        f"Già allineate: {conteggi['allineate']}   "
        f"non trovate: {conteggi['non_trovate']}   "
        f"ambigue: {conteggi['ambigue']}   escluse: {conteggi['escluse']}",
    ]
    if percorso_report:
        righe.append(f'Report: {percorso_report}')
    return '\n'.join(righe)


def aggiorna_voci_prezzario(oDoc):
    '''
    Flusso completo sul documento oDoc (passato esplicitamente, perché
    getDocument() può restituire None dopo un dialogo modale).
    '''
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name not in FOGLI_VOCI:
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Seleziona le voci in COMPUTO, VARIANTE o CONTABILITA.')
        return False
    voci, con_formula = leggi_voci_selezionate(oDoc, oSheet)
    if not voci:
        Dialogs.Exclamation(Title='ATTENZIONE!', Text='Nessuna voce selezionata.')
        return False

    percorso = Dialogs.FileSelect('Seleziona il nuovo prezzario', '*.ods')
    if not percorso:
        return False
    if LeenoUtils.findOpenDocument(percorso) is oDoc:
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Il prezzario coincide con il documento corrente.')
        return False

    oPrez, da_chiudere = apri_prezzario(percorso)
    if oPrez is None:
        Dialogs.Exclamation(Title='ATTENZIONE!', Text='Impossibile aprire il prezzario.')
        return False
    try:
        letto_prez = leggi_elenco_prezzi(oPrez)
        if letto_prez is None:
            Dialogs.Exclamation(
                Title='ATTENZIONE!',
                Text="Il prezzario non ha il foglio 'Elenco Prezzi'.")
            return False
        righe_prez = letto_prez[0]

        oDoc = LeenoUtils.resolve_document(oDoc)
        if oDoc is None:
            LeenoUtils.log_errore(
                'aggiorna_voci_prezzario: documento non più disponibile')
            return False
        letto_doc = leggi_elenco_prezzi(oDoc)
        if letto_doc is None:
            Dialogs.Exclamation(
                Title='ATTENZIONE!',
                Text="Il documento non ha il foglio 'Elenco Prezzi'.")
            return False
        righe_doc, formule_b = letto_doc

        analisi = codici_analisi(righe_doc, formule_b)
        analisi_scelte = {v.codice.upper() for v in voci
                          if v.codice.upper() in analisi or v.codice.upper() in con_formula}
        esclusi = set(analisi_scelte)
        esclusi |= costi_elementari_collegati(oDoc, analisi_scelte, analisi)

        risultato = confronta_voci(
            voci, righe_doc, righe_prez, LeenoConfig.CDPBIS_SOGLIA_VERDE,
            LeenoConfig.CDPBIS_SOGLIA_GIALLO, esclusi)
        conteggi = riepilogo_conteggi(risultato)
        esiti = ordina_per_colore(risultato.esiti)
        risultato = risultato._replace(esiti=esiti)

        if not esiti:
            percorso_report = _scrivi_report_sicuro(oDoc, percorso, risultato, set())
            Dialogs.Info(
                Title='Aggiornamento voci',
                Text='Nessuna voce da sostituire.\n'
                     + _testo_riepilogo(conteggi, 0, 0, percorso_report))
            return False

        spuntate = mostra_dialogo_conferma(esiti, conteggi)
        if spuntate is None:
            return False

        sostituite, aggiunte, insieme = applica_sostituzioni(
            oDoc, esiti, spuntate, righe_prez, righe_doc)
        percorso_report = _scrivi_report_sicuro(oDoc, percorso, risultato, insieme)
        Dialogs.Info(
            Title='Aggiornamento voci',
            Text=_testo_riepilogo(conteggi, sostituite, aggiunte, percorso_report))
        return True
    finally:
        if da_chiudere:
            chiudi_prezzario(oPrez)


def _scrivi_report_sicuro(oDoc, percorso_prezzario, risultato, sostituite):
    '''Scrive il report; in caso di errore lo registra e restituisce None.'''
    try:
        percorso = _percorso_report(oDoc, percorso_prezzario)
        return scrivi_report(percorso, righe_report(risultato, sostituite))
    except (OSError, ValueError) as _exc:
        LeenoUtils.log_errore('aggiorna_voci_prezzario: scrittura report', _exc)
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Report non scritto. Dettagli in leeno_errori.log.')
        return None


def MENU_aggiorna_voci_prezzario():
    '''
    Aggiorna le voci selezionate in COMPUTO, VARIANTE o CONTABILITA a un
    nuovo prezzario .ods con struttura LeenO.
    '''
    oDoc = LeenoUtils.getDocument()
    if oDoc is None:
        return
    try:
        aggiorna_voci_prezzario(oDoc)
    except Exception as _exc:
        LeenoUtils.log_errore('MENU_aggiorna_voci_prezzario', _exc)
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Operazione interrotta per un errore. '
                 'Dettagli in leeno_errori.log.')


def MENU_rimuovi_evidenziazioni_cdpbis():
    '''Rimuove le evidenziazioni lasciate dall'aggiornamento al prezzario.'''
    oDoc = LeenoUtils.getDocument()
    if oDoc is None:
        return
    LeenoUtils.DocumentRefresh(False, oDoc=oDoc)
    try:
        ripulite = rimuovi_evidenziazioni(oDoc)
    except Exception as _exc:
        LeenoUtils.log_errore('MENU_rimuovi_evidenziazioni_cdpbis', _exc)
        ripulite = None
    finally:
        LeenoUtils.DocumentRefresh(True, oDoc=oDoc)
    if ripulite is None:
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Operazione interrotta per un errore. '
                 'Dettagli in leeno_errori.log.')
    else:
        Dialogs.Info(
            Title='Evidenziazioni',
            Text=f'Celle ripristinate: {ripulite}')
