########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

'''
Modulo per la modifica delle impostazioni di LeenO
'''
import json
import os
import re

import uno
import unohelper
from com.sun.star.awt import XActionListener, XItemListener, XTextListener

import Dialogs
import LeenoConfig
import LeenoUtils
import DocUtils
import LeenoSheetUtils
import SheetUtils
import LeenoDialogs as DLG
import pyleeno as PL

def get_my_doc():
    return LeenoUtils.getDocument()

_JOBSETTINGSITEMS = (
    'progetto',
    'committente',
    'stazioneAppaltante',
    'rup',
    'progettista',
    'data',
    'revisione',
    'dataRevisione',
)

_PRINTSETTINGSITEMS = (
    'fileCopertine',
    'copertina',
    'intSx',
    'intCenter',
    'intDx',
    'ppSx',
    'ppCenter',
    'ppDx',
    'documenti',
)

_DOCSTRINGS = (
    '[COMMITTENTE]',
    '[DATA]',
    '[DATA_REVISIONE]',
    '[DATI_COMMITTENTE]',
    '[DATI_PROGETTISTA]',
    '[DIRETTORE_LAVORI]',
    '[FILE]',
    '[FIRME]',
    '[NUMERO_DOCUMENTO]',
    '[OGGETTO]',
    '[PAGINA]',
    '[PAGINE]',
    '[PROGETTISTA]',
    '[PROGETTO]',
    '[REVISIONE]',
    '[RUP]',
    '[SAL]',
    '[STAZIONE_APPALTANTE]',
)

# Segnaposto di intestazione e piè di pagina che si risolvono leggendo il foglio S2.
# Ogni voce: (segnaposto, etichette cercate in colonna B di S2, chiave 'Lavoro' di ripiego).
# Il valore di S2 prevale; se e' vuoto si usa l'impostazione di lavoro (MENU_JobSettings).
_SEGNAPOSTI_S2 = (
    ('[COMMITTENTE]', ('Committente',), 'committente'),
    ('[STAZIONE_APPALTANTE]', ('Stazione Appaltante',), 'stazioneAppaltante'),
    ('[PROGETTO]', ('OGGETTO',), 'progetto'),
    ('[RUP]', ('Responsabile del Progetto',), 'rup'),
    ('[PROGETTISTA]', ('Progettista',), 'progettista'),
    ('[DIRETTORE_LAVORI]', ('Direttore Lavori', 'Direttore dei Lavori'), None),
)

# Segnaposto senza dato corrispondente in S2: solo impostazioni di lavoro.
_SEGNAPOSTI_LAVORO = (
    ('[DATA]', 'data'),
    ('[REVISIONE]', 'revisione'),
    ('[DATA_REVISIONE]', 'dataRevisione'),
)

# Segnaposto [FIRME]: ovunque compaia nel piè di pagina, la riga delle firme viene aggiunta da
# sola in fondo (vedi piede_con_firme) e il segnaposto sparisce dal testo.
_SEGNAPOSTO_FIRME = '[FIRME]'

# Testo della riga delle firme: come nello stile Page_Style_Libretto_Misure2 del template.
_FIRME_TESTO = "L'IMPRESA\t\t\t\t\t\tIL DIRETTORE DEI LAVORI    "


def nome_file_documento(oDoc):
    """Nome del file del documento, senza percorso (segnaposto [FILE])."""
    return os.path.basename(oDoc.getURL().replace('%20', ' '))


def valori_segnaposto(oDoc, documento='', nSal=None, nomefile='', extra=None, lavoro=None):
    """
    Valori dei segnaposto di intestazione e piè di pagina per il documento oDoc.

    documento : titolo del documento prodotto ([OGGETTO]), es. 'Computo Metrico', 'Libretto delle Misure'
    nSal      : numero del SAL ([SAL]); None se non pertinente
    nomefile  : nome del file ([FILE]), vedi nome_file_documento
    extra     : altri segnaposto {'[NOME]': valore} che prevalgono su quelli calcolati
    lavoro    : impostazioni di lavoro; se None si caricano con loadJobSettings

    [PROGETTO] e' l'oggetto dei lavori (etichetta OGGETTO in S2); [OGGETTO] e' il titolo del
    documento. I dati di S2 si leggono per etichetta (colonna B, valore in colonna C) e
    prevalgono sulle impostazioni di lavoro. [PAGINA] e [PAGINE] non sono qui: diventano campi
    numerici del piè di pagina.
    """
    if lavoro is None:
        lavoro = loadJobSettings(oDoc)
    oS2 = oDoc.Sheets.getByName('S2')

    res = {}
    for segnaposto, etichette, chiave in _SEGNAPOSTI_S2:
        valore = ''
        for etichetta in etichette:
            valore = SheetUtils.leggi_dato_etichettato(oS2, etichetta)
            if valore:
                break
        if not valore and chiave:
            valore = str(lavoro.get(chiave, '') or '')
        res[segnaposto] = valore
    for segnaposto, chiave in _SEGNAPOSTI_LAVORO:
        res[segnaposto] = str(lavoro.get(chiave, '') or '')

    res['[OGGETTO]'] = documento
    res['[SAL]'] = '' if nSal is None else str(nSal)
    res['[FILE]'] = nomefile
    if extra:
        res.update(extra)
    return res


def sostituisci_segnaposto(testo, valori):
    """
    Sostituisce in testo i segnaposto presenti in valori, in un solo passaggio: un valore
    che contiene un segnaposto non viene espanso di nuovo. I segnaposto sconosciuti e
    [PAGINA] e [PAGINE] restano invariati.
    """
    return re.sub(r'\[[A-Z_]+\]', lambda m: valori.get(m.group(0), m.group(0)), testo)


# Documenti per i quali si configurano intestazione e piè di pagina: (chiave, titolo).
# Il titolo e' il valore di [OGGETTO].
_DOCUMENTI_STAMPA = (
    ('COMPUTO', 'Computo Metrico'),
    ('VARIANTE', 'Variante'),
    ('ELENCO', 'Elenco Prezzi'),
    ('ANALISI', 'Analisi di Prezzo'),
    ('LIBRETTO', 'Libretto delle Misure'),
    ('REGISTRO', 'Registro di Contabilità'),
    ('SAL', 'Stato di Avanzamento Lavori'),
    ('CDP', 'Certificato di Pagamento'),
)

# Il Certificato di Pagamento riporta i propri dati nel corpo: niente intestazione e piè di pagina.
_DOCUMENTI_SENZA_TESTATA = ('CDP',)

# Documenti per i quali il dialogo offre la casella della riga delle firme (il segnaposto
# [FIRME] si puo' comunque scrivere in qualsiasi piè di pagina).
_DOCUMENTI_CON_FIRME = ('LIBRETTO', 'REGISTRO')

# Le sei stringhe di intestazione (sinistra, centro, destra) e piè di pagina.
_TESTI_STAMPA = ('intSx', 'intCenter', 'intDx', 'ppSx', 'ppCenter', 'ppDx')

# Valori della versione con un solo set di stringhe per tutti i documenti.
_TESTI_PRECEDENTI = {
    'intSx': '[COMMITTENTE]',
    'intCenter': '',
    'intDx': '[PROGETTO]',
    'ppSx': '[OGGETTO]',
    'ppCenter': '',
    'ppDx': 'Pagina [PAGINA] di [PAGINE]',
}


def testi_predefiniti(chiave):
    """Le sei stringhe predefinite del documento `chiave` (copia indipendente)."""
    testi = {
        'intSx': 'Committente: [COMMITTENTE]',
        'intCenter': '[PROGETTO]',
        'intDx': '[OGGETTO]',
        'ppSx': 'realizzato con LeenO: [FILE]',
        'ppCenter': '',
        'ppDx': 'Pagina [PAGINA] di [PAGINE]',
    }
    if chiave in ('LIBRETTO', 'REGISTRO', 'SAL'):
        testi['intDx'] = '[OGGETTO] n. [SAL]'
    if chiave == 'LIBRETTO':
        testi['ppCenter'] = _SEGNAPOSTO_FIRME
    if chiave in _DOCUMENTI_SENZA_TESTATA:
        testi = dict.fromkeys(_TESTI_STAMPA, '')
    return testi


def _personalizzazioni(ps):
    """
    Stringhe personalizzate per documento, lette da ps['documenti'] (JSON): {chiave: {stringhe}}.
    Tollera dati non validi e ignora le voci sconosciute.
    """
    try:
        dati = json.loads(ps.get('documenti') or '{}')
    except (TypeError, ValueError):
        return {}
    if not isinstance(dati, dict):
        return {}
    res = {}
    for chiave, valori in dati.items():
        if isinstance(valori, dict):
            res[chiave] = {k: v for k, v in valori.items() if k in _TESTI_STAMPA and isinstance(v, str)}
    return res


def impostazioni_documento(ps, chiave):
    """
    Le sei stringhe di intestazione e piè di pagina del documento `chiave`.

    Parte dai predefiniti, applica le personalizzazioni della versione con un solo set di
    stringhe (se diverse dai vecchi valori) e infine quelle specifiche del documento.
    """
    testi = testi_predefiniti(chiave)
    if chiave in _DOCUMENTI_SENZA_TESTATA:
        return testi
    for k in _TESTI_STAMPA:
        valore = ps.get(k)
        if isinstance(valore, str) and valore != _TESTI_PRECEDENTI[k]:
            testi[k] = valore
    testi.update(_personalizzazioni(ps).get(chiave, {}))
    return testi


_PARTI_PIEDE = ('ppSx', 'ppCenter', 'ppDx')


def senza_firme(testo):
    """Il testo senza il segnaposto [FIRME] (e senza spazi o righe vuote lasciati in coda)."""
    if _SEGNAPOSTO_FIRME not in testo:
        return testo
    testo = testo.replace(_SEGNAPOSTO_FIRME, '')
    return testo.rstrip() if testo.strip() else ''


def ha_firme(testi):
    """True se una delle tre parti del piè di pagina contiene il segnaposto [FIRME]."""
    return any(_SEGNAPOSTO_FIRME in testi.get(k, '') for k in _PARTI_PIEDE)


def piede_con_firme(testi):
    """
    Le tre parti del piè di pagina (ppSx, ppCenter, ppDx) con la riga delle firme da sola in fondo.
    `testi`: le sei stringhe gia' risolte. Ritorna {'ppSx', 'ppCenter', 'ppDx'}. Se nessuna parte
    contiene [FIRME] le tre parti restano invariate.

    Le tre parti sono centrate verticalmente sulla stessa linea e le righe vuote hanno l'altezza
    delle altre (provato con LibreOffice): se hanno tutte lo stesso numero di righe, il testo
    resta in alto e le firme, ultima riga del centro, finiscono sotto "realizzato con", i numeri
    di pagina e ogni altra riga, senza sovrapporsi.
    """
    parti = {k: testi.get(k, '') for k in _PARTI_PIEDE}
    if not ha_firme(parti):
        return parti
    sx, centro, dx = (senza_firme(parti[k]) for k in _PARTI_PIEDE)

    def righe(testo):
        return testo.count('\n') + 1

    def completa(testo, n):
        return testo + '\n' * (n - righe(testo))

    occupate = [righe(t) for t in (sx, centro, dx) if t.strip()]
    if not occupate:
        return {'ppSx': sx, 'ppCenter': _FIRME_TESTO, 'ppDx': dx}
    n = max(occupate)
    return {
        'ppSx': completa(sx, n + 1),
        'ppCenter': completa(centro, n) + '\n' + _FIRME_TESTO,
        'ppDx': completa(dx, n + 1),
    }


def titolo_documento(chiave):
    """Titolo del documento `chiave` (valore di [OGGETTO]); stringa vuota se sconosciuto."""
    return dict(_DOCUMENTI_STAMPA).get(chiave, '')


def aggiorna_impostazioni_documenti(ps, valori):
    """
    Ritorna una copia di ps con le stringhe `valori` {chiave documento: {sei stringhe}}.

    Salva in ps['documenti'] (JSON) solo le differenze dai predefiniti; riporta ai vecchi
    valori le sei stringhe comuni, ormai assorbite nei singoli documenti.
    """
    nuovo = dict(ps)
    personalizzati = {}
    for chiave, testi in valori.items():
        if chiave in _DOCUMENTI_SENZA_TESTATA:
            continue
        predefiniti = testi_predefiniti(chiave)
        diff = {k: testi[k] for k in _TESTI_STAMPA if k in testi and testi[k] != predefiniti[k]}
        if diff:
            personalizzati[chiave] = diff
    nuovo['documenti'] = json.dumps(personalizzati, ensure_ascii=False, sort_keys=True)
    nuovo.update(_TESTI_PRECEDENTI)
    return nuovo


# ---------------------------------------------------------------------------------------
# Intestazione e piè di pagina a schermo (anteprima e stampa da Calc): MENU_sistema_pagine
# ---------------------------------------------------------------------------------------

# Stili di pagina di cui MENU_sistema_pagine scrive intestazione e piè di pagina.
STILI_TESTATA = frozenset((
    'PageStyle_Analisi di Prezzo',
    'PageStyle_COMPUTO_A4',
    'PageStyle_Elenco Prezzi',
))

# Stili degli atti contabili: stessa testata, ma senza le righe vuote di impaginazione, che
# alzerebbero intestazione e piè di pagina e sposterebbero i salti pagina dei SAL gia' emessi.
STILI_ATTI_CONTABILI = frozenset((
    'Page_Style_Libretto_Misure2',
    'PageStyle_REGISTRO_A4',
    'PageStyle_SAL_A4',
))

# Font e dimensioni (punti) del testo di intestazione e piè di pagina.
_FONT_TESTATA = 'Liberation Sans Narrow'
_HTXT_PREDEFINITO = 8.0
_HTXT_ANALISI = 9.0

_ATTI_CONTABILI = ('LIBRETTO', 'REGISTRO', 'SAL', 'CDP')

# Foglio del documento e foglio attivo che scelgono il documento dello stile condiviso.
_FOGLIO_DEL_DOCUMENTO = {
    'COMPUTO': 'COMPUTO',
    'VARIANTE': 'VARIANTE',
    'ELENCO': 'Elenco Prezzi',
    'ANALISI': 'Analisi di Prezzo',
}


def documento_dello_stile(nome_stile, foglio_attivo):
    """
    Chiave del documento (vedi _DOCUMENTI_STAMPA) di cui lo stile di pagina porta la testata.
    PageStyle_COMPUTO_A4 e' condiviso da COMPUTO, VARIANTE e Analisi di Prezzo: vale il foglio
    attivo, e per ogni altro foglio il Computo. None se lo stile non e' tra STILI_TESTATA o
    STILI_ATTI_CONTABILI.
    """
    if nome_stile == 'PageStyle_Elenco Prezzi':
        return 'ELENCO'
    if nome_stile == 'PageStyle_Analisi di Prezzo':
        return 'ANALISI'
    if nome_stile == 'PageStyle_COMPUTO_A4':
        return {'VARIANTE': 'VARIANTE', 'Analisi di Prezzo': 'ANALISI'}.get(foglio_attivo, 'COMPUTO')
    return {'Page_Style_Libretto_Misure2': 'LIBRETTO', 'PageStyle_REGISTRO_A4': 'REGISTRO',
            'PageStyle_SAL_A4': 'SAL'}.get(nome_stile)


def dimensione_testata(chiave):
    """Dimensione in punti del testo di intestazione e piè di pagina del documento."""
    return _HTXT_ANALISI if chiave == 'ANALISI' else _HTXT_PREDEFINITO


def titolo_a_schermo(chiave, oDoc):
    """
    Titolo del documento a schermo ([OGGETTO]): dipende dalle colonne visibili del suo foglio,
    come prima di MENU_sistema_pagine ('Incidenza MdO', 'Quadro di Raffronto'...).
    Se il foglio manca ripiega su titolo_documento.
    """
    base = titolo_documento(chiave)
    nome = _FOGLIO_DEL_DOCUMENTO.get(chiave)
    if nome is None or not oDoc.Sheets.hasByName(nome):
        return base
    oSheet = oDoc.Sheets.getByName(nome)
    if chiave == 'COMPUTO':
        visibile = oSheet.getColumns().getByIndex(30).Columns.IsVisible
        return 'Incidenza MdO' if visibile else 'Computo Metrico Estimativo'
    if chiave == 'VARIANTE':
        visibile = oSheet.getColumns().getByIndex(30).Columns.IsVisible
        return 'Incidenza MdO (Variante)' if visibile else 'Perizia di Variante'
    if chiave == 'ELENCO':
        nascoste = any(oSheet.getColumns().getByIndex(i).IsVisible == False for i in (11, 12, 13))
        return 'Quadro di Raffronto' if nascoste else 'Elenco Prezzi'
    return oSheet.Name


def testata_a_schermo(testi, valori, chiave):
    """
    Le sei stringhe di intestazione e piè di pagina pronte per lo stile di pagina del documento.

    Risolve i segnaposto con `valori`. Per Computo, Variante, Analisi ed Elenco Prezzi aggiunge
    le righe vuote che dispongono il testo come prima: committente e titolo in basso, oggetto in
    alto, "realizzato con" sotto. Gli atti contabili restano senza righe vuote: intestazione e
    piè di pagina non devono diventare piu' alti. Poi piede_con_firme porta la riga delle firme
    da sola in fondo, se c'e' [FIRME]. [PAGINA] e [PAGINE] restano: diventano campi numerici
    in scrivi_testata_stile.
    """
    t = {k: sostituisci_segnaposto(v, valori) for k, v in testi.items()}
    if chiave not in _ATTI_CONTABILI:
        t['intSx'] = '\n\n' + t['intSx']
        t['intCenter'] = t['intCenter'] + '\n\n'
        t['intDx'] = ('\n' if chiave == 'ANALISI' else '\n\n') + t['intDx']
        t['ppSx'] = '\n' + t['ppSx']
    t.update(piede_con_firme(t))
    return t


def scrivi_testata_stile(oDoc, oAktPage, testata, htxt, con_firme):
    """
    Scrive nello stile di pagina intestazione e piè di pagina `testata` (vedi testata_a_schermo)
    con font _FONT_TESTATA e dimensione `htxt`; [PAGINA] e [PAGINE] diventano campi numerici.
    Con le firme tutto il piè di pagina ha la stessa dimensione, altrimenti le righe delle tre
    parti non si allineano; senza firme la parte sinistra resta a meta' dimensione, come prima.
    """
    oHeader = oAktPage.RightPageHeaderContent
    for oTesto, chiave in ((oHeader.LeftText, 'intSx'), (oHeader.CenterText, 'intCenter'),
                           (oHeader.RightText, 'intDx')):
        oTesto.Text.String = testata[chiave]
        oTesto.Text.Text.CharFontName = _FONT_TESTATA
        oTesto.Text.Text.CharHeight = htxt
        SheetUtils.paginationFields(oDoc, oTesto.Text)
    oAktPage.RightPageHeaderContent = oHeader

    oFooter = oAktPage.RightPageFooterContent
    dimensioni = (htxt if con_firme else htxt * 0.5, htxt, htxt)
    for oTesto, chiave, dimensione in ((oFooter.LeftText, 'ppSx', dimensioni[0]),
                                       (oFooter.CenterText, 'ppCenter', dimensioni[1]),
                                       (oFooter.RightText, 'ppDx', dimensioni[2])):
        oTesto.Text.String = testata[chiave]
        oTesto.Text.Text.CharFontName = _FONT_TESTATA
        oTesto.Text.Text.CharHeight = dimensione
        SheetUtils.paginationFields(oDoc, oTesto.Text)
    oAktPage.RightPageFooterContent = oFooter


def loadJobSettings(oDoc):
    """
    Carica le impostazioni di lavoro dal documento o dalla configurazione predefinita.

    Args:
        oDoc: Il documento dal quale caricare le impostazioni.

    Returns:
        dict: Un dizionario con le impostazioni di lavoro.
    """
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'Lavoro')
    if data is None or len(data) == 0:
        data = cfg.readBlock('Lavoro', True)
    return data

def loadPageReplacements(oDoc):
    """
    Carica e restituisce le sostituzioni di pagina basate sulle impostazioni di lavoro.

    Args:
        oDoc: Il documento dal quale caricare le impostazioni.

    Returns:
        dict: Un dizionario con le sostituzioni di pagina, dove le chiavi sono i segnaposto e i valori sono i testi da inserire.
    """
    repl = loadJobSettings(oDoc)
    res = {}
    for key, val in repl.items():
        nKey = '[' + key.upper() + ']'
        if nKey in _DOCSTRINGS:
            # if simple substitution works, do it
            # so, just add [ and ] around and put to uppercase
            res[nKey] = val
        else:
            # no simple way, try to look for similar string
            # inside _DOCSTRINGS, just removing _ chars
            for v in _DOCSTRINGS:
                vr = v.replace('_', '')
                if vr == nKey:
                    res[v] = val
                    break
    return res

def storeJobSettings(oDoc, js):
    """
    Salva le impostazioni di lavoro nel documento e nella configurazione predefinita.

    Args:
        oDoc: Il documento nel quale salvare le impostazioni.
        js: Un dizionario con le impostazioni di lavoro da salvare.
    """
    cfg = LeenoConfig.Config()

    DocUtils.storeDataBlock(oDoc, 'Lavoro', js)
    cfg.writeBlock('Lavoro', js, True)

def JobSettingsDialog():
    """
    Crea e restituisce una finestra di dialogo per modificare le impostazioni di lavoro.

    Returns:
        Dialogs.Dialog: La finestra di dialogo per le impostazioni di lavoro.
    """
    # dimensione dell'icona col punto di domanda
    imgW = Dialogs.getBigIconSize()[0] * 3
    fieldW, dummy = Dialogs.getTextBox("W" * 30)

    return Dialogs.Dialog(Title='Impostazioni dati lavoro',  Horz=False, CanClose=True,  Items=[
        Dialogs.HSizer(Items=[
            Dialogs.VSizer(Items=[
                Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/books.png', MinWidth=imgW),
                Dialogs.Spacer(),
            ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Progetto:'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="progetto", FixedWidth=fieldW),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Committente:'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="committente", FixedWidth=fieldW),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Stazione appaltante'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="stazioneAppaltante"),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Responsabile del procedimento'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="rup"),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Progettista'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="progettista"),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Data'),
                Dialogs.Spacer(),
                Dialogs.DateControl(Id="data"),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Revisione'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id="revisione"),
                Dialogs.Spacer(),

                Dialogs.FixedText(Text='Data revisione'),
                Dialogs.Spacer(),
                Dialogs.DateControl(Id="dataRevisione"),
                Dialogs.Spacer(),
            ]),
        ]),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),
            Dialogs.Button(Label='Ok', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/ok.png',  RetVal=1),
            Dialogs.Spacer(),
            Dialogs.Button(Label='Annulla', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/cancel.png',  RetVal=-1),
            Dialogs.Spacer()
        ])
    ])


def MENU_JobSettings():
    """
    Mostra la finestra di dialogo per le impostazioni di lavoro e aggiorna le impostazioni se l'utente conferma.

    Carica le impostazioni di lavoro attuali, mostra la finestra di dialogo, e, se l'utente conferma, salva le nuove impostazioni.
    """
    oDoc = LeenoUtils.getDocument()
    js = loadJobSettings(oDoc)

    DLG.chi(f'js: {js}')

    dlg = JobSettingsDialog()
    dlg.setData(js)

    if dlg.run() >= 0:
        js = dlg.getData(_JOBSETTINGSITEMS)
        storeJobSettings(oDoc, js)

def fixupCover(coversPath, coverName):
    """
    Verifica se la copertina specificata è disponibile e restituisce la lista delle copertine disponibili e quella selezionata.

    Args:
        coversPath: Il percorso del file delle copertine.
        coverName: Il nome della copertina da verificare.

    Returns:
        tuple: Una tupla contenente la lista delle copertine disponibili e il nome della copertina selezionata.
    """
    covers = ()
    if coversPath is not None and coversPath != '':
        covers = SheetUtils.getSheetNames(coversPath)

    # controlla che la copertina specificata sia tra quelle disponibili
    # (uno potrebbe aver modificato il file...)
    if coverName in covers:
        return covers, coverName
    if len(covers) > 0:
        coverName = covers[0]
    else:
        coverName = ''
    return covers, coverName

def loadPrintSettings(oDoc):
    """
    Carica le impostazioni di stampa e le copertine disponibili dal documento.

    Args:
        oDoc: Il documento dal quale caricare le impostazioni.

    Returns:
        tuple: Una tupla contenente un dizionario con le impostazioni di stampa e una lista di copertine disponibili.
    """
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'ImpostazioniStampa')
    if data is None or len(data) == 0:
        data = cfg.readBlock('ImpostazioniStampa', True)

    # legge i nomi delle copertine dal file fornito, se esistente
    covers, copertina = fixupCover(data.get('fileCopertine', ''), data.get('copertina', ''))

    data['copertina'] = copertina

    return data, covers

def storePrintSettings(oDoc, js):
    """
    Salva le impostazioni di stampa nel documento e nella configurazione predefinita.

    Args:
        oDoc: Il documento nel quale salvare le impostazioni.
        js: Un dizionario con le impostazioni di stampa da salvare.
    """
    cfg = LeenoConfig.Config()

    DocUtils.storeDataBlock(oDoc, 'ImpostazioniStampa', js)
    cfg.writeBlock('ImpostazioniStampa', js, True)

_DLG_IMPOSTAZIONI_STAMPA = (
    'vnd.sun.star.script:UltimusFree2.DlgImpostazioniStampa'
    '?language=Basic&location=application')

# Promemoria dei segnaposto e nota per i documenti senza intestazione.
_LEGENDA_SEGNAPOSTI = 'Segnaposto: [COMMITTENTE] [PROGETTO] [OGGETTO] [SAL] [FILE] [FIRME] [PAGINA] [PAGINE]'
_NOTA_SENZA_TESTATA = ('Il Certificato di Pagamento riporta i propri dati nel corpo: '
                       'nessuna intestazione e nessun piè di pagina.')

# Foglio attivo -> documento mostrato all'apertura del dialogo.
_DOCUMENTO_DEL_FOGLIO = {
    'COMPUTO': 'COMPUTO',
    'VARIANTE': 'VARIANTE',
    'Elenco Prezzi': 'ELENCO',
    'Analisi di Prezzo': 'ANALISI',
    'CONTABILITA': 'LIBRETTO',
    'Registro': 'REGISTRO',
    'SAL': 'SAL',
    'CdP': 'CDP',
}


def documento_stampa_iniziale(oDoc):
    """Chiave del documento da mostrare per primo: quello del foglio attivo, altrimenti COMPUTO."""
    try:
        nome = oDoc.CurrentController.ActiveSheet.Name
    except Exception:
        return 'COMPUTO'
    return _DOCUMENTO_DEL_FOGLIO.get(nome, 'COMPUTO')


class _ControlloImpostazioniStampa:
    """
    Logica del dialogo: tiene in memoria le sei stringhe di ogni documento e mostra quelle
    del documento selezionato. La casella delle firme inserisce o toglie il segnaposto [FIRME]
    nel piè di pagina e riflette la sua presenza in una delle tre parti.
    """

    def __init__(self, oModel, ps, chiave):
        self._m = oModel
        self.valori = {c: impostazioni_documento(ps, c) for c, _ in _DOCUMENTI_STAMPA}
        self.corrente = None
        self._occupato = False
        self.mostra(chiave)

    def _modello(self, nome):
        return self._m.getByName(nome)

    def _salva_corrente(self):
        if self.corrente is None or self.corrente in _DOCUMENTI_SENZA_TESTATA:
            return
        for k in _TESTI_STAMPA:
            self.valori[self.corrente][k] = self._modello(k).Text

    def mostra(self, chiave):
        """Salva i testi del documento mostrato e mostra quelli di `chiave`."""
        self._salva_corrente()
        self.corrente = chiave
        senza = chiave in _DOCUMENTI_SENZA_TESTATA
        testi = self.valori[chiave]
        self._occupato = True
        try:
            for k in _TESTI_STAMPA:
                self._modello(k).Text = testi[k]
                self._modello(k).Enabled = not senza
                self._modello('lbl_' + k).Enabled = not senza
            self._modello('btnPredefiniti').Enabled = not senza
            self._modello('cbFirme').Enabled = chiave in _DOCUMENTI_CON_FIRME
            self._modello('cbFirme').State = 1 if ha_firme(testi) else 0
            self._modello('lblSegnaposto').Label = _NOTA_SENZA_TESTATA if senza else _LEGENDA_SEGNAPOSTI
        finally:
            self._occupato = False

    def cambia_documento(self, indice):
        if 0 <= indice < len(_DOCUMENTI_STAMPA):
            self.mostra(_DOCUMENTI_STAMPA[indice][0])

    def ripristina(self):
        """Riporta il documento mostrato ai testi predefiniti."""
        if self.corrente is None or self.corrente in _DOCUMENTI_SENZA_TESTATA:
            return
        chiave = self.corrente
        self.valori[chiave] = testi_predefiniti(chiave)
        self.corrente = None  # i testi mostrati non vanno salvati sopra i predefiniti
        self.mostra(chiave)

    def _testi_piede(self):
        return {k: self._modello(k).Text for k in _PARTI_PIEDE}

    def firme_cambiate(self):
        """Casella delle firme: aggiunge [FIRME] al centro del piè di pagina o lo toglie ovunque."""
        if self._occupato or self.corrente not in _DOCUMENTI_CON_FIRME:
            return
        self._occupato = True
        try:
            if self._modello('cbFirme').State:
                if not ha_firme(self._testi_piede()):
                    centro = self._modello('ppCenter')
                    centro.Text = (centro.Text + ' ' + _SEGNAPOSTO_FIRME).strip()
            else:
                for k in _PARTI_PIEDE:
                    self._modello(k).Text = senza_firme(self._modello(k).Text)
        finally:
            self._occupato = False

    def piede_cambiato(self):
        """Piè di pagina modificato a mano: la casella riflette la presenza di [FIRME]."""
        if self._occupato or self.corrente not in _DOCUMENTI_CON_FIRME:
            return
        self._occupato = True
        try:
            self._modello('cbFirme').State = 1 if ha_firme(self._testi_piede()) else 0
        finally:
            self._occupato = False

    def risultato(self):
        """Le sei stringhe di tutti i documenti: {chiave: {stringhe}}."""
        self._salva_corrente()
        return {c: dict(t) for c, t in self.valori.items()}


class _AscoltatoreImpostazioniStampa(unohelper.Base, XActionListener, XItemListener, XTextListener):
    """Collega i controlli del dialogo alla logica di _ControlloImpostazioniStampa."""

    def __init__(self, oDlg, controllo):
        self._oDlg = oDlg
        self._controllo = controllo

    @staticmethod
    def _nome(ev):
        return ev.Source.getModel().Name

    def actionPerformed(self, ev):
        try:
            nome = self._nome(ev)
            if nome == 'btnPredefiniti':
                self._controllo.ripristina()
            elif nome == 'btnSfoglia':
                file = Dialogs.FileSelect(est='*.ods')
                if file is not None:
                    self._oDlg.getModel().getByName('fileCopertine').Text = file
                    self._oDlg.getModel().getByName('fileCopertine').HelpText = file
        except Exception as _exc:
            LeenoUtils.log_errore('PrintSettingsDialog: azione', _exc)

    def itemStateChanged(self, ev):
        try:
            nome = self._nome(ev)
            if nome == 'docSel':
                self._controllo.cambia_documento(ev.Selected)
            elif nome == 'cbFirme':
                self._controllo.firme_cambiate()
        except Exception as _exc:
            LeenoUtils.log_errore('PrintSettingsDialog: selezione', _exc)

    def textChanged(self, ev):
        try:
            if self._nome(ev) in _PARTI_PIEDE:
                self._controllo.piede_cambiato()
        except Exception as _exc:
            LeenoUtils.log_errore('PrintSettingsDialog: testo', _exc)

    def disposing(self, ev):
        return None


def PrintSettingsDialog(ps, covers, chiave='COMPUTO', url=_DLG_IMPOSTAZIONI_STAMPA, immagini=True):
    '''
    Crea il dialogo di configurazione dell'intestazione e del piè di pagina di stampa
    (UltimusFree2.DlgImpostazioniStampa) precompilato con ps, con l'elenco covers e con il
    documento `chiave`. Restituisce (oDlg, ascoltatore, controllo): l'ascoltatore va tenuto
    vivo fino alla chiusura; il controllo fornisce le stringhe di tutti i documenti.
    '''
    ctx = LeenoUtils.getComponentContext()
    dp = ctx.ServiceManager.createInstance('com.sun.star.awt.DialogProvider')
    oDlg = dp.createDialog(url)
    oModel = oDlg.getModel()

    if immagini:
        base = LeenoUtils.LeenO_path() + '/python/pythonpath/'
        oModel.getByName('imgAnteprima').ImageURL = base + 'Icons-Big/preview.png'
        for nome, icona in (('btnOk', 'ok.png'), ('btnAnnulla', 'cancel.png'),
                            ('btnSfoglia', 'file.png')):
            oModel.getByName(nome).ImageURL = base + (
                'Icons-16x16/' if nome == 'btnSfoglia' else 'Icons-24x24/') + icona
            oModel.getByName(nome).ImageAlign = 0

    for nome in _TESTI_STAMPA:
        oDlg.getControl(nome).addItems(_DOCSTRINGS, 0)

    oDlg.getControl('docSel').addItems(tuple(titolo for _, titolo in _DOCUMENTI_STAMPA), 0)
    chiavi = [c for c, _ in _DOCUMENTI_STAMPA]
    if chiave not in chiavi:
        chiave = 'COMPUTO'
    oModel.getByName('docSel').SelectedItems = (chiavi.index(chiave),)

    percorso = ps.get('fileCopertine', '') or ''
    oModel.getByName('fileCopertine').Text = percorso
    oModel.getByName('fileCopertine').HelpText = percorso

    oList = oDlg.getControl('copertina')
    if len(covers) > 0:
        oList.addItems(tuple(covers), 0)
        corrente = ps.get('copertina', '')
        oList.selectItem(corrente if corrente in covers else covers[0], True)

    controllo = _ControlloImpostazioniStampa(oModel, ps, chiave)
    ascoltatore = _AscoltatoreImpostazioniStampa(oDlg, controllo)
    for nome in ('btnPredefiniti', 'btnSfoglia'):
        oDlg.getControl(nome).addActionListener(ascoltatore)
    oDlg.getControl('docSel').addItemListener(ascoltatore)
    oDlg.getControl('cbFirme').addItemListener(ascoltatore)
    for nome in _PARTI_PIEDE:
        oDlg.getControl(nome).addTextListener(ascoltatore)
    return oDlg, ascoltatore, controllo


def _leggiImpostazioniStampa(oDlg, ps, controllo):
    '''Impostazioni da salvare: file e copertina del dialogo, stringhe di tutti i documenti.'''
    res = dict(ps)
    res['fileCopertine'] = oDlg.getModel().getByName('fileCopertine').Text
    res['copertina'] = oDlg.getControl('copertina').getSelectedItem() or ''
    return aggiorna_impostazioni_documenti(res, controllo.risultato())


def MENU_PrintSettings():
    """
    Mostra il dialogo per modificare le impostazioni di stampa e le salva
    nel documento e nella configurazione predefinita.
    """

    oDoc = LeenoUtils.getDocument()
    ps, covers = loadPrintSettings(oDoc)

    oDlg, ascoltatore, controllo = PrintSettingsDialog(ps, covers, documento_stampa_iniziale(oDoc))
    try:
        esito = oDlg.execute()
        if esito == 1:
            ps = _leggiImpostazioniStampa(oDlg, ps, controllo)
    finally:
        oDlg.dispose()

    if esito == 1:
        storePrintSettings(oDoc, ps)

########################################################################

# Stile di pagina del foglio CdP: A4 verticale, senza intestazione e piè di pagina
# perche' il certificato riporta nel corpo i propri dati (ente, oggetto, committente).
# Non e' nel template: viene creato al bisogno da _crea_stile_pagina_cdp.
STILE_PAGINA_CDP = 'PageStyle_CdP_A4'


def _crea_stile_pagina_cdp(oDoc):
    '''
    Crea lo stile di pagina del CdP se il documento ha il foglio CdP e lo stile manca.
    Le impostazioni comuni (margini, scala) sono applicate da MENU_sistema_pagine.
    '''
    if not oDoc.getSheets().hasByName('CdP'):
        return
    oPageStyles = oDoc.StyleFamilies.getByName('PageStyles')
    if oPageStyles.hasByName(STILE_PAGINA_CDP):
        return
    oPageStyles.insertByName(STILE_PAGINA_CDP, oDoc.createInstance('com.sun.star.style.PageStyle'))
    oStile = oPageStyles.getByName(STILE_PAGINA_CDP)
    oStile.IsLandscape = False
    oStile.Width = 21000
    oStile.Height = 29700
    oStile.HeaderIsOn = False
    oStile.FooterIsOn = False


def setPageStyle(silent = False):
    '''
    Attribuisce ad ogni foglio il suo specifico stile di pagina.
    Usa assegnazione diretta (oSheet.PageStyle) senza GotoSheet/dispatch
    per evitare switch di foglio e i relativi ridisegni intermedi.
    '''
    importa_stili_pagina_non_presenti(silent=silent)
    _crea_stile_pagina_cdp(LeenoUtils.getDocument())

    stili = {
        'cP_Cop': 'Page_Style_COPERTINE',
        'COMPUTO': 'PageStyle_COMPUTO_A4',
        'VARIANTE': 'PageStyle_COMPUTO_A4',
        'Elenco Prezzi': 'PageStyle_Elenco Prezzi',
        'Analisi di Prezzo': 'PageStyle_COMPUTO_A4',
        'CONTABILITA': 'Page_Style_Libretto_Misure2',
        'Registro': 'PageStyle_REGISTRO_A4',
        'SAL': 'PageStyle_SAL_A4',
        'CdP': STILE_PAGINA_CDP,
    }

    oDoc = LeenoUtils.getDocument()
    # Cache della collezione stili: evita lookup ripetuto per ogni foglio
    oPageStyles = oDoc.StyleFamilies.getByName('PageStyles')

    for el, stile_nome in stili.items():
        try:
            if not oDoc.getSheets().hasByName(el):
                continue
            oSheet = oDoc.getSheets().getByName(el)

            # Assegnazione diretta senza navigare al foglio (nessun ridisegno)
            oSheet.PageStyle = stile_nome

            # Applica le impostazioni di scala allo stile di pagina
            if oPageStyles.hasByName(stile_nome):
                set_page_scale(oPageStyles.getByName(stile_nome))
        except Exception:
            continue



########################################################################


def npagina ():
    """
    Inserisce il numero di pagina corrente, e, per alcuni stili, aggiunge
    il conteggio totale delle pagine.
    """
    # ~ PL.importa_stili_pagina()
    oDoc = LeenoUtils.getDocument()
    # Ottieni la famiglia di stili di pagina
    page_styles = oDoc.StyleFamilies.getByName("PageStyles")

    stili = {
        # ~ 'cP_Cop': 'Page_Style_COPERTINE',
        'COMPUTO': 'PageStyle_COMPUTO_A4',
        'VARIANTE': 'PageStyle_COMPUTO_A4',
        'Elenco Prezzi': 'PageStyle_Elenco Prezzi',
        # ~'Analisi di Prezzo': 'PageStyle_Analisi di Prezzo',
        'Analisi di Prezzo': 'PageStyle_COMPUTO_A4',
        'CONTABILITA': 'Page_Style_Libretto_Misure2',
        'Registro': 'PageStyle_REGISTRO_A4',
        'SAL': 'PageStyle_SAL_A4',
    }

    for el in (stili.keys()):
        try:
            default_style = page_styles.getByName(stili[el])

            # Abilita l'intestazione
            default_style.HeaderIsOn = True
            header = default_style.RightPageHeaderContent
            footer = default_style.RightPageFooterContent
            footer.RightText.String = ""

            # Pulisci l'intestazione esistente
            header.RightText.String = ""

            # Inserisci il numero di pagina
            page_number = oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
            text_cursor = header.RightText.createTextCursor()
            text_cursor.gotoEnd(False)
            text_cursor.String = "pag. "
            text_cursor.gotoEnd(False)
            header.RightText.insertTextContent(text_cursor, page_number, True)

            if stili[el] in ('PageStyle_COMPUTO_A4', 'PageStyle_Elenco Prezzi'):

                # Inserisci il testo " di " e il conteggio totale delle pagine
                page_count = oDoc.createInstance("com.sun.star.text.TextField.PageCount")
                text_cursor.gotoEnd(False)
                text_cursor.String = " di "
                text_cursor.gotoEnd(False)
                header.RightText.insertTextContent(text_cursor, page_count, True)

            # Applica le modifiche
            default_style.RightPageHeaderContent = header
        except:
            # ~ DLG.chi(f'Stile pagina {stili[el]} inesistente.')
            pass

    return


# Geometria comune a tutti gli stili di pagina (margini e distanze in 1/100 di mm).
# Unica sorgente: la usano set_page_margins, set_page_scale e MENU_sistema_pagine.
_PAGINA_MARGINI = {
    'TopMargin': 1000,
    'BottomMargin': 1350,
    'LeftMargin': 1000,
    'RightMargin': 1000,
    'FooterLeftMargin': 0,
    'FooterRightMargin': 0,
    'HeaderLeftMargin': 0,
    'HeaderRightMargin': 0,
    'HeaderBodyDistance': 0,
    'FooterBodyDistance': 0,
}

# PageScale = 0 e ScaleToPagesX = 1: scala automatica, adatta alla larghezza
# di una pagina, altezza non specificata.
_PAGINA_SCALA = {
    'PageScale': 0,
    'CenterHorizontally': True,
    'ScaleToPagesX': 1,
    'ScaleToPagesY': 0,
}


def _imposta_proprieta(oAktPage, valori):
    """
    Scrive le proprieta' di valori {nome: valore} sullo stile di pagina con una
    sola chiamata UNO (evita le notifiche intermedie che causavano il freeze).
    I nomi sono ordinati, come richiede XMultiPropertySet; PageScale precede
    ScaleToPagesX/Y, quindi la scala automatica prevale. Se la scrittura multipla
    non e' supportata ripiega sull'assegnazione una per una.
    """
    nomi = tuple(sorted(valori))
    try:
        oAktPage.setPropertyValues(nomi, tuple(valori[n] for n in nomi))
    except Exception:
        for nome in nomi:
            setattr(oAktPage, nome, valori[nome])


def set_page_margins(oAktPage):
    """
    Imposta margini, distanze di intestazione e piè di pagina, scala di pagina
    e centratura orizzontale per lo stile di pagina oAktPage.
    """
    _imposta_proprieta(oAktPage, {**_PAGINA_MARGINI, **_PAGINA_SCALA})


def set_page_scale(oAktPage):
    """Imposta solo scala automatica e centratura orizzontale dello stile oAktPage."""
    _imposta_proprieta(oAktPage, _PAGINA_SCALA)


def set_page_borders(oAktPage):
    """Azzera i quattro bordi dello stile di pagina oAktPage."""
    bordo_zero = uno.createUnoStruct('com.sun.star.table.BorderLine2')
    bordo_zero.LineWidth = 0
    bordo_zero.OuterLineWidth = 0
    _imposta_proprieta(oAktPage, {
        'TopBorder': bordo_zero,
        'BottomBorder': bordo_zero,
        'RightBorder': bordo_zero,
        'LeftBorder': bordo_zero,
    })

def set_header(oAktPage, str1='', str2='', str3=''):
    """
    Imposta il testo dell'intestazione a destra, al centro e a sinistra
    dell'oggetto pagina fornito. Ogni parte dell'intestazione può essere
    personalizzata tramite i parametri della funzione.

    Args:
        oAktPage: Oggetto rappresentante lo stile di pagina.
        str1: Testo a sinistra nell'intestazione (default: '').
        str2: Testo al centro nell'intestazione (default: '').
        str3: Testo a destra nell'intestazione (default: '').
    """
    # ~ oDoc = LeenoUtils.getDocument()
    # ~ oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)

    # Imposta l'header della pagina
    oHeader = oAktPage.RightPageHeaderContent
    oHeader.LeftText.Text.String = str1
    oHeader.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    oHeader.CenterText.Text.String = str2
    oHeader.CenterText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    oHeader.RightText.Text.String = str3
    oHeader.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    oAktPage.RightPageHeaderContent = oHeader

def set_footer(oAktPage, str1 = '', str2 = '', str3 = ''):
    """
    Imposta il testo del pié di pagina a destra, al centro e a sinistra
    dell'oggetto pagina fornito. Ogni parte del pié di pagina può essere
    personalizzata tramite i parametri della funzione.

    Args:
        oAktPage: Oggetto rappresentante lo stile di pagina.
        str1: Testo a sinistra nell'intestazione (default: '').
        str2: Testo al centro nell'intestazione (default: '').
        str3: Testo a destra nell'intestazione (default: '').
    """

    # ~ oDoc = LeenoUtils.getDocument()
    # ~ oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)

    # Imposta il footer della pagina
    oFooter = oAktPage.RightPageFooterContent
    oFooter.LeftText.Text.String = str1
    oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oFooter.LeftText.Text.Text.CharHeight = htxt * 0.7
    oFooter.CenterText.Text.String = str2
    oFooter.CenterText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    oFooter.RightText.Text.String = str3
    oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    oAktPage.RightPageFooterContent = oFooter

def impostazioni_pagina():
    """
    Imposta stili di pagina, margini, intestazioni e piè di pagina per un documento.

    Assegna il nome del committente alle intestazioni delle pagine,
    applicando le configurazioni in base allo stile della pagina.
    Disattiva intestazioni/piè di pagina per pagine di tipo 'COPERTINE'.
    """

    oDoc = LeenoUtils.getDocument()
    stili = {
        'PageStyle_Elenco Prezzi': 'Elenco Prezzi',
        'Page_Style_Libretto_Misure2': 'Libretto delle Misure',
        'PageStyle_REGISTRO_A4': 'Registro di Contabilità',
        'PageStyle_SAL_A4': 'Stato di Avanzamento Lavori',
    }
    # ~ stili = {
        # ~ 'Analisi di Prezzo': 'Analisi di Prezzo',
        # ~ 'Elenco Prezzi': 'Elenco Prezzi',
        # ~ 'COMPUTO': 'Computo Metrico',
        # ~ 'VARIANTE': 'Variante',
        # ~ 'CONTABILITA': 'Libretto delle Misure',
        # ~ 'REGISTRO': 'Registro di Contabilità',
        # ~ 'SAL': 'Stato di Avanzamento Lavori',
    # ~ }

    try:
        committente = "Committente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String
    except Exception as e:
        DLG.chi(f"Errore nel recupero del committente: {e}")
        committente = ""

    for n in range(0, oDoc.StyleFamilies.getByName('PageStyles').Count):
        oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)
        try:
            committente_full = committente + '\n' + stili.get(oAktPage.DisplayName, '')
        except KeyError as e:
            DLG.chi(f"Stile non trovato: {e}")
            committente_full = committente


        if oAktPage.DisplayName == 'Page_Style_COPERTINE':
            oAktPage.HeaderIsOn = False
            oAktPage.FooterIsOn = False
        else:
            oAktPage.HeaderIsOn = True
            oAktPage.FooterIsOn = True

        set_page_margins(oAktPage)
        set_page_borders(oAktPage)
        set_header(oAktPage, committente_full, '', '')
        set_footer(oAktPage, '', '', '')

    npagina()

########################################################################

def importa_stili_pagina(overwrite = False, silent = False):
    """
    Importa solo gli stili di pagina dal template di LeenO,
    con la possibilità di sovrascrivere quelli esistenti.

    Args:
        overwrite (bool): se True sovrascrive gli stili esistenti.
        silent (bool): se True non avvia l'indicatore di progresso.
    """
    with LeenoUtils.DocumentRefresh(False):

        filename = PL.LeenO_path() + '/template/leeno/Computo_LeenO.ods'

        oDoc = LeenoUtils.getDocument()

        # Creare la lista di PropertyValue per le opzioni di caricamento
        loadOptions = [
            PL.crea_property_value("LoadPageStyles", True),
            PL.crea_property_value("LoadCellStyles", False),
            PL.crea_property_value("LoadTextStyles", False),
            PL.crea_property_value("LoadFrameStyles", False),
            PL.crea_property_value("LoadNumberingStyles", False),
            PL.crea_property_value("OverwriteStyles", overwrite)
        ]

        try:
            # Carica gli stili di pagina dal file di riferimento
            oDoc.StyleFamilies.loadStylesFromURL(filename, tuple(loadOptions))
        except Exception as e:
            DLG.chi(f"Errore durante l'importazione degli stili: {str(e)}")

        # attiva la progressbar
        if not silent:
            indicator = oDoc.getCurrentController().getStatusIndicator()
            indicator.start('Adattamento altezze righe in corso...', len(oDoc.Sheets.ElementNames))
        n = 1
        for el in oDoc.Sheets.ElementNames:
            indicator.Value = n
            n += 1
            oDoc.CurrentController.setActiveSheet(
                oDoc.getSheets().getByName(el))
            oSheet = oDoc.getSheets().getByName(el)
            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
        if not silent:
            indicator.end()

########################################################################

def importa_stili_pagina_non_presenti(silent = False):
    """
    Importa solo gli stili di pagina non presenti nel file corrente
    dal template di LeenO.

    Args:
        silent (bool): se True non avvia l'indicatore di progresso.
    """
    # Stili minimi richiesti da LeenO
    _STILI_LEENO = frozenset((
        'Page_Style_COPERTINE',
        'PageStyle_COMPUTO_A4',
        'PageStyle_Elenco Prezzi',
        'PageStyle_Analisi di Prezzo',
        'Page_Style_Libretto_Misure2',
        'PageStyle_REGISTRO_A4',
        'PageStyle_SAL_A4',
    ))

    oDoc = LeenoUtils.getDocument()
    existing_page_styles = oDoc.StyleFamilies.getByName('PageStyles')
    existing_style_names = set(
        existing_page_styles.getByIndex(i).Name
        for i in range(existing_page_styles.Count)
    )

    # Early-exit: se tutti gli stili LeenO sono già presenti,
    # non aprire il documento template (operazione costosa)
    if _STILI_LEENO.issubset(existing_style_names):
        return

    with LeenoUtils.DocumentRefreshContext(False):
        # Percorso del file di template
        filename = PL.LeenO_path() + '/template/leeno/Computo_LeenO.ods'

        lun_1 = len(existing_page_styles)



        try:
            # Carica il documento di template temporaneamente in modalità nascosta
            templateDoc = DocUtils.loadDocument(filename, Hidden=True)

            # Ottieni gli stili di pagina dal file di template
            template_page_styles = templateDoc.StyleFamilies.getByName('PageStyles')

            # Itera sugli stili nel template
            for i in range(template_page_styles.Count):
                template_style = template_page_styles.getByIndex(i)
                template_style_name = template_style.Name

                # Importa solo gli stili di pagina non presenti
                if template_style_name not in existing_style_names:
                    # Crea una lista di opzioni specifica per importare solo lo stile di pagina mancante
                    importOptions = [
                        PL.crea_property_value("PageStyleName", template_style_name),
                        PL.crea_property_value("OverwriteStyles", False)  # Non sovrascrivere
                    ]
                    # Importa lo stile specifico
                    oDoc.StyleFamilies.loadStylesFromURL(filename, tuple(importOptions))

            # Chiudi il documento di template
            templateDoc.close(True)

            # ~ DLG.chi("Stili di pagina mancanti importati correttamente.")
        except Exception:
            # ~ DLG.chi(f"Errore durante l'importazione degli stili: {str(e)}")
            pass
        # attiva la progressbar
        if not silent and lun_1 < len(oDoc.StyleFamilies.getByName('PageStyles')):
            indicator = oDoc.getCurrentController().getStatusIndicator()
            indicator.start('Adattamento altezze righe in corso...', len(oDoc.Sheets.ElementNames))
            n = 1
            for el in oDoc.Sheets.ElementNames:
                indicator.Value = n
                n += 1
                oDoc.CurrentController.setActiveSheet(
                    oDoc.getSheets().getByName(el))
                oSheet = oDoc.getSheets().getByName(el)
                LeenoSheetUtils.adattaAltezzaRiga(oSheet)
            indicator.end()
