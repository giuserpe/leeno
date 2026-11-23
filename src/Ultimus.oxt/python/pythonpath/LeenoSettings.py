########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

'''
Modulo per la modifica delle impostazioni di LeenO
'''
import uno
import unohelper
from com.sun.star.awt import XActionListener

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
)

_DOCSTRINGS = (
    '[COMMITTENTE]',
    '[DATA]',
    '[DATA_REVISIONE]',
    '[DATI_COMMITTENTE]',
    '[DATI_PROGETTISTA]',
    '[DIRETTORE_LAVORI]',
    '[NUMERO_DOCUMENTO]',
    '[OGGETTO]',
    '[PAGINA]',
    '[PAGINE]',
    '[PROGETTISTA]',
    '[PROGETTO]',
    '[REVISIONE]',
    '[RUP]',
    '[STAZIONE_APPALTANTE]',
)

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

_PRINTSETTINGSCOMBOS = ('intSx', 'intCenter', 'intDx', 'ppSx', 'ppCenter', 'ppDx')


class _SfogliaCopertine(unohelper.Base, XActionListener):
    '''Pulsante di selezione del file delle copertine (solo file .ods).'''
    def __init__(self, oDlg):
        self._oDlg = oDlg

    def actionPerformed(self, ev):
        try:
            file = Dialogs.FileSelect(est='*.ods')
            if file is not None:
                self._oDlg.getModel().getByName('fileCopertine').Text = file
                self._oDlg.getModel().getByName('fileCopertine').HelpText = file
        except Exception as _exc:
            LeenoUtils.log_errore('PrintSettingsDialog: selezione file copertine', _exc)

    def disposing(self, ev):
        return None


def PrintSettingsDialog(ps, covers):
    '''
    Crea il dialogo di configurazione per l'intestazione e il piè di pagina di stampa
    (UltimusFree2.DlgImpostazioniStampa) precompilato con ps e con l'elenco covers.
    Restituisce (oDlg, listener): il listener va tenuto vivo fino alla chiusura.
    '''
    ctx = LeenoUtils.getComponentContext()
    dp = ctx.ServiceManager.createInstance('com.sun.star.awt.DialogProvider')
    oDlg = dp.createDialog(_DLG_IMPOSTAZIONI_STAMPA)
    oModel = oDlg.getModel()

    base = LeenoUtils.LeenO_path() + '/python/pythonpath/'
    oModel.getByName('imgAnteprima').ImageURL = base + 'Icons-Big/preview.png'
    for nome, icona in (('btnOk', 'ok.png'), ('btnAnnulla', 'cancel.png'),
                        ('btnSfoglia', 'file.png')):
        oModel.getByName(nome).ImageURL = base + (
            'Icons-16x16/' if nome == 'btnSfoglia' else 'Icons-24x24/') + icona
        oModel.getByName(nome).ImageAlign = 0

    for nome in _PRINTSETTINGSCOMBOS:
        oDlg.getControl(nome).addItems(_DOCSTRINGS, 0)
        oModel.getByName(nome).Text = ps.get(nome, '') or ''

    percorso = ps.get('fileCopertine', '') or ''
    oModel.getByName('fileCopertine').Text = percorso
    oModel.getByName('fileCopertine').HelpText = percorso

    oList = oDlg.getControl('copertina')
    if len(covers) > 0:
        oList.addItems(tuple(covers), 0)
        corrente = ps.get('copertina', '')
        oList.selectItem(corrente if corrente in covers else covers[0], True)

    azione = _SfogliaCopertine(oDlg)
    oDlg.getControl('btnSfoglia').addActionListener(azione)
    return oDlg, azione


def _leggiImpostazioniStampa(oDlg):
    '''Legge dal dialogo i valori elencati in _PRINTSETTINGSITEMS.'''
    oModel = oDlg.getModel()
    res = {}
    for key in _PRINTSETTINGSITEMS:
        if key == 'copertina':
            res[key] = oDlg.getControl('copertina').getSelectedItem() or ''
        else:
            res[key] = oModel.getByName(key).Text
    return res


def MENU_PrintSettings():
    """
    Mostra il dialogo per modificare le impostazioni di stampa e le salva
    nel documento e nella configurazione predefinita.
    """

    oDoc = LeenoUtils.getDocument()
    ps, covers = loadPrintSettings(oDoc)

    oDlg, azione = PrintSettingsDialog(ps, covers)
    try:
        esito = oDlg.execute()
        if esito == 1:
            ps = _leggiImpostazioniStampa(oDlg)
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
