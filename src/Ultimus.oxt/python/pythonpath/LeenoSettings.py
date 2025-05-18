'''
Modulo per la modifica delle impostazioni di LeenO
'''
import Dialogs
import LeenoConfig
import LeenoUtils
import DocUtils
import LeenoSheetUtils
import SheetUtils
import LeenoDialogs as DLG
import pyleeno as PL

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

oDoc = LeenoUtils.getDocument()
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

def PrintSettingsDialog():
    # dimensione dell'icona grande
    imgW = Dialogs.getBigIconSize()[0] * 2
    fieldW, dummy = Dialogs.getTextBox("W" * 30)
    posW, dummy = Dialogs.getTextBox("SinistraXX")
    return Dialogs.Dialog(Title='Impostazioni stampa / PDF',  Horz=False, CanClose=True,  Items=[
        Dialogs.VSizer(Items=[
            Dialogs.FixedText(Text='Intestazione:'),
                # ~ Dialogs.Spacer(),
            Dialogs.HSizer(Items=[
                    Dialogs.VSizer(Items=[
                        Dialogs.FixedText(Text='Sinistra: '),
                        Dialogs.ComboBox(Id="intSx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                    ]),
                    Dialogs.Spacer(),
                    Dialogs.VSizer(Items=[
                        Dialogs.FixedText(Text='Centro: '),
                        Dialogs.ComboBox(Id="intCenter", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                    ]),
                    Dialogs.Spacer(),
                    Dialogs.VSizer(Items=[
                        Dialogs.FixedText(Text='Destra: '),
                        Dialogs.ComboBox(Id="intDx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                    ]),
            ]),
            
            # ~ Dialogs.Spacer(MinSize = 10),
            Dialogs.Spacer(MinSize = 30),
            Dialogs.HSizer(Items=[
                # ~ Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/preview.png', MinWidth=imgW * 1.5),
                # ~ Dialogs.Spacer(),
            ]),
            Dialogs.Spacer(MinSize = 30),
            
                Dialogs.FixedText(Text='Piè di pagina:'),
                # ~ Dialogs.Spacer(),
            Dialogs.HSizer(Items=[
                Dialogs.VSizer(Items=[
                    Dialogs.FixedText(Text='Sinistra: ', FixedWidth=posW),
                    # ~ Dialogs.FixedText(Text='Sinistra: '),
                    Dialogs.ComboBox(Id="ppSx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
                Dialogs.Spacer(MinSize = 45),
                Dialogs.VSizer(Items=[
                    Dialogs.FixedText(Text='Centro: '),
                    Dialogs.ComboBox(Id="ppCenter", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
                Dialogs.Spacer(MinSize = 45),
                Dialogs.VSizer(Items=[
                    Dialogs.FixedText(Text='Destra: '),
                    Dialogs.ComboBox(Id="ppDx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
            ]),
        ]),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            ]),
            # ~ Dialogs.VSizer(Items=[
            Dialogs.FixedText(Text='–' * 85),
            Dialogs.Spacer(),
            Dialogs.HSizer(Items=[
                Dialogs.FixedText(Text='Doc delle copertine: '),  # Testo fisso per indicare il file di copertina
                Dialogs.FileControl(Id="fileCopertine", Types='*.ods', MinWidth=150),  # Controllo file con larghezza minima
                Dialogs.Spacer(),  # Spaziatore per separare gli elementi
                Dialogs.FixedText(Text='Copertina in uso: ', FixedHeight=25),  # Testo fisso per la copertina selezionata
                Dialogs.ListBox(Id='copertina', MinWidth=150),  # ListBox con larghezza minima per le copertine
                # ~ Dialogs.Spacer(),  # Spaziatore per separare ulteriormente se necessario
            ]),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),  # Spaziatore per spingere i pulsanti verso destra
            Dialogs.Button(Label='Ok', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/ok.png', RetVal=1),
            # ~ Dialogs.Spacer(),  # Spaziatore per separare i pulsanti
            Dialogs.Button(Label='Annulla', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/cancel.png', RetVal=-1),
            # ~ Dialogs.Spacer(),  # Spaziatore per separare i pulsanti
        ]),
    ])


def MENU_PrintSettings():
    """
    Crea e restituisce una finestra di dialogo per modificare le impostazioni di stampa.

    Returns:
        Dialogs.Dialog: La finestra di dialogo per le impostazioni di stampa.
    """

    oDoc = LeenoUtils.getDocument()
    ps, covers = loadPrintSettings(oDoc)

    dlg = PrintSettingsDialog()
    dlg.getWidget('copertina').setList(covers)
    dlg.setData(ps)

    if dlg.run() >= 0:
        ps = dlg.getData(_PRINTSETTINGSITEMS)
        storePrintSettings(oDoc, ps)

########################################################################

def setPageStyle():
    '''
    Attribuisce ad ogni foglio il suo specifico stile di pagina.
    '''
    importa_stili_pagina_non_presenti()

    stili = {
        'cP_Cop': 'Page_Style_COPERTINE',
        'COMPUTO': 'PageStyle_COMPUTO_A4',
        'VARIANTE': 'PageStyle_COMPUTO_A4',
        'Elenco Prezzi': 'PageStyle_Elenco Prezzi',
        'Analisi di Prezzo': 'PageStyle_COMPUTO_A4',
        'CONTABILITA': 'Page_Style_Libretto_Misure2',
        'Registro': 'PageStyle_REGISTRO_A4',
        'SAL': 'PageStyle_SAL_A4',
    }

    # Ottieni il contesto e i servizi necessari
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)

    oDoc = LeenoUtils.getDocument()
    nSheet = oDoc.CurrentController.ActiveSheet.Name

    # Itera su ciascun foglio
    for el in stili.keys():
        try:
            PL.GotoSheet(el)

            # Crea le proprietà per lo stile di pagina
            oProp = [
                PL.crea_property_value("Template", stili[el]),
                PL.crea_property_value("Family", 8)
            ]

            # Applica lo stile alla pagina tramite dispatchHelper
            dispatchHelper.executeDispatch(oFrame, ".uno:StyleApply", "", 0, tuple(oProp))

            # Applica le impostazioni di zoom layout alla pagina
            oSheet = oDoc.getSheets().getByName(el)
            oSheet.PageStyle = stili[el]
            oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName(stili[el])
            oAktPage.PageScale = 0  # Imposta la scala automatica
            oAktPage.CenterHorizontally = True
            oAktPage.ScaleToPagesX = 1  # Adatta alla larghezza
            oAktPage.ScaleToPagesY = 0  # Altezza non specificata
        except Exception as e:
            importa_stili_pagina_non_presenti()
            continue

    # Ritorna al foglio originale
    PL.GotoSheet(nSheet)



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


def set_page_margins(oAktPage):
    """
    Imposta i margini, le distanze di intestazione e piè di pagina,
    la scala di pagina e il centraggio orizzontale per l'oggetto di pagina fornito.

    Args:
        oAktPage: Oggetto rappresentante lo stile di pagina.
    """
    # ~ oDoc = LeenoUtils.getDocument()
    # ~ oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)

    # Imposta i margini della pagina
    oAktPage.TopMargin = 1000
    oAktPage.BottomMargin = 1350
    oAktPage.LeftMargin = 1000
    oAktPage.RightMargin = 1000
    oAktPage.FooterLeftMargin = 0
    oAktPage.FooterRightMargin = 0
    oAktPage.HeaderLeftMargin = 0
    oAktPage.HeaderRightMargin = 0
    oAktPage.HeaderBodyDistance = 0
    oAktPage.FooterBodyDistance = 0
    oAktPage.PageScale = 0
    oAktPage.CenterHorizontally = True
    oAktPage.ScaleToPagesX = 1
    oAktPage.ScaleToPagesY = 0

def set_page_borders(oAktPage):
    """
    Rimuove i bordi superiore, inferiore, sinistro e destro 
    dell'oggetto pagina fornito, impostando sia la larghezza della linea 
    che la larghezza della linea esterna a zero.

    Args:
        oAktPage: Oggetto rappresentante lo stile di pagina.
    """
    # ~ oDoc = LeenoUtils.getDocument()
    # ~ oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByIndex(n)

    # Azzera i bordi
    borders = ['TopBorder', 'BottomBorder', 'LeftBorder', 'RightBorder']
    for border in borders:
        bordo = getattr(oAktPage, border)
        bordo.LineWidth = 0
        bordo.OuterLineWidth = 0
        setattr(oAktPage, border, bordo)

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

        htxt = 8.0 / 100 * oAktPage.PageScale

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

def importa_stili_pagina(overwrite = False):
    """
    Importa solo gli stili di pagina dal template di LeenO,
    con la possibilità di sovrascrivere quelli esistenti.

    Args:
        overwrite (bool): se True sovrascrive gli stili esistenti.
    """

    filename = PL.LeenO_path() + '/template/leeno/Computo_LeenO.ots'

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
    oProgressBar = PL.create_progress_bar(title = 'Importazione stili pagina in corso...', steps = len(oDoc.Sheets.ElementNames))
    n = 1 
    for el in oDoc.Sheets.ElementNames:
        oProgressBar.Value = n
        n += 1
        oDoc.CurrentController.setActiveSheet(
            oDoc.getSheets().getByName(el))
        oSheet = oDoc.getSheets().getByName(el)
        LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    oProgressBar.reset()
    oProgressBar.end()

########################################################################

def importa_stili_pagina_non_presenti():
    """
    Importa solo gli stili di pagina non presenti nel file corrente
    dal template di LeenO.
    """
    LeenoUtils.DocumentRefresh(False)
    # Percorso del file di template
    filename = PL.LeenO_path() + '/template/leeno/Computo_LeenO.ots'

    oDoc = LeenoUtils.getDocument()

    # Ottieni gli stili di pagina già presenti nel documento corrente
    existing_page_styles = oDoc.StyleFamilies.getByName('PageStyles')
    existing_style_names = [existing_page_styles.getByIndex(i).Name for i in range(existing_page_styles.Count)]
    lun_1 = len(existing_page_styles)

    # Creare la lista di PropertyValue per le opzioni di caricamento
    loadOptions = [
        PL.crea_property_value("LoadPageStyles", True),
        PL.crea_property_value("LoadCellStyles", False),
        PL.crea_property_value("LoadTextStyles", False),
        PL.crea_property_value("LoadFrameStyles", False),
        PL.crea_property_value("LoadNumberingStyles", False),
        PL.crea_property_value("OverwriteStyles", False)  # Non sovrascrivere per default
    ]

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
    except Exception as e:
        # ~ DLG.chi(f"Errore durante l'importazione degli stili: {str(e)}")
        pass
    # attiva la progressbar
    if lun_1 < len(oDoc.StyleFamilies.getByName('PageStyles')):
        oProgressBar = PL.create_progress_bar(title = 'Importazione stili pagina in corso...', steps = len(oDoc.Sheets.ElementNames))
        n = 1 
        for el in oDoc.Sheets.ElementNames:
            oProgressBar.Value = n
            n += 1
            oDoc.CurrentController.setActiveSheet(
                oDoc.getSheets().getByName(el))
            oSheet = oDoc.getSheets().getByName(el)
            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
        oProgressBar.reset()
        oProgressBar.end()
    LeenoUtils.DocumentRefresh(True)
