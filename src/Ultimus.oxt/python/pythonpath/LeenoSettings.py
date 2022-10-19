'''
Modulo per la modifica delle impostazioni di LeenO
'''
import Dialogs
import LeenoConfig
import LeenoUtils
import DocUtils
import SheetUtils

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
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'Lavoro')
    if data is None or len(data) == 0:
        data = cfg.readBlock('Lavoro', True)
    return data

def loadPageReplacements(oDoc):
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
    cfg = LeenoConfig.Config()

    DocUtils.storeDataBlock(oDoc, 'Lavoro', js)
    cfg.writeBlock('Lavoro', js, True)

def JobSettingsDialog():
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

    oDoc = LeenoUtils.getDocument()
    js = loadJobSettings(oDoc)

    dlg = JobSettingsDialog()
    dlg.setData(js)

    if dlg.run() >= 0:
        js = dlg.getData(_JOBSETTINGSITEMS)
        storeJobSettings(oDoc, js)

def fixupCover(coversPath, coverName):
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
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'ImpostazioniStampa')
    if data is None or len(data) == 0:
        data = cfg.readBlock('ImpostazioniStampa', True)

    # legge i nomi delle copertine dal file fornito, se esistente
    covers, copertina = fixupCover(data.get('fileCopertine', ''), data.get('copertina', ''))

    data['copertina'] = copertina

    return data, covers

def storePrintSettings(oDoc, js):
    cfg = LeenoConfig.Config()

    DocUtils.storeDataBlock(oDoc, 'ImpostazioniStampa', js)
    cfg.writeBlock('ImpostazioniStampa', js, True)

def PrintSettingsDialog():
    # dimensione dell'icona grande
    imgW = Dialogs.getBigIconSize()[0] * 3
    fieldW, dummy = Dialogs.getTextBox("W" * 30)
    posW, dummy = Dialogs.getTextBox("SinistraXX")

    return Dialogs.Dialog(Title='Impostazioni stampa / PDF',  Horz=False, CanClose=True,  Items=[
        Dialogs.VSizer(Items=[
            Dialogs.FixedText(Text='Intestazione:'),
                Dialogs.Spacer(),

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
            
            Dialogs.Spacer(MinSize = 10),
            Dialogs.HSizer(Items=[
                # ~Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/preview.png', MinWidth=imgW * 1.5),
                # ~Dialogs.Spacer(),
            ]),
            Dialogs.Spacer(MinSize = 10),
            
                Dialogs.FixedText(Text='Piè di pagina:'),
                Dialogs.Spacer(),
            Dialogs.HSizer(Items=[
                Dialogs.VSizer(Items=[
                    # ~Dialogs.FixedText(Text='Sinistra: ', FixedWidth=posW),
                    Dialogs.FixedText(Text='Sinistra: '),
                    Dialogs.ComboBox(Id="ppSx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
                Dialogs.Spacer(MinSize = 10),
                Dialogs.VSizer(Items=[
                    Dialogs.FixedText(Text='Centro: '),
                    Dialogs.ComboBox(Id="ppCenter", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
                Dialogs.Spacer(MinSize = 10),
                Dialogs.VSizer(Items=[
                    Dialogs.FixedText(Text='Destra: '),
                    Dialogs.ComboBox(Id="ppDx", List=_DOCSTRINGS, FixedHeight=20, MaxWidth=200),
                ]),
            ]),
        ]),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Documento con le copertine: '), #FixedWidth=fieldW),
                # ~Dialogs.Spacer(),
                Dialogs.FileControl(Id="fileCopertine", Types='*.ods'),
                Dialogs.Spacer(),
                Dialogs.FixedText(Text='Selezionare copertina in uso: ', FixedHeight=25),
                # ~Dialogs.Spacer(),
                Dialogs.ListBox(Id='copertina'),
                Dialogs.Spacer(MinSize = 3),
        ]),
        Dialogs.Spacer(),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),
            Dialogs.Button(Label='Ok', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/ok.png',  RetVal=1),
            Dialogs.Spacer(),
            Dialogs.Button(Label='Annulla', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/cancel.png',  RetVal=-1),
            Dialogs.Spacer(),
        ]),
    ])


def MENU_PrintSettings():

    oDoc = LeenoUtils.getDocument()
    ps, covers = loadPrintSettings(oDoc)

    dlg = PrintSettingsDialog()
    dlg.getWidget('copertina').setList(covers)
    dlg.setData(ps)

    if dlg.run() >= 0:
        ps = dlg.getData(_PRINTSETTINGSITEMS)
        storePrintSettings(oDoc, ps)
