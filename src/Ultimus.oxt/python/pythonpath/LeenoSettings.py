'''
Modulo per la modifica delle impostazioni di LeenO
'''
import Dialogs
import LeenoConfig
import LeenoUtils
import DocUtils

_JOBSETTINGSITEMS = (
    'committente',
    'stazioneAppaltante',
    'rup',
    'progettista',
    'data',
    'revisione',
    'dataRevisione',
)

_PRINTSETTINGSITEMS = (
)

def _load2(oDoc, cfg, root, name):
    # per prima cosa tenta il caricamento dal documento
    # se nullo carica dal config generale
    v = DocUtils.getDocUserDefinedAttribute(oDoc, root + '.' + name)
    if v:
        return v
    v = cfg.read(root, name)
    if v:
        return v
    return ''


def _store2(oDoc, cfg, root, name, val):
    DocUtils.setDocUserDefinedAttribute(oDoc, root + '.' + name, val)
    cfg.write(root, name, val)


def loadJobSettings(oDoc):
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'Lavoro')
    if data is None or len(data) == 0:
        data = cfg.readBlock('Lavoro', True)
    return data

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
                Dialogs.FixedText(Text='Committente'),
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


def MENU_UserSettings():
    pass


def MENU_JobSettings():

    oDoc = LeenoUtils.getDocument()
    js = loadJobSettings(oDoc)

    dlg = JobSettingsDialog()
    dlg.setData(js)

    if dlg.run() >= 0:
        js = dlg.getData(_JOBSETTINGSITEMS)
        storeJobSettings(oDoc, js)


def PrintSettingsDialog():
    # dimensione dell'icona grande
    imgW = Dialogs.getBigIconSize()[0] * 3
    fieldW, dummy = Dialogs.getTextBox("W" * 30)
    posW, dummy = Dialogs.getTextBox("SinistraXX")

    return Dialogs.Dialog(Title='Impostazioni stampa / PDF',  Horz=False, CanClose=True,  Items=[
        Dialogs.HSizer(Items=[
            Dialogs.VSizer(Items=[
                Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/printersettings.png', MinWidth=imgW),
                Dialogs.Spacer(),
            ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Documento con le copertine', FixedWidth=fieldW),
                Dialogs.Spacer(),
                Dialogs.FileControl(Id="fileCopertine", Types='*.ods'),
                Dialogs.Spacer(),
                Dialogs.FixedText(Text='Selezionare copertina in uso'),
                Dialogs.Spacer(),
                Dialogs.ListBox(List={'trulla', 'llero', 'ullalla', 'oilliiiiiiiiii'}),
                Dialogs.Spacer(),
                Dialogs.FixedText(Text='Intestazione'),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Sinistra', FixedWidth=posW),
                    Dialogs.Edit(Id="intSx", Text='TEMPORANEO'),
                ]),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Centro', FixedWidth=posW),
                    Dialogs.Edit(Id="intCenter", Text='TEMPORANEO'),
                ]),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Destra', FixedWidth=posW),
                    Dialogs.Edit(Id="intDx", Text='TEMPORANEO'),
                ]),
                Dialogs.Spacer(),
                Dialogs.FixedText(Text='Piè di pagina'),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Sinistra', FixedWidth=posW),
                    Dialogs.Edit(Id="ppSx", Text='TEMPORANEO'),
                ]),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Centro', FixedWidth=posW),
                    Dialogs.Edit(Id="ppCenter", Text='TEMPORANEO'),
                ]),
                Dialogs.Spacer(),
                Dialogs.HSizer(Items=[
                    Dialogs.FixedText(Text='Destra', FixedWidth=posW),
                    Dialogs.Edit(Id="ppDx", Text='TEMPORANEO'),
                ]),
            ]),
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

    dlg = PrintSettingsDialog()

    dlg.run()

