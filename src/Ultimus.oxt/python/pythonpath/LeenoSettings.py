'''
Modulo per la modifica delle impostazioni di LeenO
'''
import Dialogs
import LeenoConfig
import LeenoUtils
import DocUtils

def _load2(oDoc, cfg, name):
    # per prima cosa tenta il caricamento dal documento
    # se nullo carica dal config generale
    v = DocUtils.getDocUserDefinedAttribute(oDoc, 'Lavoro.' + name)
    if v:
        return v
    v = cfg.read('Lavoro', name)
    if v:
        return v
    return ''


def _store2(oDoc, cfg, name, val):
    DocUtils.setDocUserDefinedAttribute(oDoc, 'Lavoro.' + name, val)
    cfg.write('Lavoro', name, val)


def loadJobSettings(oDoc):
    cfg = LeenoConfig.Config()
    res = {}

    res['committente'] = _load2(oDoc, cfg, 'committente')
    res['stazioneAppaltante'] = _load2(oDoc, cfg, 'stazioneAppaltante')
    res['rup'] = _load2(oDoc, cfg, 'rup')
    res['progettista'] = _load2(oDoc, cfg, 'progettista')
    res['data'] = LeenoUtils.string2Date(_load2(oDoc, cfg, 'data'))
    res['revisione'] = _load2(oDoc, cfg, 'revisione')
    res['dataRevisione'] = LeenoUtils.string2Date(_load2(oDoc, cfg, 'dataRevisione'))

    return res

def storeJobSettings(oDoc, js):
    cfg = LeenoConfig.Config()

    _store2(oDoc, cfg, 'committente', js['committente'])
    _store2(oDoc, cfg, 'stazioneAppaltante', js['stazioneAppaltante'])
    _store2(oDoc, cfg, 'rup', js['rup'])
    _store2(oDoc, cfg, 'progettista', js['progettista'])
    _store2(oDoc, cfg, 'data', LeenoUtils.date2String(js['data'], 1))
    _store2(oDoc, cfg, 'revisione', js['revisione'])
    _store2(oDoc, cfg, 'dataRevisione', LeenoUtils.date2String(js['dataRevisione'], 1))

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

    dlg['committente'].setText(js['committente'])
    dlg['stazioneAppaltante'].setText(js['stazioneAppaltante'])
    dlg['rup'].setText(js['rup'])
    dlg['progettista'].setText(js['progettista'])
    dlg['data'].setDate(js['data'])
    dlg['revisione'].setText(js['revisione'])
    dlg['dataRevisione'].setDate(js['dataRevisione'])

    if dlg.run() >= 0:
        js = {}
        js['committente'] = dlg['committente'].getText()
        js['stazioneAppaltante'] = dlg['stazioneAppaltante'].getText()
        js['rup'] = dlg['rup'].getText()
        js['progettista'] = dlg['progettista'].getText()
        js['data'] = dlg['data'].getDate()
        js['revisione'] = dlg['revisione'].getText()
        js['dataRevisione'] = dlg['dataRevisione'].getDate()
        storeJobSettings(oDoc, js)

def MENU_PrintSettings():
    pass
