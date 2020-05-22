"""
    LeenO - modulo di gestione dialoghi
"""

import os
import threading
import uno

from LeenoUtils import getComponentContext, getDesktop, getDocument
import pyleeno as PL

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
# from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK_CANCEL
from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
# from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO_CANCEL
# from com.sun.star.awt.MessageBoxButtons import BUTTONS_RETRY_CANCEL
# from com.sun.star.awt.MessageBoxButtons import BUTTONS_ABORT_IGNORE_RETRY

# from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK
# from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_CANCEL
# from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_RETRY
# from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_YES
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_NO
# from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_IGNORE

from com.sun.star.awt.MessageBoxType import MESSAGEBOX
# from com.sun.star.awt.MessageBoxType import INFOBOX
# from com.sun.star.awt.MessageBoxType import WARNINGBOX
# from com.sun.star.awt.MessageBoxType import ERRORBOX
from com.sun.star.awt.MessageBoxType import QUERYBOX


# MAH...
oDialogo_attesa = None


def createUnoService(serv):
    '''
    QUESTA BISOGNA VEDERE DOVE METTERLA E SE LASCIARLA
    '''
    return getComponentContext().getServiceManager().createInstance(serv)


# rif.: https://wiki.openoffice.org/wiki/PythonDialogBox


def barra_di_stato(testo='', valore=0):
    '''Informa l'utente sullo stato progressivo dell'elaborazione.'''
    oDoc = getDocument()
    oProgressBar = oDoc.CurrentController.Frame.createStatusIndicator()
    oProgressBar.start('', 100)
    oProgressBar.Value = valore
    oProgressBar.Text = testo
    oProgressBar.reset()
    oProgressBar.end()


########################################################################
def filedia(titolo='Scegli il file...', est='*.*', mode=0):
    """
    titolo  { string }  : titolo del FilePicker
    est     { string }  : filtro di visualizzazione file
    mode    { integer } : modalità di gestione del file

    Apri file:  `mode in(0, 6, 7, 8, 9)`
    Salva file: `mode in(1, 2, 3, 4, 5, 10)`
    see:('''http://api.libreoffice.org/docs/idl/ref/
            namespacecom_1_1sun_1_1star_1_1ui_1_1
            dialogs_1_1TemplateDescription.html''' )
    see:('''http://stackoverflow.com/questions/30840736/
        libreoffice-how-to-create-a-file-dialog-via-python-macro''')
    """
    estensioni = {'*.*': 'Tutti i file(*.*)',
                  '*.odt': 'Writer(*.odt)',
                  '*.ods': 'Calc(*.ods)',
                  '*.odb': 'Base(*.odb)',
                  '*.odg': 'Draw(*.odg)',
                  '*.odp': 'Impress(*.odp)',
                  '*.odf': 'Math(*.odf)',
                  '*.xpwe': 'Primus(*.xpwe)',
                  '*.xml': 'XML(*.xml)',
                  '*.dat': 'dat(*.dat)', }
    try:
        oFilePicker = createUnoService("com.sun.star.ui.dialogs.OfficeFilePicker")
        oFilePicker.initialize((mode, ))
        oDoc = getDocument()
        oFilePicker.setDisplayDirectory(os.path.dirname(oDoc.getURL()))
        oFilePicker.Title = titolo
        app = estensioni.get(est)
        oFilePicker.appendFilter(app, est)
        if oFilePicker.execute():
            oDisp = uno.fileUrlToSystemPath(oFilePicker.getFiles()[0])
        return oDisp

    except Exception:
        MsgBox('Il file non è stato selezionato', 'ATTENZIONE!')
        return None

########################################################################


def chi(s):  # s = oggetto
    '''
    s    { object }  : oggetto da interrogare
    mostra un dialog che indica il tipo di oggetto ed i metodi ad esso applicabili
    '''
    doc = getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    s1 = str(s) + '\n\n' + str(dir(s).__str__())
    MessageBox(parentwin, str(s1), str(type(s)), 'infobox')


def DlgSiNo(s, t='Titolo'):  # s = messaggio | t = titolo
    '''
    Visualizza il menù di scelta sì/no
    restituisce 2 per sì e 3 per no
    '''
    doc = getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    # s = 'This a message'
    # t = 'Title of the box'
    # MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
    return MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO + DEFAULT_BUTTON_NO)


def MsgBox(s, t=''):  # s = messaggio | t = titolo
    '''
    Visualizza una message box
    '''
    doc = getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    # s = 'This a message'
    # t = 'Title of the box'
    # res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
    # chi(res)
    # return
    # s = res
    # t = 'Titolo'
    if t is None:
        t = 'messaggio'
    MessageBox(parentwin, str(s), t, 'infobox')


def MessageBox(ParentWin, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
    '''
    Show a message box with the UNO based toolkit
    '''
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext('com.sun.star.awt.Toolkit', ctx)
    myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)

    return myBox.execute()
# [　入手元　]


def mri(target):
    '''
    @@ DA DOCUMENTARE... MA A CHE SERVE ???
    '''
    ctx = getComponentContext()
    mrii = ctx.ServiceManager.createInstanceWithContext('mytools.Mri', ctx)
    mrii.inspect(target)
    MsgBox('MRI in corso...', 'avviso')


#######################################################################
def dlg_attesa(msg=''):
    '''
    definisce la variabile globale oDialogo_attesa
    che va gestita così negli script:

    oDialogo_attesa = dlg_attesa()
    attesa().start() #mostra il dialogo
    ...
    oDialogo_attesa.endExecute() #chiude il dialogo
    '''
    psm = uno.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    global oDialogo_attesa
    oDialogo_attesa = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgAttesa?language=Basic&location=application")

    # oDialog1Model = oDialogo_attesa.Model  # oDialogo_attesa è una variabile generale

    sString = oDialogo_attesa.getControl("Label2")
    sString.Text = msg  # 'ATTENDI...'
    oDialogo_attesa.Title = 'Operazione in corso...'
    sUrl = PL.LeenO_path() + '/icons/attendi.png'
    oDialogo_attesa.getModel().ImageControl1.ImageURL = sUrl
    return oDialogo_attesa


class attesa(threading.Thread):
    '''
    avvia il dialogo di attesa
    http://bit.ly/2fzfsT7
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        global oDialogo_attesa
        oDialogo_attesa.endExecute()  # chiude il dialogo
        oDialogo_attesa.execute()
