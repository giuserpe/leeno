"""
    LeenO - modulo di gestione dialoghi
"""

import os
import threading
import uno

from LeenoUtils import getComponentContext, getDesktop, getDocument, createUnoService
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
    ctx = getComponentContext()
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
    psm = getComponentContext().ServiceManager
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


def ScegliElaborato(titolo):
    '''
    Permetta la scelta dell'elaborato da trattare e restituisce il suo nome
    '''
    oDoc = getDocument()
    psm = getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgXLO = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dialog_XLO?language=Basic&location=application"
    )
    # oDialog1Model = oDlgXLO.Model
    oDlgXLO.Title = titolo  # Menù import XPWE'

    for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
        try:
            importo = oDoc.getSheets().getByName(el).getCellRangeByName(
                'A2').String
            if el == 'COMPUTO':
                oDlgXLO.getControl(
                    "CME_XLO").Label = '~Computo:     € ' + importo
            if el == 'VARIANTE':
                oDlgXLO.getControl(
                    "VAR_XLO").Label = '~Variante:    € ' + importo
            if el == 'CONTABILITA':
                oDlgXLO.getControl(
                    "CON_XLO").Label = 'C~ontabilità: € ' + importo
            #  else:
            #  oDlgXLO.getControl("CON_XLO").Label  = 'Contabilità: €: 0,0'
        except Exception:
            pass

    if oDlgXLO.execute() == 1:
        if oDlgXLO.getControl("CME_XLO").State:
            elaborato = 'COMPUTO'
        elif oDlgXLO.getControl("VAR_XLO").State:
            elaborato = 'VARIANTE'
        elif oDlgXLO.getControl("CON_XLO").State:
            elaborato = 'CONTABILITA'
        elif oDlgXLO.getControl("EP_XLO").State:
            elaborato = 'Elenco'
    return elaborato

