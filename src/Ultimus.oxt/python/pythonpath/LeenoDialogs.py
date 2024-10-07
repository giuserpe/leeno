"""
    LeenO - modulo di gestione dialoghi
"""

import threading

import LeenoUtils
import pyleeno as PL
import Dialogs

import inspect
import os
import traceback

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

# rif.: https://wiki.openoffice.org/wiki/PythonDialogBox

def barra_di_stato(testo='', valore=0):
    '''Informa l'utente sullo stato progressivo dell'elaborazione.'''
    oDoc = LeenoUtils.getDocument()
    oProgressBar = oDoc.CurrentController.Frame.createStatusIndicator()
    oProgressBar.start('', 100)
    oProgressBar.Value = valore
    oProgressBar.Text = testo
    oProgressBar.reset()
    oProgressBar.end()


def chi(s):
    '''
    s    { object }  : oggetto da interrogare
    Mostra un dialog che indica il tipo di oggetto ed i metodi ad esso applicabili.
    '''

    # Ottieni il frame del chiamante per recuperare il numero di linea, il nome del file e il nome della funzione
    caller_frame = inspect.stack()[1]
    line_number = caller_frame.lineno
    file_name = os.path.basename(caller_frame.filename)
    function_name = caller_frame.function  # Nome della funzione chiamante
    # Verifica che il documento sia disponibile
    doc = LeenoUtils.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow if doc else None

    if parentwin is not None:
        # Costruisci il messaggio
        s1 = (
            f'Rappresentazione dell\'oggetto:\n{str(s)}\n\n'
            f'Metodi e attributi disponibili:\n{str(dir(s))}\n\n'
            f'Nome del file chiamante: {file_name}\n'
            f'Numero di linea della chiamata: {line_number}\n'
            f'Nome della funzione chiamante: {function_name}()'
        )

        # Traccia dello stack attuale (se presente)
        stack_trace = traceback.format_exc()
        if stack_trace.strip():  # Se c'è una traccia di stack, includerla
            s1 += f'\n\nTraccia dello stack:\n{stack_trace}'

        # Mostra il messaggio in un dialogo
        MessageBox(parentwin, s1, f'Tipo di oggetto: {str(type(s))}', 'infobox')



def errore(e):
    '''
    Mostra un messaggio dettagliato dell'errore, includendo il tipo di eccezione,
    il messaggio e la traccia dello stack per facilitare il debug.

    Args:
        e (Exception): L'eccezione catturata.

    Comportamento:
        Visualizza un dialogo con il tipo di errore, il messaggio e la traccia completa dello stack.
    '''
    import traceback
    error_type = type(e).__name__
    stack_trace = traceback.format_exc()

    chi(f'Errore di tipo "{error_type}": {str(e)}\n\nTraccia dello stack:\n{stack_trace}')


def DlgSiNo(s, t='Titolo'):  # s = messaggio | t = titolo
    '''
    Visualizza il menù di scelta sì/no
    restituisce 2 per sì e 3 per no
    '''
    doc = LeenoUtils.getDocument()
    parentwin = doc.CurrentController.Frame.ContainerWindow
    # s = 'This a message'
    # t = 'Title of the box'
    # MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
    return MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO + DEFAULT_BUTTON_NO)


def MsgBox(s, t=''):  # s = messaggio | t = titolo
    '''
    Visualizza una message box
    '''
    doc = LeenoUtils.getDocument()
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
    ctx = LeenoUtils.getComponentContext()
    sm = ctx.ServiceManager
    sv = sm.createInstanceWithContext('com.sun.star.awt.Toolkit', ctx)
    myBox = sv.createMessageBox(ParentWin, MsgType, MsgButtons, MsgTitle, MsgText)

    return myBox.execute()
# [　入手元　]


def mri(target):
    '''
    Inspector https://extensions.openoffice.org/project/MRI
    '''
    ctx = LeenoUtils.getComponentContext()
    mrii = ctx.ServiceManager.createInstanceWithContext('mytools.Mri', ctx)
    mrii.inspect(target)
    MsgBox('MRI in corso...', 'avviso')


#######################################################################
def dlg_attesa(msg=''):
    '''
    definisce la variabile globale oDialogo_attesa
    che va gestita così negli script:

    dlg_attesa()
    attesa().start() #mostra il dialogo
    ...
    LeenoUtils.getGlobalVar('oDialogo_attesa').endExecute() #chiude il dialogo
    '''
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDialogo_attesa = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.DlgAttesa?language=Basic&location=application")

    # oDialog1Model = oDialogo_attesa.Model  # oDialogo_attesa è una variabile generale

    sString = oDialogo_attesa.getControl("Label2")
    sString.Text = msg  # 'ATTENDI...'
    oDialogo_attesa.Title = 'Operazione in corso...'
    sUrl = PL.LeenO_path() + '/icons/attendi.png'
    oDialogo_attesa.getModel().ImageControl1.ImageURL = sUrl
    LeenoUtils.setGlobalVar('oDialogo_attesa', oDialogo_attesa)
    return oDialogo_attesa


class attesa(threading.Thread):
    '''
    avvia il dialogo di attesa
    http://bit.ly/2fzfsT7
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        LeenoUtils.getGlobalVar('oDialogo_attesa').endExecute()  # chiude il dialogo
        LeenoUtils.getGlobalVar('oDialogo_attesa').execute()


def ScegliElaborato(titolo):
    '''
    Permetta la scelta dell'elaborato da trattare e restituisce il suo nome
    '''
    oDoc = LeenoUtils.getDocument()
    psm = LeenoUtils.getComponentContext().ServiceManager
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
                    "CME_XLO").Label = '~Computo:     ' + importo
            if el == 'VARIANTE':
                oDlgXLO.getControl(
                    "VAR_XLO").Label = '~Variante:    ' + importo
            if el == 'CONTABILITA':
                oDlgXLO.getControl(
                    "CON_XLO").Label = 'C~ontabilità: ' + importo
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


def ScegliElabDest(*, Title='', AskTarget=False, AskSort=False, Sort=False, ValComputo=None, ValVariante=None, ValContabilita=None):

    # altezza radiobutton, per fare il testo con il valore uguale
    # ed averli quindi allineati
    radioH = Dialogs.getRadioButtonHeight()

    # dimensione dell'icona col punto di domanda
    imgW = Dialogs.getBigIconSize()[0] * 2

    # i 3 valori totali
    vCmp = '' if ValComputo is None else '{:9.2f} €'.format(ValComputo)
    vVar = '' if ValVariante is None else '{:9.2f} €'.format(ValVariante)
    vCon = '' if ValContabilita is None else '{:9.2f} €'.format(ValContabilita)

    # handler per il button di info
    def infoHandler(owner,  widgetId,  widget,  cmdStr):
        if widgetId == 'sortInfo':
            Dialogs.Info(Title="Note sull'ordinamento", Text=
                "È possibile effettuare l'ordinamento delle voci di\n"
                "computo in  base  alla  struttura delle categorie.\n"
                "Se il file in origine è particolarmente disordinato,\n"
                "riceverai un messaggio che ti indica come intervenire.\n"
                "Se il risultato finale non dovesse andar bene, puoi\n"
                "ripetere l'importazione senza il riordino delle voci\n"
                "de-selezionando la casella relativa"
            )
        return False

    # prima parte fissa del dialogo
    dlg = Dialogs.Dialog(Title=Title, Handler=infoHandler, Items=[
        Dialogs.HSizer(Items=[
            Dialogs.VSizer(Items=[
                Dialogs.Spacer(),
                Dialogs.ImageControl(Id="img", Image="Icons-Big/question.png", MinWidth=imgW),
                Dialogs.Spacer()
            ]),
            Dialogs.VSizer(Id="vsizer", Items=[
                Dialogs.GroupBox(Label="Scegli elaborato", Items=[
                    Dialogs.HSizer(Items=[
                        Dialogs.RadioGroup(Id="elab", Items=[
                            "Computo",
                            "Variante",
                            "Contabilità",
                            "Elenco Prezzi"
                        ]),
                        Dialogs.Spacer(),
                        Dialogs.VSizer(Items=[
                            Dialogs.FixedText(Id="TotComputo", Text=vCmp, MinHeight=radioH),
                            Dialogs.FixedText(Id="TotVariante", Text=vVar, MinHeight=radioH),
                            Dialogs.FixedText(Id="TotContabilità", Text=vCon, MinHeight=radioH),
                            Dialogs.Spacer()
                        ])
                    ])
                ])
            ])
        ])
    ])
    if AskTarget:
        dlg["vsizer"].add(
            Dialogs.Spacer(),
            Dialogs.GroupBox(Label="Scegli destinazione", Items=[
                Dialogs.RadioGroup(Id="dest", Items=[
                    "Documento corrente",
                    "Nuovo documento"
                ])
            ])
        )

    if AskSort:
        dlg["vsizer"].add(
            Dialogs.Spacer(),
            Dialogs.GroupBox(Label="Ordinamento computo", Items=[
                Dialogs.CheckBox(Id="sort", Label="Ordina computo", State=Sort),
                Dialogs.Spacer(),
                Dialogs.Button(Id="sortInfo", Label="Info su ordinamento", Icon="Icons-24x24/info.png")
            ])
        )

    dlg["vsizer"].add(
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Button(Label="Ok", RetVal=1, Icon="Icons-24x24/ok.png"),
            Dialogs.Spacer(),
            Dialogs.Button(Label="Annulla", RetVal=-1, Icon="Icons-24x24/cancel.png"),
        ])
    )

    # check if we canceled the job
    if dlg.run() == -1:
        return None

    elab = ('COMPUTO', 'VARIANTE', 'CONTABILITA', 'Elenco')[dlg['elab'].getCurrent()]
    dest = ('CORRENTE', 'NUOVO')[dlg['dest'].getCurrent() if AskTarget else 1]
    sort = dlg['sort'].getState()

    return {'elaborato':elab, 'destinazione':dest, 'ordina':sort}
