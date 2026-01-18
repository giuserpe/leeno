"""
    LeenO - modulo di gestione dialoghi
"""

import threading

import LeenoUtils
import LeenoGlobals
import pyleeno as PL
import Dialogs

import inspect
import os
import traceback
import subprocess

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

def chi(s = 'pausa...', OFF=False):
    '''
    s    { object }  : oggetto da interrogare
    Mostra un dialog che indica il tipo di oggetto ed i metodi ad esso applicabili.
    '''

    # Ottieni il frame del chiamante per recuperare il numero di linea, il nome del file e il nome della funzione
    caller_frame = inspect.stack()[1]
    line_number = caller_frame.lineno
    full_file_path = caller_frame.filename  # Ottieni il percorso completo
    full_file_path = LeenoGlobals.dest() + full_file_path.split('LeenO.oxt')[-1]

    file_name = os.path.basename(full_file_path)  # Solo il nome del file
    function_name = caller_frame.function

    # Verifica che il documento sia disponibile
    try:
        doc = LeenoUtils.getDocument()
        if not doc:
            return
        parentwin = doc.CurrentController.Frame.ContainerWindow
    except Exception:
        return


    if parentwin:
        # Costruisci il messaggio
        s1 = (

            f'Rappresentazione dell\'oggetto:\n{str(s)}\n\n'
            f'Metodi e attributi disponibili:\n{str(dir(s))}\n\n'
            f'Nome del file chiamante: {file_name}\n'
            f'Numero di linea della chiamata: {line_number}\n'
            f'Nome della funzione chiamante: {function_name}()'
        )

        # Apri il file e vai alla riga specificata
        if OFF:
            pass
        else:
            PL.apri_con_editor(full_file_path, line_number)

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
    error_type = type(e).__name__
    stack_trace = traceback.format_exc()

    chi(f'Errore di tipo "{error_type}": {str(e)}\n\nTraccia dello stack:\n{stack_trace}')


def DlgSiNo(s, t='Titolo'):  # s = messaggio | t = titolo
    '''
    Visualizza il menù di scelta sì/no
    restituisce 2 per sì e 3 per no
    '''
    # Verifica che il documento sia disponibile
    try:
        doc = LeenoUtils.getDocument()
        if not doc:
            return
        parentwin = doc.CurrentController.Frame.ContainerWindow
    except Exception:
        return
    # s = 'This a message'
    # t = 'Title of the box'
    # MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
    return MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO + DEFAULT_BUTTON_NO)


def MsgBox(Text, Title=''):  # s = messaggio | t = titolo
    '''
    Visualizza una message box
    '''
    # Verifica che il documento sia disponibile
    try:
        doc = LeenoUtils.getDocument()
        if not doc:
            return
        parentwin = doc.CurrentController.Frame.ContainerWindow
    except Exception:
        return
    # s = 'This a message'
    # t = 'Title of the box'
    # res = MessageBox(parentwin, s, t, QUERYBOX, BUTTONS_YES_NO_CANCEL + DEFAULT_BUTTON_NO)
    # chi(res)
    # return
    # s = res
    # t = 'Titolo'
    if Title is None:
        Title = 'messaggio'
    MessageBox(parentwin, str(Text), Title, 'infobox')


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

# def ScegliElaborato(titolo):
#     '''
#     Permetta la scelta dell'elaborato da trattare e restituisce il suo nome
#     '''
#     oDoc = LeenoUtils.getDocument()
#     psm = LeenoUtils.getComponentContext().ServiceManager
#     dp = psm.createInstance("com.sun.star.awt.DialogProvider")
#     oDlgXLO = dp.createDialog(
#         "vnd.sun.star.script:UltimusFree2.Dialog_XLO?language=Basic&location=application"
#     )
#     # oDialog1Model = oDlgXLO.Model
#     oDlgXLO.Title = titolo  # Menù import XPWE'

#     for el in ("COMPUTO", "VARIANTE", "CONTABILITA"):
#         try:
#             importo = oDoc.getSheets().getByName(el).getCellRangeByName(
#                 'A2').String
#             if el == 'COMPUTO':
#                 oDlgXLO.getControl(
#                     "CME_XLO").Label = '~Computo:     ' + importo
#             if el == 'VARIANTE':
#                 oDlgXLO.getControl(
#                     "VAR_XLO").Label = '~Variante:    ' + importo
#             if el == 'CONTABILITA':
#                 oDlgXLO.getControl(
#                     "CON_XLO").Label = 'C~ontabilità: ' + importo
#         except Exception:
#             pass

#     if oDlgXLO.execute() == 1:
#         if oDlgXLO.getControl("CME_XLO").State:
#             elaborato = 'COMPUTO'
#         elif oDlgXLO.getControl("VAR_XLO").State:
#             elaborato = 'VARIANTE'
#         elif oDlgXLO.getControl("CON_XLO").State:
#             elaborato = 'CONTABILITA'
#         elif oDlgXLO.getControl("EP_XLO").State:
#             elaborato = 'Elenco'
#     return elaborato

def ScegliElaborato(Titolo="Titolo", flag="export"):
    """
    Mostra un dialogo per scegliere l'elaborato da trattare e restituisce il nome scelto.
    """
    oDoc = LeenoUtils.getDocument()
    psm = LeenoUtils.getComponentContext().ServiceManager
    dp = psm.createInstance("com.sun.star.awt.DialogProvider")
    oDlgXLO = dp.createDialog(
        "vnd.sun.star.script:UltimusFree2.Dialog_XLO?language=Basic&location=application"
    )
    oDlgXLO.Title = Titolo

    controlli = {
        "COMPUTO": "CME_XLO",
        "VARIANTE": "VAR_XLO",
        "CONTABILITA": "CON_XLO",
    }

    Image = PL.LeenO_path() + "/python/pythonpath/Icons-Big/ok.png"
    oDlgXLO.getModel().ImageControl1.ImageURL = Image

    if flag == "export":
        for nome, ctrl in controlli.items():
            try:
                importo = oDoc.getSheets().getByName(nome).getCellRangeByName("A2").String
                etichette = {
                    "COMPUTO": f"~Computo:     {importo}",
                    "VARIANTE": f"~Variante:    {importo}",
                    "CONTABILITA": f"C~ontabilità: {importo}",
                }
                oDlgXLO.getControl(ctrl).Label = etichette[nome]

                if oDoc.getSheets().hasByName("COMPUTO"):
                    oDlgXLO.getControl("CME_XLO").State = 1
                elif oDoc.getSheets().hasByName("VARIANTE"):
                    oDlgXLO.getControl("VAR_XLO").State = 1
                elif oDoc.getSheets().hasByName("CONTABILITA"):
                    oDlgXLO.getControl("CON_XLO").State = 1

            except Exception:
                oDlgXLO.getControl(ctrl).Label = f"~{nome.capitalize()}: (nessun dato)"
                oDlgXLO.getControl(ctrl).Enable = False
    else:
        etichette = {
            "CME_XLO": "~Computo - Variante",
            "VAR_XLO": "~Computo - C~ontabilità",
            "CON_XLO": "~Variante - Contabilità",
        }

        if not oDoc.Sheets.hasByName('VARIANTE'):
            oDlgXLO.getControl('CME_XLO').setEnable(0)
            oDlgXLO.getControl('CON_XLO').setEnable(0)
            oDlgXLO.getControl("VAR_XLO").State = 1

        if not oDoc.Sheets.hasByName('CONTABILITA'):
            oDlgXLO.getControl('VAR_XLO').setEnable(0)
            oDlgXLO.getControl('CON_XLO').setEnable(0)
            # oDlgXLO.getControl("CME_XLO").State = 1


        # if oDoc.getSheets().hasByName("COMPUTO"):
        #     oDlgXLO.getControl("CME_XLO").State = 1
        # elif oDoc.getSheets().hasByName("VARIANTE"):
        #     oDlgXLO.getControl("VAR_XLO").State = 1
        # elif oDoc.getSheets().hasByName("CONTABILITA"):
        #     oDlgXLO.getControl("CON_XLO").State = 1

        for ctrl, testo in etichette.items():
            oDlgXLO.getControl(ctrl).Label = testo

        if ( not oDoc.getSheets().hasByName("VARIANTE") and
            not oDoc.getSheets().hasByName("CONTABILITA")):
            Dialogs.Info(
                Title="Informazione",
                Text="Nessuna Variante o Contabilità presente per il confronto."
            )
            return

    # Esegue il dialogo
    if oDlgXLO.execute() != 1:
        oDlgXLO.dispose()
        return None

    # Mappatura scelta → elaborato
    mappa_scelte = {
        "CME_XLO": {"export": "COMPUTO", "parallelo": "computo_variante"},
        "VAR_XLO": {"export": "VARIANTE", "parallelo": "computo_contabilità"},
        "CON_XLO": {"export": "CONTABILITA", "parallelo": "variante_contabilità"},
        # "EP_XLO":  {"export": "Elenco", "parallelo": ""},
    }

    elaborato = None
    for ctrl, mappa in mappa_scelte.items():
        try:
            if oDlgXLO.getControl(ctrl).State:
                elaborato = mappa.get(flag)
                break
        except Exception:
            continue

    oDlgXLO.dispose()
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
