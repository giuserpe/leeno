'''
    PONTE TEMPORANEO TRA BASIC E PYTHON
    IN QUESTO FILE SONO CONTENUTE TUTTE LE FUNZIONI CHIAMATE DA BASIC
    ACCENTRATE PER POTERLE ELIMINARE PIAN PIANO
'''

# set this to 1 if you want a dialog on every call from basic to python
# set to 0 to avoid it
ALERT_BASIC_CALLS = 1

# dirty trick to have pythonpath added if missing
import sys, os, inspect
myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
if(myPath not in sys.path):
    sys.path.append(myPath)

import Dialogs

# builtins dictionary in portable way... sigh
if type(__builtins__) == type(sys):
    bDict = __builtins__.__dict__
else:
    bDict = __builtins__

if 'CALL_SET' not in bDict:
    bDict['CALL_SET'] = set()

def callAlert():
    if ALERT_BASIC_CALLS:
        funcName = inspect.stack()[1][3]
        if not funcName in CALL_SET:
            Dialogs.Exclamation(Title="LeenoBasicBridge", Text=f"Chiamata da basic a\n'{funcName}'")
            CALL_SET.add(funcName)

import LeenoUtils
import pyleeno as PL

def MENU_debug():
    '''
    MENU_debug
    '''
    callAlert()
    PL.MENU_debug()
    # MAH...CHE DEVE FARE ???


def DlgMain():
    '''
    DlgMain"
    '''
    callAlert()
    PL.DlgMain()


def attiva_contabilita():
    '''
    attiva_contabilita
    '''
    callAlert()
    PL.attiva_contabilita()


def genera_variante():
    '''
    genera_variante
    '''
    callAlert()
    PL.genera_variante()


def vai_a_M1():
    '''
    vai_a_M1
    '''
    callAlert()
    PL.vai_a_M1()


def vai_a_S1():
    '''
    vai_a_S1
    '''
    callAlert()
    PL.vai_a_S1()


def vai_a_S2():
    '''
    vai_a_S2
    '''
    callAlert()
    PL.vai_a_S2()


def vai_a_variabili():
    '''
    vai_a_variabili
    '''
    callAlert()
    PL.vai_a_variabili()


def vai_a_ElencoPrezzi():
    '''
    vai_a_ElencoPrezzi
    '''
    callAlert()
    PL.vai_a_ElencoPrezzi()


def inizializza_analisi():
    '''
    inizializza_analisi
    '''
    callAlert()
    PL.inizializza_analisi()


def vai_a_Computo():
    '''
    vai_a_Computo
    '''
    callAlert()
    PL.vai_a_Computo()


def vai_a_Scorciatoie():
    '''
    vai_a_Scorciatoie
    '''
    callAlert()
    PL.vai_a_Scorciatoie()


def ssUltimus():
    '''
    ssUltimus
    '''
    callAlert()
    PL.ssUltimus()


def voce_breve():
    '''
    voce_breve
    '''
    callAlert()
    PL.voce_breve()


def tante_analisi_in_ep():
    '''
    tante_analisi_in_ep
    '''
    callAlert()
    PL.tante_analisi_in_ep()


def analisi_in_ElencoPrezzi():
    '''
    analisi_in_ElencoPrezzi
    '''
    callAlert()
    PL.analisi_in_ElencoPrezzi()


def inizializza_elenco():
    '''
    inizializza_elenco
    '''
    callAlert()
    PL.inizializza_elenco()


def riordina_ElencoPrezzi():
    '''
    riordina_ElencoPrezzi
    '''
    callAlert()
    PL.riordina_ElencoPrezzi()


def struttura_Elenco(ctx):
    '''
    struttura_Elenco
    '''
    callAlert()
    PL.struttura_Elenco()


def cancella_voci_non_usate():
    '''
    cancella_voci_non_usate
    '''
    callAlert()
    PL.cancella_voci_non_usate()


def richiesta_offerta():
    '''
    richiesta_offerta
    '''
    callAlert()
    PL.richiesta_offerta()


def trova_np():
    '''
    trova_np
    '''
    callAlert()
    PL.trova_np()


def sproteggi_sheet_TUTTE():
    callAlert()
    PL.sproteggi_sheet_TUTTE()


def rigenera_tutte():
    '''
    rigenera_tutte
    '''
    callAlert()
    PL.rigenera_tutte()


def trova_ricorrenze():
    '''
    trova_ricorrenze
    '''
    callAlert()
    PL.trova_ricorrenze()


def set_larghezza_colonne():
    '''
    set_larghezza_colonne
    '''
    callAlert()
    PL.set_larghezza_colonne()


def config_default():
    '''
    config_default
    '''
    callAlert()
    PL.config_default()


def donazioni():
    '''
    donazioni
    '''
    callAlert()
    PL.donazioni()


def invia_voce():
    '''
    invia_voce
    '''
    callAlert()
    PL.invia_voce()


def rifa_nomearea(oDoc, sSheet, sRange, sName):
    '''
    rifa_nomearea
    '''
    callAlert()
    PL.rifa_nomearea(oDoc, sSheet, sRange, sName)


def autoexec():
    callAlert()
    PL.autoexec()


def autoexec_off():
    callAlert()
    PL.autoexec_off()


def struttura_off():
    callAlert()
    PL.struttura_off()


def setTabColor(color):
    callAlert()
    PL.setTabColor(color)


def adatta_altezza_riga(nSheet=None):
    callAlert()
    PL.adatta_altezza_riga(nSheet)


def paste_clip(arg=None, insCells=0):
    callAlert()
    PL.paste_clip(arg, insCells)


def copy_clip():
    callAlert()
    PL.copy_clip()


def ins_voce_elenco():
    callAlert()
    PL.ins_voce_elenco()


def Filtra_Computo_Cap():
    callAlert()
    PL.Filtra_Computo_Cap()


def Filtra_Computo_SottCap():
    callAlert()
    PL.Filtra_Computo_SottCap()


def Filtra_Computo_A():
    callAlert()
    PL.Filtra_Computo_A()


def Filtra_Computo_B():
    callAlert()
    PL.Filtra_Computo_B()


def Filtra_Computo_C():
    callAlert()
    PL.Filtra_Computo_C()


def EliminaVociDoppieElencoPrezzi():
    callAlert()
    PL.EliminaVociDoppieElencoPrezzi()


def Tutti_Subtotali():
    callAlert()
    PL.Tutti_Subtotali()


def salva_come(nomefile=None):
    callAlert()
    PL.salva_come(nomefile)


def ScriviNomeDocumentoPrincipale():
    '''
    Indica qual è il Documento Principale
    '''
    callAlert()
    PL.ScriviNomeDocumentoPrincipale()


def fissa():
    callAlert()
    PL.fissa()


def bak0():
    '''
    Fa il backup del file di lavoro all'apertura.
    '''
    callAlert()
    PL.bak0()


def numera_voci(bit=1):
    '''
    bit { integer }  : 1 rinumera tutto
                       0 rinumera dalla voce corrente in giù
    '''
    callAlert()
    PL.numera_voci(bit)


def parziale_verifica():
    '''
    Controlla l'esattezza del calcolo del parziale quanto le righe di
    misura vengono aggiunte o cancellate.
    '''
    callAlert()
    PL.parziale_verifica()


def struttura_ComputoM():
    callAlert()
    PL.struttura_ComputoM()


def struttura_Analisi():
    callAlert()
    PL.struttura_Analisi()


def voce_breve_ep():
    callAlert()
    PL.voce_breve_ep()


def inserisci_Riga_rossa():
    callAlert()
    PL.inserisci_Riga_rossa()


def Rinumera_TUTTI_Capitoli2():
    callAlert()
    PL.Rinumera_TUTTI_Capitoli2()


def copia_riga_computo(lrow):
    callAlert()
    PL.copia_riga_computo(lrow)


def ins_voce_computo():
    callAlert()
    PL.ins_voce_computo()


def ins_voce_contab(lrow=0, arg=1):
    '''
    Inserisce una nuova voce in CONTABILITA.
    '''
    callAlert()
    PL.ins_voce_contab(lrow, arg)



