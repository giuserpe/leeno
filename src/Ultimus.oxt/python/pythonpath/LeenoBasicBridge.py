'''
    PONTE TEMPORANEO TRA BASIC E PYTHON
    IN QUESTO FILE SONO CONTENUTE TUTTE LE FUNZIONI CHIAMATE DA BASIC
    ACCENTRATE PER POTERLE ELIMINARE PIAN PIANO
'''

# dirty trick to have pythonpath added if missing
import sys, os, inspect
myPath = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
if(myPath not in sys.path):
    sys.path.append(myPath)

import LeenoUtils
import pyleeno as PL

def MENU_debug():
    '''
    MENU_debug
    '''
    PL.MENU_debug()
    # MAH...CHE DEVE FARE ???


def DlgMain():
    '''
    DlgMain"
    '''
    PL.DlgMain()


def attiva_contabilita():
    '''
    attiva_contabilita
    '''
    PL.attiva_contabilita()


def genera_variante():
    '''
    genera_variante
    '''
    PL.genera_variante()


def vai_a_M1():
    '''
    vai_a_M1
    '''
    PL.vai_a_M1()


def vai_a_S1():
    '''
    vai_a_S1
    '''
    PL.vai_a_S1()


def vai_a_S2():
    '''
    vai_a_S2
    '''
    PL.vai_a_S2()


def vai_a_variabili():
    '''
    vai_a_variabili
    '''
    PL.vai_a_variabili()


def vai_a_ElencoPrezzi():
    '''
    vai_a_ElencoPrezzi
    '''
    PL.vai_a_ElencoPrezzi()


def inizializza_analisi():
    '''
    inizializza_analisi
    '''
    PL.inizializza_analisi()


def vai_a_Computo():
    '''
    vai_a_Computo
    '''
    PL.vai_a_Computo()


def vai_a_Scorciatoie():
    '''
    vai_a_Scorciatoie
    '''
    PL.vai_a_Scorciatoie()


def ssUltimus():
    '''
    ssUltimus
    '''
    PL.ssUltimus()


def voce_breve():
    '''
    voce_breve
    '''
    PL.voce_breve()


def tante_analisi_in_ep():
    '''
    tante_analisi_in_ep
    '''
    PL.tante_analisi_in_ep()


def analisi_in_ElencoPrezzi():
    '''
    analisi_in_ElencoPrezzi
    '''
    PL.analisi_in_ElencoPrezzi()


def inizializza_elenco():
    '''
    inizializza_elenco
    '''
    PL.inizializza_elenco()


def riordina_ElencoPrezzi():
    '''
    riordina_ElencoPrezzi
    '''
    PL.riordina_ElencoPrezzi()


def struttura_Elenco(ctx):
    '''
    struttura_Elenco
    '''
    PL.struttura_Elenco()


def cancella_voci_non_usate():
    '''
    cancella_voci_non_usate
    '''
    PL.cancella_voci_non_usate()


def richiesta_offerta():
    '''
    richiesta_offerta
    '''
    PL.richiesta_offerta()


def trova_np():
    '''
    trova_np
    '''
    PL.trova_np()


def sproteggi_sheet_TUTTE():
    PL.sproteggi_sheet_TUTTE()


def rigenera_tutte():
    '''
    rigenera_tutte
    '''
    PL.rigenera_tutte()


def trova_ricorrenze():
    '''
    trova_ricorrenze
    '''
    PL.trova_ricorrenze()


def set_larghezza_colonne():
    '''
    set_larghezza_colonne
    '''
    PL.set_larghezza_colonne()


def config_default():
    '''
    config_default
    '''
    PL.config_default()


def donazioni():
    '''
    donazioni
    '''
    PL.donazioni()


def invia_voce():
    '''
    invia_voce
    '''
    PL.invia_voce()


def rifa_nomearea(oDoc, sSheet, sRange, sName):
    '''
    rifa_nomearea
    '''
    PL.rifa_nomearea(oDoc, sSheet, sRange, sName)


def autoexec():
    PL.autoexec()


def autoexec_off():
    PL.autoexec_off()


def struttura_off():
    PL.struttura_off()


def setTabColor(color):
    PL.setTabColor(color)


def adatta_altezza_riga(nSheet=None):
    PL.adatta_altezza_riga(nSheet)


def paste_clip(arg=None, insCells=0):
    PL.paste_clip(arg, insCells)


def copy_clip():
    PL.copy_clip()


def ins_voce_elenco():
    PL.ins_voce_elenco()


def Filtra_Computo_Cap():
    PL.Filtra_Computo_Cap()


def Filtra_Computo_SottCap():
    PL.Filtra_Computo_SottCap()


def Filtra_Computo_A():
    PL.Filtra_Computo_A()


def Filtra_Computo_B():
    PL.Filtra_Computo_B()


def Filtra_Computo_C():
    PL.Filtra_Computo_C()


def EliminaVociDoppieElencoPrezzi():
    PL.EliminaVociDoppieElencoPrezzi()


def Tutti_Subtotali():
    PL.Tutti_Subtotali()


def salva_come(nomefile=None):
    PL.salva_come(nomefile)


def dp():
    '''
    Indica qual è il Documento Principale
    '''
    PL.dp()


def fissa():
    PL.fissa()


def bak0():
    '''
    Fa il backup del file di lavoro all'apertura.
    '''
    PL.bak0()


def numera_voci(bit=1):
    '''
    bit { integer }  : 1 rinumera tutto
                       0 rinumera dalla voce corrente in giù
    '''
    PL.numera_voci(bit)


def parziale_verifica():
    '''
    Controlla l'esattezza del calcolo del parziale quanto le righe di
    misura vengono aggiunte o cancellate.
    '''
    PL.parziale_verifica()


def struttura_ComputoM():
    PL.struttura_ComputoM()


def struttura_Analisi():
    PL.struttura_Analisi()


def voce_breve_ep():
    PL.voce_breve_ep()


def inserisci_Riga_rossa():
    PL.inserisci_Riga_rossa()


def Rinumera_TUTTI_Capitoli2():
    PL.Rinumera_TUTTI_Capitoli2()


def copia_riga_computo(lrow):
    PL.copia_riga_computo(lrow)


def ins_voce_computo():
    PL.ins_voce_computo()


def ins_voce_contab(lrow=0, arg=1):
    '''
    Inserisce una nuova voce in CONTABILITA.
    '''
    PL.ins_voce_contab(lrow, arg)



