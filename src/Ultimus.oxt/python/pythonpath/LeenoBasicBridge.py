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

print('''
#######################################################################
# IMPORTING PYLEENO FROM LEENBASICBRIDGE
#######################################################################''')

import pyleeno as PL
print('''
#######################################################################
# DONE IMPORTING PYLEENO FROM LEENBASICBRIDGE
#######################################################################''')


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


