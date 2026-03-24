import os
import sys


def getGlobalVar(name):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    return bDict.get('LEENO_GLOBAL_' + name)


def setGlobalVar(name, value):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    bDict['LEENO_GLOBAL_' + name] = value


def initGlobalVars(d):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    for key, value in d.items():
        bDict['LEENO_GLOBAL_' + key] = value


if getGlobalVar('Lmajor') is None:
    initGlobalVars({
        'Lmajor': 3,        # INCOMPATIBILITA'
        'Lminor': 25,       # NUOVE FUNZIONALITA'
        'Lsubv': "0",       # CORREZIONE BUGS
        # ~'Lmajor': 3,        # INCOMPATIBILITA'
        # ~'Lminor': 23,       # NUOVE FUNZIONALITA'
        # ~'Lsubv': "0",       # CORREZIONE BUGS

        'noVoce': ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta',
                   'comp Int_colonna', 'Ultimus_centro_bordi_lati',
                   'comp Int_colonna_R_prima'),

        'stili_cat': ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta', 
                    'comp Int_colonna_R_prima', 'comp Int_colonna'),

        'stili_computo': ('Comp Start Attributo', 'comp progress', 'comp 10 s',
                        'Comp End Attributo'),

        'stili_contab': ('Comp Start Attributo_R', 'comp 10 s_R', 'uuuuu',
                         'Comp End Attributo_R', 'Comp TOTALI'),
        'stili_analisi': ('Analisi_Sfondo', 'An.1v-Att Start', 'An-1_sigla',
                          'An-lavoraz-desc', 'An-lavoraz-Cod-sx', 'An-lavoraz-desc-CEN',
                          'An-sfondo-basso Att End'),
        'stili_elenco':  ('EP-Cs', 'EP-aS'),

        'codice_da_cercare': '',
        'sUltimus': '',

        'sblocca_computo': 0,

    })

# Variabile 'src_oxt' come riferimento per definizione di 'dest'
src_oxt = '_LeenO'  # Puoi sostituirla con il valore che desideri, o lasciarla come variabile di esempio.

def dest():
    '''Definisce il percorso di destinazione basato sul sistema operativo'''

    # Per sistema Windows
    if os.name == 'nt':
        if not os.path.exists('w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/'):
            try:
                os.makedirs(
                    os.getenv("HOMEPATH") + '\\' + src_oxt +
                    '\\leeno\\src\\Ultimus.oxt\\')
            except FileExistsError:
                pass
            return os.getenv("HOMEDRIVE") + os.getenv(
                "HOMEPATH") + '\\' + src_oxt + '\\leeno\\src\\Ultimus.oxt\\'
        else:
            return 'w:/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt'

    # Per sistemi Linux o macOS
    else:
        dest_path = '/media/giuserpe/PRIVATO/_dwg/ULTIMUSFREE/_SRC/leeno/src/Ultimus.oxt/python/pythonpath'
        if not os.path.exists(dest_path):
            try:
                dest_path = os.getenv(
                    "HOME") + '/' + src_oxt + '/leeno/src/Ultimus.oxt/'
                os.makedirs(dest_path)
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/leeno/bin/')
                os.makedirs(os.getenv("HOME") + '/' + src_oxt + '/_SRC/OXT')
            except FileExistsError:
                pass
        return dest_path
