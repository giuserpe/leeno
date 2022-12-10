import LeenoUtils

if LeenoUtils.getGlobalVar('Lmajor') is None:
    LeenoUtils.initGlobalVars({
        # ~'Lmajor': 3,        # INCOMPATIBILITA'
        # ~'Lminor': 22,       # NUOVE FUNZIONALITA'
        # ~'Lsubv': "0.dev",       # CORREZIONE BUGS
        'Lmajor': 3,        # INCOMPATIBILITA'
        'Lminor': 23,       # NUOVE FUNZIONALITA'
        'Lsubv': "0",       # CORREZIONE BUGS

        'noVoce': ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta',
                   'comp Int_colonna', 'Ultimus_centro_bordi_lati',
                   'comp Int_colonna_R_prima'),

        'stili_cat': ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta', 
                    'comp Int_colonna_R_prima'),

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

