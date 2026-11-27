# -*- coding: utf-8 -*-
########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

'''
Motore unico di impaginazione: applica a uno stile di pagina geometria, intestazione,
piè di pagina e font. Lo usano MENU_sistema_pagine (documento vivo) e l'esportazione PDF
(copia del documento), cosi' schermo e PDF non possono divergere.
'''
import LeenoSettings as LS

# Stili senza intestazione e piè di pagina comuni: la copertina e il CdP, che riporta nel
# corpo i propri dati.
_STILE_COPERTINA = 'Page_Style_COPERTINE'


def contesto_stampa(oDoc, n_sal=None):
    '''
    Dati calcolati una sola volta per tutti gli stili di pagina di `oDoc`.

    n_sal { str | None } : [SAL] degli atti contabili, l'ultimo registrato.
    '''
    printSettings, dummy = LS.loadPrintSettings(oDoc)
    return {
        'printSettings': printSettings,
        'nomefile': LS.nome_file_documento(oDoc),
        'n_sal': n_sal,
        'testate': {},  # documento -> (testata pronta, dimensione, con firme)
    }


def applica_stile_pagina(oDoc, oStile, foglio_attivo, ctx):
    '''
    Applica allo stile di pagina `oStile` geometria e, per gli stili di LeenO che la
    portano, intestazione e piè di pagina.

    oDoc { UNO } : documento a cui appartiene lo stile (vivo o copia)
    oStile { UNO } : stile di pagina
    foglio_attivo { str } : nome del foglio attivo, che sceglie il documento degli stili condivisi
    ctx { dict } : risultato di contesto_stampa
    '''
    nome = oStile.DisplayName
    senza_intestazione = nome in (_STILE_COPERTINA, LS.STILE_PAGINA_CDP)
    oStile.HeaderIsOn = not senza_intestazione
    oStile.FooterIsOn = not senza_intestazione

    # margini, scala e centratura: definizione unica in LeenoSettings
    LS.set_page_margins(oStile)

    if nome not in LS.STILI_TESTATA and nome not in LS.STILI_ATTI_CONTABILI:
        return

    chiave = LS.documento_dello_stile(nome, foglio_attivo)
    testate = ctx['testate']
    if chiave not in testate:
        testi = LS.impostazioni_documento(ctx['printSettings'], chiave)
        valori = LS.valori_segnaposto(oDoc, documento=LS.titolo_a_schermo(chiave, oDoc),
                                      nSal=ctx['n_sal'], nomefile=ctx['nomefile'])
        testate[chiave] = (LS.testata_a_schermo(testi, valori, chiave),
                           LS.dimensione_testata(chiave), LS.ha_firme(testi))
    testata, htxt, con_firme = testate[chiave]

    if nome in LS.STILI_TESTATA:
        LS.set_page_borders(oStile)
    LS.scrivi_testata_stile(oDoc, oStile, testata, htxt, con_firme)
