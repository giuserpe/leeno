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


def contesto_stampa(oDoc, n_sal=None, extra=None):
    '''
    Dati calcolati una sola volta per tutti gli stili di pagina da applicare.

    oDoc { UNO } : documento di LeenO da cui leggere dati (S2, impostazioni, colonne visibili)
    n_sal { str | None } : [SAL] degli atti contabili
    extra { dict | None } : altri segnaposto {'[NOME]': valore}, es. [NUMERO_DOCUMENTO] del PDF
    '''
    printSettings, dummy = LS.loadPrintSettings(oDoc)
    return {
        'printSettings': printSettings,
        'nomefile': LS.nome_file_documento(oDoc),
        'n_sal': n_sal,
        'extra': extra,
        'testate': {},  # documento -> (testata pronta, dimensione, con firme)
    }


def applica_stile_pagina(oDoc, oStile, foglio_attivo, ctx, oDocStili=None):
    '''
    Applica allo stile di pagina `oStile` geometria e, per gli stili di LeenO che la
    portano, intestazione e piè di pagina.

    oDoc { UNO } : documento di LeenO, sorgente dei dati
    oStile { UNO } : stile di pagina
    oDocStili { UNO | None } : documento a cui appartiene lo stile, se e' una copia di oDoc
        (esportazione PDF): i campi numerici di pagina vanno creati li'. Se None e' oDoc
    foglio_attivo { str } : nome del foglio attivo, che sceglie il documento degli stili condivisi
    ctx { dict } : risultato di contesto_stampa
    '''
    if oDocStili is None:
        oDocStili = oDoc
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
                                      nSal=ctx['n_sal'], nomefile=ctx['nomefile'],
                                      extra=ctx['extra'])
        testate[chiave] = (LS.testata_a_schermo(testi, valori, chiave),
                           LS.dimensione_testata(chiave), LS.ha_firme(testi))
    testata, htxt, con_firme = testate[chiave]

    if nome in LS.STILI_TESTATA:
        LS.set_page_borders(oStile)
    LS.scrivi_testata_stile(oDocStili, oStile, testata, htxt, con_firme)
