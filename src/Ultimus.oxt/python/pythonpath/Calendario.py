########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

"""
Module for generating dynamic dialogs
i.e. dialogs that auto-adjust their layout based on contents and
external constraints

Copyright 2020 by Massimo Del Fedele
"""
import LeenoUtils
import Dialogs


def calendario():
    '''
    Mostra un calendario da cui selezionare la data e la restituisce
    in formato gg/mm/aaaa.
    '''
    testo = Dialogs.pickDate()
    lst = str(testo).split('-')
    try:
        testo = lst[2] + '/' + lst[1] + '/' + lst[0]
    except Exception as _exc:
        LeenoUtils.log_errore('Calendario.calendario: formato data', _exc)

    return testo
