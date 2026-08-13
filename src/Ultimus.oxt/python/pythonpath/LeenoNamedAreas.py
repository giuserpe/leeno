#!/usr/bin/env python3
# -*- Mode: Python; coding: utf-8; indent-tabs-mode: nil; tab-width: 4 -*-
########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################
"""
LeenoNamedAreas.py
------------------
Centralizza la rigenerazione di tutti i NamedArea (intervalli nominati)
necessari al funzionamento di LeenO.

Struttura dei NamedArea gestiti:

  A) FISSI (range esteso all'ultima riga del foglio)
     - elenco_prezzi   -> 'Elenco Prezzi'.$A$3:$AF$<lastrow>
     - Lista           -> 'Elenco Prezzi'.$A$3:$A$<lastrow>
     - elenco_variante -> 'VARIANTE'.$A$3:$AF$<lastrow>  (se foglio esiste)
     - analisi         -> 'Analisi di Prezzo'.$A$3:$K$<lastrow>
     - blocco_analisi  -> 'S5'.$B$108:$P$133  (fisso)
     - oneri_sicurezza -> 'S5'.$B$93:$P$93    (fisso, solo se assente)

  B) FOGLI PRINCIPALI (colonne calcolate su ultima riga usata)
     - AA, BB, cEuro       -> COMPUTO
     - varAA, varBB, varEuro -> VARIANTE
     - GG, G1G1, conEuro   -> CONTABILITA

  C) DINAMICI SAL (non ricreabili da zero senza i confini strutturali
     dei SAL, noti solo ai flussi di LeenoContab).
     Questo modulo si limita a RILEVARE quali esistono e se sono corrotti,
     senza tentare di ricostruirli.

Uso tipico:
    import LeenoNamedAreas as LNA
    LNA.rigenera_tutto()        # rigenera A + B, controlla C
    LNA.rigenera_statici()      # solo gruppo A
    LNA.rigenera_fogli()        # solo gruppo B (equivalente a sistema_aree)
    LNA.verifica_dinamici()     # diagnostica gruppo C, nessuna modifica
"""

import LeenoUtils
import SheetUtils
import LeenoDialogs as DLG


# ---------------------------------------------------------------------------
# Helper interni
# ---------------------------------------------------------------------------

def _get_ep_last_row(oSheet):
    """
    Restituisce l'indice di riga (0-based) della riga sentinella
    'Fine elenco' nel foglio Elenco Prezzi.
    Se non trovata, usa getUsedArea().EndRow.
    """
    row = SheetUtils.uFindStringCol('Fine elenco', 0, oSheet)
    if row is None:
        row = SheetUtils.getUsedArea(oSheet).EndRow
    return row


def _get_analisi_last_row(oSheet):
    """
    Restituisce l'indice di riga (0-based) della riga sentinella
    'Fine ANALISI' nel foglio Analisi di Prezzo.
    Se non trovata, usa getUsedArea().EndRow.
    """
    row = SheetUtils.uFindStringCol('Fine ANALISI', 0, oSheet)
    if row is None:
        row = SheetUtils.getUsedArea(oSheet).EndRow
    return row


# ---------------------------------------------------------------------------
# GRUPPO A -- range fissi (dipendenti dall'ultima riga del foglio)
# ---------------------------------------------------------------------------

def rigenera_elenco_prezzi(oDoc=None):
    """
    Rigenera 'elenco_prezzi' e 'Lista' sul foglio 'Elenco Prezzi'.
    Ritorna True se ha effettuato la rigenerazione, False se il foglio
    non esiste o e' vuoto.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('Elenco Prezzi'):
        return False
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    last_row = _get_ep_last_row(oSheet) + 1   # converte a 1-based
    if last_row < 3:
        return False
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f'$A$3:$AF${last_row}', 'elenco_prezzi')
    SheetUtils.NominaArea(oDoc, 'Elenco Prezzi', f'$A$3:$A${last_row}', 'Lista')
    return True


def rigenera_elenco_variante(oDoc=None):
    """
    Rigenera 'elenco_variante' sul foglio 'VARIANTE'.
    Usa la stessa geometria colonne di 'elenco_prezzi' (A:AF).
    Ritorna True se ha effettuato la rigenerazione, False se il foglio
    non esiste o e' vuoto.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('VARIANTE'):
        return False
    oSheet = oDoc.getSheets().getByName('VARIANTE')
    last_row = SheetUtils.getUsedArea(oSheet).EndRow + 1
    if last_row < 3:
        return False
    SheetUtils.NominaArea(oDoc, 'VARIANTE', f'$A$3:$AF${last_row}', 'elenco_variante')
    return True


def rigenera_analisi(oDoc=None):
    """
    Rigenera 'analisi' sul foglio 'Analisi di Prezzo'.
    Ritorna True se ha effettuato la rigenerazione, False se il foglio
    non esiste.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
        return False
    oSheet = oDoc.getSheets().getByName('Analisi di Prezzo')
    last_row = _get_analisi_last_row(oSheet) + 1   # 1-based
    if last_row < 3:
        return False
    SheetUtils.NominaArea(oDoc, 'Analisi di Prezzo', f'$A$3:$K${last_row}', 'analisi')
    return True


def rigenera_blocco_analisi(oDoc=None):
    """
    Rigenera 'blocco_analisi' (range fisso $B$108:$P$133 su foglio 'S5').
    Usato da LeenoAnalysis.Inserisci_Utili.
    Ritorna True se ha agito, False se il foglio S5 non esiste.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    if not oDoc.getSheets().hasByName('S5'):
        return False
    SheetUtils.NominaArea(oDoc, 'S5', '$B$108:$P$133', 'blocco_analisi')
    return True


def rigenera_oneri_sicurezza(oDoc=None):
    """
    Aggiunge 'oneri_sicurezza' solo se assente (range fisso legacy su 'S5').
    Non sovrascrive se gia' presente.
    Ritorna True se ha creato il range, False se era gia' presente o se
    il foglio S5 non esiste.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    oRanges = oDoc.NamedRanges
    if oRanges.hasByName('oneri_sicurezza'):
        return False   # gia' presente: non toccare
    if not oDoc.getSheets().hasByName('S5'):
        return False
    oSheet = oDoc.getSheets().getByName('S5')
    oCellAddress = oSheet.getCellRangeByName('B10').getCellAddress()
    oRanges.addNewByName('oneri_sicurezza', "$S5.$B$93:$P$93", oCellAddress, 0)
    return True


def rigenera_statici(oDoc=None):
    """
    Rigenera tutti i NamedArea del GRUPPO A.
    Ritorna un dizionario {nome_descrittivo: True|False} con l'esito.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    return {
        'elenco_prezzi + Lista':  rigenera_elenco_prezzi(oDoc),
        'elenco_variante':        rigenera_elenco_variante(oDoc),
        'analisi':                rigenera_analisi(oDoc),
        'blocco_analisi':         rigenera_blocco_analisi(oDoc),
        'oneri_sicurezza':        rigenera_oneri_sicurezza(oDoc),
    }


# ---------------------------------------------------------------------------
# GRUPPO B -- colonne sui fogli COMPUTO / VARIANTE / CONTABILITA
#   Identico alla funzione sistema_aree() di pyleeno, mantenuto in sync:
#   se modifichi la mappa la', aggiorna anche qui (o rifattorizza in un
#   import condiviso).
# ---------------------------------------------------------------------------

# Mappa: nome_foglio -> [(colonna_lettera, nome_area), ...]
_FOGLI_AREE = {
    'COMPUTO': [
        ('AJ', 'AA'),
        ('J',  'BB'),
        ('AK', 'cEuro'),
    ],
    'VARIANTE': [
        ('AJ', 'varAA'),
        ('J',  'varBB'),
        ('AK', 'varEuro'),
    ],
    'CONTABILITA': [
        ('AJ', 'GG'),
        ('J',  'G1G1'),
        ('AK', 'conEuro'),
    ],
}


def rigenera_fogli(oDoc=None):
    """
    Rigenera le colonne nominate nei fogli COMPUTO, VARIANTE e CONTABILITA.
    Equivalente a sistema_aree() in pyleeno.py (mantenuto sincronizzato).
    Ritorna un dizionario {(foglio, nome_area): True|'saltato'|'vuoto'}.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()

    esiti = {}
    with LeenoUtils.DocumentRefreshContext(False):
        for nome_foglio, aree in _FOGLI_AREE.items():
            if not oDoc.getSheets().hasByName(nome_foglio):
                for _, nome_area in aree:
                    esiti[(nome_foglio, nome_area)] = 'saltato'
                continue

            oSheet = oDoc.getSheets().getByName(nome_foglio)
            row = SheetUtils.getUsedArea(oSheet).EndRow

            if row < 3:
                for _, nome_area in aree:
                    esiti[(nome_foglio, nome_area)] = 'vuoto'
                continue

            for col, nome_area in aree:
                range_str = f'${col}$3:${col}${row}'
                SheetUtils.NominaArea(oDoc, nome_foglio, range_str, nome_area)
                esiti[(nome_foglio, nome_area)] = True

    return esiti


# ---------------------------------------------------------------------------
# GRUPPO C -- diagnostica range dinamici SAL (sola lettura)
# ---------------------------------------------------------------------------

_PREFISSI_SAL = ('_Lib_', '_Reg_', '_SAL_', '_CdP_')


def verifica_dinamici(oDoc=None):
    """
    Scansiona tutti i NamedArea con prefissi SAL (_Lib_, _Reg_, _SAL_, _CdP_)
    e verifica che siano integri (ReferredCells.RangeAddress accessibile).

    Ritorna una lista di dizionari:
        [{'nome': str, 'valido': bool, 'errore': str|None}, ...]

    Non crea, non modifica, non elimina nulla.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()

    risultati = []
    oNamedRanges = oDoc.NamedRanges
    for name in oNamedRanges.getElementNames():
        if not any(name.startswith(p) for p in _PREFISSI_SAL):
            continue
        record = {'nome': name, 'valido': False, 'errore': None}
        try:
            oNamedRanges.getByName(name).ReferredCells.RangeAddress
            record['valido'] = True
        except Exception as e:
            record['errore'] = str(e)
        risultati.append(record)

    return risultati


# ---------------------------------------------------------------------------
# Punto di ingresso principale
# ---------------------------------------------------------------------------

def rigenera_tutto(oDoc=None, mostra_log=False):
    """
    Rigenera tutti i NamedArea rigenerabili (GRUPPO A + GRUPPO B).
    Il GRUPPO C (SAL) non viene modificato: viene solo verificato e
    incluso nel log se mostra_log=True.

    Parametri:
        oDoc      : documento Calc; se None usa getDocument()
        mostra_log: se True, mostra un riepilogo via DLG.chi()

    Ritorna un dizionario con le chiavi 'statici', 'fogli', 'dinamici'.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()

    esito_statici = rigenera_statici(oDoc)
    esito_fogli   = rigenera_fogli(oDoc)
    esito_din     = verifica_dinamici(oDoc) if mostra_log else []

    if mostra_log:
        righe = ['=== LeenoNamedAreas.rigenera_tutto ===', '']
        righe.append('-- GRUPPO A (range fissi) --')
        for k, v in esito_statici.items():
            stato = 'OK' if v else 'FOGLIO ASSENTE o vuoto'
            righe.append(f'  {k}: {stato}')

        righe.append('')
        righe.append('-- GRUPPO B (colonne fogli principali) --')
        for (foglio, area), v in esito_fogli.items():
            righe.append(f'  {foglio}.{area}: {v}')

        righe.append('')
        righe.append('-- GRUPPO C (SAL – solo diagnostica) --')
        if not esito_din:
            righe.append('  (nessun range SAL trovato)')
        for r in esito_din:
            stato = 'VALIDO' if r['valido'] else f'CORROTTO: {r["errore"]}'
            righe.append(f'  {r["nome"]}: {stato}')

        DLG.chi('\n'.join(righe))

    return {
        'statici':  esito_statici,
        'fogli':    esito_fogli,
        'dinamici': esito_din,
    }


# ---------------------------------------------------------------------------
# Voce di menu esposta a LibreOffice
# ---------------------------------------------------------------------------

def MENU_rigenera_NamedAreas():
    """
    Punto di ingresso chiamabile da macro/toolbar.
    Rigenera tutti i NamedArea e mostra un breve report di diagnostica.
    """
    rigenera_tutto(mostra_log=True)
