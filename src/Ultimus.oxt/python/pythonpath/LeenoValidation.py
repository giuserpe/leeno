# -*- Mode: Python; coding: utf-8 -*-
"""Validazioni foglio Calc (UNO Validation) per LeenO."""

import uno

import Dialogs
import LeenoUtils
import SheetUtils
from undo_utils import UndoContext


def debug_validation():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  DLG.mri(oDoc.CurrentSelection.Validation)

    oSheet.getCellRangeByName('L1').String = 'Ricicla da:'
    oSheet.getCellRangeByName('L1').CellStyle = 'Reg_prog'
    oCell = oSheet.getCellRangeByName('N1')
    if oCell.String not in ("COMPUTO", "VARIANTE", 'Scegli origine'):
        oCell.CellStyle = 'Menu_sfondo _input_grasBig'
        valida_cella(oCell,
                     '"COMPUTO";"VARIANTE"',
                     titoloInput='Scegli...',
                     msgInput='COMPUTO o VARIANTE',
                     err=True)
        oCell.String = 'Scegli...'


def valida_cella(oCell, lista_val, titoloInput='', msgInput='', err=False):
    '''
    Imposta un elenco di valori a cascata (Validation LIST) su una cella.

    oCell       {object}  : oggetto cella (es. oSheet.getCellByPosition(0,0))
    lista_val   {string}  : stringa valori separati da punto e virgola: '"A";"B";"C"'
    titoloInput {string}  : titolo del tooltip di aiuto
    msgInput    {string}  : messaggio del tooltip di aiuto
    err         {boolean} : se True, impedisce l'inserimento di valori non in lista
    '''
    # Recuperiamo l'oggetto Validation esistente della cella
    oTabVal = oCell.Validation

    # Configurazione Messaggio di Input (il tooltip che appare al passaggio del mouse)
    oTabVal.ShowInputMessage = True
    oTabVal.InputTitle = titoloInput
    oTabVal.InputMessage = msgInput

    # Configurazione Messaggio di Errore
    oTabVal.ShowErrorMessage = err
    oTabVal.ErrorMessage = "ERRORE: Questo valore non è consentito."
    oTabVal.ErrorAlertStyle = uno.Enum(
        "com.sun.star.sheet.ValidationAlertStyle", "STOP")

    # Definizione del tipo di validazione: LIST
    oTabVal.Type = uno.Enum("com.sun.star.sheet.ValidationType", "LIST")

    # Impostazione della formula (la lista dei valori)
    oTabVal.setFormula1(lista_val)

    # Nota importante: l'oggetto Validation va riassegnato alla cella per rendere effettive le modifiche
    oCell.Validation = oTabVal


def _uno_condition_operator_greater_equal():
    cache = _uno_condition_operator_greater_equal
    if getattr(cache, '_enum', None) is None:
        cache._enum = uno.Enum(
            "com.sun.star.sheet.ConditionOperator", "GREATER_EQUAL")
    return cache._enum


def _uno_validation_type_decimal():
    cache = _uno_validation_type_decimal
    if getattr(cache, '_enum', None) is None:
        cache._enum = uno.Enum(
            "com.sun.star.sheet.ValidationType", "DECIMAL")
    return cache._enum


def _uno_validation_alert_stop():
    cache = _uno_validation_alert_stop
    if getattr(cache, '_enum', None) is None:
        cache._enum = uno.Enum(
            "com.sun.star.sheet.ValidationAlertStyle", "STOP")
    return cache._enum


def _imposta_validazione_decimale_su_intervallo(
        oCellOrRange, skip_se_gia_decimale=False):
    """
    Applica validazione decimale a SheetCell / SheetCellRange (uso interno e macro batch).
    Evita ricreazioni ripetute dell'enum UNO per ConditionOperator.
    """
    VT_DECIMAL = 2
    if skip_se_gia_decimale:
        try:
            if int(oCellOrRange.Validation.Type) == VT_DECIMAL:
                return
        except Exception:
            pass

    op_ge = _uno_condition_operator_greater_equal()
    vt_dec = _uno_validation_type_decimal()
    alert_stop = _uno_validation_alert_stop()

    val = oCellOrRange.Validation
    val.Type = 0
    oCellOrRange.Validation = val

    val = oCellOrRange.Validation
    val.Type = vt_dec
    val.Operator = op_ge
    val.setFormula1("-1E300")
    val.IgnoreBlankCells = True
    val.ShowErrorMessage = True
    val.ErrorAlertStyle = alert_stop
    val.ErrorMessage = (
        "Sono ammessi solo numeri o formule che restituiscono numeri."
    )
    oCellOrRange.Validation = val


def applica_validazione_decimale():
    """
    Applica validazione decimale alle colonne F, G, H, I con stile di cella che inizia con 'comp'.
    Modalità batch: rileva blocchi contigui di celle per velocizzare l'operazione.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.getActiveSheet()
    lastrow = SheetUtils.getLastUsedRow(oSheet)

    first_row = 4
    cols_f_i = range(4, 9)  # colonne F..I (0-based)
    comp_prefix = 'comp'

    indicator = oDoc.CurrentController.getStatusIndicator()
    indicator.start('Applicazione validazione decimale (batch)...', len(cols_f_i))

    rng_sheet = oSheet.getCellRangeByPosition

    def apply_dec_batch(rng):
        _imposta_validazione_decimale_su_intervallo(
            rng, skip_se_gia_decimale=True)

    try:
        with LeenoUtils.DocumentRefreshContext(False):
            with UndoContext('Applicazione validazione decimale'):
                if lastrow >= first_row:
                    n_rows = lastrow - first_row + 1
                    for i, col in enumerate(cols_f_i):
                        indicator.setValue(i + 1)
                        oColRange = rng_sheet(col, first_row, col, lastrow)
                        col_cell = oColRange.getCellByPosition
                        start_row = None
                        for rel_row in range(n_rows):
                            row = first_row + rel_row
                            stile = col_cell(0, rel_row).CellStyle
                            if stile.startswith(comp_prefix):
                                if start_row is None:
                                    start_row = row
                            elif start_row is not None:
                                apply_dec_batch(
                                    rng_sheet(col, start_row, col, row - 1))
                                start_row = None
                        if start_row is not None:
                            apply_dec_batch(
                                rng_sheet(col, start_row, col, lastrow))
    finally:
        indicator.end()

    Dialogs.Info(
        Title='Avviso',
        Text=(
            "Validazione decimale applicata in modalità batch: ammessi numeri (anche 0) e formule; "
            "non ammessi testi. Il formato numerico resta quello di stile della cella."
        ),
    )


def _try_unprotect_leeno_sheet(oSheet):
    """Sblocca il foglio se protetto (password vuota o tipica template LeenO)."""
    if oSheet is None or not oSheet.isProtected():
        return
    for pwd in ("", "password"):
        try:
            oSheet.unprotect(pwd)
            return
        except Exception:
            continue


def _raccolta_celle_validazione_esplicita(doc, oCell):
    """
    Risolve oCell in una lista di SheetCell.
    Su LibreOffice/pyuno alcuni intervalli non espongono supportsService(SheetCellRange)
    ma hanno getRangeAddress: proviamo anche quel percorso.
    """
    celle = []
    if doc is None:
        return celle

    ss = getattr(oCell, "supportsService", None)
    if callable(ss):
        try:
            if oCell.supportsService("com.sun.star.sheet.SheetCell"):
                return [oCell]
        except Exception:
            pass
        try:
            if oCell.supportsService("com.sun.star.sheet.SheetCellRange"):
                addr = oCell.getRangeAddress()
                sheet = doc.Sheets[addr.Sheet]
                for r in range(addr.StartRow, addr.EndRow + 1):
                    for c in range(addr.StartColumn, addr.EndColumn + 1):
                        celle.append(sheet.getCellByPosition(c, r))
                return celle
        except Exception:
            pass

    try:
        addr = oCell.getRangeAddress()
        sheet = doc.Sheets[addr.Sheet]
        for r in range(addr.StartRow, addr.EndRow + 1):
            for c in range(addr.StartColumn, addr.EndColumn + 1):
                celle.append(sheet.getCellByPosition(c, r))
        return celle
    except Exception:
        pass

    try:
        ca = oCell.getCellAddress()
        sheet = doc.Sheets[ca.Sheet]
        return [sheet.getCellByPosition(ca.Column, ca.Row)]
    except Exception:
        return []


def valida_numeri_decimale(oCell=None, *, unprotect_if_needed=False):
    """Validazione decimale: ammessi numeri (incluso 0) e formule numeriche; escluse le stringhe.

    unprotect_if_needed: su foglio protetto Calc può ignorare l'impostazione della validità
    (dialog «Validità» resta «Ogni valore»); prova unprotect con password LeenO tipiche.
    """

    celle = []
    doc = LeenoUtils.getDocument()

    if oCell is not None:
        celle = _raccolta_celle_validazione_esplicita(doc, oCell)
        if not celle:
            return
    else:
        if doc is None or not hasattr(doc, "CurrentSelection"):
            return

        sel = doc.CurrentSelection

        if getattr(sel, "supportsService", None) and callable(sel.supportsService):
            try:
                if sel.supportsService("com.sun.star.sheet.SheetCell"):
                    celle.append(sel)
                elif sel.supportsService("com.sun.star.sheet.SheetCellRange"):
                    addr = sel.getRangeAddress()
                    sheet = doc.Sheets[addr.Sheet]
                    for r in range(addr.StartRow, addr.EndRow + 1):
                        for c in range(addr.StartColumn, addr.EndColumn + 1):
                            celle.append(sheet.getCellByPosition(c, r))
                elif sel.supportsService("com.sun.star.sheet.SheetCellRanges"):
                    for r in sel.getRangeAddresses():
                        sheet = doc.Sheets[r.Sheet]
                        for rr in range(r.StartRow, r.EndRow + 1):
                            for cc in range(r.StartColumn, r.EndColumn + 1):
                                celle.append(sheet.getCellByPosition(cc, rr))
            except Exception:
                celle = []
        if not celle:
            celle = _raccolta_celle_validazione_esplicita(doc, sel)

    if not celle:
        return

    if unprotect_if_needed:
        try:
            addr0 = celle[0].getCellAddress()
            _try_unprotect_leeno_sheet(doc.Sheets[addr0.Sheet])
        except Exception:
            pass

    for cell in celle:
        _imposta_validazione_decimale_su_intervallo(cell)


def valida_numeri_decimale_diverso_da_0():
    """Applica validazione decimale diverso da 0; non forza formato (rispetta lo stile cella)."""
    doc = LeenoUtils.getDocument()
    if not hasattr(doc, "CurrentSelection"):
        return

    sel = doc.CurrentSelection
    celle = []

    # Raccolta celle selezionate
    if sel.supportsService("com.sun.star.sheet.SheetCell"):
        celle.append(sel)
    elif sel.supportsService("com.sun.star.sheet.SheetCellRange"):
        addr = sel.getRangeAddress()
        sheet = doc.Sheets[addr.Sheet]
        for r in range(addr.StartRow, addr.EndRow + 1):
            for c in range(addr.StartColumn, addr.EndColumn + 1):
                celle.append(sheet.getCellByPosition(c, r))
    elif sel.supportsService("com.sun.star.sheet.SheetCellRanges"):
        for r in sel.getRangeAddresses():
            sheet = doc.Sheets[r.Sheet]
            for rr in range(r.StartRow, r.EndRow + 1):
                for cc in range(r.StartColumn, r.EndColumn + 1):
                    celle.append(sheet.getCellByPosition(cc, rr))
    else:
        return

    VT_DECIMAL = 2
    ALERT_STOP = 0

    for cell in celle:
        # 1. Reset completo
        val = cell.Validation
        val.Type = 0
        cell.Validation = val

        # Non impostiamo NumberFormat sulla cella: un formato diretto ha priorita'
        # sullo stile (CellStyle); finche' esiste, le modifiche ai decimali dallo stile non si applicano.

        # 2. Nuova validazione
        val = cell.Validation
        val.Type = VT_DECIMAL
        val.Operator = uno.Enum("com.sun.star.sheet.ConditionOperator", "NOT_EQUAL")
        val.setFormula1("0")
        val.IgnoreBlankCells = True
        val.ShowErrorMessage = True
        val.ErrorAlertStyle = ALERT_STOP
        val.ErrorMessage = "Sono ammessi solo numeri o formule che restituiscono numeri."

        cell.Validation = val

    Dialogs.Info(
        Title='Avviso',
        Text=(
            "Validazione decimale (diverso da 0) applicata senza modificare il formato numerico: "
            "resta quello di stile della cella."
        ),
    )
