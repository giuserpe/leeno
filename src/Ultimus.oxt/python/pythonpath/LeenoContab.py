#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# LeenoContab.py
'''
LeenoContab.py - Contabilità per Leeno
'''

from datetime import date
from com.sun.star.table import CellRangeAddress
from com.sun.star.sheet.GeneralFunction import MAX
from com.sun.star.sheet.CellFlags import \
    VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

import LeenoUtils
import LeenoGlobals
import SheetUtils
import LeenoSheetUtils
import LeenoSettings as LS
import LeenoComputo
import Dialogs
import LeenoDialogs as DLG
import pyleeno as PL
import LeenoEvents
import LeenoBasicBridge
import uno
import os
import itertools
import DocUtils
import LeenoConfig
import LeenoImport_XPWE
cfg = LeenoConfig.Config()

from collections import defaultdict
from functools import wraps


class ProgressIndicatorManager:
    """
    Context manager e decorator per gestire il ciclo di vita dell'indicatore di progresso
    e prevenire l'hijacking da parte di operazioni interne di Calc.
    """
    def __init__(self, oDoc, total_steps=4):
        self.oDoc = oDoc
        self.indicator = oDoc.getCurrentController().getStatusIndicator()
        self.total_steps = total_steps
        self.current_step = 0
        self.current_message = "In corso..."

    def start(self, message="In corso...", steps=None):
        """Avvia l'indicatore con un messaggio"""
        if steps is not None:
            self.total_steps = steps
        self.current_message = message
        self.indicator.start(message, self.total_steps)

    def update(self, step, message=None):
        """Aggiorna step e messaggio, reclamando l'indicatore"""
        self.current_step = step
        if message:
            self.current_message = message
            self.indicator.start(self.current_message, self.total_steps)
        self.indicator.setValue(step)

    def reclaim(self):
        """Reclama l'indicatore dopo operazioni che potrebbero averlo hijackato"""
        self.indicator.start(self.current_message, self.total_steps)
        self.indicator.setValue(self.current_step)

    def end(self):
        """Termina l'indicatore"""
        self.indicator.end()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.end()
        return False


def with_progress_reclaim(manager_attr='progress'):
    """
    Decorator che reclama automaticamente l'indicatore dopo l'esecuzione della funzione.
    Utile per funzioni che potrebbero triggerare operazioni interne di Calc.

    Args:
        manager_attr: nome dell'attributo che contiene il ProgressIndicatorManager
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            result = func(*args, **kwargs)
            # Cerca il manager negli argomenti o nel primo argomento (self/cls)
            if args and hasattr(args[0], manager_attr):
                manager = getattr(args[0], manager_attr)
                manager.reclaim()
            return result
        return wrapper
    return decorator


def sbloccaContabilita(oSheet, lrow):
    '''
    Controlla che non ci siano atti contabili registrati e dà il consenso a procedere.
    Ritorna True se il consenso è stato dato, False altrimenti
    '''
    if LeenoGlobals.getGlobalVar('sblocca_computo') == 1:
        return True
    if oSheet.Name != 'CONTABILITA':
        return True

    partenza = LeenoSheetUtils.cercaPartenza(oSheet, lrow)
    if partenza[2] == '#reg':
        res = Dialogs.YesNoCancelDialog(IconType="question",
           Title="Voce già registrata",
           Text= "Lavorando in questo punto del foglio,\n"
                 "comprometterai la validità degli atti contabili già emessi.\n\n"
                 "Vuoi procedere?\n\n"
                 "SCEGLIENDO SÌ DOVRAI NECESSARIAMENTE RIGENERARLI!"
        )
        if res == 1:
            LeenoGlobals.setGlobalVar('sblocca_computo', 1)
            return True
        return False
    return True


# ###############################################################


def insertVoceContabilita(oSheet, lrow):
    '''
    Inserisce una nuova voce in CONTABILITA.
    '''
    # controllo che non ci siano atti registrati
    # se ci sono, chiede conferma per poter operare
    if not sbloccaContabilita(oSheet, lrow):
        return False

    stili_contab = LeenoGlobals.getGlobalVar('stili_contab')
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    nSal = 0
    if stile == 'comp Int_colonna_R_prima':
        lrow += 1
    elif stile == 'Ultimus_centro_bordi_lati':
        i = lrow
        while i != 0:
            if oSheet.getCellByPosition(23, i).Value != 0:
                nSal = int(oSheet.getCellByPosition(23, i).Value)
                break
            i -= 1
        while oSheet.getCellByPosition(0, lrow).CellStyle == stile:
            lrow += 1
        if oSheet.getCellByPosition(0, lrow).CellStyle == 'uuuuu':
            lrow += 1

    elif stile == 'Comp TOTALI':
        pass
    elif stile in stili_contab:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        nSal = int(oSheet.getCellByPosition(23, sStRange.RangeAddress.StartRow + 1).Value)
        lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow)
    else:
        return

    oDoc = SheetUtils.getDocumentFromSheet(oSheet)
    oSheetto = oDoc.getSheets().getByName('S5')
    oRangeAddress = oSheetto.getCellRangeByPosition(0, 22, 48, 26).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
    # inserisco le righe
    oSheet.getRows().insertByIndex(lrow, 5)
    oSheet.copyRange(oCellAddress, oRangeAddress)
    oSheet.getCellRangeByPosition(0, lrow, 48, lrow + 5).Rows.OptimalHeight = True

    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    sopra = sStRange.RangeAddress.StartRow
    for n in reversed(range(0, sopra)):
        if oSheet.getCellByPosition(1, n).CellStyle == 'Ultimus_centro_bordi_lati':
            break
        if oSheet.getCellByPosition(1, n).CellStyle == 'Data_bianca':
            data = oSheet.getCellByPosition(1, n).Value
            break
    try:
        oSheet.getCellByPosition(1, sopra + 2).Value = data
    except Exception:
        oSheet.getCellByPosition(1, sopra + 2).Value = date.today().toordinal() - 693594

    # raggruppo i righi di misura
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = CellRangeAddress()
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow + 2
    oCellRangeAddr.EndRow = lrow + 2
    oSheet.group(oCellRangeAddr, 1)
    ###################################

    if oDoc.NamedRanges.hasByName('_Lib_' + str(nSal)):
        if lrow - 1 == oSheet.getCellRangeByName('_Lib_' + str(nSal)).getRangeAddress().EndRow:
            nSal += 1

    oSheet.getCellByPosition(23, sopra + 1).Value = nSal
    oSheet.getCellByPosition(23, sopra + 1).CellStyle = 'Sal'

    oSheet.getCellByPosition(35, sopra + 4).Formula = '=B' + str(sopra + 2)
    oSheet.getCellByPosition(36, sopra +4).Formula = (
       '=IF(ISERROR(P' + str(sopra + 5) + ');"";IF(P' +
       str(sopra + 5) + '<>"";P' + str(sopra + 5) + ';""))')
    oSheet.getCellByPosition(36, sopra + 4).CellStyle = "comp -controolo"

    LeenoSheetUtils.numeraVoci(oSheet, 0, True)

    '''
        @@@@ NOTA BENE : QUESTA PARTE È PER L'USO INTERATTIVO
        VEDIAMO CHE FARNE IN SEGUITO
    if cfg.read('Generale', 'pesca_auto') == '1':
        if arg == 0:
            return
        pesca_cod()
    '''

# ###############################################################

import Calendario

def imposta_data():
    """ Imposta la data scelta nelle misure selezionate."""
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    import datetime
    testo = Calendario.calendario()

    try:
        oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
    except AttributeError:
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text='''La selezione deve essere contigua.''')
        return 0

    dv_start = LeenoComputo.DatiVoce(oSheet, oRangeAddress.StartRow)
    prima_riga = dv_start.SR

    dv_end = LeenoComputo.DatiVoce(oSheet, oRangeAddress.EndRow)
    ultima_riga = dv_end.ER

    for el in range(prima_riga, ultima_riga + 1):
        cell = oSheet.getCellByPosition(1, el)  # colonna B
        if cell.CellStyle == 'Data_bianca':
            try:
                cell.String = testo
            except Exception as e:
                return
    return


# ###############################################################
def ultimo_sal():
    '''
    restituisce il numero di sal registrati
    '''
    oDoc = LeenoUtils.getDocument()
    oRanges = oDoc.NamedRanges
    lista = []
    [lista.append(str(i))                           #select
    for i in range(1, 100)                          #from
    if oRanges.hasByName("_Lib_" + str(i)) == True] #where
    return lista


def mostra_sal(uSal):
    '''
    Mostra solo gli atti relativi al SAL scelto.

    Parametri:
    uSal { integer } : numero del SAL da mostrare
    '''
    oDoc = LeenoUtils.getDocument()

    d = [
        ('CONTABILITA', '_Lib_', 11),
        ('Registro', '_Reg_', 9),
        ('SAL', '_SAL_', 5)
    ]

    listaSal = ultimo_sal()

    if uSal:
        SheetUtils.visualizza_PageBreak()
        for sal in range(1, len(listaSal) + 1):
            for el in d:
                # ~ nomearea = el[1] + str(sal)
                try:
                    nomearea = el[1] + str(sal)
                    # ~ DLG.chi(el[0])
                    oSheet = oDoc.Sheets.getByName(el[0])
                    oRanges = oDoc.NamedRanges
                    oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress

                    # Definisci i limiti dell'intervallo
                    daRiga = oNamedRange.StartRow
                    aRiga = oNamedRange.EndRow
                    daColonna = oNamedRange.StartColumn
                    aColonna = oNamedRange.EndColumn

                    oNamedRange.EndColumn = el[2]

                    oSheet.ungroup(oNamedRange, 1)
                    oSheet.group(oNamedRange, 1)

                    if sal == uSal:
                        oSheet.setPrintAreas((oNamedRange,))
                        oSheet.setPrintTitleRows(True)
                        PL.GotoSheet(oSheet.Name)
                        oSheet.getCellRangeByPosition(daColonna, daRiga, aColonna, aRiga).Rows.IsVisible = True
                        oDoc.CurrentController.setFirstVisibleRow(1)
                        PL._gotoCella(0, daRiga - 1)
                    else:
                        oSheet.getCellRangeByPosition(daColonna, daRiga, aColonna, aRiga).Rows.IsVisible = False
                except Exception as e:
                    # ~ DLG.errore(e)
                    continue

                    # ~ DLG.chi(f"Errore nell'accesso all'area nominata {nomearea}: {e}")
                    # ~ pass

    return

@with_progress_reclaim(manager_attr='progress')
def MENU_AnnullaAttiContabili():
    '''
    Annulla gli atti dell'ultimo SAL registrato (Libretto, Registro, SAL, CdP).
    '''
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oRanges = oDoc.NamedRanges

    listaSal = ultimo_sal()

    if len (listaSal) == 0:
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text="Nessun SAL registrato da eliminare.")
        return
    messaggio = 'Stai per eliminare gli atti del SAL n.' + \
    listaSal[-1] + '\n\nVuoi procedere?'
    if Dialogs.YesNoDialog(IconType="warning",Title='*** A T T E N Z I O N E ! ***',
        Text= messaggio) == 1:
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start("Annullamento atti in corso...", 5)

        # 0. Elimina CdP (NamedRange + svuota celle compilate)
        nome_cdp = '_CdP_' + listaSal[-1]
        try:
            if oRanges.hasByName(nome_cdp):
                oRanges.removeByName(nome_cdp)
            if oDoc.Sheets.hasByName('CdP'):
                oCdP = oDoc.Sheets.getByName('CdP')
                # Svuota solo le celle con formula o valore scritte da GeneraCdP
                # identificate tramite le stesse ancoraggi usati in GeneraCdP
                anchors_cdp = [
                    'Per lavori e somministrazioni',
                    'SOMMANO importi soggetti',
                    'SOMMANO importi NON soggetti',
                    'Ritenuta per infortuni',
                    'Ammontare dei Certificati',
                    'TOTALE DETRAZIONE',
                    'RISULTA IL CREDITO',
                    'I.V.A.',
                    'TOTALE GENERALE',
                ]
                for label in anchors_cdp:
                    try:
                        result = SheetUtils.uFindString(label, oCdP)
                        if result:
                            r_a = result[1]
                            for cc in range(8):
                                cell = oCdP.getCellByPosition(cc, r_a)
                                if cell.getFormula().startswith('=') or \
                                   (cell.Type.value != 'EMPTY' and
                                    cell.CellStyle not in ('comp Int_colonna_R', 'Ultimus_centro_bordi_lati')):
                                    cell.clearContents(VALUE + STRING + FORMULA)
                    except Exception:
                        pass
                # Svuota blocco certificati precedenti (N°/Data/Importo)
                try:
                    r_ncert_result = SheetUtils.uFindString('N°', oCdP)
                    if r_ncert_result:
                        r_f = r_ncert_result[1] + 1
                        r_sogg_res = SheetUtils.uFindString('SOMMANO importi soggetti', oCdP)
                        r_s = r_sogg_res[1] if r_sogg_res else r_f + 10
                        for rr in range(r_f, r_s):
                            for cc in range(6):
                                oCdP.getCellByPosition(cc, rr).clearContents(
                                    VALUE + STRING + FORMULA)
                except Exception:
                    pass
                # Ripristina etichetta IVA
                try:
                    r_iva_res = SheetUtils.uFindString('I.V.A.', oCdP)
                    if r_iva_res:
                        r_iv = r_iva_res[1]
                        for cc in range(6):
                            lbl = oCdP.getCellByPosition(cc, r_iv).String
                            if '%' in lbl and 'I.V.A.' in lbl:
                                oCdP.getCellByPosition(cc, r_iv).String = \
                                    'per I.V.A. al __%'
                                break
                except Exception:
                    pass
        except Exception:
            pass
        indicator.setValue(1)

        #elimina libretto
        oSheet = oDoc.Sheets.getByName('CONTABILITA')
        nome_area = "_Lib_" + listaSal[-1]
        oNamedRange = oRanges.getByName(nome_area).ReferredCells.RangeAddress
        oSheet.ungroup(oNamedRange, 1)
        #range del _Lib_
        daRiga = oNamedRange.StartRow
        aRiga = oNamedRange.EndRow
        # ripulisce le colonne da VALUE+STRING+FORMULA
        flags = VALUE+STRING+FORMULA
        oSheet.getCellRangeByPosition(19, daRiga, 25, aRiga).clearContents(
        flags)
        # annulla lo sbiancamento celle
        flags = HARDATTR
        oSheet.getCellRangeByPosition(0, 2, 25, aRiga).clearContents(
        flags)
        # cancella firme
        firma = PL.seleziona_voce(aRiga)
        oSheet.Rows.removeByIndex(firma[0] , firma[1] - firma[0] + 1)
        indicator.setValue(1)

        # --- CANCELLA TITOLI E FILLER (in ordine inverso per non sballare gli indici) ---
        for i in reversed(range(daRiga, aRiga + 1)):
            oCell = oSheet.getCellByPosition(2, i)
            style = oSheet.getCellByPosition(0, i).CellStyle
            content = oCell.String
            if style == "Ultimus_centro_bordi_lati" and (
                content in ("SICUREZZA (CALCOLO ANALITICO)", "LAVORI A MISURA") or
                content.startswith("–––")
            ):
                oSheet.Rows.removeByIndex(i, 1)

        # cancella riga gialla
        oSheet.Rows.removeByIndex(daRiga - 1, 1)
        oDoc.NamedRanges.removeByName(nome_area)
        # cancella area di stampa
        LeenoSheetUtils.DelPrintSheetArea()
        # importo prossimo sal
        oSheet.getCellRangeByName('Z2').Formula = (
        "=$P$2-SUBTOTAL(9;$P$2:$P$" + str(daRiga - 1) + ")"
        )
        indicator.setValue(2)

        try:
            [oDoc.Sheets.removeByName(el)   #select
            for el in ('Registro', 'SAL')   #from
            if len (listaSal) == 1]         #where
        except Exception as e:
            # ~ DLG.errore(e)
            pass

        if len(listaSal) > 1:
        #elimina registro
            # ~PL.GotoSheet('Registro')
            oSheet = oDoc.Sheets.getByName('Registro')
            nome_area = "_Reg_" + listaSal[-1]
            if len (listaSal) == 1:
                oDoc.Sheets.removeByName('Registro')
            else:
                oNamedRange = oRanges.getByName(nome_area).ReferredCells.RangeAddress
                oSheet.ungroup(oNamedRange, 1)
                #range del _Reg_
                daRiga = oNamedRange.StartRow -1
                aRiga = oNamedRange.EndRow
                #cancella registro
                oSheet.Rows.removeByIndex(daRiga, aRiga - daRiga + 1)
                #cancella area di stampa
                LeenoSheetUtils.DelPrintSheetArea()
            oDoc.NamedRanges.removeByName(nome_area)

        #elimina SAL
            oSheet = oDoc.Sheets.getByName('SAL')
            nome_area = "_SAL_" + listaSal[-1]
            if len (listaSal) == 1:
                oDoc.Sheets.removeByName('SAL')
            else:
                oNamedRange = oRanges.getByName(nome_area).ReferredCells.RangeAddress
                oSheet.ungroup(oNamedRange, 1)
                #range del _Reg_
                daRiga = oNamedRange.StartRow -1
                aRiga = oNamedRange.EndRow
                #cancella registro
                oSheet.Rows.removeByIndex(daRiga, aRiga - daRiga + 1)
                #cancella area di stampa
                LeenoSheetUtils.DelPrintSheetArea()
            oDoc.NamedRanges.removeByName(nome_area)
        # --- Pulizia SITUAZIONE CONTABILE in S2 ---
        try:
            oS2 = oDoc.getSheets().getByName('S2')
            markerS2 = SheetUtils.uFindString("SITUAZIONE CONTABILE", oS2)
            yS2, xS2 = markerS2[0], markerS2[1]
            nSalDel = int(listaSal[-1])
            col_del = yS2 + nSalDel
            # Cancella tutta la colonna del SAL (righe da +1 a +25)
            oS2.getCellRangeByPosition(col_del, xS2 + 1, col_del, xS2 + 25).clearContents(
                VALUE + DATETIME + STRING + FORMULA)
        except:
            pass

        indicator.setValue(3)
        indicator.end()
    # ~LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    oSheet = oDoc.CurrentController.ActiveSheet
    try:
        nSal = ultimo_sal()[-1]
        oSheet.getCellRangeByName('Z3').String = nSal
    except:
        oSheet.getCellRangeByName('Z3').String = ''
    oSheet.Rows.OptimalHeight = True

    if len (listaSal) == 1:
        SheetUtils.visualizza_PageBreak(False)

    try:
        nSal = int(listaSal[-1]) - 1
        mostra_sal(nSal)
        # Se c'è un SAL precedente, rigenera il CdP per quel SAL
        if nSal > 0:
            GeneraCdP(oDoc, nSal=nSal)
    except Exception as e:
        # ~ DLG.errore(e)
        pass
    PL.GotoSheet('CONTABILITA')


# ###############################################################


def Menu_svuotaContabilita():
    oDoc = LeenoUtils.getDocument()
    messaggio= """
Questa operazione svuoterà il foglio CONTABILITA e cancellerà
tutti gli elaborati contabili generati fino a questo momento.

OPERAZIONE NON REVERSIBILE!

VUOI PROCEDERE UGUALMENTE?"""
    if Dialogs.YesNoDialog(IconType="warning",Title='*** A T T E N Z I O N E ! ***',
        Text= messaggio) == 1:
        svuotaContabilita(oDoc)


def svuotaContabilita(oDoc):
    '''
    svuota_contabilita
    Ricrea il foglio di contabilità partendo da zero.
    '''
    with LeenoUtils.DocumentRefreshContext(False):
        for n in range(1, 100):
            if oDoc.NamedRanges.hasByName('_Lib_' + str(n)):
                oDoc.NamedRanges.removeByName('_Lib_' + str(n))
                oDoc.NamedRanges.removeByName('_SAL_' + str(n))
                oDoc.NamedRanges.removeByName('_Reg_' + str(n))
        for el in ('Registro', 'SAL', 'CONTABILITA'):
            if oDoc.Sheets.hasByName(el):
                oDoc.Sheets.removeByName(el)

        # Trova l'indice per CONTABILITA (a destra di COMPUTO e VARIANTE)
        try:
            # COMPUTO è sempre presente
            nIdx = oDoc.Sheets.getByName('COMPUTO').RangeAddress.Sheet + 1
            if oDoc.Sheets.hasByName('VARIANTE'):
                idx_v = oDoc.Sheets.getByName('VARIANTE').RangeAddress.Sheet
                if idx_v >= nIdx:
                    nIdx = idx_v + 1
        except:
            nIdx = 3

        oDoc.Sheets.insertNewByName('CONTABILITA', nIdx)
        PL.GotoSheet('CONTABILITA')
        oSheet = oDoc.Sheets.getByName('CONTABILITA')

        SheetUtils.setTabColor(oSheet, 16757935)
        oSheet.getCellRangeByName('C1').Formula = '=RIGHT(CELL("FILENAME"; A1); LEN(CELL("FILENAME"; A1)) - FIND("$"; CELL("FILENAME"; A1)))'
        oSheet.getCellRangeByName('C1').CellStyle = 'comp Int_colonna'
        oSheet.getCellRangeByName('C1').CellBackColor = 16757935
        oSheet.getCellRangeByName('A3').String = 'N.'
        oSheet.getCellRangeByName('B3').String = 'Articolo\nData'
        oSheet.getCellRangeByName('C3').String = 'LAVORAZIONI\nO PROVVISTE'
        oSheet.getCellRangeByName('F3').String = 'P.U.\nCoeff.'
        oSheet.getCellRangeByName('G3').String = 'Lung.'
        oSheet.getCellRangeByName('H3').String = 'Larg.'
        oSheet.getCellRangeByName('I3').String = 'Alt.\nPeso'
        oSheet.getCellRangeByName('J3').String = 'Quantità\nPositive'
        oSheet.getCellRangeByName('L3').String = 'Quantità\nNegative'
        oSheet.getCellRangeByName('N3').String = 'Prezzo\nunitario'
        oSheet.getCellRangeByName('P3').String = 'Importi'
        oSheet.getCellRangeByName('Q3').String = 'Incidenza\nsul totale'
        oSheet.getCellRangeByName('R3').String = 'Sicurezza\ninclusa'
        oSheet.getCellRangeByName('S3').String = 'senza errori'
        oSheet.getCellRangeByName('T3').String = 'Lib.\nN.'
        oSheet.getCellRangeByName('U3').String = 'Lib.\nP.'
        oSheet.getCellRangeByName('W3').String = 'flag'
        oSheet.getCellRangeByName('X3').String = 'SAL\nN.'
        oSheet.getCellRangeByName('Z3').String = 'Importi\nSAL parziali'
        oSheet.getCellRangeByName('AB3').String = 'Sicurezza\nunitaria'
        oSheet.getCellRangeByName('AC3').String = 'Materiali\ne Noli €'
        oSheet.getCellRangeByName('AD3').String = 'Incidenza\nMdO %'
        oSheet.getCellRangeByName('AE3').String = 'Importo\nMdO'
        oSheet.getCellRangeByName('AF3').String = 'Super Cat'
        oSheet.getCellRangeByName('AG3').String = 'Cat'
        oSheet.getCellRangeByName('AH3').String = 'Sub Cat'
        #  oSheet.getCellByPosition(34,2).String = 'tag B'sub Scrivi_header_moduli
        oSheet.getCellByPosition(35,2).String = 'tag C'
        oSheet.getCellRangeByName('AK3').String = 'senza errori'
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
        #  colore colonne riga di intestazione
        oSheet.getCellRangeByPosition(0, 2, 36, 2).CellStyle = 'comp Int_colonna_R'
        oSheet.getCellByPosition(0, 2).CellStyle = 'comp Int_colonna_R_prima'
        oSheet.getCellByPosition(18, 2).CellStyle = 'COnt_noP'
        oSheet.getCellRangeByPosition(0, 0, 0, 3).Rows.OptimalHeight = True
        #  riga di controllo importo
        oSheet.getCellRangeByPosition(0, 1, 36, 1).CellStyle = 'comp In testa'
        oSheet.getCellRangeByName('C2').String = 'QUESTA RIGA NON VIENE STAMPATA'
        oSheet.getCellRangeByPosition(0, 1, 1, 1).merge(True)
        oSheet.getCellRangeByName('N2').String = 'TOTALE:'
        oSheet.getCellRangeByName('U2').String = 'SAL SUCCESSIVO:'

        oSheet.getCellRangeByName('Z2').Formula = '=$P$2-SUBTOTAL(9;$P$2:$P$2)'

        oSheet.getCellRangeByName('P2').Formula = '=SUBTOTAL(9;P:P)'  # importo lavori registrati
        oSheet.getCellByPosition(0, 1).Formula = '=AK2'  # importo lavori
        oSheet.getCellByPosition(
            17, 1).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza

        oSheet.getCellByPosition(
            28, 1).Formula = '=SUBTOTAL(9;AC3:AC4)'  # importo materiali
        oSheet.getCellByPosition(29,
                                1).Formula = '=AE2/Z2/100'  # Incidenza manodopera %
        oSheet.getCellByPosition(29, 1).CellStyle = 'Comp TOTALI %'
        oSheet.getCellByPosition(
            30, 1).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
        oSheet.getCellByPosition(
            36, 1).Formula = '=SUBTOTAL(9;AK3:AK4)'  # importo certo

        # riga del totale
        oSheet.getCellByPosition(2, 3).String = 'T O T A L E'
        oSheet.getCellByPosition(15,
                                3).Formula = '=SUBTOTAL(9;P:P)'  # importo lavori registrati
        oSheet.getCellByPosition(
            17, 3).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza
        oSheet.getCellByPosition(
            30, 3).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
        oSheet.getCellRangeByPosition(0, 3, 36, 3).CellStyle = 'Comp TOTALI'
        # riga rossa
        oSheet.getCellByPosition(0, 4).String = 'Fine Computo'
        oSheet.getCellRangeByPosition(0, 4, 36, 4).CellStyle = 'Riga_rossa_Chiudi'
        PL._gotoCella(2, 2)
        LeenoSheetUtils.setLarghezzaColonne(oSheet)

        return oSheet


# ###############################################################


def generaContabilita(oDoc):
    '''
    Mostra il foglio di contabilità, se presente
    Altrimenti lo genera
    '''
    if oDoc.Sheets.hasByName('S1'):
        oDoc.Sheets.getByName('S1').getCellByPosition(7, 327).Value = 1
        if oDoc.Sheets.hasByName('CONTABILITA'):
            oSheet = oDoc.Sheets.getByName('CONTABILITA')
        else:
            #oSheet = oDoc.Sheets.insertNewByName('CONTABILITA', 5)
            oSheet = svuotaContabilita(oDoc)
            insertVoceContabilita(oSheet, 0)

            LeenoEvents.assegna()
            LeenoSheetUtils.ScriviNomeDocumentoPrincipaleInFoglio(oSheet)

    return oSheet

########################################################################
# CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA #


def attiva_contabilita():
    '''
    Se presenti, attiva e visualizza le tabelle di contabilità
    @@@ MODIFICA IN CORSO CON 'LeenoContab.generaContabilita'
    '''
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    if oDoc.Sheets.hasByName('S1'):
        oDoc.Sheets.getByName('S1').getCellByPosition(7, 327).Value = 1
        if oDoc.Sheets.hasByName('CONTABILITA'):
            for el in ('Registro', 'SAL', 'CONTABILITA'):
                if oDoc.Sheets.hasByName(el):
                    PL.GotoSheet(el)
        else:
            oDoc.Sheets.insertNewByName('CONTABILITA', 5)
            svuotaContabilita(oDoc)
            PL.GotoSheet('CONTABILITA')
            PL._gotoCella(0, 2)
        PL.GotoSheet('CONTABILITA')
    LeenoBasicBridge.ScriviNomeDocumentoPrincipale()
    LeenoEvents.assegna()
########################################################################


def partita(testo):
    '''
    Aggiunge/detrae rigo di PARTITA PROVVISORIA
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != "CONTABILITA":
        return
    x = PL.LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0, x).CellStyle == 'comp 10 s_R':
        if oSheet.getCellByPosition(2, x).Type.value != 'EMPTY':
            PL.Copia_riga_Ent()
            x += 1
        oSheet.getCellByPosition(2, x).String = testo
        oSheet.getCellRangeByPosition(2, x, 8, x).CellBackColor = 16777113
        PL._gotoCella(5, x)


def MENU_partita_aggiungi():
    '''
    @@ DA DOCUMENTARE
    '''
    partita('PARTITA PROVVISORIA')


def MENU_partita_detrai():
    '''
    @@ DA DOCUMENTARE
    '''
    partita('SI DETRAE PARTITA PROVVISORIA')


########################################################################
def struttura_CONTAB():
    '''
    Visualizza in modalità struttura i documenti contabili
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    PL.struttura_off()
    is_ctrl, is_shift = PL.GetModifiers()

    force_color_level = -1
    if is_ctrl and is_shift:
        force_color_level = 2 # Sotto Categoria
    elif is_ctrl:
        force_color_level = 0 # Super Categoria
    elif is_shift:
        force_color_level = 1 # Categoria

    if force_color_level != -1:
        for n in range(0, 4):
            PL.applica_livelli(n, vedi=False, force_color_level=force_color_level)

    oRanges = oDoc.NamedRanges

    if oSheet.Name == 'CONTABILITA':
        pref = "_Lib_"
        y = 3
        if not oDoc.NamedRanges.hasByName("_Lib_1"):
            PL.struttura_ComputoM()
    elif oSheet.Name == 'Registro':
        pref = "_Reg_"
        y = 1
    elif oSheet.Name == 'SAL':
        pref = "_SAL_"
        y = 1
    for i in range(1, 50):
        try:
            oRange=oRanges.getByName(pref + str(i)).ReferredCells.RangeAddress
            # ~oSheet.ungroup(oRange, 1)
            oSheet.group(oRange, 1)
            oSheet.getCellRangeByPosition(0, oRange.StartRow,
                11, oRange.EndRow).Rows.IsVisible = False
        except:
            try:
                oSheet.getCellRangeByPosition(0, oRange.StartRow,
                    11, oRange.EndRow).Rows.IsVisible = True
                PL._gotoCella(0, oRange.StartRow -1)
                oDoc.CurrentController.setFirstVisibleRow(y)
            except:
                # ~Dialogs.NotifyDialog(Image='Icons-Big/info.png',
                        # ~Title = 'Info',
                        # ~Text='''In questo Libretto delle Misure
# ~non ci sono misure registrate.''')
                # ~ if oSheet.Name == 'CONTABILITA':
                    # ~ PL.struttura_ComputoM()
                pass
            return

def aggiorna_S2_libretto(oDoc, nSal, aVoce, nPag):
    '''
    Aggiorna specificamente i dati del Libretto nel foglio Situazione Contabile.
    Sincronizza: Numero SAL, Data, Ultima Voce e Ultima Pagina.
    '''
    try:
        oS2 = oDoc.getSheets().getByName('S2')

        # 1. Trovo la colonna corretta (nSal)
        # Assumendo che il titolo "SITUAZIONE CONTABILE" sia in colonna A (0)
        # Il SAL 1 sarà in colonna B (1), il SAL 2 in colonna C (2), ecc.
        col_sal = nSal

        # 2. Aggiorno l'intestazione del SAL (Righe fisse in alto come da immagine)
        # Riferimenti basati sull'immagine: Riga 2 (SAL n.), Riga 3 (A tutto il)
        oS2.getCellByPosition(col_sal, 1).Value = nSal
        # Conversione data corretta per LibreOffice
        oS2.getCellByPosition(col_sal, 2).Value = date.today().toordinal() - 693594

        # 3. Aggiorno i riferimenti a fondo pagina tramite ricerca etichette
        # Questo rende il codice immune all'inserimento di nuove righe nel foglio S2
        mappa_celle = {
            "Ultima voce registrata n.": aVoce,
            "Ultima pagina libretto n.": nPag
        }

        for etichetta, valore in mappa_celle.items():
            # Cerchiamo l'etichetta nella colonna A (0)
            pos = SheetUtils.uFindStringCol(etichetta, 0, oS2)
            if pos:
                riga = int(pos)
                oS2.getCellByPosition(col_sal, riga).Value = valore

    except Exception as e:
        # Usiamo il gestore errori centralizzato di LeenoDispatcher
        handle_exception(e)

# --- All'interno di GeneraLibretto, sostituisci il vecchio blocco con: ---
# aggiorna_S2_libretto(oDoc, nSal, aVoce, nPag)





def GeneraLibretto(oDoc):
    '''
    CONTABILITA' - Genera il Libretto delle Misure.
    Include gestione analitica VDS, firme, riempimento pagina e marcatura.
    '''
    PL.chiudi_dialoghi()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'CONTABILITA':
        return

    PL.numera_voci()
    oRanges = oDoc.NamedRanges

    # 1. IDENTIFICAZIONE NUMERO NUOVO SAL
    nSal = 1
    for i in reversed(range(1, 50)):
        if oRanges.hasByName("_Lib_" + str(i)):
            nSal = i + 1
            break

    # --- DETERMINAZIONE PAGINA DI PARTENZA ---
    oS2 = oDoc.getSheets().getByName('S2')
    markerS2 = SheetUtils.uFindString("SITUAZIONE CONTABILE", oS2)
    yS2, xS2 = markerS2[0], markerS2[1]

    if nSal == 1:
        start_nPage = 0
    else:
        last_sal_page = oS2.getCellByPosition(yS2 + (nSal - 1), xS2 + 25).Value
        start_nPage = int(last_sal_page) + 1

    # 2. SUGGERIMENTO INTERVALLO VOCI (daVoce / aVoce)
    daVoceSuggerita = 1
    libretti = SheetUtils.sStrColtoList('segue Libretto delle Misure n.', 2, oSheet, start=2)
    try:
        daVoceSuggerita = int(oSheet.getCellByPosition(2, libretti[-1]).String.split('÷')[1]) + 1
    except:
        daVoceSuggerita = 1

    daVoce = PL.InputBox(str(daVoceSuggerita), f"SAL n.{nSal}: Libretto, da voce n.")
    if not daVoce: return

    try:
        lrow_start = int(SheetUtils.uFindStringCol(daVoce, 0, oSheet))
    except: return

    sStRange_start = LeenoComputo.circoscriveVoceComputo(oSheet, lrow_start)
    primariga = sStRange_start.RangeAddress.StartRow

    for _ in range(1, 10):
        if primariga > 0 and oSheet.getCellByPosition(0, primariga - 1).CellStyle in ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'):
            primariga -= 1

    last_row_contab = LeenoSheetUtils.cercaUltimaVoce(oSheet)
    aVoceMassima = 0
    for el in reversed(range(3, last_row_contab + 1)):
        s_val = oSheet.getCellByPosition(0, el).String.strip()
        if s_val.isdigit():
            aVoceMassima = int(s_val)
            break

    aVoce = PL.InputBox(str(aVoceMassima), f"SAL n.{nSal}: Libretto, a voce n.")
    if not aVoce or int(aVoce) < int(daVoce): return

    try:
        lrow_end = int(SheetUtils.uFindStringCol(aVoce, 0, oSheet))
    except: return
    ultimariga = LeenoComputo.circoscriveVoceComputo(oSheet, lrow_end).RangeAddress.EndRow


    # --- RACCOLTA DATI CUMULATIVA PER IL SAL ---
    SAL = []
    SAL_VDS = []

    # Cerchiamo la riga della prima voce assoluta (1)
    try:
        if oRanges.hasByName("_Lib_1"):
            r_start_abs = oRanges.getByName("_Lib_1").ReferredCells.RangeAddress.StartRow
        else:
            r_start_abs = int(SheetUtils.uFindStringCol("1", 0, oSheet, equal=1))
    except:
        r_start_abs = primariga

    c_i = r_start_abs
    voci_coll = set()
    while c_i <= ultimariga:
        resp = LeenoComputo.datiVoceComputo(oSheet, c_i)
        if resp:
            d_v = resp[1]
            num_v = str(d_v[0]).strip()
            if num_v not in voci_coll:
                if 'VDS_' in str(d_v[1]):
                    SAL_VDS.append(d_v)
                else:
                    SAL.append(d_v)
                voci_coll.add(num_v)
        c_i = LeenoSheetUtils.prossimaVoce(oSheet, c_i, saltaCat=True)
    # -------------------------------------------

    # Eseguiamo l'adattamento delle altezze prima del calcolo dei filler
    # per avere coordinate Y corrette
    PL.comando('CalculateHard')
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    PL.comando('CalculateHard')

    last_hard_break_y = 0.0
    if nSal > 1:
        # Per i SAL successivi al primo, il punto zero è l'inizio della sezione
        last_hard_break_y = oSheet.getCellByPosition(0, primariga).Position.Y

    voci_elaborate = set() # Per evitare duplicazioni durante gli shift di riga
    curr_i = primariga
    current_section_type = None

    while curr_i <= ultimariga:
        # In questo loop gestiamo solo la formattazione del Libretto (Titoli e Filler)
        res = LeenoComputo.datiVoceComputo(oSheet, curr_i)
        if res is None:
            curr_i += 1
            continue

        num_voce = str(res[1][0]).strip()
        if num_voce in voci_elaborate:
            curr_i = LeenoSheetUtils.prossimaVoce(oSheet, curr_i, saltaCat=True)
            continue
        voci_elaborate.add(num_voce)

        datiVoce = res[1]
        is_vds = 'VDS_' in str(datiVoce[1])
        voce_type = 'VDS' if is_vds else 'LAVORI'

        # Inserimento titoli di sezione (logica Registro)
        if voce_type != current_section_type:
            # Riempimento pagina e salto (solo se non è il primo titolo del blocco)
            if current_section_type is not None:
                # 1. Calcolo filler per finire la pagina
                PL.comando('CalculateHard')
                h_pagina_std = 25510 # Altezza Libretto
                oCellPrev = oSheet.getCellByPosition(2, curr_i - 1)
                y_pos = oCellPrev.Position.Y
                h_prev = oSheet.getRows().getByIndex(curr_i - 1).Height

                # Calcola spazio partendo dalla FINE della riga precedente
                current_y = y_pos + h_prev

                # Calcolo relatvo all'ultimo Hard Break
                dist_from_break = current_y - last_hard_break_y
                ingombro_pag = dist_from_break % h_pagina_std

                spazio_da_coprire = h_pagina_std - ingombro_pag - 1500 # margine sicuro (1.5 cm)

                if spazio_da_coprire > 1000 and (ingombro_pag > 3000):
                    num_righe_filler = int(spazio_da_coprire // 500)

                    # Inserimento batch filler
                    oSheet.getRows().insertByIndex(curr_i, num_righe_filler)
                    oFRange = oSheet.getCellRangeByPosition(0, curr_i, 11, curr_i + num_righe_filler - 1)
                    oFRange.CellStyle = "Ultimus_centro_bordi_lati"
                    oFRange.Rows.Height = 500

                    filler_text = "––––––––––––––––––––––––––––––"
                    oFTextRange = oSheet.getCellRangeByPosition(2, curr_i, 2, curr_i + num_righe_filler - 1)
                    oFTextRange.setDataArray(tuple((filler_text,) for _ in range(num_righe_filler)))

                    curr_i += num_righe_filler
                    ultimariga += num_righe_filler

                oSheet.getRows().insertByIndex(curr_i, 1)
                oSheet.getRows().getByIndex(curr_i).IsStartOfNewPage = True

                # Aggiorna il punto zero per la nuova pagina (Hard Break)
                last_hard_break_y = oSheet.getCellByPosition(0, curr_i).Position.Y
            else:
                oSheet.getRows().insertByIndex(curr_i, 1)
                if nSal > 1:
                    oSheet.getRows().getByIndex(curr_i).IsStartOfNewPage = True
                    last_hard_break_y = oSheet.getCellByPosition(0, curr_i).Position.Y
            oSheet.getCellRangeByPosition(0, curr_i, 11, curr_i).CellStyle = "Ultimus_centro_bordi_lati"
            titolo = "SICUREZZA (CALCOLO ANALITICO)" if is_vds else "LAVORI A MISURA"
            oSheet.getCellByPosition(2, curr_i).String = titolo

            current_section_type = voce_type
            curr_i += 1
            ultimariga += 1
            # Ricarichiamo i dati dopo lo shift
            res_after = LeenoComputo.datiVoceComputo(oSheet, curr_i)
            if res_after:
                datiVoce = res_after[1]

        curr_i = LeenoSheetUtils.prossimaVoce(oSheet, curr_i, saltaCat=True)

    try:
        # Calcolo somme totali
        # sic analitico = somma degli importi (indice 6) delle voci in SAL_VDS
        tot_sic = sum([float(el[6]) for el in SAL_VDS if el[6]])
        tot_mdo = sum([float(el[8]) for el in SAL if el[8]]) + sum([float(el[8]) for el in SAL_VDS if el[8]])

        # Raggruppamento per datiSAL (Lavori) e datiSAL_VDS (Sicurezza) con N. ord. sequenziale e raggruppamento per Articolo
        def raggruppa_voci(dati_lista):
            # Raggruppa per solo Articolo per garantire l'univocità
            gruppi_quant = defaultdict(float)
            gruppi_importo = defaultdict(float)
            gruppo_dati = {} # Mappa articolo -> [desc, um, prezzo, first_appearance_index]

            for i, row in enumerate(dati_lista):
                art = str(row[1]).strip()
                if not art or art == "LAVORI": # Filtra voci senza articolo o placeholder
                    continue

                q = float(row[4]) if row[4] else 0.0
                imp = float(row[6]) if row[6] else 0.0

                # Salta se sia quantità che importo sono zero (voce non eseguita/fantasma)
                if q == 0.0 and imp == 0.0:
                    continue

                gruppi_quant[art] += q
                gruppi_importo[art] += imp
                if art not in gruppo_dati:
                    gruppo_dati[art] = [row[2], row[3], float(row[5]), i]

            # Restituiamo una lista ordinata per Codice Articolo
            articoli_ordinati = sorted(gruppo_dati.keys())

            res_list = []
            for art in articoli_ordinati:
                desc, um, prezzo, _ = gruppo_dati[art]
                res_list.append([art, desc, um, gruppi_quant[art], prezzo, gruppi_importo[art]])

            return res_list

        raggruppati_lavori = raggruppa_voci(SAL)
        raggruppati_vds = raggruppa_voci(SAL_VDS)

        # Creazione dati finali con numerazione sequenziale da 1
        n_ord_global = 1

        datiSAL = []
        for row in raggruppati_lavori:
            art, desc, um, quant, prezzo, importo = row
            datiSAL.append([f"{n_ord_global}\n{art}", desc, um, quant, prezzo, importo])
            n_ord_global += 1

        datiSAL_VDS = []
        for row in raggruppati_vds:
            art, desc, um, quant, prezzo, importo = row
            datiSAL_VDS.append([f"{n_ord_global}\n{art}", desc, um, quant, prezzo, importo])
            n_ord_global += 1

        PL.comando('DeletePrintArea')
        SheetUtils.visualizza_PageBreak()

        # Annotazione SAL e Totale
        oSheet.getCellByPosition(25, ultimariga - 1).String = f"SAL n.{nSal}"
        oSheet.getCellByPosition(25, ultimariga).Formula = f"=SUBTOTAL(9;P{primariga+1}:P{ultimariga+1})"
        oSheet.getCellByPosition(25, ultimariga).CellStyle = "comp sotto Euro 3_R"

        # 5. GESTIONE FIRME
        inizioFirme = ultimariga + 1
        fineFirme = firme_libretto(inizioFirme)

        # 6. CREAZIONE AREA NOMINALE
        nomearea = f"_Lib_{nSal}"
        area_str = f"$A${primariga + 1}:$AJ${fineFirme + 1}"
        SheetUtils.NominaArea(oDoc, "CONTABILITA", area_str, nomearea)

        # 7. RIEMPIMENTO PAGINA
        insrow()

        # Recupero parametri post-filler
        oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
        daRiga = oNamedRange.StartRow
        aRiga = oNamedRange.EndRow

        # Stili firme + filler
        oSheet.getCellRangeByPosition(0, inizioFirme, 32, aRiga).CellStyle = "Ultimus_centro_bordi_lati"
        oSheet.getCellByPosition(2, inizioFirme + 1).CellStyle = "Ultimus_destra"

        # 8. IMPOSTAZIONE PAGINA (Omissis intestazioni standard)
        # ... [Qui rimangono le tue impostazioni LS.setPageStyle, header, footer] ...

        oPrintRange = oNamedRange
        oPrintRange.EndColumn = 11
        oSheet.setPrintAreas((oPrintRange,))

        # --- BUGFIX: Imposta intestazione di stampa alla riga 2 (indice 2, terza riga) ---
        oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oTitles.Sheet = oSheet.RangeAddress.Sheet
        oTitles.StartRow = 2
        oTitles.EndRow = 2
        oSheet.setTitleRows(oTitles)
        oSheet.setPrintTitleRows(True)

    except Exception as e:
        DLG.errore(e)
        return

    oSheet.getCellRangeByPosition(0, daRiga, 11, aRiga).CellBackColor = -1
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    SheetUtils.visualizza_PageBreak()

    # --- 9. RIGA DI RINVIO E MARCATURA (Metodo a due passaggi) ---
    # Inseriamo prima la riga di rinvio per far sì che Calc ricalcoli i salti con essa presente
    oSheet.getRows().insertByIndex(daRiga, 1)
    aRiga += 1 # Slittamento indici dovuto all'inserimento della riga di rinvio
    primariga += 1
    ultimariga += 1

    # Forza Calc a ricalcolare i salti pagina rinfrescando la vista
    SheetUtils.visualizza_PageBreak(False)
    SheetUtils.visualizza_PageBreak(True)

    # --- 10. ANNOTAZIONE E MARCATURA IN BATCH ---
    # Estrai i salti di pagina orizzontali (righe) da Calc
    breaks = sorted([b.Position for b in oSheet.getRowPageBreaks()])

    # Preparazione dei dati per l'inserimento batch (Colonne 19, 20, 21, 22, 23)
    # 19: nSal, 20: nPag, 21: -, 22: #reg, 23: nSal

    num_rows = aRiga - daRiga + 1
    # Recuperiamo gli stili della colonna 1 in un colpo solo per il filtraggio
    oStyleRange = oSheet.getCellRangeByPosition(1, daRiga, 1, aRiga)
    # Nota: CellStyle non si recupera in batch facilmente con DataArray,
    # ma possiamo minimizzare le chiamate.

    nPagFinale = start_nPage

    # Preparazione array per le colonne 19-23
    anno_data = [] # Conterrà liste di 5 elementi per ogni riga

    for i in range(daRiga, aRiga + 1):
        num_breaks_between = len([b for b in breaks if daRiga < b <= i])
        nPagCorrente = start_nPage + num_breaks_between

        row_data = [None] * 5 # Fallback: non sovrascrivere se non necessario

        if i == daRiga:
            # Configura riga di rinvio
            oSheet.getCellRangeByPosition(0, daRiga, 36, daRiga).CellStyle = "uuuuu"
            oSheet.getCellByPosition(2, daRiga).String = f"segue Libretto delle Misure n.{nSal} - {daVoce}÷{aVoce}"

            # Subtotale nella riga di rinvio (corretto per lo slittamento riga)
            formula_sum = f"=SUBTOTAL(9;$P${primariga + 1}:$P${ultimariga + 1})"
            for c in (15, 25):
                cell = oSheet.getCellByPosition(c, daRiga)
                cell.Formula, cell.CellStyle = formula_sum, "comp sotto Euro 3_R"

            row_data = [nSal, nPagCorrente, None, None, nSal]
        else:
            # Verifica se la riga è una riga di 'voce'
            if oSheet.getCellByPosition(1, i).CellStyle == "comp Art-EP_R":
                row_data = [nSal, nPagCorrente, None, "#reg", nSal]
                nPagFinale = nPagCorrente

        anno_data.append(tuple(row_data))

    # Scrittura batch delle annotazioni (Colonne 19-23)
    oAnnoRange = oSheet.getCellRangeByPosition(19, daRiga, 23, aRiga)
    # Filtriamo i None per non cancellare celle esistenti che non vogliamo toccare?
    # In realtà in queste colonne (T-X) solitamente non c'è altro nel libretto generato.
    # Per sicurezza, potremmo recuperare il DataArray esistente e aggiornarlo.
    current_data = list(oAnnoRange.getDataArray())
    for idx, new_row in enumerate(anno_data):
        updated_row = list(current_data[idx])
        for c_idx, val in enumerate(new_row):
            if val is not None:
                updated_row[c_idx] = val
        current_data[idx] = tuple(updated_row)

    oAnnoRange.setDataArray(tuple(current_data))

    # Scrive l'ultimo numero di pagina annotato nella riga gialla di riepilogo
    oSheet.getCellByPosition(20, daRiga).Value = nPagFinale

    # --- 11. AGGIORNAMENTO S2 ---
    oS2.getCellByPosition(yS2 + nSal, xS2 + 1).Value = nSal
    oS2.getCellByPosition(yS2 + nSal, xS2 + 2).Value = date.today().toordinal() - 693594
    oS2.getCellByPosition(yS2 + nSal, xS2 + 24).Value = int(aVoce)
    oS2.getCellByPosition(yS2 + nSal, xS2 + 25).Value = nPagFinale

    PL._gotoCella(0, daRiga)

    # RESTITUZIONE 9 PARAMETRI
    return nSal, daVoce, aVoce, primariga + 1, aRiga + 1, datiSAL, tot_sic, tot_mdo, datiSAL_VDS



#######################################################################



def scrivi_intestazioni_fisse(oSheet, nome_foglio):
    ''' Scrive i titoli delle colonne e imposta le larghezze in base al tipo di foglio '''

    # Configurazione per il Registro
    if nome_foglio == "Registro":
        cols_config = [
            ("N. ord.\nArticolo\nData", 1600),
            ("LAVORAZIONI\nE SOMMINISTRAZIONI", 7500),
            ("Lib.\nN.", 650),
            ("Lib.\nP.", 650),
            ("U.M.", 1000),
            ("Quantità\nPositive", 1600),
            ("Quantità\nNegative", 1600),
            ("Prezzo\nunitario", 1400),
            ("Importo\ndebito", 1950),
            ("Importo\npagamento", 1950)
        ]

    # Configurazione per il SAL
    elif nome_foglio == "SAL":
        cols_config = [
            ("N. ord.\nArticolo", 1600),
            ("LAVORAZIONI\nE SOMMINISTRAZIONI", 11050),
            ("U.M.", 1500),
            ("Quantità", 1800),
            ("Prezzo\nunitario", 1400),
            ("Importo", 1900)
        ]
    else:
        return # Altri fogli non gestiti qui

    # Applicazione Intestazioni (Riga 0)
    oRangeHead = oSheet.getCellRangeByPosition(0, 0, len(cols_config) - 1, 0)
    oRangeHead.CellStyle = "comp Int_colonna_R"

    for i, (titolo, width) in enumerate(cols_config):
        oCell = oSheet.getCellByPosition(i, 0)
        oCell.String = titolo
        oCell.Columns.Width = width

    # Ottimizzazione altezza riga intestazione
    oSheet.getRows().getByIndex(0).OptimalHeight = True






def setup_foglio(oDoc, nome_foglio, dopo_di="CONTABILITA"):
    ''' Crea il foglio o lo sposta se necessario dopo 'dopo_di' '''
    sheets = oDoc.getSheets()

    # Calcolo posizione di destinazione
    try:
        anchor_idx = sheets.getByName(dopo_di).RangeAddress.Sheet
        target_pos = anchor_idx + 1
    except:
        target_pos = sheets.getCount() # mette in fondo se non trova l'ancora

    if not sheets.hasByName(nome_foglio):
        sheets.insertNewByName(nome_foglio, target_pos)
        sheet = sheets.getByName(nome_foglio)
        scrivi_intestazioni_fisse(sheet, nome_foglio)
    else:
        # Se esiste, lo spostiamo nella posizione corretta
        sheets.moveByName(nome_foglio, target_pos)
        sheet = sheets.getByName(nome_foglio)
    return sheet




#######################################################################


def GeneraRegistro(oDoc, dati):
    '''
    REGISTRO - Genera il Registro di Contabilità mantenendo l'ordine esatto di CONTABILITA.
    Inserisce titoli di sezione quando cambia il tipo di voce (LAVORI ↔ VDS).
    '''
    # 0. Spacchettamento dei parametri
    nSal_corrente, daVoce, aVoce, p_riga, u_riga, _, tot_sic, _, datiSAL_VDS = dati

    oRegSheet = setup_foglio(oDoc, "Registro")
    oSheetContab = oDoc.Sheets.getByName("CONTABILITA")

    start_i = p_riga - 1
    end_i = u_riga - 1

    # 1. Recupero posizione di inserimento
    if nSal_corrente == 1:
        insRow = 1
    else:
        try:
            oPrevRange = oDoc.NamedRanges.getByName(f"_Reg_{nSal_corrente-1}").ReferredCells.RangeAddress
            insRow = oPrevRange.EndRow + 1
        except:
            insRow = SheetUtils.getLastUsedRow(oRegSheet) + 1

    # 2. Raccolta dati MANTENENDO L'ORDINE e marcando il tipo
    REG_DATA_ORDERED = []  # Lista di tuple: (dati_riga, is_vds)
    visti = set()

    for r in range(start_i, end_i + 1):
        res = LeenoComputo.datiVoceComputo(oSheetContab, r)
        if res is not None:
            dati_riga = res[0]  # REG tuple
            dati_sal = res[1]   # SAL tuple
            if str(dati_riga[1]).strip() == "" or str(dati_riga[4]).strip() == "":
                continue

            # Il codice articolo è in SAL[1]
            codice_articolo = str(dati_sal[1]).strip()
            riga_tuple = tuple(dati_riga)

            if riga_tuple not in visti:
                is_vds = codice_articolo.startswith("VDS_")
                REG_DATA_ORDERED.append((dati_riga, is_vds))
                visti.add(riga_tuple)

    if not REG_DATA_ORDERED:
        return True

    # 3. INTESTAZIONE GENERALE (solo una volta all'inizio)
    oRegSheet.getRows().insertByIndex(insRow, 2)
    oRegSheet.getCellRangeByPosition(0, insRow, 9, insRow).CellStyle = "uuuuu"
    oRegSheet.getCellByPosition(1, insRow).String = f"segue Registro n.{nSal_corrente} - {daVoce}÷{aVoce}"

    oRegSheet.getCellByPosition(1, insRow + 1).String = "R I P O R T O"
    oRegSheet.getCellByPosition(8, insRow + 1).Formula = f'=IF(SUBTOTAL(9;$I$2:$I${insRow+1})=0;"";SUBTOTAL(9;$I$2:$I${insRow+1}))'
    oRegSheet.getCellRangeByPosition(0, insRow + 1, 9, insRow + 1).CellStyle = "Ultimus_Bordo_sotto"

    current_row = insRow + 2
    prima_riga_dati = current_row

    # 4. INSERIMENTO VOCI CON SEZIONI DINAMICHE IN BATCH
    current_section_type = None
    current_section_start = None

    # Raggruppa le voci per tipo (LAVORI o VDS) per inserirle in blocchi
    def get_type(entry):
        return 'VDS' if entry[1] else 'LAVORI'

    for voce_type, group in itertools.groupby(REG_DATA_ORDERED, key=get_type):
        block = [entry[0] for entry in group]
        num_voci = len(block)
        is_vds = (voce_type == 'VDS')

        # A. Gestione cambio sezione
        if voce_type != current_section_type:
            # Chiudi sezione precedente con parziale
            if current_section_type is not None:
                section_end_row = current_row - 1
                oRegSheet.getRows().insertByIndex(current_row, 2)
                oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row + 1).CellStyle = "Ultimus_centro_bordi_lati"
                current_row += 1

                testo_parziale = "Parziale della Sicurezza €" if current_section_type == 'VDS' else "Parziale dei Lavori a Misura €"
                oRegSheet.getCellByPosition(1, current_row).String = testo_parziale
                oRegSheet.getCellByPosition(1, current_row).CellStyle = "Ultimus_destra"
                oRegSheet.getCellByPosition(8, current_row).Formula = f"=SUBTOTAL(9;I{current_section_start+1}:I{section_end_row+1})"
                oRegSheet.getCellByPosition(8, current_row).CellStyle = "Ultimus_destra_totali"
                current_row += 1

                # Filler tra sezioni
                num_filler = _riempi_pagina(oRegSheet, current_row, col=1, last_col=9, h_pagina=25810, margine=2000)
                current_row += num_filler + 1

            # Titolo nuova sezione
            oRegSheet.getRows().insertByIndex(current_row, 1)
            if current_section_type is not None:
                oRegSheet.getRows().getByIndex(current_row).IsStartOfNewPage = True
            titolo = "SICUREZZA (CALCOLO ANALITICO)" if is_vds else "LAVORI A MISURA"
            oRegSheet.getCellByPosition(1, current_row).String = titolo
            oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row).CellStyle = "Ultimus_centro_bordi_lati"
            current_row += 1

            current_section_type = voce_type
            current_section_start = current_row

        # B. Inserimento batch del blocco voci
        oRegSheet.getRows().insertByIndex(current_row, num_voci)
        oRange = oRegSheet.getCellRangeByPosition(0, current_row, 8, current_row + num_voci - 1)
        oRange.setDataArray(tuple(tuple(d) for d in block))

        # Styling batch
        oRegSheet.getCellRangeByPosition(0, current_row, 1, current_row + num_voci - 1).CellStyle = "List-stringa-sin"
        oRegSheet.getCellRangeByPosition(2, current_row, 4, current_row + num_voci - 1).CellStyle = "List-num-centro"
        oRegSheet.getCellRangeByPosition(5, current_row, 9, current_row + num_voci - 1).CellStyle = "List-num-euro"

        current_row += num_voci

    # Chiudi l'ultima sezione con parziale
    if current_section_type is not None:
        section_end_row = current_row - 1

        # Riga vuota prima del parziale
        oRegSheet.getRows().insertByIndex(current_row, 1)
        oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row + 1).CellStyle = "Ultimus_centro_bordi_lati"
        current_row += 1

        # Riga parziale
        oRegSheet.getRows().insertByIndex(current_row, 1)
        testo_parziale = "Parziale della Sicurezza €" if current_section_type == 'VDS' else "Parziale dei Lavori a Misura €"
        oRegSheet.getCellByPosition(1, current_row).String = testo_parziale
        oRegSheet.getCellByPosition(1, current_row).CellStyle = "Ultimus_destra"
        oRegSheet.getCellByPosition(8, current_row).Formula = f"=SUBTOTAL(9;I{current_section_start+1}:I{section_end_row+1})"
        oRegSheet.getCellByPosition(8, current_row).CellStyle = "Ultimus_destra_totali"
        current_row += 2  # Spazio prima delle firme

    # 6. TOTALE GENERALE E FIRME
    lastRowWithData = current_row - 2

    # --- BUGFIX: Aggiunge l'importo totale alla riga gialla di riepilogo (insRow) ---
    oRegSheet.getCellByPosition(8, insRow).Formula = f"=SUBTOTAL(9;$I${prima_riga_dati+1}:$I${lastRowWithData+1})"
    oRegSheet.getCellByPosition(8, insRow).CellStyle = "comp sotto Euro 3_R"

    num_righe_firme = 22
    oRegSheet.getRows().insertByIndex(current_row, num_righe_firme)

    # Stile blocco firme
    oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row + num_righe_firme - 1).CellStyle = "Ultimus_centro_bordi_lati"

    # Totale generale
    oRegSheet.getCellByPosition(1, current_row).String = "Lavori a tutto il ___/___/_________ - T O T A L E  €"
    oRegSheet.getCellByPosition(1, current_row).CellStyle = "Ultimus_destra"
    oRegSheet.getCellByPosition(8, current_row).Formula = f"=SUBTOTAL(9;$I${prima_riga_dati+1}:$I${lastRowWithData+1})"
    oRegSheet.getCellByPosition(8, current_row).CellStyle = "Ultimus_destra_totali"

    # Dati per firme
    oSheet_S2 = oDoc.getSheets().getByName("S2")
    data_str = oSheet_S2.getCellRangeByName('$S2.C4').String.split(' ')[-1]
    datafirme = (data_str + ", ") if data_str else "Data, "
    nome_dl = oSheet_S2.getCellRangeByName("$S2.C16").String
    nome_impresa = oSheet_S2.getCellRangeByName("$S2.C17").String

    # Posizionamento firme
    riga_base_firme = current_row + 4
    oRegSheet.getCellByPosition(1, riga_base_firme).CellStyle = "Ultimus_destra"
    oRegSheet.getCellByPosition(1, riga_base_firme).Formula = f'=CONCATENATE("{datafirme}";TEXT(NOW();"GG/mm/aaaa"))'

    oRegSheet.getCellByPosition(1, riga_base_firme + 2).Formula = f'L\'Impresa esecutrice\n({nome_impresa})'
    oRegSheet.getCellByPosition(1, riga_base_firme + 6).Formula = f'Il Direttore dei Lavori\n({nome_dl})'

    # Certificato di Pagamento
    nSal_Cert = 1
    for i in reversed(range(1, 51)):
        if oDoc.NamedRanges.hasByName(f"_Lib_{i}"):
            nSal_Cert = i
            break

    oRegSheet.getCellByPosition(1, riga_base_firme + 10).CellStyle = "Ultimus_destra"
    oRegSheet.getCellByPosition(1, riga_base_firme + 10).Formula = f'=CONCATENATE("In data __/__/____ è stato emesso il CERTIFICATO DI PAGAMENTO n.{nSal_Cert} per un importo di €")'
    oRegSheet.getCellByPosition(9, riga_base_firme + 10).CellStyle = "List-num-euro"

    # Seconda firma del DL
    oRegSheet.getCellByPosition(1, riga_base_firme + 12).Formula = f'Il Direttore dei Lavori\n({nome_dl})'

    # 7. CHIUSURA (A RIPORTARE)
    riga_riportare = current_row + num_righe_firme
    oRegSheet.getCellByPosition(1, riga_riportare).String = "A   R I P O R T A R E"
    oRegSheet.getCellByPosition(8, riga_riportare).Formula = f'=IF(SUBTOTAL(9;$I$2:$I${riga_riportare})=0;"";SUBTOTAL(9;$I$2:$I${riga_riportare}))'
    oRegSheet.getCellRangeByPosition(0, riga_riportare, 9, riga_riportare).CellStyle = "Ultimus_Bordo_sotto"

    # 8. RIEMPIMENTO PAGINA finale
    num_filler = _riempi_pagina(oRegSheet, riga_riportare, col=1, last_col=9, h_pagina=25810, margine=2000, max_filler=20)
    riga_riportare += num_filler

    # 9. OTTIMIZZAZIONE E LAYOUT
    PL.comando('CalculateHard') # Forza ricalcolo layout
    LeenoSheetUtils.adattaAltezzaRiga(oRegSheet)
    SheetUtils.visualizza_PageBreak(True)

    # 10. AREA NOMINALE E STAMPA
    area_rif = f"$A${insRow+2}:$J${riga_riportare+1}"
    nome_area = f"_Reg_{nSal_corrente}"
    SheetUtils.NominaArea(oDoc, "Registro", area_rif, nome_area)

    oNamedRange = oDoc.NamedRanges.getByName(nome_area).ReferredCells.RangeAddress
    oRegSheet.setPrintAreas((oNamedRange,))

    # Ottimizzazione altezze
    oRegSheet.getCellRangeByPosition(0, riga_base_firme, 9, riga_base_firme + 18).Rows.OptimalHeight = True
    LeenoSheetUtils.adattaAltezzaRiga(oRegSheet)

    return True




def setup_intestazioni_registro(oSheet, nSal, oDoc):
    ''' Configura intestazioni, larghezze colonne e testata del Registro '''

    # --- 1. Intestazioni di Colonna ---
    # Definiamo titoli e larghezze in un'unica struttura per scorrere velocemente
    # Formato: (Titolo, Larghezza in 1/100mm)
    cols_config = [
        ("N. ord.\nArticolo\nData", 1600),
        ("LAVORAZIONI\nE SOMMINISTRAZIONI", 7500),
        ("Lib.\nN.", 650),
        ("Lib.\nP.", 650),
        ("U.M.", 1000),
        ("Quantità\nPositive", 1600),
        ("Quantità\nNegative", 1600),
        ("Prezzo\nunitario", 1400),
        ("Importo\ndebito", 1950),
        ("Importo\npagamento", 1950)
    ]

    # Applichiamo lo stile alla riga 0 (Intestazione)
    oRangeHead = oSheet.getCellRangeByPosition(0, 0, len(cols_config)-1, 0)
    oRangeHead.CellStyle = "comp Int_colonna_R"

    for i, (titolo, width) in enumerate(cols_config):
        oCell = oSheet.getCellByPosition(i, 0)
        oCell.String = titolo
        oCell.Columns.Width = width

    # --- 2. Configurazione Pagina e Header ---
    # Recuperiamo i dati dal foglio S2 (Configurazione LeenO)
    try:
        oSheetS2 = oDoc.Sheets.getByName('S2')
        committente = oSheetS2.getCellRangeByName("C6").String
        oggetto_lavori = oSheetS2.getCellRangeByName("C7").String
    except:
        committente = "Committente non definito"
        oggetto_lavori = ""

    # Applichiamo lo stile di pagina (deve esistere nel template)
    style_name = 'PageStyle_REGISTRO_A4'
    if oDoc.StyleFamilies.getByName('PageStyles').hasByName(style_name):
        oSheet.PageStyle = style_name
        oStyle = oDoc.StyleFamilies.getByName('PageStyles').getByName(style_name)

        # Costruiamo il testo per l'header
        testo_header = (f"Committente: {committente}\n"
                        f"Lavori: {oggetto_lavori}\n"
                        f"REGISTRO DI CONTABILITÀ n. {nSal}")

        # Usiamo l'helper di LeenO per impostare l'header
        LS.set_header(oStyle, testo_header, '', '')
        LS.npagina() # Gestione numerazione pagine

    # --- 3. Righe da ripetere in stampa ---
    # Impostiamo la riga 0 come riga di intestazione fissa su ogni pagina stampata
    iSheet = oSheet.RangeAddress.Sheet
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet
    oTitles.StartRow = 0
    oTitles.EndRow = 0
    oSheet.setTitleRows(oTitles)
    oSheet.setPrintTitleRows(True)



def setup_intestazioni_sal(oSheet, nSal, oDoc):
    ''' Configura colonne e intestazioni specifiche per il foglio SAL '''
    cols_config = [
        ("N. ord.\nArticolo", 1600),
        ("LAVORAZIONI\nE SOMMINISTRAZIONI", 11050),
        ("U.M.", 1500),
        ("Quantità", 1800),
        ("Prezzo\nunitario", 1400),
        ("Importo", 1900)
    ]

    oRangeHead = oSheet.getCellRangeByPosition(0, 0, len(cols_config)-1, 0)
    oRangeHead.CellStyle = "comp Int_colonna_R"

    for i, (titolo, width) in enumerate(cols_config):
        oCell = oSheet.getCellByPosition(i, 0)
        oCell.String = titolo
        oCell.Columns.Width = width


def GeneraSAL(oDoc, dati):
    # Unpack dei dati
    nSal, daVoce, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati

    if not datiSAL and not datiSAL_VDS:
        return

    oSalSheet = setup_foglio(oDoc, "SAL", dopo_di="Registro")
    PL.GotoSheet('SAL')

    # --- 1. Calcolo riga di inserimento ---
    if nSal == 1:
        insRow = 1
        setup_intestazioni_sal(oSalSheet, nSal, oDoc)
    else:
        nome_precedente = f"_SAL_{nSal-1}"
        if oDoc.NamedRanges.hasByName(nome_precedente):
            oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
            insRow = oPrevRange.EndRow + 1
        else:
            insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

    # --- 2. Inserimento Intestazione "segue SAL" ---
    oSalSheet.getRows().insertByIndex(insRow, 1)
    oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu"
    # oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"
    oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - 1÷{aVoce}"

    # --- 3. Scrittura Voci SAL per sezioni ---
    current_row = insRow + 1
    stili_colonne = ["List-stringa-sin", "List-stringa-sin", "List-num-centro", "List-num-euro", "List-num-euro", "List-num-euro"]

    subtotalStartRow = current_row + 1
    foundFirstData = False
    riga_sic_partial = None

    sections = [
        ("LAVORI A MISURA", datiSAL, "Parziale dei Lavori a Misura €"),
        ("SICUREZZA (CALCOLO ANALITICO)", datiSAL_VDS, "Parziale della Sicurezza €")
    ]

    for title, data, partial_label in sections:
        if not data: continue

        # Titolo sezione
        oSalSheet.getRows().insertByIndex(current_row, 1)
        if foundFirstData:
            oSalSheet.getRows().getByIndex(current_row).IsStartOfNewPage = True
        oSalSheet.getCellRangeByPosition(0, current_row, 5, current_row).CellStyle = "Ultimus_centro_bordi_lati"
        oSalSheet.getCellByPosition(1, current_row).String = title
        current_row += 1

        # Voci
        numVoci = len(data)
        dataStartRow = current_row
        totaltStartRow = current_row
        lastDataRowSec = dataStartRow + numVoci - 1

        if not foundFirstData:
            subtotalStartRow = dataStartRow
            foundFirstData = True

        oSalSheet.getRows().insertByIndex(dataStartRow, numVoci)
        oSalSheet.getCellRangeByPosition(0, dataStartRow, 5, lastDataRowSec).setDataArray(tuple(data))

        for col_idx, nome_stile in enumerate(stili_colonne):
            oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRowSec).CellStyle = nome_stile

        oSalSheet.getCellRangeByPosition(0, dataStartRow, 0, lastDataRowSec).Rows.OptimalHeight = True
        current_row = lastDataRowSec + 1

        # Inserimento parziale di sezione (come nel Registro)
        oSalSheet.getRows().insertByIndex(current_row, 2) # Riga vuota + Riga parziale
        oSalSheet.getCellRangeByPosition(0, current_row, 5, current_row + 1).CellStyle = "Ultimus_centro_bordi_lati"

        current_row += 1
        oSalSheet.getCellByPosition(1, current_row).String = partial_label
        oSalSheet.getCellByPosition(1, current_row).CellStyle = "Ultimus_destra"
        oSalSheet.getCellByPosition(5, current_row).Formula = f"=SUBTOTAL(9;F{dataStartRow+1}:F{lastDataRowSec+1})"
        oSalSheet.getCellByPosition(5, current_row).CellStyle = "Ultimus_destra_totali"

        if partial_label == "Parziale della Sicurezza €":
            riga_sic_partial = current_row

        if partial_label == "Parziale dei Lavori a Misura €":
            current_row += 1
            oSalSheet.getRows().insertByIndex(current_row, 1)
            oSalSheet.getCellRangeByPosition(0, current_row, 5, current_row).CellStyle = "Ultimus_centro_bordi_lati"
            oSalSheet.getCellByPosition(1, current_row).Formula = '=CONCATENATE("RIBASSO del ";TEXT(VLOOKUP("Ribasso:";$S2.$B$1:$C$1000;2;0)*100;"#.##0,000");"% da applicare su €")'
            oSalSheet.getCellByPosition(1, current_row).CellStyle = "Ultimus_destra_1"

            oSalSheet.getCellByPosition(5, current_row).Formula = f"=SUBTOTAL(9;F{dataStartRow+1}:F{lastDataRowSec+1})"
            oSalSheet.getCellByPosition(5, current_row).CellStyle = "Ultimus_destra_totali"

        current_row += 1

        # RIEMPIMENTO PAGINA (filler) tra sezioni (solo se ne seguono altre)
        is_last_section = (title == sections[-1][0])
        if not is_last_section:
            num_filler = _riempi_pagina(oSalSheet, current_row, col=1, last_col=5, h_pagina=25850)
            current_row += num_filler


    lastDataRow = current_row - 1

    # --- BUGFIX: Aggiunge l'importo totale alla riga gialla di riepilogo (insRow) ---
    oSalSheet.getCellByPosition(5, insRow).Formula = f"=SUBTOTAL(9;$F${subtotalStartRow+1}:$F${lastDataRow+1})"
    oSalSheet.getCellByPosition(5, insRow).CellStyle = "comp sotto Euro 3_R"

    # --- 4. Riepilogo dopo le voci ---
    r = current_row
    oSalSheet.getRows().insertByIndex(r, 3)  # 3 righe: vuota + totale + chiusura

    # Riga vuota di separazione
    oSalSheet.getCellRangeByPosition(0, r, 5, r).CellStyle = "Ultimus_centro_bordi_lati"
    r += 1

    # Parziale complessivo (Rapporto tra Lavori e Sicurezza)
    # oSalSheet.getCellByPosition(1, r).String = "Parziale dei Lavori a Misura €"
    # oSalSheet.getCellByPosition(1, r).CellStyle = "Ultimus_destra"
    # SUBTOTAL(9;...) ignora le righe che contengono a loro volta SUBTOTAL,
    # quindi la somma finale su tutto il range è corretta.
    # oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;F{subtotalStartRow+1}:F{lastDataRow+1})"
    # oSalSheet.getCellByPosition(5, r).CellStyle = "Ultimus_destra_totali"
    riga_parziale = r  # Salva per passarla al riepilogo
    # r += 1

    # Lavori a tutto il __/__/____ - TOTALE
    # oSalSheet.getCellByPosition(1, r).String = "Lavori a tutto il ___/___/_________ - T O T A L E  €"
    oSalSheet.getCellByPosition(1, r).String = "T O T A L E  €"
    oSalSheet.getCellByPosition(1, r).CellStyle = "Ultimus_destra"
    # oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;$F$2:$F${r})"
    oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;$F${subtotalStartRow}:$F${r})"
    oSalSheet.getCellByPosition(5, r).CellStyle = "Ultimus_destra_totali"
    r += 1

    # Riga vuota di chiusura
    oSalSheet.getCellRangeByPosition(0, r, 5, r).CellStyle = "Ultimus_centro_bordi_lati"

    # --- 5. CHIUSURA CON FILLER E RIEPILOGO ---
    oDoc.calculate()
    fineFirme, insRowRiepilogo = firme_contabili_sal(oDoc, oSalSheet, r + 1, sic, mdo, riga_parziale, riga_sic_partial)

    # --- 5. NamedRange ---
    # Escludiamo la riga "segue SAL" (insRow) dall'area del NamedRange
    # insRow è 0-indexed, quindi la riga dati successiva è insRow + 1
    # Per Calc $A$2 è riga 1, quindi insRow+2 è la coordinata corretta se insRow=1.
    area_sal = f"$A${insRow+2}:$F${fineFirme+1}"
    LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")

    # --- 6. Aggiornamento SITUAZIONE CONTABILE in S2 ---
    aggiorna_S2_sal(oDoc, nSal, insRowRiepilogo, mdo)

    # Altezza ottimale finale per la chiusura
    oSalSheet.getCellRangeByPosition(0, lastDataRow + 1, 0, fineFirme).Rows.OptimalHeight = True

def firme_contabili_sal(oDoc, oSheet, startRow, sic, mdo, riga_subtotale, riga_sic=None):
    '''
    Genera la pagina di riepilogo del SAL con filler dinamico.
    Traduzione dal VBasic originale: usa calcolo posizionale Y
    per riempire fino a fine pagina, poi scrive il riepilogo.
    '''
    fcol = 0
    ncol = "F"  # Colonna degli importi (indice 5)
    h_pagina_std = 25850  # Altezza pagina SAL in 1/100 mm

    # --- 1. FILLER DINAMICO (riempimento bordi fino a fine pagina) ---
    currRow = startRow
    num_filler = _riempi_pagina(oSheet, currRow, col=1, last_col=5, h_pagina=25850)
    currRow += num_filler

    # Riga di chiusura con stile "comp Descr" come nel VBasic originale
    oSheet.getRows().insertByIndex(currRow, 1)
    oSheet.getCellRangeByPosition(fcol, currRow, fcol + 5, currRow).CellStyle = "comp Descr"
    currRow += 1

    # --- 2. PAGINA DI RIEPILOGO ---
    # Il riepilogo inizia sulla nuova pagina
    insRow = currRow  # Salviamo il punto di partenza del riepilogo

    # Inseriamo le righe necessarie per il riepilogo (14 righe)
    oSheet.getRows().insertByIndex(currRow, 14)
    # Impostiamo il salto pagina DOPO l'inserimento per evitare che venga spostato
    oSheet.getRows().getByIndex(insRow).IsStartOfNewPage = True
    oSheet.getCellRangeByPosition(fcol, insRow, fcol + 5, insRow + 13).CellStyle = "Ultimus_centro_bordi_lati"

    # Titolo
    oSheet.getCellByPosition(fcol + 1, insRow + 1).String = "R I E P I L O G O   S A L"
    oSheet.getCellByPosition(fcol + 1, insRow + 1).CellStyle = "Ultimus_centro_Dsottolineato"

    # Info Appalto
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 3, fcol + 1, insRow + 4).CellStyle = "Ultimus_sx_italic"
    oSheet.getCellByPosition(fcol + 1, insRow + 3).String = "Appalto: a misura"
    oSheet.getCellByPosition(fcol + 1, insRow + 4).String = "Offerta: unico ribasso"

    # --- 3. LOGICA ECONOMICA (Valori e Formule) ---
    # Stile colonna importi
    oSheet.getCellRangeByPosition(5, insRow + 6, 5, insRow + 13).CellStyle = "ULTIMUS"

    # Riga del subtotale dei dati (riga_subtotale è 0-indexed)
    Row_Misura = riga_subtotale  # riga 0-indexed dove finiscono i dati

    # Lavori a Misura
    oSheet.getCellByPosition(fcol + 1, insRow + 6).String = "Lavori a Misura €"
    oSheet.getCellByPosition(fcol + 1, insRow + 6).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(5, insRow + 6).Formula = f"=${ncol}${Row_Misura + 1}"

    # Detrazione Sicurezza (negativa)
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 7, fcol + 1, insRow + 8).CellStyle = "Ultimus_destra_1"
    oSheet.getCellByPosition(fcol + 1, insRow + 7).String = "Di cui importo per la Sicurezza €"
    if riga_sic is not None:
        oSheet.getCellByPosition(5, insRow + 7).Formula = f"={ncol}${riga_sic + 1}"
    else:
        oSheet.getCellByPosition(5, insRow + 7).Value = sic

    # Detrazione Manodopera (negativa)
    # oSheet.getCellByPosition(fcol + 1, insRow + 8).String = "Di cui importo per la Manodopera"
    # oSheet.getCellByPosition(5, insRow + 8).CellStyle = "ULTIMUS"
    # oSheet.getCellByPosition(5, insRow + 8).Value = mdo * -1

    # Base Ribasso = somma delle 3 righe precedenti
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 8, fcol + 1, insRow + 9).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(fcol + 1, insRow + 8).String = "Importo dei Lavori a Misura su cui applicare il ribasso €"
    oSheet.getCellByPosition(5, insRow + 8).Formula = f"={ncol}{insRow + 7}-{ncol}{insRow + 8}"

    # Ribasso (testo dinamico + calcolo)
    oSheet.getCellByPosition(fcol + 1, insRow + 9).Formula = \
        '=CONCATENATE("RIBASSO del ";TEXT(VLOOKUP("Ribasso:";$S2.$B$1:$C$1000;2;0)*100;"#.##0,000");"%")'
    oSheet.getCellByPosition(5, insRow + 9).Formula = f"={ncol}{insRow + 9}*-VLOOKUP(\"Ribasso:\";$S2.$B$1:$C$1000;2;0)"

    # Re-integro Sicurezza e Manodopera (positivi)
    # oSheet.getCellRangeByPosition(fcol + 1, insRow + 10, fcol + 1, insRow + 11).CellStyle = "Ultimus_sx_bold"
    # oSheet.getCellByPosition(fcol + 1, insRow + 10).String = "Importo per la Sicurezza €"
    # oSheet.getCellByPosition(5, insRow + 10).Value = sic

    # oSheet.getCellByPosition(fcol + 1, insRow + 12).String = "Importo per la Manodopera"
    # oSheet.getCellByPosition(5, insRow + 12).CellStyle = "ULTIMUS"
    # oSheet.getCellByPosition(5, insRow + 12).Value = mdo

    # Totale Parziale Lavori a Misura
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 11, fcol + 1, insRow + 11).CellStyle = "Ultimus_destra_1"
    oSheet.getCellByPosition(fcol + 1, insRow + 11).String = "PER I LAVORI A MISURA €"
    oSheet.getCellByPosition(5, insRow + 11).Formula = f"=SUM({ncol}{insRow + 9}:{ncol}{insRow + 10})"

    # TOTALE GENERALE
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 13, fcol + 1, insRow + 13).CellStyle = "Ultimus_destra_1"
    oSheet.getCellByPosition(fcol + 1, insRow + 13).String = "T O T A L E  €"
    oSheet.getCellByPosition(5, insRow + 13).CellStyle = "Ultimus_destra_totali"
    oSheet.getCellByPosition(5, insRow + 13).Formula = f"={ncol}{insRow + 8}+{ncol}{insRow + 12}"

    currRow = insRow + 14

    # --- 4. FILLER FINALE (fino a fine pagina del riepilogo) ---
    num_filler = _riempi_pagina(oSheet, currRow, col=1, last_col=5, h_pagina=25850)
    currRow += num_filler


    # Riga finale di chiusura (senza tratteggio, con bordi) richiesto dall'utente
    oSheet.getRows().insertByIndex(currRow, 1)
    # oSheet.getCellRangeByPosition(fcol, currRow, fcol + 5, currRow).CellStyle = "Ultimus_centro_bordi_lati"
    oSheet.getCellRangeByPosition(fcol, currRow, fcol + 5, currRow).CellStyle = "Ultimus_"
    currRow += 1

    # Restituisce (ultima riga del blocco, inizio riepilogo) per NamedRange e S2
    return currRow - 1, insRow


def aggiorna_S2_sal(oDoc, nSal, insRowRiepilogo, mdo):
    '''
    Popola la SITUAZIONE CONTABILE nel foglio S2 con formule
    che puntano alle celle del riepilogo SAL.

    Parametri
    ---------
    oDoc : documento
    nSal : int – numero del SAL corrente
    insRowRiepilogo : int – riga 0-indexed di inizio del riepilogo SAL
    mdo : float – totale manodopera
    '''
    try:
        oS2 = oDoc.getSheets().getByName('S2')
        markerS2 = SheetUtils.uFindString("SITUAZIONE CONTABILE", oS2)
        yS2, xS2 = markerS2[0], markerS2[1]
        col = yS2 + nSal  # Colonna del SAL corrente (F per SAL1, G per SAL2, …)
        ncol = _col_letter(col)

        # Riga 1-indexed del riepilogo SAL per le formule Calc
        R = insRowRiepilogo + 1  # conversione 0-indexed → 1-indexed

        # Mappatura: (offset da xS2, formula o valore)
        # Le formule cross-sheet usano il formato $SAL.F$XX
        dati = [
            # (+8)  Lavori e somministrazioni a MISURA = Riepilogo insRow+6
            (8,  f"=$SAL.$F${R + 6}"),
            # (+4)  Quota sicurezza non soggetta a ribasso = Riepilogo insRow+7
            (4,  f"=$SAL.$F${R + 7}"),
            # (+9)  Quota sicurezza (ripetuta) = Riepilogo insRow+7
            (9,  f"=$SAL.$F${R + 7}"),
            # (+12) Importo su cui applicare il ribasso = Riepilogo insRow+8
            (12, f"=$SAL.$F${R + 8}"),
            # (+13) Ribasso = Riepilogo insRow+9
            (13, f"=$SAL.$F${R + 9}"),
            # (+14) Importo ribassato (PER I LAVORI A MISURA) = Riepilogo insRow+11
            (14, f"=$SAL.$F${R + 11}"),
            # ritenute per infortuni
            (16, f"=({ncol}17+{ncol}12)*$S2.$C$85"),
            # recupero anticipazione
            (17, f"=({ncol}17+{ncol}12)*$S2.$C$80"),
            # detrazioni
            (19, f"={ncol}19+{ncol}20"),
            # (+20) Importo Certificato di pagamento = Riepilogo insRow+13 (TOTALE)
            (20, f"={ncol}12+{ncol}17-{ncol}22"),
        ]

        for offset, formula in dati:
            oS2.getCellByPosition(col, xS2 + offset).Formula = formula

        # Quota MDO (valore diretto, non presente nel riepilogo SAL)
        oS2.getCellByPosition(col, xS2 + 5).Value = mdo   # Quota MDO non sogg.
        oS2.getCellByPosition(col, xS2 + 10).Value = mdo  # Quota MDO (ripetuta)

    except Exception as e:
        # Non bloccante: errore nel popolamento S2 non deve interrompere il SAL
        try:
            DLG.errore(f"Errore aggiornamento S2: {e}")
        except:
            pass


########################################################################

def _riempi_pagina(oSheet, insertAt, col=2, last_col=9, h_pagina=25510, margine=2000, max_filler=10):
    """
    Riempie lo spazio residuo nella pagina corrente con righe tratteggiate.
    Usa la stessa logica del Registro: calcolo posizionale Y con cap massimo.

    Ritorna il numero di righe filler effettivamente inserite.
    """
    PL.comando('CalculateHard')

    y_pos = oSheet.getCellByPosition(col, insertAt - 1).Position.Y
    h_row = oSheet.getRows().getByIndex(insertAt - 1).Height
    ingombro_pag = (y_pos + h_row) % h_pagina
    spazio_da_coprire = h_pagina - ingombro_pag - margine

    if spazio_da_coprire <= 500:
        return 0

    num_righe = min(max_filler, int(spazio_da_coprire // 500))
    if num_righe <= 0:
        return 0

    # Inserimento batch di tutte le righe filler
    oSheet.getRows().insertByIndex(insertAt, num_righe)

    # Formattazione batch dell'intera area filler
    oRange = oSheet.getCellRangeByPosition(0, insertAt, last_col, insertAt + num_righe - 1)
    oRange.CellStyle = "Ultimus_centro_bordi_lati"

    # Inserimento batch delle stringhe di filler
    filler = "––––––––––––––––––––––––––––––"
    oFillerRange = oSheet.getCellRangeByPosition(col, insertAt, col, insertAt + num_righe - 1)
    oFillerRange.setDataArray(tuple((filler,) for _ in range(num_righe)))

    return num_righe


def insrow():
    """
    Riempie l'ultima pagina con righe tratteggiate.
    Usa il metodo VBasic collaudato: inserisce righe una alla volta finché
    Calc non segnala un salto pagina sulla riga appena inserita.
    Funziona per CONTABILITA (Libretto), Registro e SAL.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oRanges = oDoc.NamedRanges

    nSh = {'CONTABILITA': '_Lib_', 'Registro': '_Reg_', 'SAL': '_SAL_'}

    prefix = nSh.get(oSheet.Name)
    if not prefix: return

    try:
        last_indices = ultimo_sal()
        if not last_indices: return
        nSal = last_indices[-1]
        nomearea = prefix + str(nSal)
    except: return

    if not oRanges.hasByName(nomearea): return

    oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
    sRow = oNamedRange.StartRow
    iRow = oNamedRange.EndRow
    insertAt = iRow

    col = 2 if oSheet.Name == 'CONTABILITA' else 1
    last_col = 9  # Colonne A-J per CONTABILITA

    num_righe = _riempi_pagina(oSheet, insertAt, col=col, last_col=last_col)

    if num_righe > 0:
        # Aggiorna l'area nominale per includere le nuove righe
        area_rif = f"$A${sRow+1}:$AJ${iRow + num_righe + 1}"
        SheetUtils.NominaArea(oDoc, oSheet.Name, area_rif, nomearea)



def firme_libretto(lrowF=None, oSheet=None):
    """
    Inserisce i dati per le firme nel foglio specificato o in quello attivo,
    con spaziatura uniforme. Funziona per Contabilità, Registro e SAL.
    """
    oDoc = LeenoUtils.getDocument()

    # Se non passiamo il foglio, prendiamo quello attivo
    if oSheet is None:
        oSheet = oDoc.CurrentController.ActiveSheet

    oSheet_S2 = oDoc.getSheets().getByName("S2")

    # --- 1. Recupero dati da S2 ---
    luogo_raw = oSheet_S2.getCellRangeByName("$S2.C4").String
    ultimo_token = luogo_raw.split(" ")[-1] if luogo_raw else ""
    luogo = f"{ultimo_token}, " if ultimo_token else "Data, "

    # --- 2. Gestione Riga di Partenza ---
    if lrowF is None:
        lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    # --- 3. Composizione Lista Firme ---
    firme = []
    firme.append(f"{luogo} ___/___/_________") # Data

    impresa = oSheet_S2.getCellRangeByName("$S2.C17").String
    firme.append(f"L'Impresa esecutrice\n({impresa})")

    contabile = oSheet_S2.getCellRangeByName("$S2.C14").String
    if contabile:
        firme.append(f"Il Direttore Operativo Contabile\n({contabile})")

    cse = oSheet_S2.getCellRangeByName("$S2.C15").String
    if cse:
        firme.append(f"Visto: il C.S.E.\n({cse})")

    direttore = oSheet_S2.getCellRangeByName("$S2.C16").String
    firme.append(f"Il Direttore dei Lavori\n({direttore})")

    # --- 4. Inserimento Righe e Scrittura ---
    # Calcoliamo la colonna di destinazione in base al foglio
    # Registro usa colonna I (8), SAL usa colonna F (5), Contabilità colonna C (2)
    col = 2 # Default (CONTABILITA)
    if oSheet.Name == "Registro": col = 8
    elif oSheet.Name == "SAL": col = 5

    # Inserimento spazio fisico
    num_righe_firme = len(firme) * 3
    oSheet.getRows().insertByIndex(lrowF, num_righe_firme)

    riga_corrente = lrowF + 1
    for firma in firme:
        oCell = oSheet.getCellByPosition(col, riga_corrente)
        oCell.String = firma # Usiamo String invece di Formula per evitare errori con i nomi

        # Formattazione minima: allineamento a destra per Registro/SAL
        if col > 2:
            oCell.HoriJustify = 3 # Right

        riga_corrente += 3

    # Inserisce un ulteriore spazio finale prima del limite area stampa
    oSheet.getRows().insertByIndex(riga_corrente - 2, 2)

    # RESTITUISCE l'indice dell'ultima riga (fondamentale per area_sal e area_reg)
    return riga_corrente







########################################################################
@with_progress_reclaim(manager_attr='progress')
def GeneraAttiContabili():
    oDoc = LeenoUtils.getDocument()
    EseguiContabilita(oDoc)
    return

def EseguiContabilita(oDoc):
    ''' Coordina la generazione degli atti contabili con barra di stato visibile '''
    indicator = oDoc.getCurrentController().getStatusIndicator()
    try:
        # Blocca l'interfaccia per evitare sfarfallio e velocizzare
        oDoc.lockControllers()

        # Avvia l'indicatore (totale passi: 5)
        indicator.start("Inizializzazione contabilità...", 5)
        indicator.setValue(1)

        # 1. Genera il Libretto
        PL.struttura_off()
        indicator.setText("Generazione Libretto delle Misure...")
        dati = GeneraLibretto(oDoc)
        if not dati:
            indicator.end()
            return

        indicator.setValue(2)

        # 2. Passa i dati al Registro
        indicator.setText("Aggiornamento Registro di Contabilità...")
        GeneraRegistro(oDoc, dati)
        indicator.setValue(3)

        # 3. Passa i dati al SAL
        indicator.setText("Compilazione Stato Avanzamento Lavori (SAL)...")
        GeneraSAL(oDoc, dati)
        indicator.setValue(4)

        # 4. Genera il Certificato di Pagamento
        indicator.setText("Compilazione Certificato di Pagamento...")
        try:
            GeneraCdP(oDoc, dati)
        except Exception as e_cdp:
            # Non bloccante: il CdP è un atto integrativo
            try:
                DLG.errore(f'Attenzione: CdP non generato: {e_cdp}')
            except Exception:
                pass
        indicator.setValue(5)

        # Mostra l'ultimo SAL generato
        listaSal = ultimo_sal()
        try:
            nSal = int(listaSal[-1])
            mostra_sal(nSal)
        except Exception:
            pass

        Dialogs.Info(Text="Atti contabili (Libretto, Registro, SAL e CdP) aggiornati con successo.")

    except Exception as e:
        DLG.errore(f"Errore durante l'esecuzione: {str(e)}")
    finally:
        # Molto importante: sblocca sempre i controller e chiudi l'indicatore
        indicator.end()
        if oDoc.hasControllersLocked():
            oDoc.unlockControllers()



# CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA #
########################################################################
########################################################################
# CERTIFICATO DI PAGAMENTO ## CdP ## CdP ## CdP ## CdP ## CdP ## CdP #
########################################################################

def _col_letter(n):
    '''Converte indice colonna 0-based in lettera/e (es. 0→A, 25→Z, 26→AA).'''
    result = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _intero_in_lettere(n):
    '''Converte un intero non-negativo in lettere italiane.'''
    if n == 0:
        return 'zero'
    uni = ['', 'uno', 'due', 'tre', 'quattro', 'cinque', 'sei',
           'sette', 'otto', 'nove', 'dieci', 'undici', 'dodici',
           'tredici', 'quattordici', 'quindici', 'sedici',
           'diciassette', 'diciotto', 'diciannove']
    deci = ['', '', 'venti', 'trenta', 'quaranta', 'cinquanta',
            'sessanta', 'settanta', 'ottanta', 'novanta']

    def _sotto_mille(m):
        if m == 0:
            return ''
        if m < 20:
            return uni[m]
        d, u = divmod(m, 10)
        tens = deci[d]
        if u in (1, 8):
            tens = tens[:-1]
        return tens + (uni[u] if u else '')

    def _centinaia(m):
        if m == 0:
            return ''
        c, resto = divmod(m, 100)
        pref = 'cento' if c == 1 else (uni[c] + 'cento' if c else '')
        return pref + _sotto_mille(resto)

    res = ''
    if n >= 1_000_000_000:
        q, n = divmod(n, 1_000_000_000)
        res += ('unmiliardo' if q == 1 else _centinaia(q) + 'miliardi')
    if n >= 1_000_000:
        q, n = divmod(n, 1_000_000)
        res += ('unmilione' if q == 1 else _centinaia(q) + 'milioni')
    if n >= 1_000:
        q, n = divmod(n, 1_000)
        res += ('mille' if q == 1 else _centinaia(q) + 'mila')
    res += _centinaia(n)
    return res


def numero_in_lettere_euro(importo):
    '''
    Converte un importo in euro in lettere italiane.
    Esempio: 1234.56 → "milleduecentotrentaquattro/56"
    '''
    if importo < 0:
        return 'meno ' + numero_in_lettere_euro(-importo)
    euro = int(importo)
    cent = round((importo - euro) * 100)
    return f'{_intero_in_lettere(euro)}/{cent:02d}'


def _trova_riva(oCdP, testo, col_hint=None):
    '''
    Cerca `testo` nel foglio CdP e restituisce (row, col) 0-indexed.
    Se col_hint è fornito, cerca solo in quella colonna.
    Ritorna (None, None) se non trovato.
    '''
    try:
        if col_hint is not None:
            row = SheetUtils.uFindStringCol(testo, col_hint, oCdP)
            if row is not None:
                return int(row), col_hint
        result = SheetUtils.uFindString(testo, oCdP)
        if result:
            return result[1], result[0]
    except Exception:
        pass
    return None, None


def setup_foglio_CdP(oDoc):
    '''
    Trova il foglio CdP nel documento.
    Se non presente, lo copia dal template Computo_LeenO.ods.
    Ritorna l'oggetto sheet o None se fallisce.
    '''
    if not oDoc.getSheets().hasByName('CdP'):
        # Costruisce il percorso del template
        template_path = os.path.join(LeenoGlobals.dest(), 'template', 'leeno', 'Computo_LeenO.ods')

        # Carica il template in modalità nascosta
        oTemplate = DocUtils.loadDocument(template_path, Hidden=True)
        if not oTemplate:
            # Fallback se il caricamento fallisce
            Dialogs.Exclamation(
                Title='Certificato di Pagamento',
                Text=f'Impossibile caricare il template:\n{template_path}'
            )
            return None

        try:
            # Determina la posizione di inserimento: a destra di SAL o alla fine
            pos = oDoc.getSheets().Count
            if oDoc.getSheets().hasByName('SAL'):
                pos = oDoc.getSheets().getByName('SAL').getRangeAddress().Sheet + 1

            # Importa il foglio CdP
            oDoc.getSheets().importSheet(oTemplate, 'CdP', pos)
        except Exception as e:
            Dialogs.Exclamation(
                Title='Certificato di Pagamento',
                Text=f'Errore durante l\'importazione del foglio CdP:\n{str(e)}'
            )
            return None
        finally:
            # Chiude il template
            if oTemplate:
                oTemplate.close(True)

    return oDoc.getSheets().getByName('CdP')


def _leggi_iva_da_S2(oS2):
    '''
    Cerca l'aliquota IVA in S2: riga successiva a "Ritenute per infortuni".
    Ritorna il valore come float (es. 0.22 per 22%) oppure 0.22 di default.
    '''
    try:
        row_rit = SheetUtils.uFindStringCol('Ritenute per infortuni', 0, oS2)
        if row_rit is not None:
            iva_val = oS2.getCellByPosition(2, int(row_rit) + 1).Value
            if iva_val and iva_val > 0:
                return iva_val
        # secondo tentativo: cerca con label parziale
        row_rit = SheetUtils.uFindStringCol('Ritenute per infortuni', 1, oS2)
        if row_rit is not None:
            iva_val = oS2.getCellByPosition(2, int(row_rit) + 1).Value
            if iva_val and iva_val > 0:
                return iva_val
    except Exception:
        pass
    return 0.22  # default IVA 22%


def GeneraCdP(oDoc, dati=None, nSal=None):
    '''
    CERTIFICATO DI PAGAMENTO - Popola il foglio CdP per il SAL corrente.
    Usa la struttura fissa del foglio template; individua le celle tramite
    ancoraggi testuali per massima robustezza.
    '''
    if dati:
        nSal, daVoce, aVoce, _, _, _, tot_sic, tot_mdo, _ = dati
    elif nSal is None:
        listaSal = ultimo_sal()
        if not listaSal:
            return
        nSal = int(listaSal[-1])

    oCdP = setup_foglio_CdP(oDoc)
    if oCdP is None:
        return

    # ── Dati S2 ──────────────────────────────────────────────────────────
    oS2 = oDoc.getSheets().getByName('S2')
    markerS2 = SheetUtils.uFindString('SITUAZIONE CONTABILE', oS2)
    yS2, xS2 = markerS2[0], markerS2[1]
    col_sal = yS2 + nSal          # colonna S2 del SAL corrente (0-based)
    s2_col  = _col_letter(col_sal) # lettera corrispondente

    aliquota_iva = _leggi_iva_da_S2(oS2)
    perc_iva_str = f'{aliquota_iva * 100:.0f}'

    # Dati anagrafici
    committente = oS2.getCellByPosition(2, 5).String   # C6
    oggetto     = oS2.getCellByPosition(2, 6).String   # C7
    impresa     = oS2.getCellByPosition(2, 16).String  # C17
    nome_dl     = oS2.getCellByPosition(2, 15).String  # C16
    luogo_raw   = oS2.getCellByPosition(2, 3).String   # C4
    luogo       = luogo_raw.split(' ')[-1] if luogo_raw else ''

    # ── Ancoraggi nel foglio CdP ─────────────────────────────────────────
    def R(testo, col=None):
        r, _ = _trova_riva(oCdP, testo, col)
        return r

    r_comm     = R('COMMITTENTE')
    r_imp      = R('IMPRESA')
    r_Ncert    = R('N°')               # prima riga header blocco sinistra
    r_lavori   = R('Per lavori e somministrazioni')
    r_sogg     = R('SOMMANO importi soggetti')
    r_nonsogg  = R('SOMMANO importi NON soggetti')
    r_ritenuta = R('Ritenuta per infortuni')
    r_certprec = R('Ammontare dei Certificati')
    r_totdet   = R('TOTALE DETRAZIONE')
    r_credito  = R('RISULTA IL CREDITO')
    r_iva      = R('I.V.A.')
    r_totgen   = R('TOTALE GENERALE')
    r_certifica = R('CERTIFICA')

    # Colonna valori destra (F = 5 di default, verifica dalla riga sogg)
    val_col = 5
    if r_sogg is not None:
        for c in range(7, -1, -1):
            t = oCdP.getCellByPosition(c, r_sogg).Type.value
            if t != 'EMPTY':
                val_col = c
                break

    # Colonna importo blocco sinistra (D = 3 di default)
    imp_col_sx = 3
    if r_Ncert is not None:
        for c in range(5, 0, -1):
            t = oCdP.getCellByPosition(c, r_Ncert).Type.value
            if t != 'EMPTY':
                imp_col_sx = c
                break

    v = _col_letter(val_col)      # es. 'F'
    s = _col_letter(imp_col_sx)   # es. 'D'

    # ── 1. Intestazione ──────────────────────────────────────────────────
    if r_comm is not None:
        # cerca la cella editabile alla destra dell'ancora COMMITTENTE
        for c in range(1, 6):
            if oCdP.getCellByPosition(c, r_comm).Type.value == 'EMPTY' or \
               oCdP.getCellByPosition(c, r_comm).CellStyle in ('Default', 'ULTIMUS', ''):
                oCdP.getCellByPosition(c, r_comm).String = committente
                break
    if r_imp is not None:
        for c in range(1, 6):
            if oCdP.getCellByPosition(c, r_imp).Type.value == 'EMPTY' or \
               oCdP.getCellByPosition(c, r_imp).CellStyle in ('Default', 'ULTIMUS', ''):
                oCdP.getCellByPosition(c, r_imp).String = impresa
                break

    # Numero certificato e rata nel titolo (r_Ncert - 2)
    if r_Ncert is not None and r_Ncert >= 2:
        r_titolo = r_Ncert - 2
        # Cerca cella "N. ___" nel titolo e scrive il numero
        for c in range(6):
            cell = oCdP.getCellByPosition(c, r_titolo)
            if 'CERTIFICATO' in cell.String.upper():
                # Trova le celle a destra con placeholder numerico
                for cc in range(c + 1, 7):
                    cv = oCdP.getCellByPosition(cc, r_titolo)
                    if cv.Type.value == 'EMPTY' or cv.Value == 0:
                        cv.Value = nSal
                        break
                break

    # ── 2. Blocco sinistro: certificati precedenti ───────────────────────
    if r_Ncert is not None:
        r_first = r_Ncert + 1   # prima riga dati cert LIST
        for i in range(1, nSal):
            r = r_first + (i - 1)
            if r_sogg is not None and r >= r_sogg:
                break
            col_i      = yS2 + i
            s2_col_i   = _col_letter(col_i)
            # Data SAL i (offset +2)
            data_i = oS2.getCellByPosition(col_i, xS2 + 2).Value
            oCdP.getCellByPosition(0, r).Value  = i          # N°
            oCdP.getCellByPosition(1, r).Value = data_i     # Data
            # Importo: punta a S2 offset+20 = Importo Certificato di Pagamento i
            oCdP.getCellByPosition(imp_col_sx, r).Formula = \
                f'=$S2.${s2_col_i}${xS2 + 21}'

        # TOTALE anticipazione (stessa riga dell'IVA, blocco sinistro)
        if r_iva is not None and r_first <= r_iva:
            oCdP.getCellByPosition(imp_col_sx, r_iva).Formula = \
                f'=SUBTOTAL(9;{s}{r_first + 1}:{s}{r_iva})'

    # ── 3. Blocco destro: importi ─────────────────────────────────────────
    # "Per lavori e somministrazioni" → SAL corrente TOTALE (S2 offset +20)
    if r_lavori is not None:
        oCdP.getCellByPosition(val_col, r_lavori).Formula = \
            f'=$S2.${s2_col}${xS2 + 10}+$S2.${s2_col}${xS2 + 15}'

    # "Per materiali giacenti in cantiere" → riga sotto r_lavori, lascia editabile
    # (non scriviamo nulla: cella già vuota nel template)

    # SOMMANO importi soggetti a ritenute
    if r_sogg is not None and r_lavori is not None:
        oCdP.getCellByPosition(val_col, r_sogg).Formula = \
            f'=SUBTOTAL(9;{v}{r_lavori + 1}:{v}{r_sogg})'

    # SOMMANO importi NON soggetti a ritenute
    if r_nonsogg is not None and r_sogg is not None:
        r_ns_start = r_sogg + 2   # prima riga editabile sezione non soggetti
        oCdP.getCellByPosition(val_col, r_nonsogg).Formula = \
            f'=SUBTOTAL(9;{v}{r_ns_start + 1}:{v}{r_nonsogg})'

    # a) Ritenuta per infortuni 0,5%
    if r_ritenuta is not None and r_sogg is not None:
        oCdP.getCellByPosition(val_col, r_ritenuta).Formula = \
            f'={v}{r_sogg + 1}*0.005'

    # b) Ammontare Certificati precedenti → TOTALE anticipazione (blocco sinistra)
    if r_certprec is not None and r_iva is not None:
        oCdP.getCellByPosition(val_col, r_certprec).Formula = \
            f'={s}{r_iva + 1}'

    # TOTALE DETRAZIONE
    if r_totdet is not None and r_ritenuta is not None:
        oCdP.getCellByPosition(val_col, r_totdet).Formula = \
            f'=SUM({v}{r_ritenuta + 1}:{v}{r_totdet})'

    # RISULTA IL CREDITO DELL'IMPRESA
    if r_credito is not None and r_sogg is not None and r_totdet is not None:
        oCdP.getCellByPosition(val_col, r_credito).Formula = \
            f'={v}{r_sogg + 1}+{v}{r_nonsogg + 1}-{v}{r_totdet + 1}'

    # per I.V.A. al __%
    if r_iva is not None and r_credito is not None:
        # Aggiorna etichetta con percentuale letta da S2
        for c in range(val_col - 1, -1, -1):
            lbl = oCdP.getCellByPosition(c, r_iva).String
            if 'I.V.A.' in lbl or 'IVA' in lbl.upper():
                oCdP.getCellByPosition(c, r_iva).String = \
                    f'per I.V.A. al {perc_iva_str}%'
                break
        oCdP.getCellByPosition(val_col, r_iva).Formula = \
            f'={v}{r_credito + 1}*{aliquota_iva}'

    # TOTALE GENERALE
    if r_totgen is not None and r_credito is not None and r_iva is not None:
        oCdP.getCellByPosition(val_col, r_totgen).Formula = \
            f'={v}{r_credito + 1}+{v}{r_iva + 1}'

    # ── 4. Sezione CERTIFICA ─────────────────────────────────────────────
    if r_certifica is not None and r_totgen is not None:
        # Importo finale (valore numerico per lettere)
        try:
            oDoc.calculate()
            importo_finale = oCdP.getCellByPosition(val_col, r_totgen).Value
        except Exception:
            importo_finale = 0.0
        in_lettere = numero_in_lettere_euro(importo_finale)

        # Scrive la riga "CHE al termine dell'articolo..."
        r_che = r_certifica + 1
        for c in range(6):
            cell = oCdP.getCellByPosition(c, r_che)
            if 'CHE' in cell.String.upper() or cell.String.strip() == '':
                pass  # la riga è già nel template con formula/testo fisso

        # Riga importo in lettere
        r_dicitura = r_certifica + 2
        oCdP.getCellByPosition(1, r_dicitura).String = \
            f'Diconsi: (euro {in_lettere}).'

    # ── 5. Firme ─────────────────────────────────────────────────────────
    if r_certifica is not None:
        r_firma = r_certifica + 3
        # Luogo e data (colonna sinistra)
        oCdP.getCellByPosition(1, r_firma).String = \
            f'{luogo}, ___/___/_________'
        # Il Responsabile del Procedimento (colonna destra)
        # Cerca la cella "Responsabile" nel template
        r_resp, c_resp = _trova_riva(oCdP, 'Responsabile')
        if r_resp is not None:
            # Scrive il nome DL nella riga sotto
            oCdP.getCellByPosition(c_resp, r_resp + 1).String = \
                f'({nome_dl})'

    # ── 6. Area nominata e stampa ────────────────────────────────────────
    ultimo_row = SheetUtils.getLastUsedRow(oCdP) + 1
    area_cdp = f'$A$1:${_col_letter(val_col)}${ultimo_row + 1}'
    SheetUtils.NominaArea(oDoc, 'CdP', area_cdp, f'_CdP_{nSal}')

    # Area di stampa = tutto il foglio
    addr = oCdP.getCellRangeByPosition(0, 0, val_col, ultimo_row).getRangeAddress()
    oCdP.setPrintAreas((addr,))

    LeenoSheetUtils.adattaAltezzaRiga(oCdP)
    PL.GotoSheet('CdP')
    return True



@with_progress_reclaim(manager_attr='progress')
def MENU_GeneraCdP():
    '''
    Rigenera il solo Certificato di Pagamento per l'ultimo SAL registrato.
    Utile per aggiornare l'IVA o i dati anagrafici senza rigenerare tutti gli atti.
    '''
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    listaSal = ultimo_sal()
    if not listaSal:
        Dialogs.Exclamation(
            Title='Certificato di Pagamento',
            Text='Nessun SAL registrato. Generare prima gli atti contabili.'
        )
        return

    nSal = int(listaSal[-1])

    # Utilizziamo l'helper per la rigenerazione
    try:
        indicator = oDoc.getCurrentController().getStatusIndicator()
        indicator.start('Compilazione Certificato di Pagamento...', 1)

        if GeneraCdP(oDoc):
            indicator.setValue(1)
            indicator.end()
            Dialogs.Info(Text='Certificato di Pagamento aggiornato.')
        else:
            indicator.end()
            Dialogs.Exclamation(Text='Errore durante la rigenerazione del CdP.')
    except Exception as e:
        DLG.errore(e)


########################################################################
# g_exportedScripts = GeneraAttiContabili
def MENU_trasferimento_onfly():
    '''
    Trasferisce i dati da COMPUTO/VARIANTE a CONTABILITA on-the-fly.
    '''
    # 1. Esecuzione logica pesante (locked)
    if not _MENU_trasferimento_onfly_core():
        return

    # 2. Finalizzazione UI (unlocked - refresh attivo)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('CONTABILITA')
    
    # Dialogs.Ok(Text='Trasferimento completato con successo!')
    LeenoUtils.DocumentRefresh(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet, all=True)



@LeenoUtils.no_refresh
def _MENU_trasferimento_onfly_core():
    ''' Logica core del trasferimento '''
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()

    # Scegli elaborato sorgente
    try:
        source_name = DLG.ScegliElaborato(Titolo='Scegli foglio sorgente',
                                         flag='export')
    except Exception:
        return False

    if source_name == 'CONTABILITA':
        Dialogs.Exclamation(Title='ATTENZIONE!',
                            Text='Il foglio sorgente non può essere la CONTABILITA.')
        return False

    if not oDoc.getSheets().hasByName(source_name):
        return False

    # Raccolta dati
    data = get_transfer_data(source_name)

    # Preparazione CONTABILITA
    generaContabilita(oDoc)

    # Importazione
    LeenoImport_XPWE.compilaComputo(
        oDoc,
        'CONTABILITA',
        data['capitoliCategorie'],
        data['elencoPrezzi'],
        data['listaMisure']
    )
    return True


def get_transfer_data(source_sheet_name):
    ''' Scansiona il foglio sorgente e restituisce le strutture dati per l'import '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(source_sheet_name)
    lastRow = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    capitoliCategorie = {
        'SuperCategorie': [],
        'Categorie': [],
        'SottoCategorie': []
    }
    listaspcat, listacat, listasbcat = [], [], []
    listaMisure = []
    diz_articoli = {}

    # 1. Ricostruzione categorie per mapping
    for n in range(0, lastRow):
        cell_1 = oSheet.getCellByPosition(1, n)
        cell_2 = oSheet.getCellByPosition(2, n)

        if cell_1.CellStyle == 'Livello-0-scritta':
            desc = cell_2.String
            if desc not in listaspcat:
                listaspcat.append(desc)
                capitoliCategorie['SuperCategorie'].append(
                    {'id_sc': str(len(listaspcat)), 'dessintetica': desc})
        elif cell_2.CellStyle == 'Livello-1-scritta mini':
            desc = cell_2.String
            if desc not in listacat:
                listacat.append(desc)
                capitoliCategorie['Categorie'].append(
                    {'id_sc': str(len(listacat)), 'dessintetica': desc})
        elif cell_2.CellStyle == 'livello2_':
            desc = cell_2.String
            if desc not in listasbcat:
                listasbcat.append(desc)
                capitoliCategorie['SottoCategorie'].append(
                    {'id_sc': str(len(listasbcat)), 'dessintetica': desc})

    # 2. Scansione voci e misure
    nVCItem = 2
    for n in range(0, lastRow):
        if oSheet.getCellByPosition(0,
                                    n).CellStyle in ('Comp Start Attributo',
                                                     'Comp Start Attributo_R'):
            sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, n)
            sopra = sStRange.RangeAddress.StartRow
            sotto = sStRange.RangeAddress.EndRow

            # Voce
            v_data = LeenoComputo.datiVoceComputo(oSheet, sopra)
            tariffa = v_data[1]
            diz_articoli[tariffa] = {'tariffa': tariffa}

            # Misure
            lista_righe = []
            for m in range(sopra + 2, sotto):
                desc_riga = oSheet.getCellByPosition(2, m).String
                idvv = '-2'
                if '- vedi voce n.' in desc_riga:
                    try:
                        idvv = str(int(desc_riga.split('- vedi voce n.')[1].split(' ')[0]) + 1)
                        desc_riga = ''
                    except:
                        pass

                riga = (
                    desc_riga,
                    '',
                    '',
                    PL.valuta_cella(oSheet.getCellByPosition(5, m)),
                    PL.valuta_cella(oSheet.getCellByPosition(6, m)),
                    PL.valuta_cella(oSheet.getCellByPosition(7, m)),
                    PL.valuta_cella(oSheet.getCellByPosition(8, m)),
                    str(oSheet.getCellByPosition(9, m).Value),
                    '0',
                    idvv,
                )
                lista_righe.append(riga)

            listaMisure.append({
                'id_vc': str(nVCItem),
                'id_ep': tariffa,
                'quantita': str(oSheet.getCellByPosition(9, sotto).Value),
                'datamis': PL.oggi(),
                'flags': '134217728' if 'VDS_' in tariffa else '0',
                'idspcat': oSheet.getCellByPosition(31, sotto).String or '0',
                'idcat': oSheet.getCellByPosition(32, sotto).String or '0',
                'idsbcat': oSheet.getCellByPosition(33, sotto).String or '0',
                'lista_rig': lista_righe
            })
            nVCItem += 1

    return {
        'listaMisure': listaMisure,
        'capitoliCategorie': capitoliCategorie,
        'elencoPrezzi': {
            'DizionarioArticoli': diz_articoli
        }
    }
