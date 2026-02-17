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
# import itertools
# import operator
import LeenoConfig
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
    if LeenoUtils.getGlobalVar('sblocca_computo') == 1:
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
            LeenoUtils.setGlobalVar('sblocca_computo', 1)
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

    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
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
    Annulla gli atti dell'ultimo SAL rgistrato.
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
        indicator.start("Annullamento atti in corso...", 4)
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
        nSal = int(listaSal[-1]) -1
        mostra_sal(nSal)
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

        oDoc.Sheets.insertNewByName('CONTABILITA', 3)
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


    SAL = []
    SAL_VDS = [] # Nuova lista per voci della sicurezza
    voci_elaborate = set() # Per evitare duplicazioni durante gli shift di riga
    curr_i = primariga
    current_section_type = None

    # Eseguiamo l'adattamento delle altezze prima del calcolo dei filler
    # per avere coordinate Y corrette
    PL.comando('CalculateHard')
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    PL.comando('CalculateHard')

    last_hard_break_y = 0.0
    if nSal > 1:
        # Per i SAL successivi al primo, il punto zero è l'inizio della sezione
        last_hard_break_y = oSheet.getCellByPosition(0, primariga).Position.Y

    while curr_i <= ultimariga:
        # Recupera dati voce: (num, art, desc, um, quant, prezzo, importo, sic, mdo)
        res = LeenoComputo.datiVoceComputo(oSheet, curr_i)
        if res is None:
            curr_i += 1
            continue

        # Salvaguardia contro il doppio conteggio
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
                    for _ in range(num_righe_filler):
                        oSheet.getRows().insertByIndex(curr_i, 1)
                        oSheet.getCellRangeByPosition(0, curr_i, 11, curr_i).CellStyle = "Ultimus_centro_bordi_lati"
                        oSheet.getCellByPosition(2, curr_i).String = "––––––––––––––––––––––––––––––"
                        # Forza altezza riga filler
                        oSheet.getRows().getByIndex(curr_i).Height = 500
                        curr_i += 1
                        ultimariga += 1

                oSheet.getRows().insertByIndex(curr_i, 1)
                oSheet.getRows().getByIndex(curr_i).IsStartOfNewPage = True

                # Aggiorna il punto zero per la nuova pagina (Hard Break)
                # Attenzione: la nuova pagina inizia a curr_i (titolo sezione)
                last_hard_break_y = oSheet.getCellByPosition(0, curr_i).Position.Y
            else:
                oSheet.getRows().insertByIndex(curr_i, 1)
                # Forza il salto pagina se è un nuovo SAL
                if nSal > 1:
                    oSheet.getRows().getByIndex(curr_i).IsStartOfNewPage = True
                    last_hard_break_y = oSheet.getCellByPosition(0, curr_i).Position.Y
            oSheet.getCellRangeByPosition(0, curr_i, 11, curr_i).CellStyle = "Ultimus_centro_bordi_lati"
            titolo = "SICUREZZA (CALCOLO ANALITICO)" if is_vds else "LAVORI A MISURA"
            oSheet.getCellByPosition(2, curr_i).String = titolo # Colonna C descrizione

            current_section_type = voce_type
            curr_i += 1
            ultimariga += 1
            # Ricarichiamo i dati della riga corrente dopo lo shift
            datiVoce = LeenoComputo.datiVoceComputo(oSheet, curr_i)[1]

        # Smistamento basato sul prefisso "VDS_"
        if voce_type == 'VDS':
            SAL_VDS.append(datiVoce)
        else:
            SAL.append(datiVoce)

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
                gruppi_quant[art] += float(row[4])
                gruppi_importo[art] += float(row[6]) # Importo (indice 6)
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

    # Estrai i salti di pagina orizzontali (righe) da Calc
    breaks = sorted([b.Position for b in oSheet.getRowPageBreaks()])

    nPagFinale = start_nPage # Valore di fallback

    # Ciclo di annotazione su tutto il blocco del SAL (inclusa la nuova riga 'daRiga')
    for i in range(daRiga, aRiga + 1):
        # Il numero di pagina è start_nPage + numero di salti che avvengono DOPO daRiga e PRIMA/SULLA riga i
        num_breaks_between = len([b for b in breaks if daRiga < b <= i])
        nPagCorrente = start_nPage + num_breaks_between

        if i == daRiga:
            # Configura riga di rinvio (segue Libretto...)
            oSheet.getCellRangeByPosition(0, daRiga, 36, daRiga).CellStyle = "uuuuu"
            oSheet.getCellByPosition(2, daRiga).String = f"segue Libretto delle Misure n.{nSal} - {daVoce}÷{aVoce}"
            oSheet.getCellByPosition(20, daRiga).Value = nPagCorrente
            oSheet.getCellByPosition(19, daRiga).Value = nSal
            oSheet.getCellByPosition(23, daRiga).Value = nSal

            # Subtotale nella riga di rinvio (corretto per lo slittamento riga)
            formula_sum = f"=SUBTOTAL(9;$P${primariga + 1}:$P${ultimariga + 1})"
            for c in (15, 25):
                cell = oSheet.getCellByPosition(c, daRiga)
                cell.Formula, cell.CellStyle = formula_sum, "comp sotto Euro 3_R"
            continue

        # Verifica se la riga è una riga di 'voce' (misura) che richiede la marcatura della pagina
        style = oSheet.getCellByPosition(1, i).CellStyle
        if style == "comp Art-EP_R":
             # Annotazione SAL, Registro e Pagina
             oSheet.getCellByPosition(19, i).Value = nSal
             oSheet.getCellByPosition(22, i).String = "#reg"
             oSheet.getCellByPosition(23, i).Value = nSal
             oSheet.getCellByPosition(20, i).Value = nPagCorrente
             nPagFinale = nPagCorrente

    # Scrive l'ultimo numero di pagina annotato nella riga gialla di riepilogo (daRiga)
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

    # 4. INSERIMENTO VOCI CON SEZIONI DINAMICHE
    # Inserimento parziali immediato quando cambia sezione
    current_section_type = None
    current_section_start = None

    for dati_riga, is_vds in REG_DATA_ORDERED:
        voce_type = 'VDS' if is_vds else 'LAVORI'

        # Se cambia il tipo di voce, chiudi la sezione precedente e apri una nuova
        if voce_type != current_section_type:
            # Chiudi sezione precedente con parziale (se esiste)
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
                current_row += 1

                # RIEMPIMENTO PAGINA tra parziale e prossima sezione
                PL.comando('CalculateHard')
                h_pagina_std = 25810
                y_pos = oRegSheet.getCellByPosition(1, current_row - 1).Position.Y
                ingombro_pag = y_pos % h_pagina_std
                spazio_da_coprire = h_pagina_std - ingombro_pag - 2000

                if spazio_da_coprire > 500:
                    num_righe_filler = min(10, int(spazio_da_coprire // 500))
                    for _ in range(num_righe_filler):
                        oRegSheet.getRows().insertByIndex(current_row, 1)
                        oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row).CellStyle = "Ultimus_centro_bordi_lati"
                        oRegSheet.getCellByPosition(1, current_row).String = "––––––––––––––––––––––––––––––"
                        current_row += 1

                current_row += 1  # Spazio prima della prossima sezione

            # Inserisci titolo nuova sezione
            oRegSheet.getRows().insertByIndex(current_row, 1)
            if current_section_type is not None:
                oRegSheet.getRows().getByIndex(current_row).IsStartOfNewPage = True
            titolo = "SICUREZZA (CALCOLO ANALITICO)" if is_vds else "LAVORI A MISURA"
            oRegSheet.getCellByPosition(1, current_row).String = titolo
            oRegSheet.getCellRangeByPosition(0, current_row, 9, current_row).CellStyle = "Ultimus_centro_bordi_lati"
            current_row += 1

            # Inizia nuova sezione
            current_section_type = voce_type
            current_section_start = current_row

        # Inserisci la voce
        oRegSheet.getRows().insertByIndex(current_row, 1)
        oRange = oRegSheet.getCellRangeByPosition(0, current_row, 8, current_row)
        oRange.setDataArray((tuple(dati_riga),))

        oRegSheet.getCellRangeByPosition(0, current_row, 1, current_row).CellStyle = "List-stringa-sin"
        oRegSheet.getCellRangeByPosition(2, current_row, 4, current_row).CellStyle = "List-num-centro"
        oRegSheet.getCellRangeByPosition(5, current_row, 9, current_row).CellStyle = "List-num-euro"

        current_row += 1

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
    PL.comando('CalculateHard')
    h_pagina_std = 25810
    y_pos = oRegSheet.getCellByPosition(1, riga_riportare - 1).Position.Y
    ingombro_pag = y_pos % h_pagina_std
    spazio_da_coprire = h_pagina_std - ingombro_pag - 2000

    if spazio_da_coprire > 500:
        num_righe_filler = int(spazio_da_coprire // 500)
        oRegSheet.getRows().insertByIndex(riga_riportare, num_righe_filler)
        for r in range(riga_riportare, riga_riportare + num_righe_filler):
            oRegSheet.getCellRangeByPosition(0, r, 9, r).CellStyle = "Ultimus_centro_bordi_lati"
            oRegSheet.getCellByPosition(1, r).String = "––––––––––––––––––––––––––––––"
        riga_riportare += num_righe_filler

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


# def firme_contabili_sal(oDoc, oSheet, startRow, sic, mdo, rigaMisura):
#     fcol = 0  # Colonna A
#     insRow = startRow

#     # --- 1. GESTIONE SALTO PAGINA E RIEMPIMENTO RIGHE VUOTE ---
#     currRow = insRow
#     try:
#         # Mentre non siamo all'inizio di una nuova pagina...
#         while not oSheet.getRows().getByIndex(currRow).IsStartOfNewPage:
#             oSheet.getRows().insertByIndex(currRow, 1)

#             # Applichiamo lo stile con i bordi laterali alla riga appena creata
#             # Questo crea l'effetto "tabella continua" fino a fine pagina
#             oRangeVuoto = oSheet.getCellRangeByPosition(fcol, currRow, fcol + 5, currRow)
#             oRangeVuoto.CellStyle = "Ultimus_centro_bordi_lati"

#             if currRow > insRow + 60: break # Sicurezza per evitare loop infiniti
#             currRow += 1

#         # Se abbiamo inserito righe, puliamo l'eccesso per far spazio al titolo
#         if currRow > insRow:
#             oSheet.getRows().removeByIndex(currRow - 1, 1)
#             currRow -= 1
#     except:
#         pass

#     # --- 2. RIEPILOGO SAL (Titolo e Info) ---
#     # Il titolo del riepilogo viene posto subito dopo le righe con i bordi
#     oSheet.getCellByPosition(fcol + 1, currRow + 1).String = "R I E P I L O G O   S A L"
#     oSheet.getCellByPosition(fcol + 1, currRow + 1).CellStyle = "Ultimus_centro_Dsottolineato"

#     # Info Appalto
#     oSheet.getCellByPosition(fcol + 1, currRow + 3).String = "Appalto: a misura"
#     oSheet.getCellByPosition(fcol + 1, currRow + 4).String = "Offerta: unico ribasso"
#     oSheet.getCellRangeByPosition(fcol + 1, currRow + 3, fcol + 1, currRow + 4).CellStyle = "Ultimus_sx_italic"

#     # --- 3. LOGICA ECONOMICA (Valori e Formule) ---
#     # Prepariamo la colonna degli importi (F = indice 5)
#     oSheet.getCellRangeByPosition(5, currRow + 6, 5, currRow + 15).CellStyle = "ULTIMUS"
#     ncol = "F"

#     # Lavori a Misura
#     oSheet.getCellByPosition(fcol + 1, currRow + 6).String = "Lavori a Misura €"
#     oSheet.getCellByPosition(fcol + 1, currRow + 6).CellStyle = "Ultimus_sx_bold"
#     # Punta alla riga del totale calcolato in precedenza (rigaMisura + 1 per indice umano)
#     oSheet.getCellByPosition(5, currRow + 6).Formula = f"=${ncol}${rigaMisura + 1}"

#     # Detrazioni Sicurezza e Manodopera (per calcolo base ribasso)
#     oSheet.getCellByPosition(fcol + 1, currRow + 7).String = "Di cui importo per la Sicurezza"
#     oSheet.getCellByPosition(5, currRow + 7).Value = sic * -1

#     oSheet.getCellByPosition(fcol + 1, currRow + 8).String = "Di cui importo per la Manodopera"
#     oSheet.getCellByPosition(5, currRow + 8).Value = mdo * -1
#     oSheet.getCellRangeByPosition(fcol + 1, currRow + 7, fcol + 1, currRow + 8).CellStyle = "Ultimus_sx"

#     # Base Ribasso
#     oSheet.getCellByPosition(fcol + 1, currRow + 9).String = "Importo dei Lavori a Misura su cui applicare il ribasso"
#     oSheet.getCellByPosition(5, currRow + 9).Formula = f"=SUM({ncol}{currRow + 7}:{ncol}{currRow + 9})"
#     oSheet.getCellByPosition(fcol + 1, currRow + 9).CellStyle = "Ultimus_destra"

#     # Calcolo Ribasso (Testo e Valore)
#     # Nota: $S2.$C$78 è il riferimento standard LeenO per il ribasso d'asta
#     oSheet.getCellByPosition(fcol + 1, currRow + 10).Formula = \
#         '=CONCATENATE("RIBASSO del ";TEXT($S2.$C$78*100;"#.##0,00");"%")'
#     oSheet.getCellByPosition(5, currRow + 10).Formula = f"=-{ncol}{currRow + 10}*$S2.$C$78"
#     oSheet.getCellByPosition(fcol + 1, currRow + 10).CellStyle = "Ultimus_destra"

#     # Re-integro Sicurezza e Manodopera
#     oSheet.getCellByPosition(fcol + 1, currRow + 11).String = "Importo per la Sicurezza"
#     oSheet.getCellByPosition(5, currRow + 11).Value = sic

#     oSheet.getCellByPosition(fcol + 1, currRow + 12).String = "Importo per la Manodopera"
#     oSheet.getCellByPosition(5, currRow + 12).Value = mdo
#     oSheet.getCellRangeByPosition(fcol + 1, currRow + 11, fcol + 1, currRow + 12).CellStyle = "Ultimus_sx"

#     # Totale Parziale Lavori
#     oSheet.getCellByPosition(fcol + 1, currRow + 13).String = "PER I LAVORI A MISURA €"
#     oSheet.getCellByPosition(5, currRow + 13).Formula = f"=SUM({ncol}{currRow + 11}:{ncol}{currRow + 14})"
#     oSheet.getCellByPosition(fcol + 1, currRow + 13).CellStyle = "Ultimus_destra_bold"

#     # TOTALE GENERALE SAL
#     oSheet.getCellByPosition(fcol + 1, currRow + 15).String = "T O T A L E  €"
#     oSheet.getCellByPosition(5, currRow + 15).Formula = f"={ncol}{currRow + 14}"
#     oSheet.getCellByPosition(fcol + 1, currRow + 15).CellStyle = "Ultimus_destra_bold"
#     oSheet.getCellByPosition(5, currRow + 15).CellStyle = "Ultimus_destra_totali"

#     # Ritorna l'ultima riga utilizzata per definire il NamedRange nel chiamante
#     return currRow + 16



# def GeneraSAL(oDoc, dati):
#     # Unpack dei 9 valori passati da GeneraLibretto
#     nSal, _, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati

#     # Il riepilogo SAL comprende sia le voci Lavori sia quelle di Sicurezza
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento (Risoluzione Errore) ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc) # Crea testata se nSal=1
#     else:
#         # Tenta di recuperare dal NamedRange, altrimenti cerca l'ultima riga libera
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).ReferredCells.RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             # Fallback: trova l'ultima riga che contiene dati
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento Righe di Intestazione ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - 1÷{aVoce}"
#     oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu" # Riga gialla

#     # --- 3. Scrittura Dati (DataArray) ---
#     dataStartRow = insRow + 1
#     lastDataRow = dataStartRow + len(datiSAL_Riepilogo) - 1
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # --- 4. Inserimento Formule (Prezzi e Importi) ---
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         # VLOOKUP su elenco_prezzi e calcolo prodotto
#         formule.append([
#             f'=VLOOKUP(A{r};elenco_prezzi;5;0)',
#             f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'
#         ])
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # --- 5. Riepilogo Economico (Sotto le voci) ---
#     r = lastDataRow + 2
#     oSalSheet.getCellByPosition(1, r).String = "PARZIALE LAVORI A MISURA €"
#     oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;F{dataStartRow+1}:F{lastDataRow+1})"

#     oSalSheet.getCellByPosition(1, r+1).String = "di cui SICUREZZA (non soggetta a ribasso) €"
#     oSalSheet.getCellByPosition(5, r+1).Value = sic

#     oSalSheet.getCellByPosition(1, r+2).String = "RIBASSO D'ASTA (da Situazione Contabile)"
#     # Formula LeenO: -(Importo - Sicurezza) * Ribasso
#     oSalSheet.getCellByPosition(5, r+2).Formula = f"=-(F{r+1}-F{r+2}) * $S2.$C$78"

#     oSalSheet.getCellByPosition(1, r+4).String = "TOTALE NETTO SAL €"
#     oSalSheet.getCellByPosition(5, r+4).Formula = f"=F{r+1}+F{r+3}"
#     oSalSheet.getCellByPosition(5, r+4).CellStyle = "Ultimus_destra_totali"

#     # --- 6. Firme e NamedRange ---
#     fineFirme = firme_contabili(r + 6)
#     area_sal = f"$A${insRow+1}:$F${fineFirme+1}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")

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

#     # Configurazione testata SAL
#     style_name = 'PageStyle_SAL_A4'
#     if oDoc.StyleFamilies.getByName('PageStyles').hasByName(style_name):
#         oSheet.PageStyle = style_name
#         oStyle = oDoc.StyleFamilies.getByName('PageStyles').getByName(style_name)

#         try:
#             committente = oDoc.Sheets.getByName('S2').getCellRangeByName("C6").String
#         except: committente = ""

#         testo_header = f"Committente: {committente}\nSTATO AVANZAMENTO LAVORI n. {nSal}"
#         LS.set_header(oStyle, testo_header, '', '')

#     # Righe da ripetere
#     oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
#     oTitles.Sheet = oSheet.RangeAddress.Sheet
#     oTitles.StartRow = 0
#     oTitles.EndRow = 0
#     oSheet.setTitleRows(oTitles)
#     oSheet.setPrintTitleRows(True)


# def GeneraSAL(oDoc, dati):
#     # Unpack dei 9 valori passati da GeneraLibretto
#     nSal, _, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati

#     # Il riepilogo SAL comprende sia le voci Lavori sia quelle di Sicurezza
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc)
#     else:
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento e Stile Riga Intestazione (Gialla) ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oRangeGiallo = oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow)
#     oRangeGiallo.CellStyle = "uuuuu"
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - 1÷{aVoce}"

# # --- 3. Scrittura Dati e Applicazione Stili Specifici per Colonna ---
#     dataStartRow = insRow + 1
#     numVoci = len(datiSAL_Riepilogo)
#     lastDataRow = dataStartRow + numVoci - 1

#     # Scrittura dei valori (colonne 0-3)
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # Definizione dello schema stili richiesto
#     # 0: List-stringa-sin | 1: List-stringa-sin | 2: List-num-centro
#     # 3: List-num-euro    | 4: List-num-euro    | 5: List-num-euro
#     stili_colonne = [
#         "List-stringa-sin", "List-stringa-sin", "List-num-centro",
#         "List-num-euro",    "List-num-euro",    "List-num-euro"
#     ]

#     # Applicazione massiva degli stili per colonna
#     for col_idx, nome_stile in enumerate(stili_colonne):
#         oRangeCol = oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRow)
#         oRangeCol.CellStyle = nome_stile

#     # --- 4. Inserimento Formule (Prezzi e Importi) ---
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         formule.append([
#             f'=VLOOKUP(A{r};elenco_prezzi;5;0)',
#             f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'
#         ])

#     # Inserimento formule nelle colonne 4 e 5 (Prezzo e Importo)
#     # Gli stili sono già stati applicati nel ciclo precedente
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # Affiniamo gli stili per le colonne numeriche (allineamento e decimali)
#     oSalSheet.getCellRangeByPosition(3, dataStartRow, 5, lastDataRow).CellStyle = "comp sotto destra"

#     # --- 5. Riepilogo Economico e Stili Totali ---
#     r = lastDataRow + 2
#     # Applichiamo uno stile più marcato per le etichette del riepilogo
#     oSalSheet.getCellRangeByPosition(1, r, 1, r+4).CellStyle = "comp_testa_R"

#     oSalSheet.getCellByPosition(1, r).String = "PARZIALE LAVORI A MISURA €"
#     oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;F{dataStartRow+1}:F{lastDataRow+1})"
#     oSalSheet.getCellByPosition(5, r).CellStyle = "Ultimus_destra_totali"

#     oSalSheet.getCellByPosition(1, r+1).String = "di cui SICUREZZA (non soggetta a ribasso) €"
#     oSalSheet.getCellByPosition(5, r+1).Value = sic
#     oSalSheet.getCellByPosition(5, r+1).CellStyle = "comp sotto destra"

#     oSalSheet.getCellByPosition(1, r+2).String = "RIBASSO D'ASTA (da Situazione Contabile)"
#     # Recupero ribasso da S2
#     oSalSheet.getCellByPosition(5, r+2).Formula = f"=-(F{r+1}-F{r+2}) * $S2.$C$78"
#     oSalSheet.getCellByPosition(5, r+2).CellStyle = "comp sotto destra"

#     oSalSheet.getCellByPosition(1, r+4).String = "TOTALE NETTO SAL €"
#     oSalSheet.getCellByPosition(5, r+4).Formula = f"=F{r+1}+F{r+3}"
#     oSalSheet.getCellByPosition(5, r+4).CellStyle = "Ultimus_destra_totali"

#     # --- 6. Firme e NamedRange ---
#     # fineFirme = firme_contabili(r + 6)
#     # area_sal = f"$A${insRow+1}:$F${fineFirme+1}"
#     # # Creazione del range nominato per gestire l'accodamento del SAL successivo
#     # LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")
# # ... (parte precedente di GeneraSAL) ...

#     # r è la riga dove hai inserito "PARZIALE LAVORI A MISURA €" nel punto 5
#     riga_totale_precedente = r

#     # Chiamata alla nuova funzione (passando sic e mdo ricevuti nell'unpack iniziale)
#     fineFirme = firme_contabili_sal(oDoc, oSalSheet, r + 6, sic, mdo, riga_totale_precedente)

#     # Definizione area finale
#     area_sal = f"$A${insRow+1}:$F${fineFirme}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")



# def GeneraSAL(oDoc, dati):
#     # Unpack dei 9 valori passati da GeneraLibretto
#     nSal, daVoce, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc)
#     else:
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento Riga Intestazione (Gialla) ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu"
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"

#     # --- 3. Scrittura Dati e Stili Colonne ---
#     dataStartRow = insRow + 1
#     numVoci = len(datiSAL_Riepilogo)
#     lastDataRow = dataStartRow + numVoci - 1

#     # Inseriamo le righe necessarie per i dati
#     oSalSheet.getRows().insertByIndex(dataStartRow, numVoci)

#     # Scrittura DataArray (colonne 0-3)
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # Applicazione stili granulari come da schema richiesto
#     stili_colonne = [
#         "List-stringa-sin", "List-stringa-sin", "List-num-centro",
#         "List-num-euro",    "List-num-euro",    "List-num-euro"
#     ]
#     for col_idx, nome_stile in enumerate(stili_colonne):
#         oSalSheet.getCellByPosition(col_idx, dataStartRow).Columns.Width # (opzionale)
#         oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRow).CellStyle = nome_stile

#     # --- 4. Formule Prezzi e Importi ---
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         formule.append([
#             f'=VLOOKUP(A{r};elenco_prezzi;5;0)',
#             f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'
#         ])
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # --- 5. Applicazione stili "a tappeto" e Firme ---
#     # riga_inizio_riepilogo è subito dopo i dati
#     r_inizio_firme = lastDataRow + 1

#     # Chiamata alla funzione che gestisce il riempimento bordi, riepilogo e firme
#     fineFirme = firme_contabili_sal(oDoc, oSalSheet, r_inizio_firme, sic, mdo, lastDataRow)

#     # --- 6. Chiusura NamedRange ---
#     area_sal = f"$A${insRow+1}:$F${fineFirme}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")
#     oSalSheet.getRows().OptimalHeight = True





# def GeneraSAL(oDoc, dati):
#     # Unpack dei dati
#     nSal, daVoce, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc)
#     else:
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento Intestazione "segue SAL" ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu"
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"

#     # --- 3. Scrittura Voci SAL ---
#     dataStartRow = insRow + 1
#     numVoci = len(datiSAL_Riepilogo)
#     lastDataRow = dataStartRow + numVoci - 1

#     oSalSheet.getRows().insertByIndex(dataStartRow, numVoci)
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # Stili per colonna
#     stili_colonne = ["List-stringa-sin", "List-stringa-sin", "List-num-centro", "List-num-euro", "List-num-euro", "List-num-euro"]
#     for col_idx, nome_stile in enumerate(stili_colonne):
#         oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRow).CellStyle = nome_stile

#     # Formule Prezzi/Importi
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         formule.append([f'=VLOOKUP(A{r};elenco_prezzi;5;0)', f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'])
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # --- 4. CHIUSURA: BORDURA A TAPPETO, RIEPILOGO E FIRME ---
#     # r_inizio_chiusura è la riga subito dopo l'ultima voce dati
#     r_inizio_chiusura = lastDataRow + 1

#     # Chiamata alla funzione di chiusura
#     fineFirme = firme_contabili_sal(oDoc, oSalSheet, r_inizio_chiusura, sic, mdo, lastDataRow)

#     # --- 5. NamedRange ---
#     area_sal = f"$A${insRow+1}:$F${fineFirme}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")
#     oSalSheet.getRows().OptimalHeight = True


# def GeneraSAL(oDoc, dati):
#     # Unpack dei dati
#     nSal, daVoce, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc)
#     else:
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento Intestazione "segue SAL" ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu"
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"

#     # --- 3. Scrittura Voci SAL ---
#     dataStartRow = insRow + 1
#     numVoci = len(datiSAL_Riepilogo)
#     lastDataRow = dataStartRow + numVoci - 1

#     oSalSheet.getRows().insertByIndex(dataStartRow, numVoci)
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # Schema stili richiesto per colonna
#     stili_colonne = ["List-stringa-sin", "List-stringa-sin", "List-num-centro", "List-num-euro", "List-num-euro", "List-num-euro"]
#     for col_idx, nome_stile in enumerate(stili_colonne):
#         oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRow).CellStyle = nome_stile

#     # Formule Prezzi/Importi
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         formule.append([f'=VLOOKUP(A{r};elenco_prezzi;5;0)', f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'])
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # --- 4. CHIUSURA: RIEPILOGO E FIRME ---
#     # Gestione altezza ottimale per la tabella dati
#     oSalSheet.getRows().getByIndex(dataStartRow).OptimalHeight = True

#     # Chiamata alla funzione di chiusura
#     fineFirme = firme_contabili_sal(oDoc, oSalSheet, lastDataRow + 1, sic, mdo, lastDataRow)

#     # --- 5. NamedRange ---
#     area_sal = f"$A${insRow+1}:$F${fineFirme}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_str=area_sal, nomearea=f"_SAL_{nSal}")

#     # Altezza ottimale finale per tutto il blocco
#     oSalSheet.getRows().getByIndex(insRow, fineFirme).OptimalHeight = True



# def GeneraSAL(oDoc, dati):
#     # Unpack dei dati
#     nSal, daVoce, aVoce, _, _, datiSAL, sic, mdo, datiSAL_VDS = dati
#     datiSAL_Riepilogo = datiSAL + datiSAL_VDS

#     if not datiSAL_Riepilogo:
#         return

#     oSalSheet = setup_foglio(oDoc, "SAL")
#     PL.GotoSheet('SAL')

#     # --- 1. Calcolo riga di inserimento ---
#     if nSal == 1:
#         insRow = 1
#         setup_intestazioni_sal(oSalSheet, nSal, oDoc)
#     else:
#         nome_precedente = f"_SAL_{nSal-1}"
#         if oDoc.NamedRanges.hasByName(nome_precedente):
#             oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).getReferredCells().RangeAddress
#             insRow = oPrevRange.EndRow + 1
#         else:
#             insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

#     # --- 2. Inserimento Intestazione "segue SAL" ---
#     oSalSheet.getRows().insertByIndex(insRow, 1)
#     oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu"
#     oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"

#     # --- 3. Scrittura Voci SAL ---
#     dataStartRow = insRow + 1
#     numVoci = len(datiSAL_Riepilogo)
#     lastDataRow = dataStartRow + numVoci - 1

#     oSalSheet.getRows().insertByIndex(dataStartRow, numVoci)
#     oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

#     # Schema stili richiesto per colonna
#     stili_colonne = ["List-stringa-sin", "List-stringa-sin", "List-num-centro", "List-num-euro", "List-num-euro", "List-num-euro"]
#     for col_idx, nome_stile in enumerate(stili_colonne):
#         oSalSheet.getCellRangeByPosition(col_idx, dataStartRow, col_idx, lastDataRow).CellStyle = nome_stile

#     # Formule Prezzi/Importi
#     formule = []
#     for x in range(dataStartRow, lastDataRow + 1):
#         r = x + 1
#         formule.append([f'=VLOOKUP(A{r};elenco_prezzi;5;0)', f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'])
#     oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

#     # Applica OptimalHeight alle voci inserite
#     oSalSheet.getRows().getByIndex(dataStartRow, lastDataRow).OptimalHeight = True

#     # --- 4. CHIUSURA CON FILLER (insrow dinamico) ---
#     # Passiamo lastDataRow + 1 come punto di partenza per il riempimento
#     fineFirme = firme_contabili_sal(oDoc, oSalSheet, lastDataRow + 1, sic, mdo, lastDataRow)

#     # --- 5. NamedRange ---
#     area_sal = f"$A${insRow+1}:$F${fineFirme}"
#     LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")

#     # Altezza ottimale finale per il blocco chiusura
#     oSalSheet.getRows().getByIndex(lastDataRow + 1, fineFirme).OptimalHeight = True

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
    oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - {daVoce}÷{aVoce}"

    # --- 3. Scrittura Voci SAL per sezioni ---
    current_row = insRow + 1
    stili_colonne = ["List-stringa-sin", "List-stringa-sin", "List-num-centro", "List-num-euro", "List-num-euro", "List-num-euro"]

    subtotalStartRow = current_row
    foundFirstData = False

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
        current_row += 1

        # RIEMPIMENTO PAGINA (filler) tra sezioni (solo se ne seguono altre)
        is_last_section = (title == sections[-1][0])
        if not is_last_section:
            num_filler = _riempi_pagina(oSalSheet, current_row, col=1, last_col=5, h_pagina=25850)
            current_row += num_filler


    lastDataRow = current_row - 1

    # --- 4. Riepilogo dopo le voci ---
    r = current_row
    oSalSheet.getRows().insertByIndex(r, 4)  # 4 righe: vuota + parziale + totale + vuota

    # Riga vuota di separazione
    oSalSheet.getCellRangeByPosition(0, r, 5, r).CellStyle = "Ultimus_centro_bordi_lati"
    r += 1

    # Parziale complessivo (Rapporto tra Lavori e Sicurezza)
    oSalSheet.getCellByPosition(1, r).String = "Parziale dei Lavori a Misura €"
    oSalSheet.getCellByPosition(1, r).CellStyle = "Ultimus_destra"
    # SUBTOTAL(9;...) ignora le righe che contengono a loro volta SUBTOTAL,
    # quindi la somma finale su tutto il range è corretta.
    oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;F{subtotalStartRow+1}:F{lastDataRow+1})"
    oSalSheet.getCellByPosition(5, r).CellStyle = "Ultimus_destra_totali"
    riga_parziale = r  # Salva per passarla al riepilogo
    r += 1

    # Lavori a tutto il __/__/____ - TOTALE
    oSalSheet.getCellByPosition(1, r).String = "Lavori a tutto il ___/___/_________ - T O T A L E  €"
    oSalSheet.getCellByPosition(1, r).CellStyle = "Ultimus_destra"
    oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;$F$2:$F${r})"
    oSalSheet.getCellByPosition(5, r).CellStyle = "Ultimus_destra_totali"
    r += 1

    # Riga vuota di chiusura
    oSalSheet.getCellRangeByPosition(0, r, 5, r).CellStyle = "Ultimus_centro_bordi_lati"

    # --- 5. CHIUSURA CON FILLER E RIEPILOGO ---
    oDoc.calculate()
    fineFirme = firme_contabili_sal(oDoc, oSalSheet, r + 1, sic, mdo, riga_parziale)

    # --- 5. NamedRange ---
    # Escludiamo la riga "segue SAL" (insRow) dall'area del NamedRange
    # insRow è 0-indexed, quindi la riga dati successiva è insRow + 1
    # Per Calc $A$2 è riga 1, quindi insRow+2 è la coordinata corretta se insRow=1.
    area_sal = f"$A${insRow+2}:$F${fineFirme+1}"
    LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")

    # Altezza ottimale finale per la chiusura
    oSalSheet.getCellRangeByPosition(0, lastDataRow + 1, 0, fineFirme).Rows.OptimalHeight = True

def firme_contabili_sal(oDoc, oSheet, startRow, sic, mdo, riga_subtotale):
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

    # Inseriamo le righe necessarie per il riepilogo (16 righe)
    oSheet.getRows().insertByIndex(currRow, 16)
    # Impostiamo il salto pagina DOPO l'inserimento per evitare che venga spostato
    oSheet.getRows().getByIndex(insRow).IsStartOfNewPage = True
    oSheet.getCellRangeByPosition(fcol, insRow, fcol + 5, insRow + 15).CellStyle = "Ultimus_centro_bordi_lati"

    # Titolo
    oSheet.getCellByPosition(fcol + 1, insRow + 1).String = "R I E P I L O G O   S A L"
    oSheet.getCellByPosition(fcol + 1, insRow + 1).CellStyle = "Ultimus_centro_Dsottolineato"

    # Info Appalto
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 3, fcol + 1, insRow + 4).CellStyle = "Ultimus_sx_italic"
    oSheet.getCellByPosition(fcol + 1, insRow + 3).String = "Appalto: a misura"
    oSheet.getCellByPosition(fcol + 1, insRow + 4).String = "Offerta: unico ribasso"

    # --- 3. LOGICA ECONOMICA (Valori e Formule) ---
    # Stile colonna importi
    oSheet.getCellRangeByPosition(5, insRow + 6, 5, insRow + 15).CellStyle = "ULTIMUS"

    # Riga del subtotale dei dati (riga_subtotale è 0-indexed)
    Row_Misura = riga_subtotale  # riga 0-indexed dove finiscono i dati

    # Lavori a Misura
    oSheet.getCellByPosition(fcol + 1, insRow + 6).String = "Lavori a Misura €"
    oSheet.getCellByPosition(fcol + 1, insRow + 6).CellStyle = "Ultimus_sx_bold"
    oSheet.getCellByPosition(5, insRow + 6).Formula = f"=${ncol}${Row_Misura + 1}"

    # Detrazione Sicurezza (negativa)
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 7, fcol + 1, insRow + 8).CellStyle = "Ultimus_sx"
    oSheet.getCellByPosition(fcol + 1, insRow + 7).String = "Di cui importo per la Sicurezza"
    oSheet.getCellByPosition(5, insRow + 7).Value = sic * -1

    # Detrazione Manodopera (negativa)
    oSheet.getCellByPosition(fcol + 1, insRow + 8).String = "Di cui importo per la Manodopera"
    oSheet.getCellByPosition(5, insRow + 8).CellStyle = "ULTIMUS"
    oSheet.getCellByPosition(5, insRow + 8).Value = mdo * -1

    # Base Ribasso = somma delle 3 righe precedenti
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 9, fcol + 1, insRow + 10).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(fcol + 1, insRow + 9).String = "Importo dei Lavori a Misura su cui applicare il ribasso"
    oSheet.getCellByPosition(5, insRow + 9).Formula = f"=SUM({ncol}{insRow + 7}:{ncol}{insRow + 9})"

    # Ribasso (testo dinamico + calcolo)
    oSheet.getCellByPosition(fcol + 1, insRow + 10).Formula = \
        '=CONCATENATE("RIBASSO del ";TEXT($S2.$C$78*100;"#.##0,00");"%")'
    oSheet.getCellByPosition(5, insRow + 10).Formula = f"=-{ncol}{insRow + 10}*$S2.$C$78"

    # Re-integro Sicurezza e Manodopera (positivi)
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 11, fcol + 1, insRow + 12).CellStyle = "Ultimus_sx"
    oSheet.getCellByPosition(fcol + 1, insRow + 11).String = "Importo per la Sicurezza"
    oSheet.getCellByPosition(5, insRow + 11).Value = sic

    oSheet.getCellByPosition(fcol + 1, insRow + 12).String = "Importo per la Manodopera"
    oSheet.getCellByPosition(5, insRow + 12).CellStyle = "ULTIMUS"
    oSheet.getCellByPosition(5, insRow + 12).Value = mdo

    # Totale Parziale Lavori a Misura
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 13, fcol + 1, insRow + 13).CellStyle = "Ultimus_destra_bold"
    oSheet.getCellByPosition(fcol + 1, insRow + 13).String = "PER I LAVORI A MISURA €"
    oSheet.getCellByPosition(5, insRow + 13).Formula = f"=SUM({ncol}{insRow + 10}:{ncol}{insRow + 13})"

    # TOTALE GENERALE
    oSheet.getCellRangeByPosition(fcol + 1, insRow + 15, fcol + 1, insRow + 15).CellStyle = "Ultimus_destra_bold"
    oSheet.getCellByPosition(fcol + 1, insRow + 15).String = "T O T A L E  €"
    oSheet.getCellByPosition(5, insRow + 15).CellStyle = "Ultimus_destra_totali"
    oSheet.getCellByPosition(5, insRow + 15).Formula = f"=SUM({ncol}{insRow + 10}:{ncol}{insRow + 13})"

    currRow = insRow + 16

    # --- 4. FILLER FINALE (fino a fine pagina del riepilogo) ---
    num_filler = _riempi_pagina(oSheet, currRow, col=1, last_col=5, h_pagina=25850)
    currRow += num_filler


    # Riga finale di chiusura (senza tratteggio, con bordi) richiesto dall'utente
    oSheet.getRows().insertByIndex(currRow, 1)
    oSheet.getCellRangeByPosition(fcol, currRow, fcol + 5, currRow).CellStyle = "Ultimus_centro_bordi_lati"
    currRow += 1

    # Restituisce l'ultima riga del blocco per il NamedRange
    return currRow - 1


########################################################################

def _riempi_pagina(oSheet, insertAt, col=2, last_col=9, h_pagina=25510, margine=2000, max_filler=10):
    """
    Riempie lo spazio residuo nella pagina corrente con righe tratteggiate.
    Usa la stessa logica del Registro: calcolo posizionale Y con cap massimo.

    Ritorna il numero di righe filler effettivamente inserite.
    """
    PL.comando('CalculateHard')

    filler = "––––––––––––––––––––––––––––––"
    y_pos = oSheet.getCellByPosition(col, insertAt - 1).Position.Y
    h_row = oSheet.getRows().getByIndex(insertAt - 1).Height
    ingombro_pag = (y_pos + h_row) % h_pagina
    spazio_da_coprire = h_pagina - ingombro_pag - margine

    if spazio_da_coprire <= 500:
        return 0

    num_righe = min(max_filler, int(spazio_da_coprire // 500))
    if num_righe <= 0:
        return 0

    oSheet.getRows().insertByIndex(insertAt, num_righe)
    for r in range(insertAt, insertAt + num_righe):
        oSheet.getCellRangeByPosition(0, r, last_col, r).CellStyle = "Ultimus_centro_bordi_lati"
        oSheet.getCellByPosition(col, r).String = filler

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

        # Avvia l'indicatore (totale passi: 4)
        indicator.start("Inizializzazione contabilità...", 4)
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

        # Mostra l'ultimo SAL generato
        listaSal = ultimo_sal()
        try:
            nSal = int(listaSal[-1])
            mostra_sal(nSal)
        except:
            pass

        Dialogs.Info(Text="Atti contabili (Libretto, Registro e SAL) aggiornati con successo.")

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
# g_exportedScripts = GeneraAttiContabili
