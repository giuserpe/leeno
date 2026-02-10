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
        # cancella riga gialla
        oSheet.Rows.removeByIndex(daRiga - 1, 1)
        oDoc.NamedRanges.removeByName(nome_area)
        # cancella area di stampa
        LeenoSheetUtils.DelPrintSheetArea()
        # importo prossimo sal
        oSheet.getCellRangeByName('Z2').Formula = (
        "=$P$2-SUBTOTAL(9;$P$2:$P$" + str(daRiga - 1) + ")"
        )

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
        start_nPage = 1
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

    # 4. RACCOLTA DATI ECONOMICI (LOGICA ORIGINALE ADATTATA)
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator: indicator.start("Generazione Libretto delle Misure...", 7)

    SAL = []
    SAL_VDS = [] # Nuova lista per voci della sicurezza
    curr_i = primariga

    while curr_i <= ultimariga:
        # Recupera dati voce: (num, art, desc, um, quant, prezzo, importo, sic, mdo)
        datiVoce = LeenoComputo.datiVoceComputo(oSheet, curr_i)[1]

        # Smistamento basato sul prefisso "VDS_"
        if 'VDS_' in str(datiVoce[1]):
            SAL_VDS.append(datiVoce)
        else:
            SAL.append(datiVoce)

        curr_i = LeenoSheetUtils.prossimaVoce(oSheet, curr_i, saltaCat=True)

    try:
        # Calcolo somme totali
        # sic analitico = somma degli importi (indice 6) delle voci in SAL_VDS
        tot_sic = sum([float(el[6]) for el in SAL_VDS if el[6]])
        tot_mdo = sum([float(el[8]) for el in SAL if el[8]]) + sum([float(el[8]) for el in SAL_VDS if el[8]])

        # Raggruppamento per datiSAL (Lavori)
        gruppi = defaultdict(float)
        for row in SAL:
            key = (row[1], row[2], row[3])
            gruppi[key] += float(row[4])
        datiSAL = sorted([list(k) + [v] for k, v in gruppi.items()], key=lambda x: x[0])

        # Raggruppamento per datiSAL_VDS (Sicurezza)
        gruppi_vds = defaultdict(float)
        for row in SAL_VDS:
            key = (row[1], row[2], row[3])
            gruppi_vds[key] += float(row[4])
        datiSAL_VDS = sorted([list(k) + [v] for k, v in gruppi_vds.items()], key=lambda x: x[0])

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
        indicator.setValue(2)
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

    except Exception as e:
        DLG.errore(e)
        return

    indicator.setValue(4)
    oSheet.getCellRangeByPosition(0, daRiga, 11, aRiga).CellBackColor = -1
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    SheetUtils.visualizza_PageBreak()

    # --- 9. MARCATURA E CALCOLO PAGINE PROGRESSIVE ---
    indicator.setValue(6)
    row_breaks = oSheet.RowPageBreaks
    nPagFinale = start_nPage

    for i in range(primariga, aRiga + 1):
        if oSheet.getCellByPosition(1, i).CellStyle == "comp Art-EP_R":
            oSheet.getCellByPosition(19, i).Value = nSal
            oSheet.getCellByPosition(22, i).String = "#reg"
            oSheet.getCellByPosition(23, i).Value = nSal

            offset_interno = 0
            for pb in row_breaks:
                if primariga < pb.Position <= i:
                    offset_interno += 1
            nPagCorrente = start_nPage + offset_interno
            oSheet.getCellByPosition(20, i).Value = nPagCorrente
            nPagFinale = nPagCorrente

    # --- 10. RIGA DI RINVIO (Omissis logica inserimento daRiga) ---
    oSheet.getRows().insertByIndex(daRiga, 1)
    oSheet.getCellRangeByPosition(0, daRiga, 36, daRiga).CellStyle = "uuuuu"
    oSheet.getCellByPosition(2, daRiga).String = f"segue Libretto delle Misure n.{nSal} - {daVoce}÷{aVoce}"
    oSheet.getCellByPosition(20, daRiga).Value = nPagFinale
    oSheet.getCellByPosition(19, daRiga).Value = nSal
    oSheet.getCellByPosition(23, daRiga).Value = nSal

    formula_sum = f"=SUBTOTAL(9;$P${primariga + 2}:$P${ultimariga + 2})"
    for c in (15, 25):
        cell = oSheet.getCellByPosition(c, daRiga)
        cell.Formula, cell.CellStyle = formula_sum, "comp sotto Euro 3_R"
    # --- 11. AGGIORNAMENTO S2 ---
    oS2.getCellByPosition(yS2 + nSal, xS2 + 1).Value = nSal
    oS2.getCellByPosition(yS2 + nSal, xS2 + 2).Value = date.today().toordinal() - 693594
    oS2.getCellByPosition(yS2 + nSal, xS2 + 24).Value = int(aVoce)
    oS2.getCellByPosition(yS2 + nSal, xS2 + 25).Value = nPagFinale

    if indicator: indicator.end()
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






def setup_foglio(oDoc, nome_foglio):
    ''' Crea il foglio se non esiste e restituisce l'oggetto sheet '''
    sheets = oDoc.getSheets()
    if not sheets.hasByName(nome_foglio):
        # Inserisce il nuovo foglio dopo 'CONTABILITA'
        pos = sheets.getByName('CONTABILITA').RangeAddress.Sheet + 1
        sheets.insertNewByName(nome_foglio, pos)
        sheet = sheets.getByName(nome_foglio)
        # Qui potresti chiamare una funzione per scrivere le intestazioni fisse
        scrivi_intestazioni_fisse(sheet, nome_foglio)
    else:
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

    # 9. AREA NOMINALE E STAMPA
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


def GeneraSAL(oDoc, dati):
    # Unpack degli 8 valori passati da GeneraLibretto
    nSal, _, aVoce, _, _, datiSAL_Riepilogo, sic, mdo = dati

    oSalSheet = setup_foglio(oDoc, "SAL")
    PL.GotoSheet('SAL')

    # --- 1. Calcolo riga di inserimento (Risoluzione Errore) ---
    if nSal == 1:
        insRow = 1
        setup_intestazioni_sal(oSalSheet, nSal, oDoc) # Crea testata se nSal=1
    else:
        # Tenta di recuperare dal NamedRange, altrimenti cerca l'ultima riga libera
        nome_precedente = f"_SAL_{nSal-1}"
        if oDoc.NamedRanges.hasByName(nome_precedente):
            oPrevRange = oDoc.NamedRanges.getByName(nome_precedente).ReferredCells.RangeAddress
            insRow = oPrevRange.EndRow + 1
        else:
            # Fallback: trova l'ultima riga che contiene dati
            insRow = SheetUtils.getLastUsedRow(oSalSheet) + 1

    # --- 2. Inserimento Righe di Intestazione ---
    oSalSheet.getRows().insertByIndex(insRow, 1)
    oSalSheet.getCellByPosition(1, insRow).String = f"segue SAL n.{nSal} - 1÷{aVoce}"
    oSalSheet.getCellRangeByPosition(0, insRow, 5, insRow).CellStyle = "uuuuu" # Riga gialla

    # --- 3. Scrittura Dati (DataArray) ---
    dataStartRow = insRow + 1
    lastDataRow = dataStartRow + len(datiSAL_Riepilogo) - 1
    oSalSheet.getCellRangeByPosition(0, dataStartRow, 3, lastDataRow).setDataArray(tuple(datiSAL_Riepilogo))

    # --- 4. Inserimento Formule (Prezzi e Importi) ---
    formule = []
    for x in range(dataStartRow, lastDataRow + 1):
        r = x + 1
        # VLOOKUP su elenco_prezzi e calcolo prodotto
        formule.append([
            f'=VLOOKUP(A{r};elenco_prezzi;5;0)',
            f'=IF(C{r}="%";D{r}*E{r}/100;D{r}*E{r})'
        ])
    oSalSheet.getCellRangeByPosition(4, dataStartRow, 5, lastDataRow).setFormulaArray(tuple(formule))

    # --- 5. Riepilogo Economico (Sotto le voci) ---
    r = lastDataRow + 2
    oSalSheet.getCellByPosition(1, r).String = "PARZIALE LAVORI A MISURA €"
    oSalSheet.getCellByPosition(5, r).Formula = f"=SUBTOTAL(9;F{dataStartRow+1}:F{lastDataRow+1})"

    oSalSheet.getCellByPosition(1, r+1).String = "di cui SICUREZZA (non soggetta a ribasso) €"
    oSalSheet.getCellByPosition(5, r+1).Value = sic

    oSalSheet.getCellByPosition(1, r+2).String = "RIBASSO D'ASTA (da Situazione Contabile)"
    # Formula LeenO: -(Importo - Sicurezza) * Ribasso
    oSalSheet.getCellByPosition(5, r+2).Formula = f"=-(F{r+1}-F{r+2}) * $S2.$C$78"

    oSalSheet.getCellByPosition(1, r+4).String = "TOTALE NETTO SAL €"
    oSalSheet.getCellByPosition(5, r+4).Formula = f"=F{r+1}+F{r+3}"
    oSalSheet.getCellByPosition(5, r+4).CellStyle = "Ultimus_destra_totali"

    # --- 6. Firme e NamedRange ---
    fineFirme = firme_contabili(r + 6)
    area_sal = f"$A${insRow+1}:$F${fineFirme+1}"
    LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area_sal, f"_SAL_{nSal}")

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

    # Configurazione testata SAL
    style_name = 'PageStyle_SAL_A4'
    if oDoc.StyleFamilies.getByName('PageStyles').hasByName(style_name):
        oSheet.PageStyle = style_name
        oStyle = oDoc.StyleFamilies.getByName('PageStyles').getByName(style_name)

        try:
            committente = oDoc.Sheets.getByName('S2').getCellRangeByName("C6").String
        except: committente = ""

        testo_header = f"Committente: {committente}\nSTATO AVANZAMENTO LAVORI n. {nSal}"
        LS.set_header(oStyle, testo_header, '', '')

    # Righe da ripetere
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = oSheet.RangeAddress.Sheet
    oTitles.StartRow = 0
    oTitles.EndRow = 0
    oSheet.setTitleRows(oTitles)
    oSheet.setPrintTitleRows(True)


########################################################################


# def insrow():
#     """
#     Inserisce righe nel foglio attivo basandosi su ultima area nominata
#     e altezza della pagina.

#     Aggiunge righe finché l'altezza della pagina non viene superata.
#     """
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     oRanges = oDoc.NamedRanges

#     nSh = {
#         'CONTABILITA': '_Lib_',
#         'Registro': '_Reg_',
#         'SAL' : '_SAL_'
#     }

#     nome = nSh.get(oSheet.Name)
#     nSal = ultimo_sal()[-1]
#     nomearea = nome + str(nSal)

#     if oRanges.hasByName(nomearea):
#         oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
#         sRow = oNamedRange.StartRow
#         iRow = oNamedRange.EndRow
#         # ~ return [iRow, iRow - sRow]

#     if oSheet.Name == 'CONTABILITA':
#         col = 2
#     else:
#         col = 1
#     hattuale = oSheet.getCellByPosition(col, iRow).Position.Y - \
#     oSheet.getCellByPosition(col, sRow).Position.Y

#     if oSheet.Name == 'CONTABILITA':
#         hpagina = (len(oSheet.RowPageBreaks) - 1) * 25510
#     elif oSheet.Name == 'Registro':
#         hpagina = (len(oSheet.RowPageBreaks) - 1) * 25810
#     elif oSheet.Name == 'SAL':
#         hpagina = (len(oSheet.RowPageBreaks) - 1) * 25850


#     for i in range(50):
#         oSheet.getRows().insertByIndex(iRow, 1)
#         oSheet.getCellByPosition(col, iRow).String = '––––––––––––––––––––––––––––––' #+ str(i)
#         iRow += 1
#         # Verifica se la cella supera l'altezza pagina e interrompe il ciclo se necessario
#         hattuale = oSheet.getCellByPosition(col, iRow).Position.Y - \
#         oSheet.getCellByPosition(col, sRow).Position.Y

#         # ~ DLG.chi(f'hattuale: {hattuale}\nhpagina: {hpagina}')

#         if hattuale >= hpagina:
#             break
#     return



# def insrow():
#     """
#     Inserisce righe nel foglio attivo basandosi su ultima area nominata
#     e altezza della pagina.

#     Aggiunge righe finché l'altezza della pagina non viene superata.
#     """
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     oRanges = oDoc.NamedRanges

#     nSh = {
#         'CONTABILITA': '_Lib_',
#         'Registro': '_Reg_',
#         'SAL': '_SAL_'
#     }

#     # Costanti altezze pagina
#     hPage = {
#         'CONTABILITA': 25510,
#         'Registro': 25810,
#         'SAL': 25850
#     }

#     nome = nSh.get(oSheet.Name)
#     if not nome:
#         return  # foglio non gestito

#     nSal = ultimo_sal()[-1]
#     nomearea = nome + str(nSal)

#     if not oRanges.hasByName(nomearea):
#         return  # nessuna area trovata

#     oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
#     sRow = oNamedRange.StartRow
#     iRow = oNamedRange.EndRow

#     # Colonna di riferimento
#     col = 2 if oSheet.Name == 'CONTABILITA' else 1

#     # Altezza disponibile
#     hpagina = (len(oSheet.RowPageBreaks) - 1) * hPage[oSheet.Name]

#     # Linea di riempimento
#     filler = "––––––––––––––––––––––––––––––"

#     for _ in range(50):
#         oSheet.getRows().insertByIndex(iRow, 1)
#         oSheet.getCellByPosition(col, iRow).String = filler
#         iRow += 1

#         hattuale = (
#             oSheet.getCellByPosition(col, iRow).Position.Y -
#             oSheet.getCellByPosition(col, sRow).Position.Y
#         )

#         if hattuale >= hpagina:
#             break



# def insrow():
#     """
#     Riempie l'ultima pagina del libretto calcolando lo spazio residuo
#     tramite l'ingombro relativo (modulo altezza pagina).
#     """
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     oRanges = oDoc.NamedRanges

#     nSh = {'CONTABILITA': '_Lib_', 'Registro': '_Reg_', 'SAL': '_SAL_'}
#     hPage = {'CONTABILITA': 25510, 'Registro': 25810, 'SAL': 25850}

#     prefix = nSh.get(oSheet.Name)
#     if not prefix: return

#     try:
#         nSal = ultimo_sal()[-1]
#         nomearea = prefix + str(nSal)
#     except: return

#     if not oRanges.hasByName(nomearea): return

#     oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
#     sRow = oNamedRange.StartRow
#     iRow = oNamedRange.EndRow # Riga delle firme

#     col = 2 if oSheet.Name == 'CONTABILITA' else 1
#     filler = "––––––––––––––––––––––––––––––"
#     h_pagina_std = hPage[oSheet.Name]

#     # 1. Calcolo altezza attuale del libretto
#     y_inizio = oSheet.getCellByPosition(col, sRow).Position.Y
#     y_fine = oSheet.getCellByPosition(col, iRow).Position.Y
#     altezza_occupata = y_fine - y_inizio

#     # 2. Calcolo quanto spazio occupa l'ultima pagina parziale
#     # Il modulo (%) ci dice quanti mm occupa il libretto nell'ultima pagina
#     ingombro_ultima_pagina = altezza_occupata % h_pagina_std

#     # 3. Lo spazio da coprire è la differenza per arrivare a h_pagina_std
#     # Togliamo un margine di sicurezza (1500) per evitare salti pagina indesiderati
#     spazio_da_coprire = (h_pagina_std - ingombro_ultima_pagina) - 1500

#     if spazio_da_coprire <= 500: # Se manca meno di una riga, usciamo
#         return

#     # 4. Calcolo numero righe
#     altezza_riga = oSheet.getRows().getByIndex(iRow).Height
#     if altezza_riga <= 0: altezza_riga = 500

#     num_righe = int(spazio_da_coprire // altezza_riga)

#     if num_righe > 0:
#         # Inserimento massivo all'interno dell'area (alla riga iRow, spingendo giù le firme)
#         oSheet.getRows().insertByIndex(iRow, num_righe)

#         # Formattazione
#         for r in range(iRow, iRow + num_righe):
#             oSheet.getCellRangeByPosition(0, r, 11, r).CellStyle = "Ultimus_centro_bordi_lati"
#             oSheet.getCellByPosition(col, r).String = filler





def insrow():
    """
    Riempie la pagina calcolando lo spazio residuo.
    Funziona per CONTABILITA (Libretto) e Registro.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oRanges = oDoc.NamedRanges

    # Mappatura prefissi e altezze pagina (espresse in centesimi di mm)
    nSh = {'CONTABILITA': '_Lib_', 'Registro': '_Reg_', 'SAL': '_SAL_'}
    hPage = {'CONTABILITA': 25510, 'Registro': 25810, 'SAL': 25850}

    prefix = nSh.get(oSheet.Name)
    if not prefix: return

    try:
        # Recupera l'ultimo SAL generato
        nSal = ultimo_sal()[-1]
        nomearea = prefix + str(nSal)
    except: return

    if not oRanges.hasByName(nomearea): return

    oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
    sRow = oNamedRange.StartRow
    iRow = oNamedRange.EndRow

    # Nel Registro vogliamo inserire PRIMA dell'ultima riga (A RIPORTARE)
    # per far sì che la chiusura finisca a fondo pagina
    insertAt = iRow

    col = 2 if oSheet.Name == 'CONTABILITA' else 1
    filler = "––––––––––––––––––––––––––––––"
    h_pagina_std = hPage[oSheet.Name]

    # 1. Calcolo posizione Y assoluta
    y_inizio = oSheet.getCellByPosition(col, sRow).Position.Y
    y_fine = oSheet.getCellByPosition(col, iRow).Position.Y
    altezza_occupata = y_fine - y_inizio

    # 2. Calcolo ingombro nell'ultima pagina
    ingombro_ultima_pagina = altezza_occupata % h_pagina_std

    # 3. Calcolo spazio da coprire
    # Aumentiamo il margine di sicurezza per il Registro a 2000 per gestire meglio i bordi
    margine = 2000 if oSheet.Name == 'Registro' else 1500
    spazio_da_coprire = (h_pagina_std - ingombro_ultima_pagina) - margine

    if spazio_da_coprire <= 800: # Evitiamo righe troppo sottili
        return

    # 4. Determinazione altezza riga e numero righe
    altezza_riga = oSheet.getRows().getByIndex(insertAt).Height
    if altezza_riga <= 0: altezza_riga = 500

    num_righe = int(spazio_da_coprire // altezza_riga)

    if num_righe > 0:
        # Inserimento righe
        oSheet.getRows().insertByIndex(insertAt, num_righe)

        # Formattazione e riempimento
        for r in range(insertAt, insertAt + num_righe):
            # Applica lo stile dei bordi laterali
            oSheet.getCellRangeByPosition(0, r, 9, r).CellStyle = "Ultimus_centro_bordi_lati"
            oSheet.getCellByPosition(col, r).String = filler

        # 5. IMPORTANTE: Aggiorna l'area nominale per includere le nuove righe
        # Se non lo facciamo, l'area di stampa rimane quella vecchia (corta)
        area_rif = f"$A${sRow+1}:$J${iRow + num_righe + 1}"
        SheetUtils.NominaArea(oDoc, oSheet.Name, area_rif, nomearea)





# def insrow_old():
#     """
#     Versione Dinamica: calcola l'altezza pagina reale dal foglio
#     e riempie lo spazio residuo nell'ultima pagina.
#     """
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     oRanges = oDoc.NamedRanges

#     nSh = {'CONTABILITA': '_Lib_', 'Registro': '_Reg_', 'SAL': '_SAL_'}
#     prefix = nSh.get(oSheet.Name)
#     if not prefix: return

#     try:
#         nSal = ultimo_sal()[-1]
#         nomearea = prefix + str(nSal)
#     except: return

#     if not oRanges.hasByName(nomearea): return

#     oAddr = oRanges.getByName(nomearea).ReferredCells.RangeAddress
#     sRow = oAddr.StartRow
#     iRow = oAddr.EndRow # Riga delle firme

#     col = 2 if oSheet.Name == 'CONTABILITA' else 1
#     filler = "––––––––––––––––––––––––––––––"

#     # 1. RILEVAMENTO DINAMICO DELL'ALTEZZA PAGINA
#     # Recuperiamo i salti pagina orizzontali (RowPageBreaks)
#     pb = oSheet.RowPageBreaks
#     DLG.chi(len(pb))
#     if len(pb) < 1:
#         # Se non ci sono salti pagina, non possiamo calcolare l'altezza dinamica
#         # usiamo un fallback o usciamo
#         return

#     # Calcoliamo l'altezza media della pagina basandoci sull'ultimo salto disponibile
#     # Y dell'ultimo break / numero di break = altezza media reale della pagina
#     last_pb_pos = pb[-1].Position # Posizione riga dell'ultimo salto
#     y_last_pb = oSheet.getCellByPosition(col, last_pb_pos).Position.Y
#     h_pagina_dinamica = y_last_pb / len(pb)

#     # 2. CALCOLO INGOMBRO E SPAZIO RESIDUO
#     y_inizio = oSheet.getCellByPosition(col, sRow).Position.Y
#     y_fine = oSheet.getCellByPosition(col, iRow).Position.Y

#     # Altezza occupata dal libretto
#     altezza_libretto = y_fine - y_inizio

#     # Spazio occupato nell'ultima pagina (Modulo)
#     ingombro_ultima_pag = altezza_libretto % h_pagina_dinamica

#     # Spazio da colmare (con margine di sicurezza di circa 1cm)
#     spazio_da_coprire = (h_pagina_dinamica - ingombro_ultima_pag) - 1000

#     if spazio_da_coprire < 500: return

#     # 3. INSERIMENTO RIGHE
#     altezza_riga = oSheet.getRows().getByIndex(iRow).Height
#     if altezza_riga <= 0: altezza_riga = 450 # Altezza minima di sicurezza

#     num_righe = int(spazio_da_coprire // altezza_riga)

#     if num_righe > 0:
#         # Inseriamo all'interno dell'area (alla riga iRow) per espanderla
#         oSheet.getRows().insertByIndex(iRow, num_righe)

#         # Formattazione massiva
#         for r in range(iRow, iRow + num_righe):
#             oSheet.getCellRangeByPosition(0, r, 11, r).CellStyle = "Ultimus_centro_bordi_lati"
#             oSheet.getCellByPosition(col, r).String = filler





########################################################################
# def firme_contabili(lrowF=None):
#     """
#     Inserisce i dati necessari alle firme nel foglio "CONTABILITA",
#     con spaziatura uniforme tra i blocchi.
#     """
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     oSheet_S2 = oDoc.getSheets().getByName("S2")

#     # Ricava il luogo dall'intestazione del foglio S2
#     luogo_raw = oSheet_S2.getCellRangeByName("$S2.C4").String
#     ultimo_token = luogo_raw.split(" ")[-1] if luogo_raw else ""
#     luogo = f"{ultimo_token}, " if ultimo_token else "Data, "

#     if oSheet.Name != "CONTABILITA":
#         return

#     # Se non viene passata una riga, calcola l'ultima disponibile
#     if lrowF is None:
#         lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

#     firme = []

#     # Progettista (luogo + data)
#     firme.append(f"{luogo} ___/___/_________")

#     # Impresa esecutrice
#     impresa = oSheet_S2.getCellRangeByName("$S2.C17").String
#     firme.append(f"L'Impresa esecutrice\n({impresa})")

#     # Direttore Operativo Contabile (solo se presente)
#     contabile = oSheet_S2.getCellRangeByName("$S2.C14").String
#     if contabile:
#         firme.append(f"Il Direttore Operativo Contabile\n({contabile})")

#     # CSE (solo se presente)
#     cse = oSheet_S2.getCellRangeByName("$S2.C15").String
#     if cse:
#         firme.append(f"Visto: il C.S.E.\n({cse})")

#     # Direttore Lavori
#     direttore = oSheet_S2.getCellRangeByName("$S2.C16").String
#     firme.append(f"Il Direttore dei Lavori\n({direttore})")

#     # Numero righe da inserire = blocchi × 3
#     oSheet.getRows().insertByIndex(lrowF, len(firme) * 3)

#     riga_corrente = lrowF + 1
#     for firma in firme:
#         oSheet.getCellByPosition(2, riga_corrente).Formula = firma
#         riga_corrente += 3 # avanza sempre di 3 righe

#     oSheet.getRows().insertByIndex(riga_corrente -2, 3)

#     return riga_corrente +1

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
def EseguiContabilita(oDoc):
    ''' Coordina la generazione degli atti contabili (Libretto, Registro, SAL) '''
    try:
        # 1. Genera il Libretto e recupera gli 8 parametri fondamentali
        PL.struttura_off()
        dati = GeneraLibretto(oDoc)
        if not dati: return

        # 2. Passa i dati al Registro
        GeneraRegistro(oDoc, dati)

        # 3. Passa i dati al SAL
        # GeneraSAL(oDoc, dati)

        listaSal = ultimo_sal()
        try:
            nSal = int(listaSal[-1])
            mostra_sal(nSal)
        except Exception as e:
            DLG.errore(e)

        Dialogs.Info(Title="Atti contabili (Libretto, Registro e SAL) aggiornati con successo.")

    except Exception as e:
        DLG.errore(e)

def GeneraAttiContabili():
    oDoc = LeenoUtils.getDocument()
    EseguiContabilita(oDoc)
    return
    '''
    Genera atti contabili.
    '''
    # with LeenoUtils.DocumentRefreshContext(False):

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != "CONTABILITA":
        return

    # la generazione del libretto è inclusa in GeneraRegistro()
    GeneraLibretto(oDoc)
    # GeneraRegistro(oDoc)

    listaSal = ultimo_sal()
    try:
        nSal = int(listaSal[-1])
        mostra_sal(nSal)
    except Exception as e:
        # ~ DLG.errore(e)
        pass
    PL.GotoSheet('CONTABILITA')

    # ~Dialogs.Info(Title = 'Voci registrate!',
        # ~Text="La generazione degli allegati contabili è stata completata.")


# CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA #
########################################################################
########################################################################
# g_exportedScripts = GeneraAttiContabili
