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


def imposta_data():
    """ Imposta la data scelta nelle misure selezionate."""
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    testo = PL.calendario()

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

        oSheet.getCellRangeByName('P2').Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori registrati
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
                                3).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori registrati
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

def GeneraLibretto(oDoc):
    '''
    CONTABILITA' - Si ottiene una riga gialla con l'indicazione delle
    voci di misurazione registrate ed un parziale dell'importo del SAL a
    cui segue la visualizzazione in struttura delle voci registrate nel
    Libretto delle Misure.
    '''

    # ~oDoc = LeenoUtils.getDocument()
    PL.chiudi_dialoghi()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'CONTABILITA':
        return

    PL.numera_voci()
    oRanges = oDoc.NamedRanges

    #trovo il numero del nuovo sal
    nSal = 1
    for i in reversed(range(1, 50)):
        if oRanges.hasByName("_Lib_" + str(i)) == True:
            nSal = i +1
            break
    daVoce = 1
    old_nPage = 1

    libretti = SheetUtils.sStrColtoList('segue Libretto delle Misure n.', 2, oSheet, start=2)
    try:
        daVoce = int(oSheet.getCellByPosition(2, libretti[-1]
        ).String.split('÷')[1]) + 1
    except:
        daVoce = 1
    oCellRange = oSheet.getCellRangeByPosition(0, 3, 0,
        SheetUtils.getUsedArea(oSheet).EndRow - 2)

    nomearea="_Lib_" + str(nSal)

    #  Recupero la prima riga non registrata
    daVoce = PL.InputBox(str(daVoce), f"SAL n.{nSal}: Libretto delle Misure, da voce n.")
    if len(daVoce) ==0:
        return

    try:
        lrow = int(SheetUtils.uFindStringCol(daVoce, 0, oSheet))
    except TypeError:
        return

    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    primariga = sStRange.RangeAddress.StartRow

    # include nel range del SAL eventuali titoli di categoria
    for el in range(1, 10):
        if oSheet.getCellByPosition(0, primariga - 1).CellStyle in ('Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'):
            primariga -= 1

    #  ULTIMA VOCE
    oCellRange = oSheet.getCellRangeByPosition(
        0, 3, 0, SheetUtils.getUsedArea(oSheet).EndRow - 2)
    aVoce = int(oCellRange.computeFunction(MAX))

    aVoce = PL.InputBox(str(aVoce), f"SAL n.{nSal}: Libretto delle Misure, a voce n.")

    if len(aVoce) == 0:
        return
    if int(aVoce) < int(daVoce):
        Dialogs.Exclamation(Title='ATTENZIONE!', Text=f"Il range di voci scelto ({daVoce} ÷ {aVoce}) non è valido.")
        return

    try:
        lrow = int(SheetUtils.uFindStringCol(aVoce, 0, oSheet))
    except TypeError:
        return
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    ultimariga = sStRange.RangeAddress.EndRow
        
    #################################
    #################################
    #################################
    for i in range(primariga, ultimariga):
        # ~ oDoc.CurrentController.select(oSheet.getCellByPosition(2, i))
        if "SICUREZZA" in oSheet.getCellByPosition(2, i).String:
            # ~ insRighe()
            oSheet.getCellByPosition(0, i).Rows.IsManualPageBreak = True
            break
    #################################
    #################################
    #################################

    # attiva la progressbar
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator is not None:
        indicator.start("Generazione Libretto delle Misure...", 6)
    indicator.setValue(1)


    # Recupero i dati per il SAL
    # ottengo datiSAL = [art,  desc, um, quant] in cui quant è la "somma a tutto il"
    # ottengo datiSAL_VDS = [art,  desc, um, quant] in cui quant è la "somma a tutto il"
    SAL = []
    SAL_VDS = []
    i = 5
    while i < ultimariga + 1:
        '''
        SAL = (art,  desc, um, quant, prezzo, importo, sic, mdo)
        EP = elenco articoli
        '''
        # ~if 'VDS_' in LeenoComputo.datiVoceComputo(oSheet, i)[1][0]:
            # ~datiSAL_VDS = LeenoComputo.datiVoceComputo(oSheet, i)[1] #SAL = (num, art,  desc, um, quant, prezzo, importo, sic, mdo)
            # ~SAL_VDS.append(datiSAL_VDS)
        # ~else:
            # ~datiSAL = LeenoComputo.datiVoceComputo(oSheet, i)[1] #SAL = (num, art,  desc, um, quant, prezzo, importo, sic, mdo)
            # ~SAL.append(datiSAL)
        datiSAL = LeenoComputo.datiVoceComputo(oSheet, i)[1] #SAL = (num, art,  desc, um, quant, prezzo, importo, sic, mdo)
        if 'VDS_' in datiSAL[1]:
            SAL_VDS.append(datiSAL)
        else:
            SAL.append(datiSAL)
        i= LeenoSheetUtils.prossimaVoce(oSheet, i, saltaCat=True)

    # SAL = list(set(SAL))
    # totale_SAL = sum(voce[6] for voce in SAL)

    # SAL_VDS = list(set(SAL_VDS))
    # totale_VDS = sum(voce[6] for voce in SAL_VDS)

    # DLG.chi(f"totale_VDS = {totale_VDS}")
    # DLG.chi(f"totale_SAL = {totale_SAL}")
    # return


    try:
        sic = []
        mdo = []
        for el in SAL:
            sic.append(el[7])
            mdo.append(el[8])
        sic = sum(sic)
        mdo = sum(mdo)




        from collections import defaultdict
        datiSAL = []
        gruppi = defaultdict(float)  # Dizionario per sommare direttamente i valori
        for row in SAL:
            key = (row[1], row[2], row[3])  # Chiave di raggruppamento
            gruppi[key] += float(row[4])  # Somma diretta

        # Converti il dizionario in lista
        datiSAL = [list(k) + [v] for k, v in gruppi.items()]
        datiSAL = sorted(datiSAL, key=lambda x: x[0])

        PL.comando ('DeletePrintArea')
        SheetUtils.visualizza_PageBreak()

        oSheet.getCellByPosition(25, ultimariga - 1).String = "SAL n." + str(nSal)
        oSheet.getCellByPosition(25, ultimariga).Formula = (
            "=SUBTOTAL(9;P" + str(primariga + 1) + ":P" + str(ultimariga+1) + ")" )
        oSheet.getCellByPosition(25, ultimariga).CellStyle = "comp sotto Euro 3_R"

        # immetti le firme
        inizioFirme = ultimariga + 1

        # DLG.chi(inizioFirme)
        # return

        fineFirme = firme_contabili(inizioFirme) # riga di inserimento
        # fineFirme = inizioFirme + 10


        indicator.setValue(2)
        area="$A$" + str(primariga + 1) + ":$AJ$" + str(fineFirme + 1)

        SheetUtils.NominaArea(oDoc, "CONTABILITA", area, nomearea)

        #applico gli stili corretti ad alcuni dati della firma
        oSheet.getCellRangeByPosition(0, inizioFirme, 32, fineFirme).CellStyle = "Ultimus_centro_bordi_lati"
        oSheet.getCellByPosition(2, inizioFirme + 1).CellStyle = "Ultimus_destra"

        oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

        #range del _Lib_
        daRiga = oNamedRange.StartRow
        aRiga = oNamedRange.EndRow
        daColonna = oNamedRange.StartColumn
        aColonna = oNamedRange.EndColumn
        
        LS.importa_stili_pagina_non_presenti()
        LS.setPageStyle()
        oSheet.PageStyle = "Page_Style_Libretto_Misure2"
        LeenoSheetUtils.setLarghezzaColonne(oSheet)
        oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName('Page_Style_Libretto_Misure2')
        committente = "Committente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String + '\nLibretto delle Misure n.' + str(nSal)
        LS.set_header(oAktPage, committente, '', '')
        LS.npagina()
        LS.set_footer(oAktPage, '', "L'IMPRESA \t\t\t\t\t\t\t IL DIRETTORE DEI LAVORI", '')

        iSheet = oSheet.RangeAddress.Sheet

        # imposta area di stampa e riga da ripetere
        oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        oTitles.Sheet = iSheet
        oTitles.StartColumn = 0
        oTitles.StartRow = 2
        oTitles.EndColumn = 11
        oTitles.EndRow = 2
        oSheet.setTitleRows(oTitles)
        oNamedRange.EndColumn = 11
        oSheet.setPrintAreas((oNamedRange,))
        oSheet.setPrintTitleRows(True)
    except Exception as e:
        DLG.chi("ERRORE")
        DLG.errore(e)

    indicator.setValue(3)

    # sbianco l'area di stampa
    oSheet.getCellRangeByPosition(daColonna, daRiga, 11, aRiga).CellBackColor = -1

    x = fineFirme

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    SheetUtils.visualizza_PageBreak()
    oSheet.removeAllManualPageBreaks()

    #=lib===================
    insrow()

    # #cancello la prima riga per aumentare lo spazio per la firma
    # oSheet.getCellByPosition(2, x).String = ""
    # oSheet.getCellByPosition(2, x + 1).String = ""

    indicator.setValue(4)
    # ----------------------------------------------------------------------
    # QUESTA DEVE DIVENTARE UN'OPZIONE A SCELTA DELL'UTENTE
    # in caso di libretto unico questo if è da attivare
    # in modo che la numerazione delle pagine non ricominci da capo
    # ~if nSal > 1:
        # ~nLib = 1
    inumPag = 1 + old_nPage
    nLib = nSal

##########
    # COMPILO LA SITUAZIONE CONTABILE IN "S2" 1di2
    oS2 = oDoc.getSheets().getByName('S2')
    # trovo la posizione del titolo
    oEnd=SheetUtils.uFindString("SITUAZIONE CONTABILE", oS2)
    xS2=oEnd[1]
    yS2=oEnd[0]

    oS2.getCellByPosition(yS2 + nSal, xS2 + 1).Value = nSal
    oS2.getCellByPosition(yS2 + nSal, xS2 + 2).Value = date.today().toordinal() - 693594    #data
    oS2.getCellByPosition(yS2 + nSal, xS2 + 24).Value = aVoce        #ultima voce libretto
    oS2.getCellByPosition(yS2 + nSal, xS2 + 25).Value = inumPag      #ultima pagina libretto
##########

#  inumPag = 0'+ old_nPage 'SE IL LIBRETTO è UNICO

    #inserisco i dati
    LeenoUtils.setGlobalVar('sblocca_computo', 0) #registrando gli atti contabili, bisogna inibire alcune modifiche
    indicator.setValue(5)
    for i in range(primariga, fineFirme):
        if oSheet.getCellByPosition(1, i).CellStyle == "comp Art-EP_R":
            if primariga == 0:
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
                primariga = sStRange.RangeAddress.StartRow
            oSheet.getCellByPosition(19, i).Value= nLib     #numero libretto
            oSheet.getCellByPosition(22, i).String =  "#reg"  #flag registrato
            oSheet.getCellByPosition(23, i).Value= nSal     #numero SAL
            for nPag in range(0, len(oSheet.RowPageBreaks)):
                if i < oSheet.RowPageBreaks[nPag].Position:
                    oSheet.getCellByPosition(20, i).Value = nPag   #pagina
                    break

    indicator.setValue(6)
    # annoto ultimo numero di pagina
    oSheet.getCellByPosition(20 , fineFirme).Value = nPag
    oSheet.getCellByPosition(20 , fineFirme).CellStyle = "num centro"
#  inumPag = nPag ' + old_nPage 'SE IL LIBRETTO è UNICO

    # ~SheetUtils.visualizza_PageBreak(False)

    # inserisco la prima riga GIALLA del LIBRETTO
    oSheet.getRows().insertByIndex(daRiga, 1)
    oSheet.getCellRangeByPosition (0, daRiga, 36, daRiga).CellStyle = "uuuuu"

    #range del _Lib_
    oSheet.getCellByPosition(2,  daRiga).String = (
        "segue Libretto delle Misure n." + str(nSal) +
        " - " + str(daVoce) + "÷" + str(aVoce)
        )
    oSheet.getCellByPosition(20, daRiga).Value =  nPag  #Pagina
    oSheet.getCellByPosition(19, daRiga).Value= nLib    #Libretto
    oSheet.getCellByPosition(23, daRiga).Value= nSal    #SAL
    oSheet.getCellByPosition(15, daRiga).Formula =(
        "=SUBTOTAL(9;$P$" + str(primariga + 2) + ":$P$" + str(
        ultimariga + 2) + ")"
        )
    oSheet.getCellByPosition(15, daRiga).CellStyle = "comp sotto Euro 3_R"
    oSheet.getCellByPosition(25, daRiga).Formula =(
        "=SUBTOTAL(9;$P$" + str(primariga + 2) + ":$P$" + str(
        ultimariga + 2) + ")"
        )
    oSheet.getCellByPosition(25, daRiga).CellStyle = "comp sotto Euro 3_R"

    # annoto il sal corrente sulla riga di intestazione
    oSheet.getCellRangeByName('Z3').Value = nSal
    oSheet.getCellRangeByName('Z3').CellStyle = "Menu_sfondo _input_grasBig"
    oSheet.getCellRangeByName('Z2').Formula = (
        "=$P$2-SUBTOTAL(9;$P$2:$P$" + str(ultimariga + 2) + ")"
        )

    PL._gotoCella(0, daRiga)
    indicator.end()

    return nSal, daVoce, aVoce, primariga+1, ultimariga+1, datiSAL, sic, mdo


########################################################################


def GeneraRegistro(oDoc):
    '''
    CONTABILITA' - genera un nuovo foglio 'Registro'. Si ottiene una riga
    gialla con l'indicazione delle voci di misurazione registrate ed un
    parziale dell'importo del SAL a cui segue la visualizzazione in
    struttura delle relative voci registrate nel Libretto delle Misure.
    '''

    try:
        nSal, daVoce, aVoce, primariga, ultimariga, datiSAL, sic, mdo = GeneraLibretto(oDoc)
    except Exception as e:
        DLG.errore(e)
        return

    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator is not None:
        indicator.start("Generazione Registro di Contabilità...", 5)
    indicator.setValue(1)

# Recupero i dati per il Registro
    oSheet = oDoc.Sheets.getByName("CONTABILITA")
    REG = []
    i = primariga
    while i < ultimariga:
        '''
        REG = ((num + '\n' + art + '\n' + data), desc, Nlib, Plib, um,
            quantP, quantN, prezzo, importo)
        '''
        reg = LeenoComputo.datiVoceComputo(oSheet, i)[0]
        REG.append(reg)
        # i= LeenoSheetUtils.prossimaVoce(oSheet, i)
        i= LeenoSheetUtils.prossimaVoce(oSheet, i, saltaCat=True)
    try:
        oDoc.getSheets().insertNewByName('Registro',oSheet.RangeAddress.Sheet + 1)
        PL.GotoSheet('Registro')
        oSheet = oDoc.Sheets.getByName('Registro')

    # riga di intestazione
        oSheet.getCellRangeByPosition(0,0,9,0).CellStyle="comp Int_colonna_R"
        oSheet.getCellByPosition(0,0).String = ("N. ord.\nArticolo\nData")
        oSheet.getCellByPosition(1,0).String = ("LAVORAZIONI\nE SOMMINISTRAZIONI")
        oSheet.getCellByPosition(2,0).String = ("Lib.\nN.")
        oSheet.getCellByPosition(3,0).String = ("Lib.\nP.")
        oSheet.getCellByPosition(4,0).String = ("U.M.")
        oSheet.getCellByPosition(5,0).String = ("Quantità\nPositive")
        oSheet.getCellByPosition(6,0).String = ("Quantità\nNegative")
        oSheet.getCellByPosition(7,0).String = ("Prezzo\nunitario")
        oSheet.getCellByPosition(8,0).String = ("Importo\ndebito")
        oSheet.getCellByPosition(9,0).String = ("Importo\npagamento")
        # ~oSheet.getCellByPosition(10,0).String = ("Num.\nPag.")
    # larghezza colonne
        oSheet.getCellByPosition(0,0).Columns.Width = 1600 #'N. ord.
        oSheet.getCellByPosition(1,0).Columns.Width = 7500 #'LAVORAZIONI
        oSheet.getCellByPosition(2,0).Columns.Width = 650 #'Lib.N.
        oSheet.getCellByPosition(3,0).Columns.Width = 650 #'Lib.P.
        oSheet.getCellByPosition(4,0).Columns.Width = 1000 #'U.M.
        oSheet.getCellByPosition(5,0).Columns.Width = 1600 #'Positive
        oSheet.getCellByPosition(6,0).Columns.Width = 1600 #'Negative
        oSheet.getCellByPosition(7,0).Columns.Width = 1400 #'Prezzo
        oSheet.getCellByPosition(8,0).Columns.Width = 1950 #'debito
        oSheet.getCellByPosition(9,0).Columns.Width = 1950 #'pagamento
        # ~oSheet.getCellByPosition(0, 2).Rows.OptimalHeight = True
        # ~oSheet.getCellByPosition(10,0).Columns.OptimalWidth = True #'n.pag.
        insRow = 1 #'prima riga inserimento in Registro
    except Exception as e:
        # ~ DLG.errore(e)

        # recupera il registro precedente
        PL.GotoSheet('Registro')
        oSheet= oDoc.Sheets.getByName("Registro")
        # ~DLG.chi("_Reg_" + str(nSal - 1))
        oRanges = oDoc.NamedRanges
        oPrevRange = oRanges.getByName("_Reg_" + str(nSal - 1)).ReferredCells.RangeAddress

        fRow = oPrevRange.StartRow
        lRow = oPrevRange.EndRow
        insRow = oPrevRange.EndRow + 1

    indicator.setValue(2)

# compilo il Registro
    reg =[]
    for el in REG:
        if el not in reg:
            reg.append(el)
    lastRow = insRow + len(reg) -1
    oRange = oSheet.getCellRangeByPosition(0, insRow, 8, insRow + len(reg) - 1)
    oRange.setDataArray(tuple(reg))
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

# do gli stili al Registro
    oSheet.getCellRangeByPosition(0, insRow, 1, lastRow).CellStyle = "List-stringa-sin"
    oSheet.getCellRangeByPosition(2, insRow, 4, lastRow).CellStyle = "List-num-centro"
    oSheet.getCellRangeByPosition(5, insRow, 6, lastRow).CellStyle = "comp 1a"
    oSheet.getCellRangeByPosition(7, insRow, 9, lastRow).CellStyle = "List-num-euro"

# inserisco la prima riga GIALLA nel REGISTRO
    oSheet.getRows().insertByIndex(insRow, 1)
    oSheet.getCellRangeByPosition (0, insRow, 9, insRow).CellStyle = "uuuuu"
    PL.fissa()
    # ci metto le informazioni
    oSheet.getCellByPosition(1, insRow).String = "segue Registro n." + str(nSal) + " - " + str(daVoce) + "÷" + str(aVoce)
    oSheet.getCellByPosition(2, insRow).Value= nSal        #numero libretto
    oSheet.getCellByPosition(3, insRow).Value = REG[-1][3] #ultimo numero pagina
    # indico il parziale del SAL relativo:
    oSheet.getCellByPosition(8, insRow).Formula = (
        "=SUBTOTAL(9;I" + str(insRow +2) + ":I" + str(lastRow +2) + ")")
    oSheet.getCellByPosition(8, insRow).CellStyle = "comp sotto Euro 3_R"

    # RIGA RIPORTO
    insRow += 1
    oSheet.getRows().insertByIndex(insRow, 1)
    oSheet.getCellByPosition(1, insRow).String = "R I P O R T O"
    #debito
    oSheet.getCellByPosition(8, insRow).Formula = (
        '=IF(SUBTOTAL(9;$I$2:$I$' + str(insRow) + ')=0;"";SUBTOTAL(9;$I$2:$I$' + str(insRow))
    #pagamento
    oSheet.getCellByPosition(9, insRow).Formula = (
        '=IF(SUBTOTAL(9;$J$2:$J$' + str(insRow) + ')=0;"";SUBTOTAL(9;$J$2:$J$' + str(insRow))

    oSheet.getCellRangeByPosition (0, insRow, 9, insRow).CellStyle = "Ultimus_Bordo_sotto"

    insRow += 1
    oSheet.getRows().insertByIndex(insRow, 1)

    oSheet.getCellByPosition(1, insRow).String = "LAVORI A MISURA"
    oSheet.getCellRangeByPosition(0, insRow, 9, insRow).CellStyle = "Ultimus_centro_bordi_lati"
    PL._gotoCella(1, insRow)

    lastRow = insRow + len(reg)

    inizioFirme = lastRow + 5
    firme_contabili (inizioFirme) # riga di inserimento
    fineFirme = inizioFirme + 18

    indicator.setValue(3)

# set area del REGISTRO
    area="$A$" + str(insRow) + ":$J$" + str(fineFirme + 1)
    nomearea = "_Reg_" + str(nSal)
    LeenoBasicBridge.rifa_nomearea(oDoc, "Registro", area , nomearea)

    oRanges = oDoc.NamedRanges
    oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

    #range del _Reg_
    # ~daRiga = oNamedRange.StartRow
    # ~aRiga = oNamedRange.EndRow
    # ~daColonna = oNamedRange.StartColumn
    # ~aColonna = oNamedRange.EndColumn

    LS.setPageStyle()
    oSheet.PageStyle = 'PageStyle_REGISTRO_A4'
    oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName('PageStyle_REGISTRO_A4')
    committente = "Committente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String + \
    '\nRegistro di Contabilità n.' + str(nSal)
    LS.set_header(oAktPage, committente, '', '')
    LS.npagina()
    
    iSheet = oSheet.RangeAddress.Sheet

    # imposta riga da ripetere
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet
    # ~oTitles.StartColumn = 0
    oTitles.StartRow = 0
    # ~oTitles.EndColumn = 9
    # ~oTitles.EndRow = 0
    oSheet.setTitleRows(oTitles)
    oSheet.setPrintAreas((oNamedRange,))
    oSheet.setPrintTitleRows(True)

    # ~oPrintArea = oSheet.getPrintAreas()
    # ~oSheet.group(oPrintArea[0], 1)

    oSheet.getCellRangeByPosition(0, lastRow +1, 9, fineFirme).CellStyle = "Ultimus_centro_bordi_lati"

    #torno su a completare...
    oSheet.getCellByPosition(1, lastRow + 2).String = "Parziale dei Lavori a Misura €"
    oSheet.getCellByPosition(1, lastRow + 2).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(8, lastRow + 2).Formula = (
        '=SUBTOTAL(9;$I$2:$I$' + str(inizioFirme))
    oSheet.getCellByPosition(8, lastRow + 2).CellStyle = "Ultimus_destra_totali"

    # oSheet.getCellByPosition(1, lastRow + 4).String = 'Lavori a tutto il ' + PL.oggi() + ' - T O T A L E   €'
    oSheet.getCellByPosition(1, lastRow + 4).String = 'Lavori a tutto il ___/___/_________ - T O T A L E   €'
    oSheet.getCellByPosition(1, lastRow + 4).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(8, lastRow + 4).Formula = (
        '=SUBTOTAL(9;$I$2:$I$' + str(inizioFirme))
    oSheet.getCellByPosition(8, lastRow + 4).CellStyle = "Ultimus_destra_totali"

    #applico gli stili corretti ad alcuni dati della firma
    oSheet.getCellByPosition(1, lastRow + 6).CellStyle = "Ultimus_destra"

    #riga del certificato di pagamento
    oSheet.getCellByPosition(1, lastRow + 16).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(9, lastRow + 16).CellStyle = "Comp-Bianche in mezzo bordate_R"
    oSheet.getCellByPosition(9, lastRow + 16).String = "inserisci qui il CP"

    indicator.setValue(4)

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    x = fineFirme
    #=reg===================
    # ~ insrow()

    fineFirme = SheetUtils.getLastUsedRow(oSheet)

    oSheet.getCellByPosition(1,fineFirme -1).String = ""
    oSheet.getCellByPosition(1,fineFirme).String = "A   R I P O R T A R E"
    oSheet.getCellByPosition(8, fineFirme).Formula = ('=IF(SUBTOTAL(9;$I$2:$I$' + str(fineFirme) + ')=0;"";SUBTOTAL(9;$I$2:$I$' + str(fineFirme))
    oSheet.getCellByPosition(9, fineFirme).Formula = ('=IF(SUBTOTAL(9;$J$2:$J$' + str(fineFirme) + ')=0;"";SUBTOTAL(9;$J$2:$J$' + str(fineFirme))
    oSheet.getCellRangeByPosition (0, fineFirme, 9, fineFirme).CellStyle = "Ultimus_Bordo_sotto"

    indicator.setValue(5)
    indicator.end()

# ~def GeneraSAL (oDoc):
    indicator = oDoc.getCurrentController().getStatusIndicator()
    if indicator is not None:
        indicator.start("Generazione Stato di Avanzamento Lavori...", 8)
    indicator.setValue(1)

    try:
        oDoc.getSheets().insertNewByName('SAL',oSheet.RangeAddress.Sheet + 1)
        PL.GotoSheet('SAL')
        oSheet = oDoc.Sheets.getByName('SAL')

    # riga di intestazione
        oSheet.getCellRangeByPosition(0,0,6,0).CellStyle="comp Int_colonna_R"
        oSheet.getCellByPosition(0,0).String = ("N. ord.\nArticolo")
        oSheet.getCellByPosition(1,0).String = ("LAVORAZIONI\nE SOMMINISTRAZIONI")
        oSheet.getCellByPosition(2,0).String = ("U.M.")
        oSheet.getCellByPosition(3,0).String = ("Quantità")
        oSheet.getCellByPosition(4,0).String = ("Prezzo\nunitario")
        oSheet.getCellByPosition(5,0).String = ("Importo")
        oSheet.getCellByPosition(6,0).String = ("Pagine")
    # larghezza colonne
        oSheet.getCellByPosition(0,0).Columns.Width = 1600 #'N. ord.
        oSheet.getCellByPosition(1,0).Columns.Width = 11050 #'LAVORAZIONI
        oSheet.getCellByPosition(2,0).Columns.Width = 1500 #'U.M.
        oSheet.getCellByPosition(3,0).Columns.Width = 1800 #'Quantità
        oSheet.getCellByPosition(4,0).Columns.Width = 1400 #'Prezzo
        oSheet.getCellByPosition(5,0).Columns.Width = 1900 #'Importo
        oSheet.getCellByPosition(6,0).Columns.OptimalWidth = True #'n.pag.
        oSheet.getCellByPosition(0, 2).Rows.OptimalHeight = True
        insRow = 1 #'prima riga inserimento in Registro
    except Exception as e:
        # DLG.errore(e)

        # recupera il registro precedente
        PL.GotoSheet('SAL')
        oSheet= oDoc.Sheets.getByName("SAL")
        oRanges = oDoc.NamedRanges
        oPrevRange = oRanges.getByName("_SAL_" + str(nSal - 1)).ReferredCells.RangeAddress

        fRow = oPrevRange.StartRow
        lRow = oPrevRange.EndRow
        insRow = oPrevRange.EndRow + 1

    indicator.setValue(2)

    # compilo il SAL
    lastRow = insRow + len(datiSAL) -1
    oRange = oSheet.getCellRangeByPosition(0, insRow, 3, lastRow)

    sal = tuple(datiSAL)

    oRange.setDataArray(sal)

    formule = []
    for x in range(insRow, lastRow + 1):
        formule.append(['=VLOOKUP(A' + str(x + 1) + ';elenco_prezzi;5;FALSE())',
            '=IF(C' + str(x + 1) + '="%";D' + str(x + 1) + '*E' + str(x + 1) + '/100;D' + str(x + 1) + '*E' + str(x + 1) + ')'])

    indicator.setValue(3)

# do gli stili al SAL
    oSheet.getCellRangeByPosition(0, insRow, 1, lastRow).CellStyle = "List-stringa-sin"
    oSheet.getCellRangeByPosition(2, insRow, 2, lastRow).CellStyle = "List-num-centro"
    oSheet.getCellRangeByPosition(3, insRow, 3, lastRow).CellStyle = "comp 1a"
    oSheet.getCellRangeByPosition(4, insRow, 5, lastRow).CellStyle = "List-num-euro"

# completo il SAL inserendo le formule
    oRange = oSheet.getCellRangeByPosition(4, insRow, 5, lastRow)
    formule = tuple(formule)
    oRange.setFormulaArray(formule)

    nOrd = 1
    for x in range(insRow, lastRow + 1):
        oSheet.getCellByPosition(4, x).Value = oSheet.getCellByPosition(4, x).Value
        oSheet.getCellByPosition(0, x).String = str(nOrd) \
            + '\n' + oSheet.getCellByPosition(0, x).String
        nOrd += 1
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    indicator.setValue(4)

# inserisco la prima riga GIALLA nel SAL
    oSheet.getRows().insertByIndex(insRow, 1)
    oSheet.getCellRangeByPosition (0, insRow, 9, insRow).CellStyle = "uuuuu"
    PL.fissa()
    # ci metto le informazioni
    # ~ oSheet.getCellByPosition(1, insRow).String = "segue Stato di Avanzamento Lavori n." + str(nSal) + " - " + str(daVoce) + "÷" + str(aVoce)
    oSheet.getCellByPosition(1, insRow).String = "segue Stato di Avanzamento Lavori n." + str(nSal) + " - 1÷" + str(aVoce)
    # ~oSheet.getCellByPosition(2, insRow).Value = nSal        #numero libretto

    # parziale del SAL relativo:
    oSheet.getCellByPosition(5, insRow).Formula = (
        "=SUBTOTAL(9;$F$" + str(insRow +2) + ":F" + str(lastRow +2) + ")")
    oSheet.getCellByPosition(5, insRow).CellStyle = "comp sotto Euro 3_R"

    lastRow = insRow + len(datiSAL)

#torno su a completare...
    oSheet.getCellByPosition(1, lastRow + 2).String = "Parziale dei Lavori a Misura €"
    oSheet.getCellByPosition(5, lastRow + 2).Formula = (
        '=SUBTOTAL(9;$F$' + str(insRow) + ':$F$' + str(lastRow + 2))
    rigaPsal = lastRow + 2

    oSheet.getCellByPosition(1, lastRow + 4).String = 'Lavori a tutto il ' + PL.oggi() + ' - T O T A L E   €'
    oSheet.getCellByPosition(5, lastRow + 4).Formula = (
        '=SUBTOTAL(9;$F$' + str(insRow) + ':$F$' + str(lastRow + 2))

    indicator.setValue(5)

    PL._gotoCella(0, lastRow)

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    fineFirme = lastRow + 5

    LS.setPageStyle()
    oSheet.PageStyle = 'PageStyle_SAL_A4'
    oAktPage = oDoc.StyleFamilies.getByName('PageStyles').getByName('PageStyle_SAL_A4')
    committente = "Committente: " + oDoc.getSheets().getByName('S2').getCellRangeByName("C6").String + \
    "\nStato di Avanzamento Lavori n." + str(nSal) + " a tutto il " + PL.oggi()
    LS.set_header(oAktPage, committente, '', '')
    LS.npagina()
    
    # ~ oHeader = oAktPage.RightPageHeaderContent
    # ~oAktPage.PageScale = 95
    # ~ oHLText = oHeader.LeftText.Text.String = committente
    # ~ oHRText = oHeader.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oHRText = oHeader.LeftText.Text.Text.CharHeight = htxt

    # ~ oHLText = oHeader.CenterText.Text.String = oggetto
    # ~ oHRText = oHeader.CenterText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oHRText = oHeader.CenterText.Text.Text.CharHeight = htxt

    # ~ oHRText = oHeader.RightText.Text.String = '{Page}'
    # ~ oHRText = oHeader.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oHRText = oHeader.RightText.Text.Text.CharHeight = htxt

    # ~ oAktPage.RightPageHeaderContent = oHeader
    # ~FOOTER
    # ~ oFooter = oAktPage.RightPageFooterContent
    # ~ oHLText = oFooter.CenterText.Text.String = ''
    # ~ nomefile = oDoc.getURL().replace('%20',' ')
    # ~ oHLText = oFooter.LeftText.Text.String = "\nrealizzato con LeenO: " + os.path.basename(nomefile)
    # ~ oHLText = oFooter.LeftText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oHLText = oFooter.LeftText.Text.Text.CharHeight = htxt * 0.5
    # ~ oHLText = oFooter.RightText.Text.Text.CharFontName = 'Liberation Sans Narrow'
    # ~ oHLText = oFooter.RightText.Text.Text.CharHeight = htxt

# set area di stampa del SAL
    area="$A$" + str(insRow + 2) + ":$F$" + str(fineFirme + 1)
    nomearea = "_SAL_" + str(nSal)
    LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area , nomearea)

    oRanges = oDoc.NamedRanges
    oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

# imposta riga da ripetere con area di stampa
    iSheet = oSheet.RangeAddress.Sheet
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet
    oTitles.StartRow = 0
    oSheet.setTitleRows(oTitles)
    oSheet.setPrintAreas((oNamedRange,))
    oSheet.setPrintTitleRows(True)

    # =sal===================
    # ~ insrow()
    fineFirme = SheetUtils.getLastUsedRow(oSheet) +1

#applico gli stili corretti ad alcuni dati della firma
    oSheet.getCellRangeByPosition(
        0, lastRow +1, 5, fineFirme + 1).CellStyle = "Ultimus_centro_bordi_lati"
    # ~ oSheet.getCellRangeByPosition (0, fineFirme, 5, fineFirme).CellStyle = "comp Descr"
    oSheet.getCellByPosition(1, lastRow + 2).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(1, lastRow + 4).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(5, lastRow + 2).CellStyle = "Ultimus_destra_totali"
    oSheet.getCellByPosition(5, lastRow + 4).CellStyle = "Ultimus_destra_totali"

    lastRow = fineFirme + 2

    oSheet.getCellByPosition(1, lastRow + 1).String = "R I E P I L O G O   S A L"
    oSheet.getCellByPosition(1, lastRow + 3).String = ("Appalto: a misura")
    oSheet.getCellByPosition(1, lastRow + 4).String = ("Offerta: unico ribasso")
    oSheet.getCellByPosition(1, lastRow + 6).String = ("Lavori a Misura €")
    oSheet.getCellByPosition(5, lastRow + 6).Formula = "=$F$" + str(rigaPsal + 1)
    oSheet.getCellByPosition(1, lastRow + 7).String = ("Di cui importo per la Sicurezza")
    oSheet.getCellByPosition(5, lastRow + 7).Value = - sic
    oSheet.getCellByPosition(1, lastRow + 8).String = ("Di cui importo per la Manodopera")
    oSheet.getCellByPosition(5, lastRow + 8).Value = - mdo
    oSheet.getCellByPosition(1, lastRow + 9).String =  "Importo dei Lavori a Misura su cui applicare il ribasso"
    oSheet.getCellByPosition(5, lastRow + 9).Formula = "=SUM(F" + str(lastRow + 7) + ":F" + str(lastRow + 9) + ")"
    oSheet.getCellByPosition(1, lastRow + 10).Formula = (
    '''=CONCATENATE("RIBASSO del ";TEXT($S2.$C$78*100;"#.##0,000");"%")''')
    oSheet.getCellByPosition(5, lastRow + 10).Formula = "=-$F$" + str(lastRow + 10) + "*$S2.$C$78" # RIBASSO
    oSheet.getCellByPosition(1, lastRow + 11).String = ("Importo per la Sicurezza")
    oSheet.getCellByPosition(5, lastRow + 11).Value = sic
    oSheet.getCellByPosition(1, lastRow + 12).String = ("Importo per la Manodopera")
    oSheet.getCellByPosition(5, lastRow + 12).Value = mdo
    oSheet.getCellByPosition(1, lastRow + 13).String =  "PER I LAVORI A MISURA €"
    oSheet.getCellByPosition(5, lastRow + 13).Formula = "=SUM($F$" + str(lastRow + 10) + ":$F$" + str(lastRow + 13) + ")"

# IL TOTALE ANDRA' RISISTEMATO QUANDO SARANNO PRONTE LE ALTRE MODALITA' DI APPALTO: IN ECONOMIA E A CORPO
    oSheet.getCellByPosition(1, lastRow + 15).String =  "T O T A L E  €"
    oSheet.getCellByPosition(5, lastRow + 15).Formula = "=SUM($F$" + str(lastRow + 10) + ":$F$" + str(lastRow + 13) + ")"

# set area di stampa del SAL
    area="$A$" + str(insRow + 2) + ":$F$" + str(lastRow + 18)
    nomearea = "_SAL_" + str(nSal)
    LeenoBasicBridge.rifa_nomearea(oDoc, "SAL", area , nomearea)
    oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

    indicator.setValue(6)

# ~# le firme
    inizioFirme = lastRow + 17
    firme_contabili (inizioFirme) # riga di inserimento
    fineFirme = inizioFirme + 12

    oSheet.getCellRangeByPosition(
        0, lastRow, 5, fineFirme +2).CellStyle = "Ultimus_centro_bordi_lati"
    oSheet.getCellByPosition(1, lastRow + 1).CellStyle = "Ultimus_centro_Dsottolineato"
    oSheet.getCellRangeByPosition(1, lastRow + 3, 1, lastRow + 4).CellStyle = "Ultimus_sx_italic"
    oSheet.getCellRangeByPosition (5, lastRow + 6,5, lastRow + 15).CellStyle = "ULTIMUS"
    oSheet.getCellByPosition(1, lastRow + 6).CellStyle = "Ultimus_sx_bold"
    oSheet.getCellRangeByPosition(1, lastRow + 7, 1, lastRow + 8).CellStyle = "Ultimus_sx"
    oSheet.getCellByPosition(5, lastRow + 8).CellStyle = "Ultimus_"
    oSheet.getCellRangeByPosition(1, lastRow + 9, 1, lastRow + 10).CellStyle = "Ultimus_destra"
    oSheet.getCellRangeByPosition(1, lastRow + 11, 1, lastRow + 12).CellStyle = "Ultimus_sx"
    oSheet.getCellByPosition(5, lastRow + 12).CellStyle = "Ultimus_"
    oSheet.getCellRangeByPosition(1, lastRow + 13, 1, lastRow + 13).CellStyle = "Ultimus_destra_bold"
    oSheet.getCellRangeByPosition(1, lastRow + 15, 1, lastRow + 15).CellStyle = "Ultimus_destra_bold"
    oSheet.getCellByPosition(5, lastRow + 15).CellStyle = "Ultimus_destra_totali"

    LeenoSheetUtils.adattaAltezzaRiga(oSheet)

    indicator.setValue(7)

    #ridefinisci area di stampa
    oSheet.setPrintAreas((oNamedRange,))

    #=sal===================
    # ~ insrow()
    indicator.setValue(8)
    indicator.end()

    return


########################################################################


def insrow():
    """
    Inserisce righe nel foglio attivo basandosi su ultima area nominata
    e altezza della pagina.

    Aggiunge righe finché l'altezza della pagina non viene superata.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oRanges = oDoc.NamedRanges

    nSh = {
        'CONTABILITA': '_Lib_',
        'Registro': '_Reg_',
        'SAL' : '_SAL_'
    }

    nome = nSh.get(oSheet.Name)
    nSal = ultimo_sal()[-1]
    nomearea = nome + str(nSal)

    if oRanges.hasByName(nomearea):
        oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
        sRow = oNamedRange.StartRow
        iRow = oNamedRange.EndRow
        # ~ return [iRow, iRow - sRow]

    if oSheet.Name == 'CONTABILITA':
        col = 2
    else:
        col = 1
    hattuale = oSheet.getCellByPosition(col, iRow).Position.Y - \
    oSheet.getCellByPosition(col, sRow).Position.Y
    
    if oSheet.Name == 'CONTABILITA':
        hpagina = (len(oSheet.RowPageBreaks) - 1) * 25510
    elif oSheet.Name == 'Registro':
        hpagina = (len(oSheet.RowPageBreaks) - 1) * 25810
    elif oSheet.Name == 'SAL':
        hpagina = (len(oSheet.RowPageBreaks) - 1) * 25850


    for i in range(50):
        oSheet.getRows().insertByIndex(iRow, 1)
        oSheet.getCellByPosition(col, iRow).String = '––––––––––––––––––––––––––––––' #+ str(i)
        iRow += 1
        # Verifica se la cella supera l'altezza pagina e interrompe il ciclo se necessario
        hattuale = oSheet.getCellByPosition(col, iRow).Position.Y - \
        oSheet.getCellByPosition(col, sRow).Position.Y

        # ~ DLG.chi(f'hattuale: {hattuale}\nhpagina: {hpagina}')

        if hattuale >= hpagina:
            break
    return
def insrow():
    """
    Inserisce righe nel foglio attivo basandosi su ultima area nominata
    e altezza della pagina.

    Aggiunge righe finché l'altezza della pagina non viene superata.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oRanges = oDoc.NamedRanges

    nSh = {
        'CONTABILITA': '_Lib_',
        'Registro': '_Reg_',
        'SAL': '_SAL_'
    }

    # Costanti altezze pagina
    hPage = {
        'CONTABILITA': 25510,
        'Registro': 25810,
        'SAL': 25850
    }

    nome = nSh.get(oSheet.Name)
    if not nome:
        return  # foglio non gestito

    nSal = ultimo_sal()[-1]
    nomearea = nome + str(nSal)

    if not oRanges.hasByName(nomearea):
        return  # nessuna area trovata

    oNamedRange = oRanges.getByName(nomearea).ReferredCells.RangeAddress
    sRow = oNamedRange.StartRow
    iRow = oNamedRange.EndRow

    # Colonna di riferimento
    col = 2 if oSheet.Name == 'CONTABILITA' else 1

    # Altezza disponibile
    hpagina = (len(oSheet.RowPageBreaks) - 1) * hPage[oSheet.Name]

    # Linea di riempimento
    filler = "––––––––––––––––––––––––––––––"

    for _ in range(50):
        oSheet.getRows().insertByIndex(iRow, 1)
        oSheet.getCellByPosition(col, iRow).String = filler
        iRow += 1

        hattuale = (
            oSheet.getCellByPosition(col, iRow).Position.Y -
            oSheet.getCellByPosition(col, sRow).Position.Y
        )

        if hattuale >= hpagina:
            break


########################################################################
def firme_contabili(lrowF=None):
    """
    Inserisce i dati necessari alle firme nel foglio "CONTABILITA",
    con spaziatura uniforme tra i blocchi.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet_S2 = oDoc.getSheets().getByName("S2")

    # Ricava il luogo dall'intestazione del foglio S2
    luogo_raw = oSheet_S2.getCellRangeByName("$S2.C4").String
    ultimo_token = luogo_raw.split(" ")[-1] if luogo_raw else ""
    luogo = f"{ultimo_token}, " if ultimo_token else "Data, "

    if oSheet.Name != "CONTABILITA":
        return

    # Se non viene passata una riga, calcola l'ultima disponibile
    if lrowF is None:
        lrowF = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 1

    firme = []

    # Progettista (luogo + data)
    firme.append(f"{luogo} ___/___/_________")

    # Impresa esecutrice
    impresa = oSheet_S2.getCellRangeByName("$S2.C17").String
    firme.append(f"L'Impresa esecutrice\n({impresa})")

    # Direttore Operativo Contabile (solo se presente)
    contabile = oSheet_S2.getCellRangeByName("$S2.C14").String
    if contabile:
        firme.append(f"Il Direttore Operativo Contabile\n({contabile})")

    # CSE (solo se presente)
    cse = oSheet_S2.getCellRangeByName("$S2.C15").String
    if cse:
        firme.append(f"Visto: il C.S.E.\n({cse})")

    # Direttore Lavori
    direttore = oSheet_S2.getCellRangeByName("$S2.C16").String
    firme.append(f"Il Direttore dei Lavori\n({direttore})")

    # Numero righe da inserire = blocchi × 3
    oSheet.getRows().insertByIndex(lrowF, len(firme) * 3)

    riga_corrente = lrowF + 1
    for firma in firme:
        oSheet.getCellByPosition(2, riga_corrente).Formula = firma
        riga_corrente += 3 # avanza sempre di 3 righe

    oSheet.getRows().insertByIndex(riga_corrente -2, 3)

    return riga_corrente +1


########################################################################


def GeneraAttiContabili():
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
g_exportedScripts = GeneraAttiContabili
