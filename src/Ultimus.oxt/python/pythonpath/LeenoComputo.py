import uno
from com.sun.star.table import CellRangeAddress
import SheetUtils
import LeenoUtils
import LeenoSheetUtils
import LeenoConfig

import pyleeno as PL


def circoscriveVoceComputo(oSheet, lrow):
    '''
    lrow    { int }  : riga di riferimento per
                        la selezione dell'intera voce

    Circoscrive una voce di COMPUTO, VARIANTE o CONTABILITÀ
    partendo dalla posizione corrente del cursore
    '''
    #  lrow = LeggiPosizioneCorrente()[1]
    #  if oSheet.Name in('VARIANTE', 'COMPUTO','CONTABILITA'):
    if oSheet.getCellByPosition(
            0,
            lrow).CellStyle in ('comp progress', 'comp 10 s',
                                'Comp Start Attributo', 'Comp End Attributo',
                                'Comp Start Attributo_R', 'comp 10 s_R',
                                'Comp End Attributo_R', 'Livello-0-scritta',
                                'Livello-1-scritta', 'livello2 valuta'):
        y = lrow
        while oSheet.getCellByPosition(
                0, y).CellStyle not in ('Comp End Attributo',
                                        'Comp End Attributo_R'):
            y += 1
        lrowE = y
        y = lrow
        while oSheet.getCellByPosition(
                0, y).CellStyle not in ('Comp Start Attributo',
                                        'Comp Start Attributo_R'):
            y -= 1
        lrowS = y
    celle = oSheet.getCellRangeByPosition(0, lrowS, 250, lrowE)
    return celle

def insertVoceComputoGrezza(oSheet, lrow):

    # lrow = LeggiPosizioneCorrente()[1]
    ########################################################################
    # questo sistema eviterebbe l'uso della sheet S5 da cui copiare i range campione
    # potrei svuotare la S5 ma allungando di molto il codice per la generazione della voce
    # per ora lascio perdere

    #  insRows(lrow,4) #inserisco le righe
    #  oSheet.getCellByPosition(0,lrow).CellStyle = 'Comp Start Attributo'
    #  oSheet.getCellRangeByPosition(0,lrow,30,lrow).CellStyle = 'Comp-Bianche sopra'
    #  oSheet.getCellByPosition(2,lrow).CellStyle = 'Comp-Bianche sopraS'
    #
    #  oSheet.getCellByPosition(0,lrow+1).CellStyle = 'comp progress'
    #  oSheet.getCellByPosition(1,lrow+1).CellStyle = 'comp Art-EP'
    #  oSheet.getCellRangeByPosition(2,lrow+1,8,lrow+1).CellStyle = 'Comp-Bianche in mezzo Descr'
    #  oSheet.getCellRangeByPosition(2,lrow+1,8,lrow+1).merge(True)
    ########################################################################

    oDoc = SheetUtils.getDocumentFromSheet(oSheet)

    # vado alla vecchia maniera ## copio il range di righe computo da S5 ##
    oSheetto = oDoc.getSheets().getByName('S5')

    oRangeAddress = oSheetto.getCellRangeByPosition(0, 8, 42, 11).getRangeAddress()
    oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()

    oSheet.getRows().insertByIndex(lrow, 4)
    oSheet.copyRange(oCellAddress, oRangeAddress)

    # raggruppo i righi di misura
    iSheet = oSheet.RangeAddress.Sheet
    oCellRangeAddr = CellRangeAddress()
    oCellRangeAddr.Sheet = iSheet
    oCellRangeAddr.StartColumn = 0
    oCellRangeAddr.EndColumn = 0
    oCellRangeAddr.StartRow = lrow + 2
    oCellRangeAddr.EndRow = lrow + 2
    oSheet.group(oCellRangeAddr, 1)

    # correggo alcune formule
    oSheet.getCellByPosition(13, lrow + 3).Formula = '=J' + str(lrow + 4)
    oSheet.getCellByPosition(35, lrow + 3).Formula = '=B' + str(lrow + 2)

    if oSheet.getCellByPosition(31,
                                lrow - 1).CellStyle in ('livello2 valuta',
                                                        'Livello-0-scritta',
                                                        'Livello-1-scritta',
                                                        'compTagRiservato'):
        oSheet.getCellByPosition(31,
                                 lrow + 3).Value = oSheet.getCellByPosition(
                                     31, lrow - 1).Value
        oSheet.getCellByPosition(32,
                                 lrow + 3).Value = oSheet.getCellByPosition(
                                     32, lrow - 1).Value
        oSheet.getCellByPosition(33,
                                 lrow + 3).Value = oSheet.getCellByPosition(
                                     33, lrow - 1).Value
    PL._gotoCella(1, lrow + 1)


# TROPPO LENTA
def ins_voce_computo():
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    noVoce = LeenoUtils.getGlobalVar('noVoce')
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    lrow = PL.LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0, lrow).CellStyle in (noVoce + stili_computo):
        lrow = PL.next_voice(lrow, 1)
    else:
        return
    insertVoceComputoGrezza(oSheet, lrow)
    #PL.numera_voci(0)
    LeenoSheetUtils.numeraVoci(oSheet, lrow + 1, False)
    if LeenoConfig.Config().read('Generale', 'pesca_auto') == '1':
        PL.pesca_cod()
