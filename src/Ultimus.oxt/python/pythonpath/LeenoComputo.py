from com.sun.star.table import CellRangeAddress
import SheetUtils
import LeenoUtils
import LeenoSheetUtils
import LeenoConfig
import LeenoDialogs as DLG

import pyleeno as PL

def datiVoceComputo (oSheet, lrow):
    '''
    Ricava i dati dalla voce di COMPUTO / CONTABILITA
    '''

    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~lrow = PL.LeggiPosizioneCorrente()[1]

    sStRange = circoscriveVoceComputo(oSheet, lrow)
    i = sStRange.RangeAddress.StartRow
    f = sStRange.RangeAddress.EndRow
    # ~DLG.chi((i, f))
    # ~return
    num      = oSheet.getCellByPosition(0,  i+1).String
    art      = oSheet.getCellByPosition(1,  i+1).String 
    desc     = oSheet.getCellByPosition(2,  i+1).String
    quantP   = oSheet.getCellByPosition(9,    f).Value
    mdo      = oSheet.getCellByPosition(30,   f).Value
    sic      = oSheet.getCellByPosition(17,   f).Value
    voce = []
    REG = []
    SAL = []
    if oSheet.Name in ('CONTABILITA'):
        quantN = ''
        if quantP < 0:
            quantN = quantP
            quantP = ''
        data     = oSheet.getCellByPosition(1,  i+2).String
        um       = oSheet.getCellByPosition(9,  i+1).String
        Nlib     = int(oSheet.getCellByPosition(19, i+1).Value)
        Plib     = int(oSheet.getCellByPosition(20, i+1).Value)
        flag     = oSheet.getCellByPosition(22, i+1).String
        nSal     = int(oSheet.getCellByPosition(23, i+1).Value)
        prezzo   = oSheet.getCellByPosition(13,   f).Value
        importo  = oSheet.getCellByPosition(15,   f).Value
        sic  = oSheet.getCellByPosition(17,   f).Value
        mdo  = oSheet.getCellByPosition(30,   f).Value

        REG = ((num + '\n' + art + '\n' + data), desc, Nlib, Plib, um, quantP,
            quantN, prezzo, importo)#, sic, mdo, flag, nSal)
        if quantP != '':
            quant = quantP
        else:
            quant = quantN
        SAL = (art,  desc, um, quant, prezzo, importo, sic, mdo)
        return REG, SAL
    elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
        um = oSheet.getCellByPosition(8, f).String.split('[')[-1].split('[')[0]
        prezzo   = oSheet.getCellByPosition(11,   f).Value
        importo  = oSheet.getCellByPosition(18,   f).Value
        voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)
        return voce


def circoscriveVoceComputo(oSheet, lrow):
    '''
    lrow    { int }  : riga di riferimento per
                        la selezione dell'intera voce

    Circoscrive una voce di COMPUTO, VARIANTE o CONTABILITÀ
    partendo dalla posizione corrente del cursore
    '''
    # li predefinisco... @@@
    lrowS = lrow
    lrowE = lrow

    #  lrow = LeggiPosizioneCorrente()[1]
    #  if oSheet.Name in('VARIANTE', 'COMPUTO','CONTABILITA'):
    if oSheet.getCellByPosition(0, lrow).CellStyle in (
       'comp progress', 'comp 10 s',
       'Comp Start Attributo', 'Comp End Attributo',
       'Comp Start Attributo_R', 'comp 10 s_R',
       'Comp End Attributo_R', 'Livello-0-scritta',
       'Livello-1-scritta', 'livello2 valuta'):
        y = lrow
        while oSheet.getCellByPosition(0, y).CellStyle not in ('Comp End Attributo', 'Comp End Attributo_R'):
            y += 1
        lrowE = y
        y = lrow
        try:
            while oSheet.getCellByPosition(0, y).CellStyle not in ('Comp Start Attributo', 'Comp Start Attributo_R'):
                y -= 1
        except:
            return
        lrowS = y
    #trova il range di firme in CONTABILITA
    elif oSheet.getCellByPosition(0, lrow).CellStyle == 'Ultimus_centro_bordi_lati':
        for y in reversed (range(0, lrow)):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                lrowS = y + 1
                break
        for y in range(lrow, SheetUtils.getLastUsedRow(oSheet)):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                lrowE = y - 1 
                break
    elif 'ULTIMUS' in oSheet.getCellByPosition(0, lrow).CellStyle:
        lrowS = LeenoSheetUtils.cercaUltimaVoce(oSheet) +2
        lrowE = LeenoSheetUtils.rRow(oSheet) -1
    else:
        return
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

    if oSheet.getCellByPosition(31,lrow - 1).CellStyle in (
       'livello2 valuta',
       'Livello-0-scritta',
       'Livello-1-scritta',
       'compTagRiservato'):
        oSheet.getCellByPosition(31, lrow + 3).Value = oSheet.getCellByPosition(31, lrow - 1).Value
        oSheet.getCellByPosition(32, lrow + 3).Value = oSheet.getCellByPosition(32, lrow - 1).Value
        oSheet.getCellByPosition(33,lrow + 3).Value = oSheet.getCellByPosition(33, lrow - 1).Value


# TROPPO LENTA
def ins_voce_computo(cod=None):
    '''
    cod    { string }
    Se cod è presente, viene usato come codice di voce
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    noVoce = LeenoUtils.getGlobalVar('noVoce')
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')
    lrow = PL.LeggiPosizioneCorrente()[1]
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    if stile in stili_cat:
        lrow += 1
    elif stile  in (noVoce + stili_computo):
        lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)
    else:
        return
    if lrow == 2:
        lrow += 1
    insertVoceComputoGrezza(oSheet, lrow)
    if cod:
        oSheet.getCellByPosition(1, lrow + 1).String = cod
    # @@ PROVVISORIO !!!
    PL._gotoCella(1, lrow + 1)

    #PL.numera_voci(0)
    LeenoSheetUtils.numeraVoci(oSheet, lrow + 1, False)
    if LeenoConfig.Config().read('Generale', 'pesca_auto') == '1':
        PL.pesca_cod()

