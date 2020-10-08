'''
Funzioni relative alla gestione delle analisi di prezzi
'''

import pyleeno as PL
import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoEvents

def inizializzaAnalisi(oDoc):
    '''
    Se non presente, crea il foglio 'Analisi di Prezzo' ed inserisce la prima scheda
    Ritorna l'oggetto oSheet del foglio contenente le analisi
    e la riga da cui iniziare la compilazione dell'analisi corrente
    '''
    SheetUtils.NominaArea(oDoc, 'S5', '$B$108:$P$133', 'blocco_analisi')
    if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
        oDoc.getSheets().insertNewByName('Analisi di Prezzo', 1)
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oSheet.getCellRangeByPosition(0, 0, 15, 0).CellStyle = 'Analisi_Sfondo'
        oSheet.getCellByPosition(0, 1).Value = 0
        oSheet.TabColor = 12189608
        oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oCellAddress = oSheet.getCellByPosition(0, SheetUtils.getLastUsedRow(oSheet)).getCellAddress()

        # questa è l' UNICA funzione che non può prescindere dal controller
        # probabilmente una dimenticanza degli sviluppatori di LO
        # controllare se in seguito cambierà qualcosa...
        LeenoSheetUtils.setLarghezzaColonne(oSheet)

        # la riga dalla quale iniziare a scrivere
        startRow = 2

        LeenoEvents.assegna()
        LeenoSheetUtils.ScriviNomeDocumentoPrincipaleInFoglio(oSheet)

    else:
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')

        lrow = LeenoSheetUtils.cercaUltimaVoce(oSheet) - 5
        urow = SheetUtils.getLastUsedRow(oSheet)
        for n in range(lrow, urow):
            if oSheet.getCellByPosition(0, n).CellStyle == 'An-sfondo-basso Att End':
                break
        oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oSheet.getRows().insertByIndex(n + 2, 26)
        oCellAddress = oSheet.getCellByPosition(0, n + 2).getCellAddress()

        # la riga dalla quale iniziare a scrivere
        startRow = n + 2 + 1

    oSheet.copyRange(oCellAddress, oRangeAddress)
    LeenoSheetUtils.inserisciRigaRossa(oSheet)

    return oSheet, startRow


def circoscriveAnalisi(oSheet, lrow):
    '''
    lrow    { int }  : riga di riferimento per
                        la selezione dell'intera voce
    Circoscrive una voce di analisi
    partendo dalla posizione corrente del cursore
    '''
    stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_analisi:
        for el in reversed(range(0, lrow)):
            #  chi(oSheet.getCellByPosition(0, el).CellStyle)
            if oSheet.getCellByPosition(0, el).CellStyle == 'Analisi_Sfondo':
                SR = el
                break
        for el in range(lrow, SheetUtils.getLastUsedRow(oSheet)):
            if oSheet.getCellByPosition(0, el).CellStyle == 'An-sfondo-basso Att End':
                ER = el
                break
    celle = oSheet.getCellRangeByPosition(0, SR, 250, ER)
    return celle


def copiaRigaAnalisi(oSheet, lrow):
    '''
    Inserisce una nuova riga di misurazione in analisi di prezzo
    '''
    stile = oSheet.getCellByPosition(0, lrow).CellStyle
    if stile in ('An-lavoraz-desc', 'An-lavoraz-Cod-sx'):
        lrow = lrow + 1
        oSheet.getRows().insertByIndex(lrow, 1)
        # imposto gli stili
        oSheet.getCellByPosition(0, lrow).CellStyle = 'An-lavoraz-Cod-sx'
        oSheet.getCellRangeByPosition(1, lrow, 5, lrow).CellStyle = 'An-lavoraz-generica'
        oSheet.getCellByPosition(3, lrow).CellStyle = 'An-lavoraz-input'
        oSheet.getCellByPosition(6, lrow).CellStyle = 'An-senza'
        oSheet.getCellByPosition(7, lrow).CellStyle = 'An-senza-DX'
        # ci metto le formule
        #  oDoc.enableAutomaticCalculation(False)
        oSheet.getCellByPosition(1, lrow).Formula = (
           '=IF(A' + str(lrow + 1) +
           '="";"";CONCATENATE("  ";VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;2;FALSE());' '))')
        oSheet.getCellByPosition(2, lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";"";VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;3;FALSE()))')
        oSheet.getCellByPosition(3, lrow).Value = 0
        oSheet.getCellByPosition(4,lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";0;VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;5;FALSE()))')
        oSheet.getCellByPosition(5, lrow).Formula = (
           '=D' + str(lrow + 1) + '*E' + str(lrow + 1))
        oSheet.getCellByPosition(8, lrow).Formula = (
           '=IF(A' + str(lrow + 1) + '="";"";IF(VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;6;FALSE())="";"";(VLOOKUP(A' +
           str(lrow + 1) + ';elenco_prezzi;6;FALSE()))))')
        oSheet.getCellByPosition(9, lrow).Formula = (
           '=IF(I' + str(lrow + 1) + '="";"";I' +
           str(lrow + 1) + '*F' + str(lrow + 1) + ')')
        if oSheet.getCellByPosition(1, lrow - 1).CellStyle == 'An-lavoraz-dx-senza-bordi':
            oRangeAddress = oSheet.getCellByPosition(0, lrow + 1).getRangeAddress()
            oCellAddress = oSheet.getCellByPosition(0, lrow).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        oSheet.getCellByPosition(0, lrow).String = 'Cod. Art.?'

