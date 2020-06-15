'''
Funzioni relative alla gestione delle analisi di prezzi
'''
import uno
import SheetUtils
import LeenoSheetUtils

def inizializzaAnalisi(oDoc):
    '''
    Se non presente, crea il foglio 'Analisi di Prezzo' ed inserisce la prima scheda
    Ritorna l'oggetto oSheet del foglio contenente le analisi
    '''
    rifa_nomearea(oDoc, 'S5', '$B$108:$P$133', 'blocco_analisi')
    if not oDoc.getSheets().hasByName('Analisi di Prezzo'):
        oDoc.getSheets().insertNewByName('Analisi di Prezzo', 1)
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oSheet.getCellRangeByPosition(0, 0, 15, 0).CellStyle = 'Analisi_Sfondo'
        oSheet.getCellByPosition(0, 1).Value = 0
        oSheet.TabColor = 12189608
        oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oCellAddress = oSheet.getCellByPosition(0, getLastUsedCell(oSheet).EndRow).getCellAddress()
        oDoc.CurrentController.select(oSheet.getCellByPosition(0, 2))
        # unselect
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
        set_larghezza_colonne()
    else:
        GotoSheet('Analisi di Prezzo')
        oSheet = oDoc.Sheets.getByName('Analisi di Prezzo')
        oDoc.CurrentController.setActiveSheet(oSheet)
        lrow = Range2Cell()[1]
        urow = getLastUsedCell(oSheet).EndRow
        if lrow >= urow:
            lrow = ultima_voce(oSheet) - 5
        for n in range(lrow, getLastUsedCell(oSheet).EndRow):
            if oSheet.getCellByPosition(
                    0, n).CellStyle == 'An-sfondo-basso Att End':
                break
        oRangeAddress = oDoc.NamedRanges.blocco_analisi.ReferredCells.RangeAddress
        oSheet.getRows().insertByIndex(n + 2, 26)
        oCellAddress = oSheet.getCellByPosition(0, n + 2).getCellAddress()
        oDoc.CurrentController.select(oSheet.getCellByPosition(0, n + 2 + 1))
        # unselect
        oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
    oSheet.copyRange(oCellAddress, oRangeAddress)
    basic_LeenO("Menu.eventi_assegna")
    inserisci_Riga_rossa()
    dp()
