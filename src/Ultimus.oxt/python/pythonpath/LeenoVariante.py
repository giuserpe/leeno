
import SheetUtils
import LeenoSheetUtils
import LeenoComputo
import pyleeno as PL


def generaVariante(oDoc, clear):
    '''
    Genera il foglio di VARIANTE a partire dal COMPUTO
    oDoc    documento di lavoro
    clear   boolean, se True cancella la variante,
            se false copia dal computo
    ritorna il foglio contenente la variante
    '''
    if not oDoc.getSheets().hasByName('VARIANTE'):
        if oDoc.NamedRanges.hasByName("AA"):
            oDoc.NamedRanges.removeByName("AA")
            oDoc.NamedRanges.removeByName("BB")
        oDoc.Sheets.copyByName('COMPUTO', 'VARIANTE', 4)
        oSheet = oDoc.getSheets().getByName('COMPUTO')
        lrow = SheetUtils.getUsedArea(oSheet).EndRow
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AJ$3:$AJ$' + str(lrow), 'AA')
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$N$3:$N$' + str(lrow), "BB")
        SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AK$3:$AK$' + str(lrow), "cEuro")
        oSheet = oDoc.getSheets().getByName('VARIANTE')
        SheetUtils.setTabColor(oSheet, 16777062)
        oSheet.getCellByPosition(2, 0).String = "VARIANTE"
        oSheet.getCellByPosition(2, 0).CellStyle = "comp Int_colonna"
        oSheet.getCellRangeByName("C1").CellBackColor = 16777062
        oSheet.getCellRangeByPosition(0, 2, 42, 2).CellBackColor = 16777062

        # se richiesto, svuota la variante appena generata
        if clear:
            lrow = SheetUtils.uFindStringCol('TOTALI COMPUTO', 2, oSheet) - 3
            oSheet.Rows.removeByIndex(3, lrow)
            LeenoComputo.insertVoceComputoGrezza(oSheet, 2)

            # @@ PROVVISORIO !!!
            PL._gotoCella(1, 2 + 1)

            LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    else:
        oSheet = oDoc.getSheets().getByName('VARIANTE')

    return oSheet
