
import SheetUtils
import LeenoSheetUtils
import LeenoComputo
import pyleeno as PL
import LeenoUtils
import LeenoEvents

# def generaVariante(oDoc, clear):
#     '''
#     Genera il foglio di VARIANTE a partire dal COMPUTO
#     oDoc    documento di lavoro
#     clear   boolean, se True cancella la variante,
#             se false copia dal computo
#     ritorna il foglio contenente la variante
#     '''
#     if not oDoc.getSheets().hasByName('VARIANTE'):
#         if oDoc.NamedRanges.hasByName("AA"):
#             oDoc.NamedRanges.removeByName("AA")
#             oDoc.NamedRanges.removeByName("BB")

#         oSheetComputo = oDoc.getSheets().getByName("COMPUTO")
#         with LeenoUtils.ProtezioneFoglioContext("COMPUTO", oDoc=oDoc) as oSheetComputo:
#             oSheet = oDoc.getSheets().getByName('COMPUTO')
#             idx = oSheet.RangeAddress.Sheet + 1
#             oDoc.Sheets.copyByName('COMPUTO', 'VARIANTE', idx)

#         oSheet = oDoc.getSheets().getByName('COMPUTO')
#         lrow = SheetUtils.getUsedArea(oSheet).EndRow
#         SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AJ$3:$AJ$' + str(lrow), 'AA')
#         SheetUtils.NominaArea(oDoc, 'COMPUTO', '$N$3:$N$' + str(lrow), "BB")
#         SheetUtils.NominaArea(oDoc, 'COMPUTO', '$AK$3:$AK$' + str(lrow), "cEuro")
#         oSheet = oDoc.getSheets().getByName('VARIANTE')
#         SheetUtils.setTabColor(oSheet, 16777062)
#         oSheet.getCellByPosition(2, 0).String = "VARIANTE"
#         oSheet.getCellByPosition(2, 0).CellStyle = "comp Int_colonna"
#         oSheet.getCellRangeByName("C1").CellBackColor = 16777062
#         oSheet.getCellRangeByPosition(0, 2, 42, 2).CellBackColor = 16777062

#         # se richiesto, svuota la variante appena generata
#         if clear:
#             lrow = SheetUtils.uFindStringCol('TOTALI COMPUTO', 2, oSheet) - 3
#             oSheet.Rows.removeByIndex(3, lrow)
#             LeenoComputo.insertVoceComputoGrezza(oSheet, 2)

#             # @@ PROVVISORIO !!!
#             PL._gotoCella(1, 2 + 1)

#             LeenoSheetUtils.adattaAltezzaRiga(oSheet)
#     else:
#         oSheet = oDoc.getSheets().getByName('VARIANTE')

#     return oSheet

import Dialogs


import LeenoDialogs as DLG
def MENU_generaVariante():
    oDoc = LeenoUtils.getDocument()
    clear = False
    if Dialogs.YesNoDialog(
        IconType="question",
        Title='AVVISO!',
Text='''Vuoi svuotare la VARIANTE appena generata?

Se decidi di continuare, cancellerai tutte le voci di misurazione \
eventualmente già presenti nel foglio di destinazione.

Procedo con lo svuotamento?'''
    ) == 1:
        clear = True
    generaVariante(oDoc, clear)

@LeenoUtils.no_refresh
def generaVariante(oDoc, clear):
    '''
    Genera il foglio di VARIANTE a partire dal COMPUTO.
    clear: se True svuota il foglio dai righi esistenti.
    '''
    sheets = oDoc.getSheets()

    if not sheets.hasByName('VARIANTE'):
        # Pulizia NamedRanges obsoleti
        for name in ("AA", "BB"):
            if oDoc.NamedRanges.hasByName(name):
                oDoc.NamedRanges.removeByName(name)

        # Copia sicura del foglio COMPUTO
        with LeenoUtils.ProtezioneFoglioContext("COMPUTO", oDoc=oDoc):
            oSheetComputo = sheets.getByName('COMPUTO')
            idx = oSheetComputo.RangeAddress.Sheet + 1
            sheets.copyByName('COMPUTO', 'VARIANTE', idx)
            oDoc.CurrentController.select(sheets.getByName('VARIANTE').getCellRangeByName("B5"))
            oDoc.CurrentController.select(oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))
            # LeenoSheetUtils.adattaAltezzaRiga(sheets.getByName('VARIANTE'))


        # Definizione nuove aree nominate (Named Ranges)
        oSheetVar = sheets.getByName('VARIANTE')
        lrow = SheetUtils.getUsedArea(oSheetComputo).EndRow

        SheetUtils.NominaArea(oDoc, 'COMPUTO', f'$AJ$3:$AJ${lrow}', 'AA')
        SheetUtils.NominaArea(oDoc, 'COMPUTO', f'$N$3:$N${lrow}', "BB")
        SheetUtils.NominaArea(oDoc, 'COMPUTO', f'$AK$3:$AK${lrow}', "cEuro")

        # Estetica del foglio VARIANTE
        color_variante = 16777062 # Giallo tenue
        SheetUtils.setTabColor(oSheetVar, color_variante)

        # Titolo dinamico in C1 (cella 2,0)
        oSheetVar.getCellByPosition(2, 0).String = "VARIANTE"
        oSheetVar.getCellByPosition(2, 0).CellStyle = "comp Int_colonna"

        # Colora intestazioni
        oSheetVar.getCellRangeByName("C1").CellBackColor = color_variante
        oSheetVar.getCellRangeByPosition(0, 2, 42, 2).CellBackColor = color_variante

        if clear:
            # Trova la riga dei totali e svuota l'area misurazioni
            row_totali = SheetUtils.uFindStringCol('TOTALI COMPUTO', 2, oSheetVar)
            if row_totali > 3:
                oSheetVar.Rows.removeByIndex(3, row_totali - 3)

            # Inserisce un rigo vuoto iniziale
            LeenoComputo.insertVoceComputoGrezza(oSheetVar, 3)
            LeenoSheetUtils.adattaAltezzaRiga(oSheetVar)

    else:
        oSheetVar = sheets.getByName('VARIANTE')

    # Attivazione e finalizzazione
    PL.GotoSheet('VARIANTE')
    PL.ScriviNomeDocumentoPrincipale()
    LeenoEvents.assegna()

    return oSheetVar
