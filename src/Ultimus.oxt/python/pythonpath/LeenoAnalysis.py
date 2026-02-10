'''
Funzioni relative alla gestione delle analisi di prezzi
'''

import pyleeno as PL
import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoEvents

import LeenoDialogs as DLG
from undo_utils import with_undo

@with_undo("Inserimento Oneri Sicurezza")
def Inserisci_Utili():
    oDoc = LeenoUtils.getDocument()
    oSheets = oDoc.getSheets()
    oRanges = oDoc.NamedRanges

    oSheetAP = oSheets.getByName("Analisi di Prezzo")

    # 1. Gestione del Range Nominato "oneri_sicurezza"
    if not oRanges.hasByName("oneri_sicurezza"):
        oCellAddress = oSheetAP.getCellRangeByName("B10").getCellAddress()
        oRanges.addNewByName("oneri_sicurezza", "$S5.$B$93:$P$93", oCellAddress, 0)

        # if oSheets.hasByName("S5"):
        #     oSheetS5 = oSheets.getByName("S5")
        #     r = 92 # Riga 93

        #     # Applica stili e contenuti
        #     oSheetS5.getCellByPosition(1, r).CellStyle = "An-lavoraz-desc-CEN"
        #     oSheetS5.getCellByPosition(2, r).CellStyle = "An-lavoraz-Utili"
        #     oSheetS5.getCellByPosition(3, r).CellStyle = "An-lavoraz-Utili"

        #     oCellDesc = oSheetS5.getCellByPosition(4, r)
        #     oCellDesc.CellStyle = "An-lavoraz-Utili desc"
        #     oCellDesc.String = "di cui Sicurezza afferenti l'Impresa"

        #     # Formule in inglese (compatibili Calc)
        #     oCellPerc = oSheetS5.getCellByPosition(5, r)
        #     oDoc.CurrentController.select(oCellPerc)
        #     oCellPerc.CellStyle = "An-lavoraz-%"

        #     oCellPerc.Formula = (
        #         "=IF(OR(E93=\"Spese Generali\";E93=\"Spese Generali (calcolate su F)\"); $S1.$H$320;"
        #         "IF(OR(E93=\"utili d'impresa\";E93=\"Utili d'Impresa (calcolata su F+G)\");$S1.$H$321;"
        #         "IF(E93=\"Spese Generali e Utili (sulle voci precedenti)\";$S1.$H$322;"
        #         "IF(OR(E93=\"Di cui sicurezza afferenti l'impresa\";LEFT(E93;18)=\"Oneri di Sicurezza\");$S1.$H$319;"
        #         "IF(E93=\"Sconto\";$S1.$H$324;"
        #         "IF(E93=\"Maggiorazione\";$S1.$H$326;\"\"))))))"
        #     )

        #     oSheetS5.getCellByPosition(6, r).Formula = "=SUM(H77:H89)*F93"
        #     oSheetS5.getCellByPosition(6, r).CellStyle = "An-lavoraz-Utili-num sin"

        #     oCellSafe = oSheetS5.getCellByPosition(11, r)
        #     oCellSafe.CellStyle = "An-lavoraz-dx"
        #     oCellSafe.Formula = '=IF(ISERROR(SEARCH("sicurezza";E93));"";G93)'

        #     # Altri stili
        #     styles = {7: "An-senza", 8: "An-senza-DX", 9: "An-lavoraz-dx%", 10: "An-lavoraz-generica"}
        #     for col, style in styles.items():
        #         oSheetS5.getCellByPosition(col, r).CellStyle = style
        #     oSheetS5.getCellRangeByPosition(12, r, 15, r).CellStyle = "Analisi_Sfondo"

    # 2. Correzione Cursore (Calc Style)
    # Prendiamo la cella attiva dalla selezione corrente
    oSel = oDoc.CurrentController.getSelection()
    if hasattr(oSel, "CellAddress"):
        lrow = oSel.CellAddress.Row
    elif hasattr(oSel, "RangeAddress"):
        lrow = oSel.RangeAddress.StartRow
    else:
        lrow = 0

    # Trova "Fine ANALISI"
    search_desc = oSheetAP.createSearchDescriptor()
    search_desc.SearchString = "Fine ANALISI"
    found = oSheetAP.findFirst(search_desc)

    if found:
        lrowFine = found.RangeAddress.EndRow
        if lrow > lrowFine:
            lrow = lrowFine - 5

    # Ricerca punto di inserimento (Riga 'I')
    target_row = -1
    for i in range(lrow, min(lrow + 51, oSheetAP.Rows.Count)):
        cell_a = oSheetAP.getCellByPosition(0, i).String
        cell_d = oSheetAP.getCellByPosition(3, i).String

        if cell_a == "L" or "Sicurezza" in cell_d or cell_a == "----":
            DLG.chi("Riprova partendo almeno dalla riga 'I' o riga già inserita.")
            return

        if cell_a == "H" and oSheetAP.getCellByPosition(0, i+1).String == "" and \
           "Sicurezza" in oSheetAP.getCellByPosition(3, i+1).String:
            DLG.chi("La riga degli oneri per la sicurezza è già inserita!")
            oDoc.CurrentController.select(oSheetAP.getCellByPosition(4, i+1))
            return

        if cell_a == "I":
            target_row = i
            break

    # 3. Esecuzione
    oSheetAP.getRows().insertByIndex(target_row, 1)

    oNamedRange = oRanges.getByName("oneri_sicurezza")
    oRangeAddress = oNamedRange.ReferredCells.getRangeAddress()
    oDestAddress = oSheetAP.getCellByPosition(0, target_row).getCellAddress()

    oSheetAP.copyRange(oDestAddress, oRangeAddress)

    # Seleziona la cella di descrizione appena inserita
    oDoc.CurrentController.select(oSheetAP.getCellByPosition(4, target_row))


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
    Circoscrive una voce di analisi partendo dalla posizione corrente del cursore

    Args:
        oSheet (object): Foglio di lavoro
        lrow (int): Riga di riferimento per la selezione dell'intera voce

    Returns:
        object: Intervallo di celle che rappresenta l'analisi
    '''
    # Pre-carica gli stili necessari
    stili_analisi = LeenoUtils.getGlobalVar('stili_analisi')
    cell_style = oSheet.getCellByPosition(0, lrow).CellStyle

    # Variabili per i limiti
    start_row = 0
    end_row = SheetUtils.getLastUsedRow(oSheet)

    # Trova inizio analisi (cerca all'indietro)
    if cell_style in stili_analisi:
        # Ottimizzazione: usa xrange in Python 2 o range in Python 3
        for row in reversed(range(lrow)):
            if oSheet.getCellByPosition(0, row).CellStyle == 'Analisi_Sfondo':
                start_row = row
                break

        # Trova fine analisi (cerca in avanti)
        for row in range(lrow, end_row + 1):
            if oSheet.getCellByPosition(0, row).CellStyle == 'An-sfondo-basso Att End':
                end_row = row
                break

    # Restituisci l'intervallo trovato (250 colonne è un valore arbitrario)
    return oSheet.getCellRangeByPosition(0, start_row, 250, end_row)

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


def MENU_impagina_analisi():
    '''
    PL.set_area_stampa()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name != 'Analisi di Prezzo':
        return
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    oSheet.removeAllManualPageBreaks()
    for el in range (1, lr):
        if oSheet.getCellByPosition(0, el).String == '----':
            if oSheet.getCellByPosition(0, el + 2).CellStyle != 'Ultimus_centro':
                oSheet.getCellByPosition(0, el + 2).Rows.IsStartOfNewPage = True
    '''

    PL.set_area_stampa()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # Esegui solo se il foglio attivo è 'Analisi di Prezzo'
    if oSheet.Name != 'Analisi di Prezzo':
        return

    # Calcola l'ultima riga utilizzata
    last_row = SheetUtils.getLastUsedRow(oSheet) + 1

    # Rimuovi tutte le interruzioni di pagina manuali
    oSheet.removeAllManualPageBreaks()

    # Imposta una nuova interruzione di pagina dopo ogni sezione individuata
    for row in range(1, last_row):
        cell = oSheet.getCellByPosition(0, row)
        if cell.String == '----':
            next_cell = oSheet.getCellByPosition(0, row + 2)
            if next_cell.CellStyle != 'Ultimus_centro':
                next_cell.Rows.IsStartOfNewPage = True
