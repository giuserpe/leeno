from datetime import date
from com.sun.star.table import CellRangeAddress

import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoComputo
import Dialogs
import pyleeno as PL


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
        res = Dialogs.YesNoCancel(
           Title="Voce già registrata",
           Text= "Lavorando in questo punto del foglio,\n"
                 "comprometterai la validità degli atti contabili già emessi.\n\n"
                 "Vuoi procedere?\n\n"
                 "SCEGLIENDO SI' SARAI COSTRETTO A RIGENERARLI!"
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
            #  nSal += 1
        #  else
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

    # @@@ TO REMOVE !!!
    #_gotoCella(1, lrow + 1)

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
    ########################################################################

    if oDoc.NamedRanges.hasByName('#Lib#' + str(nSal)):
        if lrow - 1 == oSheet.getCellRangeByName('#Lib#' + str(nSal)).getRangeAddress().EndRow:
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


def svuotaContabilita(oDoc):
    '''
    Ricrea il foglio di contabilità partendo da zero.
    '''
    for n in range(1, 20):
        if oDoc.NamedRanges.hasByName('#Lib#' + str(n)):
            oDoc.NamedRanges.removeByName('#Lib#' + str(n))
            oDoc.NamedRanges.removeByName('#SAL#' + str(n))
            oDoc.NamedRanges.removeByName('#Reg#' + str(n))
    for el in ('Registro', 'SAL', 'CONTABILITA'):
        if oDoc.Sheets.hasByName(el):
            oDoc.Sheets.removeByName(el)

    oDoc.Sheets.insertNewByName('CONTABILITA', 3)
    oSheet = oDoc.Sheets.getByName('CONTABILITA')

    SheetUtils.setTabColor(oSheet, 16757935)
    oSheet.getCellRangeByName('C1').String = 'CONTABILITA'
    oSheet.getCellRangeByName('C1').CellStyle = 'comp Int_colonna'
    oSheet.getCellRangeByName('C1').CellBackColor = 16757935
    oSheet.getCellByPosition(0, 2).String = 'N.'
    oSheet.getCellByPosition(1, 2).String = 'Articolo\nData'
    oSheet.getCellByPosition(2, 2).String = 'LAVORAZIONI\nO PROVVISTE'
    oSheet.getCellByPosition(5, 2).String = 'P.U.\nCoeff.'
    oSheet.getCellByPosition(6, 2).String = 'Lung.'
    oSheet.getCellByPosition(7, 2).String = 'Larg.'
    oSheet.getCellByPosition(8, 2).String = 'Alt.\nPeso'
    oSheet.getCellByPosition(9, 2).String = 'Quantità\nPositive'
    oSheet.getCellByPosition(11, 2).String = 'Quantità\nNegative'
    oSheet.getCellByPosition(13, 2).String = 'Prezzo\nunitario'
    oSheet.getCellByPosition(15, 2).String = 'Importi'
    oSheet.getCellByPosition(16, 2).String = 'Incidenza\nsul totale'
    oSheet.getCellByPosition(17, 2).String = 'Sicurezza\ninclusa'
    oSheet.getCellByPosition(18, 2).String = 'importo totale\nsenza errori'
    oSheet.getCellByPosition(19, 2).String = 'Lib.\nN.'
    oSheet.getCellByPosition(20, 2).String = 'Lib.\nP.'
    oSheet.getCellByPosition(22, 2).String = 'flag'
    oSheet.getCellByPosition(23, 2).String = 'SAL\nN.'
    oSheet.getCellByPosition(25, 2).String = 'Importi\nSAL parziali'
    oSheet.getCellByPosition(27, 2).String = 'Sicurezza\nunitaria'
    oSheet.getCellByPosition(28, 2).String = 'Materiali\ne Noli €'
    oSheet.getCellByPosition(29, 2).String = 'Incidenza\nMdO %'
    oSheet.getCellByPosition(30, 2).String = 'Importo\nMdO'
    oSheet.getCellByPosition(31, 2).String = 'Super Cat'
    oSheet.getCellByPosition(32, 2).String = 'Cat'
    oSheet.getCellByPosition(33, 2).String = 'Sub Cat'
    #  oSheet.getCellByPosition(34,2).String = 'tag B'sub Scrivi_header_moduli
    #  oSheet.getCellByPosition(35,2).String = 'tag C'
    oSheet.getCellByPosition(36, 2).String = 'Importi\nsenza errori'
    oSheet.getCellByPosition(0, 2).Rows.Height = 800
    #  colore colonne riga di intestazione
    oSheet.getCellRangeByPosition(0, 2, 36, 2).CellStyle = 'comp Int_colonna_R'
    oSheet.getCellByPosition(0, 2).CellStyle = 'comp Int_colonna_R_prima'
    oSheet.getCellByPosition(18, 2).CellStyle = 'COnt_noP'
    oSheet.getCellRangeByPosition(0, 0, 0, 3).Rows.OptimalHeight = True
    #  riga di controllo importo
    oSheet.getCellRangeByPosition(0, 1, 36, 1).CellStyle = 'comp In testa'
    oSheet.getCellByPosition(2, 1).String = 'QUESTA RIGA NON VIENE STAMPATA'
    oSheet.getCellRangeByPosition(0, 1, 1, 1).merge(True)
    oSheet.getCellByPosition(13, 1).String = 'TOTALE:'
    oSheet.getCellByPosition(20, 1).String = 'SAL SUCCESSIVO:'

    oSheet.getCellByPosition(25, 1).Formula = '=$P$2-SUBTOTAL(9;$P$2:$P$2)'

    oSheet.getCellByPosition(15,
                             1).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori
    oSheet.getCellByPosition(0, 1).Formula = '=AK2'  # importo lavori
    oSheet.getCellByPosition(
        17, 1).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza

    oSheet.getCellByPosition(
        28, 1).Formula = '=SUBTOTAL(9;AC3:AC4)'  # importo materiali
    oSheet.getCellByPosition(29,
                             1).Formula = '=AE2/Z2'  # Incidenza manodopera %
    oSheet.getCellByPosition(29, 1).CellStyle = 'Comp TOTALI %'
    oSheet.getCellByPosition(
        30, 1).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
    oSheet.getCellByPosition(
        36, 1).Formula = '=SUBTOTAL(9;AK3:AK4)'  # importo certo

    #  rem riga del totale
    oSheet.getCellByPosition(2, 3).String = 'T O T A L E'
    oSheet.getCellByPosition(15,
                             3).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori
    oSheet.getCellByPosition(
        17, 3).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza
    oSheet.getCellByPosition(
        30, 3).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
    oSheet.getCellRangeByPosition(0, 3, 36, 3).CellStyle = 'Comp TOTALI'
    #  rem riga rossa
    oSheet.getCellByPosition(0, 4).String = 'Fine Computo'
    oSheet.getCellRangeByPosition(0, 4, 36, 4).CellStyle = 'Riga_rossa_Chiudi'

    # @@_gotoCella(0, 2)

    LeenoSheetUtils.setLarghezzaColonne(oSheet)

    return oSheet


# ###############################################################


def generaContabilita(oDoc):
    '''
    Ritorna il foglio di contabilità, se presente
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

            PL.eventi_assegna()
            LeenoSheetUtils.ScriviNomeDocumentoPrincipaleInFoglio(oSheet)

    return oSheet
