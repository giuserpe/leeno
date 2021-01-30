from datetime import date
from com.sun.star.table import CellRangeAddress
from com.sun.star.sheet.GeneralFunction import MAX

import LeenoUtils
import SheetUtils
import LeenoSheetUtils
import LeenoComputo
import Dialogs
import LeenoDialogs as DLG
import pyleeno as PL
import LeenoEvents
import LeenoBasicBridge
import uno


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
    svuota_contabilita
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

    # riga del totale
    oSheet.getCellByPosition(2, 3).String = 'T O T A L E'
    oSheet.getCellByPosition(15,
                             3).Formula = '=SUBTOTAL(9;P3:P4)'  # importo lavori
    oSheet.getCellByPosition(
        17, 3).Formula = '=SUBTOTAL(9;R3:R4)'  # importo sicurezza
    oSheet.getCellByPosition(
        30, 3).Formula = '=SUBTOTAL(9;AE3:AE4)'  # importo manodopera
    oSheet.getCellRangeByPosition(0, 3, 36, 3).CellStyle = 'Comp TOTALI'
    # riga rossa
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
            PL.GotoSheet('CONTABILITA')
            svuotaContabilita(oDoc)
            PL.ins_voce_contab()
        PL.GotoSheet('CONTABILITA')
    LeenoBasicBridge.ScriviNomeDocumentoPrincipale()
    LeenoEvents.assegna()
########################################################################


def partita(testo):
    '''
    Aggiunge/detrae rigo di PARTITA PROVVISORIA
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
    if oSheet.Name != "CONTABILITA":
        return
    x = LeggiPosizioneCorrente()[1]
    if oSheet.getCellByPosition(0, x).CellStyle == 'comp 10 s_R':
        if oSheet.getCellByPosition(2, x).Type.value != 'EMPTY':
            Copia_riga_Ent()
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
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
    PL.struttura_off()
    oRanges = oDoc.NamedRanges
    
    if oSheet.Name == 'CONTABILITA':
        pref = "#Lib#"
        y = 3
    elif oSheet.Name == 'Registro':
        pref = "#Reg#"
        y = 1
    elif oSheet.Name == 'SAL':
        pref = "#SAL#"
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
                Dialogs.NotifyDialog(Image='Icons-Big/info.png',
                        Title = 'Info',
                        Text='''In questo Libretto delle Misure
non ci sono misure registrate.''')
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
    #  DLG.mri(oDoc.StyleFamilies.getByName("CellStyles").getByName('comp 1-a PU'))
    #  return
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
    if oSheet.Name != 'CONTABILITA':
        return
    PL.numera_voci()
    oRanges = oDoc.NamedRanges

    # ~try:
        # ~oRanges.removeByName("#Lib#1")
    # ~except:
        # ~pass

    # ~return
    #trovo il numero del nuovo sal
    nSal = 0
    for i in reversed(range(1, 50)):
        if oRanges.hasByName("#Lib#" + str(i)) == True:
            nSal = i +1
            break
        else:
            nSal = 1
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
    if daVoce >= int(oCellRange.computeFunction(MAX)):
        Dialogs.NotifyDialog(Image='Icons-Big/exclamation.png',
                Title = 'ATTENZIONE!',
                Text='''Tutte le voci di questo Libretto delle Misure
sono già registrate.''')
        return
    # ~try:
        # ~oRanges.hasByName("#Lib#1")
        # ~daVoce = int(oSheet.getCellByPosition(2, 2 + nSal - 1
                # ~).String.split('÷')[1]) + 1
        # ~DLG.chi(daVoce)
        # ~oCellRange = oSheet.getCellRangeByPosition(
            # ~0, 3, 0,
            # ~SheetUtils.getUsedArea(oSheet).EndRow - 2)

        # ~if daVoce >= int(oCellRange.computeFunction(MAX)):
            # ~DLG.MsgBox('Non ci sono voci di misurazione da registrare.', 'ATTENZIONE!')
            # ~return
    # ~except:
        # ~pass

        # ~nSal = 1
        # ~daVoce = 1
        # ~old_nPage = 1

    nomearea="#Lib#" + str(nSal)

    #  Recupero la prima riga non registrata


    # ~if nSal > 0:
        # ~oNamedRange = oRanges.getByName("#Lib#" +
                                        # ~str(nSal)).ReferredCells.RangeAddress
        # ~frow = oNamedRange.StartRow
        # ~lrow = oNamedRange.EndRow
        # ~daVoce = oNamedRange.EndRow + 2
    #  recupero l'ultimo numero di pagina usato (serve in caso di libretto unico)
    #  oSheet = oDoc.Sheets.getByName('CONTABILITA')
        # ~old_nPage = int(oSheet.getCellByPosition(20, lrow).Value)
        # ~daVoce = int(oSheet.getCellByPosition(0, daVoce).Value)
        # ~if daVoce == 0:
            # ~DLG.MsgBox('Non ci sono voci di misurazione da registrare.', 'ATTENZIONE!')
            # ~return
        # ~oCell = oSheet.getCellRangeByPosition(0, frow, 25, lrow)
    #  Raggruppa_righe
        # ~oCell.Rows.IsVisible = False
        #cerca prima voce da registrare
        # ~for el in reversed(range(0, lrow)):
            # ~if oSheet.getCellByPosition(0, el).Value > 0:
                # ~daVoce = int(oSheet.getCellByPosition(0, el).Value + 1)
                # ~break
    # ~else:
        # ~nSal = 1
        # ~daVoce = 1
        # ~old_nPage = 1
    #############
    # PRIMA VOCE


    # ~DLG.chi(2 + nSal - 1)
    # ~return

    # ~if nSal > 1:
        # ~DLG.chi(int(oSheet.getCellByPosition(2, 2 + nSal - 1).String.split('÷')[1]) +1)
        # ~oLibNamedRange = oRanges.getByName("#Lib#" + str(nSal - 1)).ReferredCells.RangeAddress
        # ~oLibNamedRange.EndRow
        # ~DLG.chi(oLibNamedRange.EndRow)
        # ~for el in reversed(range(0, oLibNamedRange.EndRow)):
            # ~DLG.chi(oSheet.getCellByPosition(0, el).Value)
            # ~if oSheet.getCellByPosition(0, el).Value > 0:
                # ~daVoce = oSheet.getCellByPosition(0, el).Value
            # ~break

    daVoce = PL.InputBox(str(daVoce), "Registra voci Libretto da n.")
    if len(daVoce) ==0:
        return

    try:
        lrow = int(SheetUtils.uFindStringCol(daVoce, 0, oSheet))
    except TypeError:
        return
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    primariga = sStRange.RangeAddress.StartRow

    #  ULTIMA VOCE
    oCellRange = oSheet.getCellRangeByPosition(
        0, 3, 0, SheetUtils.getUsedArea(oSheet).EndRow - 2)
    aVoce = int(oCellRange.computeFunction(MAX))

    aVoce = PL.InputBox(str(aVoce), "A voce n.:")
    # ~aVoce = str(int(daVoce) + 1)
    if len(aVoce) == 0:
        return

    # attiva la progressbar
    progress = Dialogs.Progress(Title='Generazione elaborato...', Text="Libretto delle Misure")
    progress.setLimits(1, 6)
    progress.setValue(0)
    progress.show()
    progress.setValue(1)

    try:
        lrow = int(SheetUtils.uFindStringCol(aVoce, 0, oSheet))
    except TypeError:
        return
    sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    ultimariga = sStRange.RangeAddress.EndRow

    lrowDown = SheetUtils.uFindStringCol("T O T A L E", 2, oSheet)

    PL.comando ('DeletePrintArea')
    # ~oSheet.removeAllManualPageBreaks()
    SheetUtils.visualizza_PageBreak()

    oSheet.getCellByPosition(25, ultimariga - 1).String = "SAL n." + str(nSal)
    oSheet.getCellByPosition(25, ultimariga).Formula = (
        "=SUBTOTAL(9;P" + str(primariga + 1) + ":P" + str(ultimariga+1) + ")" )
    oSheet.getCellByPosition(25, ultimariga).CellStyle = "comp sotto Euro 3_R"
    # immetti le firme
    inizioFirme = ultimariga + 1

    PL.MENU_firme_in_calce (inizioFirme) # riga di inserimento
    fineFirme = inizioFirme + 10

    progress.setValue(2)
    area="$A$" + str(primariga + 1) + ":$AJ$" + str(fineFirme + 1)
#  'Print area
    LeenoBasicBridge.rifa_nomearea(oDoc, "CONTABILITA", area , nomearea)

    oSheet.getCellRangeByPosition(0, inizioFirme, 32, fineFirme).CellStyle = "Ultimus_centro_bordi_lati"
    #applico gli stili corretti ad alcuni dati della firma
    oSheet.getCellByPosition(2, inizioFirme + 1).CellStyle = "Ultimus_destra"

    oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

    #range del #Lib#
    daRiga = oNamedRange.StartRow
    aRiga = oNamedRange.EndRow
    daColonna = oNamedRange.StartColumn
    aColonna = oNamedRange.EndColumn

    iSheet = oSheet.RangeAddress.Sheet
    # imposta area di stampa
    # ~oCellRangeAddr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    # ~oCellRangeAddr.Sheet = iSheet
    # ~oCellRangeAddr.StartColumn = daColonna
    # ~oCellRangeAddr.StartRow = daRiga
    # ~oCellRangeAddr.EndColumn = 11
    # ~oCellRangeAddr.EndRow = aRiga
    
    # imposta riga da ripetere
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet
    oTitles.StartColumn = 0
    oTitles.StartRow = 2
    oTitles.EndColumn = 11
    oTitles.EndRow = 2
    oSheet.setTitleRows(oTitles)
    # ~oSheet.setPrintAreas((oCellRangeAddr,))
    oNamedRange.EndColumn = 11
    oSheet.setPrintAreas((oNamedRange,))
    oSheet.setPrintTitleRows(True)
    
    oSheet.PageStyle = "Page_Style_Libretto_Misure2" 
    
    progress.setValue(3)

    # sbianco l'area di stampa
    oSheet.getCellRangeByPosition(daColonna, daRiga, 11, aRiga).CellBackColor = -1

    progress.setValue(4)

    for i in range(0, 50):
        oSheet.getRows().insertByIndex(fineFirme, 1)
        oSheet.getCellByPosition(2, fineFirme).String = "===================="
        fineFirme += 1
        if oSheet.getCellByPosition(1, fineFirme).Rows.IsStartOfNewPage == True:
            fineFirme -= 1
            oSheet.getRows().removeByIndex(fineFirme, 1)
            break
    # ~oSheet.getRows().removeByIndex(fineFirme, 1)
    # ~fineFirme -=1

    oBordo = oSheet.getCellRangeByPosition(0, fineFirme, 32, fineFirme)
    bordo = oBordo.BottomBorder
    bordo.LineWidth = 2
    bordo.OuterLineWidth = 2
    oBordo.BottomBorder = bordo

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
    progress.setValue(5)
    for i in range(primariga, fineFirme):
        if oSheet.getCellByPosition(1, i).CellStyle == "comp Art-EP_R":
            if primariga == 0:
                sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, i)
                primariga = sStRange.RangeAddress.StartRow
            oSheet.getCellByPosition(19, i).Value= nLib     #numero libretto
            oSheet.getCellByPosition(22, i).String= "#reg"  #flag registrato
            oSheet.getCellByPosition(23, i).Value= nSal     #numero SAL

            for nPag in range(0, len(oSheet.RowPageBreaks)):
                if i < oSheet.RowPageBreaks[nPag].Position:
                    oSheet.getCellByPosition(20, i).Value = nPag   #pagina
                    break

    progress.setValue(6)
    # annoto ultimo numero di pagina
    oSheet.getCellByPosition(20 , fineFirme).Value = nPag
    oSheet.getCellByPosition(20 , fineFirme).CellStyle = "num centro"
#  inumPag = nPag ' + old_nPage 'SE IL LIBRETTO è UNICO

    # ~SheetUtils.visualizza_PageBreak(False)

    # inserisco la prima riga GIALLA del LIBRETTO
    oSheet.getRows().insertByIndex(daRiga, 1)
    oSheet.getCellRangeByPosition (0, daRiga, 36, daRiga).CellStyle = "uuuuu"

    #range del #Lib#

    oSheet.getCellByPosition(2,  daRiga).String = (
        "segue Libretto delle Misure n." + str(nSal) +
        " - " + str(daVoce) + "÷" + str(aVoce)
        )
    oSheet.getCellByPosition(20, daRiga).Value =  nPag  #Pagina
    oSheet.getCellByPosition(19, daRiga).Value= nLib    #Libretto
    oSheet.getCellByPosition(23, daRiga).Value= nSal    #SAL
    oSheet.getCellByPosition(15, daRiga).Formula = "=Z" + str(daRiga + 1)
    oSheet.getCellByPosition(15, daRiga).CellStyle = "comp sotto Euro 3_R"
    oSheet.getCellByPosition(25, daRiga).Formula =(
        "=SUBTOTAL(9;$P$" + str(primariga + 2) + ":$P$" + str(
        ultimariga + 2) + ")"
        )
    oSheet.getCellByPosition(25, daRiga).CellStyle = "comp sotto Euro 3_R"

    # annoto il sal corrente sulla riga di intestazione
    oSheet.getCellByPosition(25, 2).Value = nSal
    oSheet.getCellByPosition(25, 2).CellStyle = "Menu_sfondo _input_grasBig"
    oSheet.getCellByPosition(25, 1).Formula = (
        "=$P$2-SUBTOTAL(9;$P$2:$P$" + str(ultimariga + 2) + ")"
        )
    # ~oDoc.CurrentController.select(oSheet.getCellByPosition(25, 1))

    # corregge raggruppamento libretto precedente dopo inserimento riga gialle
    # ~try:
        # ~oPrevRange=oRanges.getByName("#Lib#" + str(nSal - 1)).ReferredCells.RangeAddress
        # ~oSheet.ungroup(oPrevRange, 1)
        # ~oSheet.group(oPrevRange, 1)
        # ~oSheet.getCellRangeByPosition(oPrevRange.StartColumn,
            # ~oPrevRange.StartRow, oPrevRange.EndColumn, oPrevRange.EndRow
            # ~).Rows.IsVisible = False
    # ~except:
        # ~pass

    # ~oPrintArea = oSheet.getPrintAreas()
    # ~oSheet.group(oPrintArea[0], 1)

    PL._gotoCella(0, daRiga)
    # ~_gotoCella(0, inizioFirme)
    progress.hide()

#  Protezione_area ("CONTABILITA",nomearea)
#  Struttura_Contab ("#Lib#")
#  Genera_REGISTRO
    struttura_CONTAB()
    return nSal, daVoce, aVoce, primariga+1, ultimariga+1


########################################################################
#  
#  name: GeneraRegistro
#  @param
#  @return
#  
def GeneraRegistro(oDoc):
    '''
    CONTABILITA' - genera un nuovo foglio 'Registro'. Si ottiene una riga
    gialla con l'indicazione delle voci di misurazione registrate ed un
    parziale dell'importo del SAL a cui segue la visualizzazione in
    struttura delle relative voci registrate nel Libretto delle Misure.
    '''
    # ~zoom = oDoc.CurrentController.ZoomValue
    # ~oDoc.CurrentController.ZoomValue = 400
    # ~oDoc = LeenoUtils.getDocument()
    # ~Annulla_atto_contabile("Registro")
    # ~lista = GeneraLibretto(oDoc)
    # ~DLG.chi(lista)
    # ~return
    try:
        nSal, daVoce, aVoce, primariga, ultimariga = GeneraLibretto(oDoc)
    except:
        return
    # ~DLG.chi((nSal, daVoce, aVoce, primariga, ultimariga))
    # ~return
    
    
    # Recupero i dati dal libretto
    oSheet = oDoc.Sheets.getByName("CONTABILITA")
    REG = []
    SAL = []
    i = primariga
    while i < ultimariga:
        '''
        voce = (num, art, data, desc, Nlib, Plib, um, quantP, quantN,
                prezzo, importo, sic, mdo, flag, nSal)
        prendo solo i dati utili [0:11]
        '''
        reg, sal = LeenoComputo.datiVoceComputo(oSheet, i)[0:11]
        REG.append(reg)
        SAL.append(sal)
        i= LeenoSheetUtils.prossimaVoce(oSheet, i)
    # ~DLG.chi(REG)
    # ~DLG.chi(SAL)
    
    try:
        oDoc.getSheets().insertNewByName('Registro',5)
        PL.GotoSheet('Registro')
        oSheet = oDoc.Sheets.getByName('Registro')

    # riga di intestazione
        oSheet.getCellRangeByPosition(0,0,9,0).CellStyle="An.1v-Att Start"
        oSheet.getCellByPosition(0,0).String = ("N. ord."+ chr(13) +"Articolo"+ chr(13) +"Data")
        oSheet.getCellByPosition(1,0).String = ("LAVORAZIONI"+ chr(13) + "E SOMMINISTRAZIONI")
        oSheet.getCellByPosition(2,0).String = ("Lib." + chr(13) +"N.")
        oSheet.getCellByPosition(3,0).String = ("Lib." + chr(13) +"P.")
        oSheet.getCellByPosition(4,0).String = ("U.M.")
        oSheet.getCellByPosition(5,0).String = ("Quantità" + chr(13) + "Positive")
        oSheet.getCellByPosition(6,0).String = ("Quantità" + chr(13) + "Negative")
        oSheet.getCellByPosition(7,0).String = ("Prezzo" + chr(13) + "unitario")
        oSheet.getCellByPosition(8,0).String = ("Importo" + chr(13) + "debito")
        oSheet.getCellByPosition(9,0).String = ("Importo" + chr(13) + "pagamento")
        # ~oSheet.getCellByPosition(10,0).String = ("Num." + chr(13) +"Pag.")
    # larghezza colonne
        oSheet.getCellByPosition(0,0).Columns.Width = 1600 #'N. ord.
        oSheet.getCellByPosition(1,0).Columns.Width = 6600 #'LAVORAZIONI
        oSheet.getCellByPosition(2,0).Columns.Width = 650 #'Lib.N.
        oSheet.getCellByPosition(3,0).Columns.Width = 650 #'Lib.P.
        oSheet.getCellByPosition(4,0).Columns.Width = 1000 #'U.M.
        oSheet.getCellByPosition(5,0).Columns.Width = 1600 #'Positive
        oSheet.getCellByPosition(6,0).Columns.Width = 1600 #'Negative
        oSheet.getCellByPosition(7,0).Columns.Width = 1400 #'Prezzo
        oSheet.getCellByPosition(8,0).Columns.Width = 1950 #'debito
        oSheet.getCellByPosition(9,0).Columns.Width = 1950 #'pagamento
        # ~oSheet.getCellByPosition(10,0).Columns.OptimalWidth = True #'n.pag.
        insRow = 1 #'prima riga inserimento in Registro
    except:
        # recupera il registro precedente
        PL.GotoSheet('Registro')
        oSheet= oDoc.Sheets.getByName("Registro")
        # ~DLG.chi("#Reg#" + str(nSal - 1))
        oRanges = oDoc.NamedRanges
        oPrevRange = oRanges.getByName("#Reg#" + str(nSal - 1)).ReferredCells.RangeAddress

        fRow = oPrevRange.StartRow
        lRow = oPrevRange.EndRow
        insRow = oPrevRange.EndRow + 1
        
        # chiudo il registro precedente
        oCell = oSheet.getCellRangeByPosition(0,fRow,11,lRow)
        oCell.Rows.IsVisible=False

    oSheet.PageStyle = 'PageStyle_REGISTRO_A4'

    # compilo il Registro
    lastRow = insRow + len(REG) -1
    oRange = oSheet.getCellRangeByPosition(0, insRow, 8, lastRow)
    # ~oDoc.CurrentController.select(oRange)
    reg = tuple(REG)
    # ~DLG.chi(insRow)
    # ~DLG.chi(len (reg[0]))
    oRange.setDataArray(reg)
    
# do gli stili al Registro
    oSheet.getCellRangeByPosition(0, insRow, 1, lastRow).CellStyle = "List-stringa-sin"
    oSheet.getCellRangeByPosition(2, insRow, 4, lastRow).CellStyle = "List-num-centro"
    oSheet.getCellRangeByPosition(5, insRow, 6, lastRow).CellStyle = "comp 1a"
    oSheet.getCellRangeByPosition(7, insRow, 9, lastRow).CellStyle = "List-num-euro"

# inserisco la prima riga GIALLA
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
    oSheet.getCellByPosition(8, insRow).Formula = (
        '=IF(SUBTOTAL(9;$I$2:$I$' + str(insRow) + ')=0;"";SUBTOTAL(9;$I$2:$I$' + str(insRow))
    oSheet.getCellByPosition(9, insRow).Formula = (
        '=IF(SUBTOTAL(9;$J$2:$J$' + str(insRow) + ')=0;"";SUBTOTAL(9;$J$2:$J$' + str(insRow))
        
    oSheet.getCellRangeByPosition (0, insRow, 9, insRow).CellStyle = "Ultimus_Bordo_sotto"

    insRow += 1
    oSheet.getRows().insertByIndex(insRow, 1)

    oSheet.getCellByPosition(1, insRow).String = "LAVORI A MISURA"
    oSheet.getCellRangeByPosition(0, insRow, 9, insRow).CellStyle = "Ultimus_centro_bordi_lati"
    PL._gotoCella(1, insRow)
    
    lastRow = insRow + len(REG)

    inizioFirme = lastRow + 5
    PL.MENU_firme_in_calce (inizioFirme) # riga di inserimento
    fineFirme = inizioFirme + 18

# set area del REGISTRO
    area="$A$" + str(insRow) + ":$J$" + str(fineFirme + 1)
    nomearea = "#Reg#" + str(nSal)
    LeenoBasicBridge.rifa_nomearea(oDoc, "Registro", area , nomearea)

    oRanges = oDoc.NamedRanges
    oNamedRange=oRanges.getByName(nomearea).ReferredCells.RangeAddress

    #range del #Reg#
    # ~daRiga = oNamedRange.StartRow
    # ~aRiga = oNamedRange.EndRow
    # ~daColonna = oNamedRange.StartColumn
    # ~aColonna = oNamedRange.EndColumn

    iSheet = oSheet.RangeAddress.Sheet

    # imposta riga da ripetere
    oTitles = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
    oTitles.Sheet = iSheet
    # ~oTitles.StartColumn = 0
    oTitles.StartRow = 0
    # ~oTitles.EndColumn = 9
    # ~oTitles.EndRow = 0
    oSheet.setTitleRows(oTitles)
    # ~oSheet.setPrintAreas((oCellRangeAddr,))
    oSheet.setPrintAreas((oNamedRange,))
    oSheet.setPrintTitleRows(True)

    # ~oPrintArea = oSheet.getPrintAreas()
    # ~oSheet.group(oPrintArea[0], 1)

    oSheet.getCellRangeByPosition(0, lastRow +1, 9, fineFirme).CellStyle = "Ultimus_centro_bordi_lati"

    #torno su a completare...
    oSheet.getCellByPosition(1, lastRow + 2).String = "Parziale dei Lavori a Misura €"
    oSheet.getCellByPosition(1, lastRow + 2).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(8, lastRow + 2).Formula = (
        '=SUBTOTAL(9;$I$2:$I$' + str(fineFirme))
    oSheet.getCellByPosition(8, lastRow + 2).CellStyle = "Ultimus_destra_totali"
    
    oSheet.getCellByPosition(1, lastRow + 4).String = 'Lavori a tutto il ' + PL.oggi() + ' - T O T A L E   €'
    oSheet.getCellByPosition(1, lastRow + 4).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(8, lastRow + 4).Formula = (
        '=SUBTOTAL(9;$I$2:$I$' + str(fineFirme))
    oSheet.getCellByPosition(8, lastRow + 4).CellStyle = "Ultimus_destra_totali"

    #applico gli stili corretti ad alcuni dati della firma
    oSheet.getCellByPosition(1, lastRow + 6).CellStyle = "Ultimus_destra"
    oSheet.getCellByPosition(1, lastRow + 16).CellStyle = "Ultimus_destra"

    
    # ~LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    for i in range(0, 50):
        oSheet.getRows().insertByIndex(fineFirme, 1)
        oSheet.getCellByPosition(1, fineFirme).String = "===================="
        fineFirme += 1
        if oSheet.getCellByPosition(1, fineFirme).Rows.IsStartOfNewPage == True:
            fineFirme -= 1
            oSheet.getRows().removeByIndex(fineFirme, 1)
            break
    # ~oSheet.getCellRangeByPosition(0, fineFirme, 9, fineFirme).CellStyle = "comp Descr"
    oSheet.getCellByPosition(1,fineFirme).String = "A   R I P O R T A R E"
    oSheet.getCellByPosition(8, fineFirme).Formula = (
        '=IF(SUBTOTAL(9;$I$2:$I$' + str(fineFirme) + ')=0;"";SUBTOTAL(9;$I$2:$I$' + str(fineFirme))
    oSheet.getCellByPosition(9, fineFirme).Formula = (
        '=IF(SUBTOTAL(9;$J$2:$J$' + str(fineFirme) + ')=0;"";SUBTOTAL(9;$J$2:$J$' + str(fineFirme))
    oSheet.getCellRangeByPosition (0, fineFirme, 9, fineFirme).CellStyle = "Ultimus_Bordo_sotto"

    struttura_CONTAB()
    return


    

    # ~ThisComponent.CurrentController.setFirstVisibleRow(daRiga)
    # ~oDoc.CurrentController.ZoomValue = zoom

def rrrrrrrrrrrrrrrrrrrrr():# ~sub Genera_REGISTRO '(C) Giuseppe Vizziello 2014

# ~fissa (0,idxrow+1)
    # ~oSheet = ThisComponent.Sheets.getByName("CONTABILITA")
    # ~ThisComponent.CurrentController.Select(oSheet)'.getCellRangeByPosition(0,0,0,0))
# ----------------------------------------------------------------------
# Recupero i dati dal libretto


# ----------------------------------------------------------------------
        # ~AppendItem(elVoci(), array (num, art, data, desc, um, Nlib, Plib, quant, prezzo, importo, sicurezza, mdo))
# ~'        Barra_Apri_Chiudi_5("                                Restano "& aRiga-i &" righe...", 0)
    # ~Next
# ----------------------------------------------------------------------
# completo la lista voci per la preparazione del SAL prima di passare alla compilazione del REGISTRO
    # ~vociSAL()=elVoci()
    
    # ~aRiga=daRiga-1
    # ~daRiga=0    'inizio dalla prima riga indipendendemente dal numero di sal corrente
    # ~for i = daRiga to aRiga
        # ~if oSheet.getCellByPosition(idxcol+0, i).cellstyle = "Comp Start Attributo_R" then 'i = inizio voce
            # ~sStRange = Circoscrive_Voce_Computo_Att (i)

            # ~f = sStRange.RangeAddress.EndRow        'fine voce
            # ~num        = oSheet.getCellByPosition(idxcol+0, i+1).getstring 'num. voce
            # ~art        = oSheet.getCellByPosition(idxcol+1, i+1).getstring 'articolo
            # ~data    = oSheet.getCellByPosition(idxcol+1, i+2).getstring 'data
            # ~desc    = oSheet.getCellByPosition(idxcol+2, i+1).getstring 'descrizione
            # ~um        = oSheet.getCellByPosition(idxcol+9, i+1).getstring 'unità
            # ~Plib    = oSheet.getCellByPosition(idxcol+20, i+1).getvalue 'Lib. P.
            # ~quant    = oSheet.getCellByPosition(idxcol+9, f).getvalue 'quantità
            # ~prezzo    = oSheet.getCellByPosition(idxcol+13, f).getvalue 'prezzo
            # ~importo    = oSheet.getCellByPosition(idxcol+15, f).getvalue 'importo
            # ~sicurezza= oSheet.getCellByPosition(idxcol+17, f).getvalue 'sicurezza
            # ~mdo        = oSheet.getCellByPosition(idxcol+30, f).getvalue 'mdo
            
# ----------------------------------------------------------------------
# approssimazione dei valori a idxdec dopo la virgola
# ~'        quant=myround(quant,idxdec)
# ~'        importo=myround(importo,idxdec)
# ~'        sicurezza=myround(sicurezza,idxdec)
# ~'        mdo=myround(mdo,idxdec)
# ----------------------------------------------------------------------
            # ~AppendItem(vociSAL(), array (num, art, data, desc, um, Nlib, Plib, quant, prezzo, importo, sicurezza, mdo))
            # ~i=f
        # ~end if
        
    # ~Next


# ----------------------------------------------------------------------
# PREPARO I DATI PER IL SAL
    # ~Dim articoli() 'lista articoli
    # ~Dim lista() 'LISTA D'APPOGGIO

    # ~For Each el In vociSAL()
        # ~Appenditem (lista(), el(1))
    # ~next

    # ~If NOT GlobalScope.BasicLibraries.isLibraryLoaded( "Tools" ) Then GlobalScope.BasicLibraries.LoadLibrary( "Tools" ) 
    # ~lista()=BubbleSortlist(BubbleSortlist(lista())) ' riordino
# ----------------------------------------------------------------------
# serve ad allungare la lista
# non viene preso in considerazione  per il FOR successivo, ma è necessario che ci sia per non mandarlo in errore
# evito il resume next perché impedisce il debug
    # ~'On Error Resume Next 
    # ~Appenditem (lista(), "ultimavocelistasolodiservizio")
# ELIMINA I DOPPIONI e trasferisco i dati puliti in articoli()
# ~'xray lista
    # ~For I = 0 To UBound(lista) -1
        # ~If lista(I) <> lista(I + 1) Then     Appenditem (articoli(), lista(I))
    # ~Next I
# ----------------------------------------------------------------------
# CALCOLO IMPORTI TOTALI
    # ~For Each el In vociSAL()
        # ~SALamisura = SALamisura+el(9)
        # ~SALsicurezza= SALsicurezza+el(10)
        # ~SALmdo = SALmdo+el(11)
    # ~Next
# ~lista()=vociSAL()
# ~ReDim vociSAL(0)
# ----------------------------------------------------------------------
# sommo i valori degli articoli ricorrenti
    # ~For Each art In articoli()
        # ~quant=0
        # ~importo=0
        # ~sicurezza=0
        # ~mdo=0
# ~'Print "art " & art
        # ~For Each i In lista()
            # ~If art=i(1) Then
                # ~desc = i(3)
                # ~um = i(4)
                # ~quant=quant+i(7)
                # ~prezzo = i(8)
                # ~importo=importo+i(9)
                # ~sicurezza=sicurezza+i(10)
                # ~mdo=mdo+i(11)
# ~'    Print i(7) &" - "& i(9) &" - "& i(10) &" - "& i(11)    
# ~'    Print  i(1) &" - "& quant &" - "& importo &" - "& sicurezza &" - "& mdo
            # ~End If
# ~'Print i(1) &" - "& quant &" - "& importo &" - "& sicurezza &" - "& mdo
        # ~Next
        # ~AppendItem (vociSAL(), array (art, desc, um,  quant, prezzo, importo, sicurezza, mdo))
# ~'        AppendItem (vociSAL(), array (art, i(3), i(4),  quant, i(8), importo, sicurezza, mdo)
    # ~Next
# ~'Print ubound (vociSAL())
# ~' Print "vocisal"
# ~'xray (vocisal())
# ~'For Each el In vocisal
# ~'print (el(0))
# ~'next
# ~'Exit Sub 
# ~'0        1        2        3        4        5        6        7        8        9            10            11
# ~'num,    art,    data,    desc,     um,     Nlib,    Plib,    quant,    prezzo, importo,    sicurezza,    mdo 
# ----------------------------------------------------------------------
# elimino i doppioni
    # ~lista()=vociSAL()
    # ~ReDim vociSAL()
    # ~nn = UBound(lista)+1
    # ~ReDim Preserve lista(nn)
# ~'xray lista
# ~'Print UBound(lista()) 
# ~i=0
    # ~Do While I < UBound(lista) '-1
# ~'        Print lista(I)(0)
        # ~If lista(I)(0) = lista(I+1)(0) Then
            # ~If Not isempty (lista(I+1)) Then Appenditem (vocisal(), lista(I+1))
            # ~i=i+1
        # ~Else 
            # ~If Not isempty (lista(I)) Then Appenditem (vocisal(), lista(I))    
        # ~'    If lista(I) <> lista(I + 1) Then     Appenditem (articoli(), lista(I))
        # ~End If
        # ~i=i+1
    # ~Loop     
# ~'xray vocisal()
# ~'Print ubound (vociSAL())
# ~'Exit Sub
# ~'#########################################################################
# COMPILO LA SITUAZIONE CONTABILE IN "S2" 2di2
    # ~oS2 = ThisComponent.Sheets.getByName("S2")
# TROVO LA POSIZIONE DEL TITOLO
    # ~oEnd=uFindString("SITUAZIONE CONTABILE", oS2)
    # ~yS2=oEnd.RangeAddress.EndRow        'riga
    # ~xS2=oEnd.RangeAddress.EndColumn    'colonna
    # ~sCol = ColumnNameOf(xS2+nSal)
    # ~oS2.getCellByPosition(xS2+nSal,yS2+4).formula = "="& sCol & yS2+4 & "*$C$74" ' iincidenza sicurezza su lavori a CORPO
    # ~oS2.getCellByPosition(xS2+nSal,yS2+5).formula = "="& sCol & yS2+4 & "*$C$76" ' iincidenza sicurezza su lavori a CORPO
    # ~oS2.getCellByPosition(xS2+nSal,yS2+8).value = SALamisura ' importo lavori a misura
    # ~oS2.getCellByPosition(xS2+nSal,yS2+9).formula = "="& sCol & yS2+9 & "*$C$74" ' iincidenza sicurezza su lavori a MISURA
    # ~oS2.getCellByPosition(xS2+nSal,yS2+10).formula = "="& sCol & yS2+9 & "*$C$76" ' iincidenza sicurezza su lavori a MISURA
    # ~sFormula = "="& sCol & yS2+9 &"+"& sCol & yS2+4 & "-" & sCol & yS2+5 & "-" & sCol & yS2+6 & "-" & sCol & yS2+10 & "-" & sCol & yS2+11
    # ~oS2.getCellByPosition(xS2+nSal,yS2+12).formula = sFormula ' quota da ribassare
    # ~oS2.getCellByPosition(xS2+nSal,yS2+13).formula = "=" & sCol & yS2+13 &"*$C$78" ' ribasso
    
    # ~oS2.getCellByPosition(xS2+nSal,yS2+14).formula = "=" & sCol & yS2+9 &"-"& sCol & yS2+14  ' Importo netto
    
    # ~oS2.getCellByPosition(xS2+nSal,yS2+15).formula = "=" & sCol & yS2+15 &"*$C$84" ' Ritenute per garanzia
    # ~oS2.getCellByPosition(xS2+nSal,yS2+16).formula = "=" & sCol & yS2+15 &"*$C$85" ' Ritenute per infortuni
    # ~oS2.getCellByPosition(xS2+nSal,yS2+17).formula = "=" & sCol & yS2+15 &"*$C$80" ' Recupero anticipazioni
    # ~oS2.getCellByPosition(xS2+nSal,yS2+18).formula = "=subtotal(9;" & ColumnNameOf(xS2+1)& yS2+19 &":"& sCol & yS2+19 &")" ' Certificati precedenti / Ultimo riporto
    # ~oS2.getCellByPosition(xS2+nSal,yS2+19).formula = "=SUM("& sCol & yS2+16 &":"& sCol & yS2+19 & ")"' totale detrazioni
    # ~sFormula = "="& sCol & yS2+15 &"-"& sCol & yS2+20' & "+" & sCol & yS2+5 & "+" & sCol & yS2+6 & "+" & sCol & yS2+10 & "+" & sCol & yS2+11
    # ~oS2.getCellByPosition(xS2+nSal,yS2+20).formula = sFormula ' Importo Certificato di pagamento
    
# ~'    oS2.getCellByPosition(xS2+nSal,yS2+8).formula = "=(ROUND(SUBTOTAL(9;"& ColumnNameOf(xS2+nSal) & yS2+6 &":"& ColumnNameOf(xS2+nSal) & yS2+9 &");2)"
# ~'    oS2.getCellByPosition(xS2+nSal,yS2+10).value = SALsicurezza ' di cui importo per la sicurezza
# ~'    oS2.getCellByPosition(xS2+nSal,yS2+11).value = SALmdo ' di cui importo per la mdo
# ~'#########################################################################

# ~initRegistro: # INIZIO COMPILAZIONE REGISTRO
# ~'GoTo initSal: ' salto direttamente al SAL evitando il registro se già presente nel documento - solo debug

    # ~oSheetReg= thisComponent.Sheets.getByName("Registro")
    # ~ThisComponent.CurrentController.Select(oSheetReg)
    # ~ScriptPy("LeenoBasicBridge.py","setTabColor", 16769505)
    
    
    # ~ThisComponent.CurrentController.Select(oSheetReg.getCellRangeByPosition(0,0,0,0))
    # ~'fcol = 0 'prima colonna inserimento in Registro
# RIGA RIPORTO
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow).setSTRING("R I P O R T O")
    
    
    # ~oSheetReg.getCellByPosition(fcol+8, InsRow).setformula("=IF(ROUND(SUBTOTAL(9;$"& ColumnNameOf(fcol+8) &"$2:$"& ColumnNameOf(fcol+8) &"$" & InsRow & ");2)=0;"""";(ROUND(SUBTOTAL(9;$"& ColumnNameOf(fcol+8) &"$2:$"& ColumnNameOf(fcol+8) &"$" & InsRow & ");2))")
    # ~oSheetReg.getCellByPosition(fcol+9, InsRow).setformula("=IF(ROUND(SUBTOTAL(9;$"& ColumnNameOf(fcol+9) &"$2:$"& ColumnNameOf(fcol+9) &"$" & InsRow & ");2)=0;"""";(ROUND(SUBTOTAL(9;$"& ColumnNameOf(fcol+9) &"$2:$"& ColumnNameOf(fcol+9) &"$" & InsRow & ");2))")
    # ~oSheetReg.getCellRangeByPosition (fcol+0,InsRow,fcol+9,InsRow).CellStyle = "Ultimus_Bordo_sotto"
    # ~InsRow=InsRow+1
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow).setSTRING("LAVORI A MISURA")
    # ~oSheetReg.getCellRangeByPosition (fcol+0,InsRow,fcol+9,InsRow).CellStyle = "Ultimus_centro_bordi_lati"
    # ~InsRow=InsRow+1
# ----------------------------------------------------------------------
# compilo il REGISTRO
    # ~for each el in elVoci()
# ~'        for each n in el()
            # ~oSheetReg.getCellByPosition(fcol, InsRow).setstring(el(0)+ chr(13) + el(1)+ chr(13) + el(2))' num, art, data
# ~'            oSheetReg.getCellByPosition(fcol+1, InsRow).setstring(el(1)+ chr(13) + el(2))
            # ~oSheetReg.getCellByPosition(fcol+1, InsRow).setstring(el(3)) 'descrizione
            # ~oSheetReg.getCellByPosition(fcol+2, InsRow).setvalue(el(5)) 'Nlib
            # ~oSheetReg.getCellByPosition(fcol+3, InsRow).setvalue(el(6)) 'Plib
            # ~oSheetReg.getCellByPosition(fcol+4, InsRow).setstring(el(4)) 'um
            # ~if el(7)>0 then
                # ~oSheetReg.getCellByPosition(fcol+5, InsRow).setvalue(el(7)) 'quantità
                # ~else
                # ~oSheetReg.getCellByPosition(fcol+6, InsRow).setvalue(el(7)) 'quantità in meno
            # ~end if
            # ~oSheetReg.getCellByPosition(fcol+7, InsRow).setvalue(el(8)) 'prezzo
# gli importi vanno tutti nella colonna DEBITO, anche se negativi
            # ~oSheetReg.getCellByPosition(fcol+8, InsRow).setvalue(el(9)) 'debito
            # ~InsRow=InsRow+1
    # ~Next
    # ~InsRow=InsRow+1
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow).setstring("Parziale dei Lavori a Misura €")
    # ~ncol=ColumnNameOf(fcol+8)
    # ~oSheetReg.getCellByPosition(fcol+8, InsRow).setformula("=SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & InsRow & ")")
    # .formula o .setformula() è uguale
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow+2).formula = "=CONCATENATE(""Lavori a tutto il "";TEXT(NOW();""DD/MM/YYYY"");"" - T O T A L E   €"")"

    # ~ThisComponent.CurrentController.Select(oSheetReg.getCellByPosition(fcol+1, InsRow+2))
# ~copy_clip
# ~consolida_clip ' consolida la data
    # ~ncol=ColumnNameOf(fcol+8)
    # ~unSelect 'deseleziona
    # ~oSheetReg.getCellByPosition(fcol+8, InsRow+2).setformula("=SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & InsRow+2 & ")")
    
# ~'    fineMisure = InsRow
    # ~inizioFirme = InsRow+3
# ~firme (inizioFirme)
# ~'print
    # ~fineFirme = getLastUsedRow(oSheetReg)
    # ~If precRowReg<fRow Then precRowReg =fRow
# ----------------------------------------------------------------------
# set area del REGISTRO
# ~ncol=ColumnNameOf(fcol+9)
    # ~area="$A$" & precRowReg+2 & ":$"& ncol &"$"&fineFirme+1
    # ~nomearea = "#Reg#" & nSal
    # ~ScriptPy("LeenoBasicBridge.py","rifa_nomearea", ThisComponent, "Registro", area , nomearea)
# set area di stampa
        # ~oNamedRange=oRanges.getByName(nomearea).referredCells
        # ~With oNamedRange.RangeAddress
            # ~daRiga = .StartRow
            # ~aRiga = .EndRow
            # ~daColonna = .StartColumn
            # ~aColonna = .EndColumn
        # ~End With

        # ~ThisComponent.CurrentController.setFirstVisibleRow(daRiga)
# ----------------------------------------------------------------------
# ~rem    gli do il colore REGISTRO
        # ~oSheetReg.getCellRangeByPosition(fcol+0, daRiga+2, fcol+1, InsRow).cellstyle = "List-stringa-sin"    'descr.
        # ~oSheetReg.getCellRangeByPosition(fcol+2, daRiga+2, fcol+4, InsRow).cellstyle = "List-num-centro"    'Lib. N. P.
        # ~oSheetReg.getCellRangeByPosition(fcol+5, daRiga+2, fcol+6, InsRow).cellstyle = "comp 1a"            'quant
        # ~oSheetReg.getCellRangeByPosition(fcol+7, daRiga+2, fcol+9, InsRow).cellstyle = "List-num-euro"        'soldi
# ----------------------------------------------------------------------
# area di stampa
    # ~Dim selArea(0) as new com.sun.star.table.CellRangeAddress
        # ~selArea(0).StartColumn = daColonna
        # ~selArea(0).StartRow = daRiga
        # ~selArea(0).EndColumn = aColonna
        # ~selArea(0).EndRow = aRiga
# ~'        xray selArea()
# ~'        xxx() = oNamedRange.RangeAddress()
# ~'        xray xxx()
# set intestazione area di stampa
        # ~oTitles = createUnoStruct("com.sun.star.table.CellRangeAddress")
        # ~oTitles.startRow = 2 ' riga dell'intestazione
        # ~oTitles.EndRow = 2 
        # ~oTitles.startColumn = daColonna
        # ~oTitles.EndColumn = aColonna
        # ~oSheetReg.setTitleRows(oTitles)
        # ~oSheetReg.setPrintareas(selArea())
        # ~oSheetReg.setPrintTitleRows(True)
# ----------------------------------------------------------------------
# ~Visualizza_PageBreak
        # ~fineFirme = getLastUsedRow(oSheetReg)

    # ~oSheetReg.getCellRangeByPosition (fcol+0,inizioFirme-4,fcol+9,inizioFirme).CellStyle = "ULTIMUS"
# sistemo i totali REGISTRO
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow).CellStyle = "Ultimus_destra"
    # ~oSheetReg.getCellByPosition(fcol+1, InsRow+2).CellStyle = "Ultimus_destra"
    # ~oSheetReg.getCellByPosition(fcol+8, InsRow).CellStyle = "Ultimus_destra_totali"
    # ~oSheetReg.getCellByPosition(fcol+8, InsRow+2).CellStyle = "Ultimus_destra_totali"

# il settaggio degli stili, messo qui e ripetuto qualche riga sotto, serve a regolare bene l'altezza delle celle
# ~adatta_altezza: 
    # ~oCell=oSheetReg.getCellRangeByPosition(fcol+0, daRiga, fcol+9, fineFirme)
    # ~ThisComponent.CurrentController.Select(oCell)
# ~'    ScriptPy("LeenoBasicBridge.py","adatta_altezza_riga")

# ~i=1
    # ~Do While oSheetReg.getCellByPosition(fcol+0,fineFirme).rows.IsStartOfNewPage = False
# ~'        oSheetReg.getCellByPosition(fcol+1 , fineFirme).setstring("Sto sistemando il Registro...")
        # ~insRows (fineFirme,1) 'insertByIndex non funziona
        # ~If i=3 Then
            # ~oSheetReg.getCellByPosition(fcol+1, fineFirme).setstring("====================")
            # ~daqui=fineFirme
        # ~End If
        # ~fineFirme = fineFirme+1
        # ~i=i+1
    # ~Loop
    # ~fineFirme = fineFirme-1
# ~'    ThisComponent.CurrentController.Select(oSheetReg.getCellByPosition(fcol+1, daqui))
# ~'copy_clip
# ~'    ThisComponent.CurrentController.Select(oSheetReg.getCellRangeByPosition(fcol+1, daqui, fcol+1, fineFirme-2))
# ~'ScriptPy("LeenoBasicBridge.py","paste_clip")
    # ~area="$A$" & precRowReg+2 & ":$J$"&getLastUsedRow(oSheetReg)'-1
# ~ScriptPy("LeenoBasicBridge.py","rifa_nomearea", ThisComponent, "Registro", area , nomearea)

# ~'    oCell=oSheetReg.getCellRangeByPosition(0,precRowReg+1,11,getLastUsedRow(oSheetReg))
# ~'    ThisComponent.CurrentController.Select(oCell)
# ~'Raggruppa_righe
# ~'MOSTRA_righe ("off")
# ~'    ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener) 'deseleziona
# ~'    oCell.Rows.IsVisible=True ' lascia aperto il gruppo

    # ~oSheetReg.rows.removeByIndex (getLastUsedRow(oSheetReg), 1)
    # ~oSheetReg.rows.removeByIndex (getLastUsedRow(oSheetReg), 1)

# LA RIPETIZIONE DEL SETTAGGIO DEGLI STILI E' VOLUTA - VEDI # DI SOPRA
    # ~oSheetReg.getCellRangeByPosition (fcol+0,inizioFirme,fcol+9,fineFirme-1).CellStyle = "Ultimus_centro_bordi_lati"
    # ~oSheetReg.getCellByPosition(fcol+1 , inizioFirme+1).CellStyle = "ULTIMUS"
    # ~oSheetReg.getCellByPosition(fcol+1 , inizioFirme+11).CellStyle = "ULTIMUS"
    # ~oSheetReg.getCellRangeByPosition (fcol+0,inizioFirme,fcol+9,inizioFirme).CellStyle = "ULTIMUS"
# ULTIMA RIGA:
    # ~oSheetReg.getCellByPosition(fcol+1, fineFirme-1).setSTRING("A   R I P O R T A R E")
    # ~ncol=ColumnNameOf(fcol+8)
    # ~oSheetReg.getCellByPosition(fcol+8, fineFirme-1).setformula("=IF(ROUND(SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & fineFirme-1 & ");2)=0;"""";(ROUND(SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & fineFirme-1 & ");2))")
    # ~ncol=ColumnNameOf(fcol+9)
    # ~oSheetReg.getCellByPosition(fcol+9, fineFirme-1).setformula("=IF(ROUND(SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & fineFirme-1 & ");2)=0;"""";(ROUND(SUBTOTAL(9;$"& ncol &"$2:$"& ncol &"$" & fineFirme-1 & ");2))")
    # ~oSheetReg.getCellRangeByPosition (fcol+0,fineFirme-1,fcol+9,fineFirme-1).CellStyle = "Ultimus_Bordo_sotto"
    
# trovo l'ultimo effettivo numero di pagina 
    # ~If inumPag =0 Then    inumPag = 1
    # ~For i = precRowReg+1 to getLastUsedRow(oSheetReg)
        # ~if oSheetReg.getCellByPosition(0,i).rows.IsStartOfNewPage = True then 
            # ~inumPag = inumPag+1
        # ~end If
    # ~Next
# ~inumPag = inumPag-1 'ultimo numero pagina ESCLUSA la copertina
# ~'    Print inumPag
# annoto ultimo numero di pagina 
# ~'    oSheetReg.getCellByPosition(fcol+10 , fineFirme-1).value = inumPag-1
# ~'    oSheetReg.getCellByPosition(fcol+10 , fineFirme-1).CellStyle = "num centro"
# ~'end Sub
# ~'Sub GENERA_SAL
# ----------------------------------------------------------------------
# ~barra_chiudi
# ----------------------------------------------------------------------
# inserisco la prima riga GIALLA del REGISTRO
# ~'    Print "GIALLO REGISTRO"
    # ~oNamedRange=oRanges.getByName(nomearea).referredCells
    # ~ins = oNamedRange.RangeAddress.StartRow
    # ~insRows (ins, 1) 'insertByIndex non funziona
    # ~oSheetReg.getCellRangeByPosition (0,ins,10,ins).CellStyle = "uuuuu" '"Ultimus_Bordo_sotto"
# ~fissa (0, ins+1)
# ----------------------------------------------------------------------
# ci metto un po' di informazioni
    # ~davoce=elVoci(0)(0) 'ultima voce
    # ~avoce=elVoci(ubound(elvoci()))(0) 'ultima voce
    # ~oSheetReg.getCellByPosition(1,ins).string = "segue Registro n." & nSal & " - " & davoce & "÷" & avoce
    # ~oSheetReg.getCellByPosition(2,ins).value= nLib '1 ' numero libretto
    # ~oSheetReg.getCellByPosition(3,ins).value = inumPag 'ultimo numero pagina
# ~'    oSheetReg.getCellByPosition(23, ins).value= nSal ' numero SAL
    # ~oSheetReg.getCellByPosition(8, ins).formula = "=SUBTOTAL(9;I" & primariga+1 & ":I" & fineFirme & ")"
    # ~oSheetReg.getCellByPosition(8, ins).cellstyle = "comp sotto Euro 3_R"
# ----------------------------------------------------------------------
# ~Struttura_contab ("#Reg#")
# ~'Struttura
# ----------------------------------------------------------------------
# ~Ripristina_statusLine 
# ~'Exit Sub ' mi fermo al registro - solo debug
# ~initSal:
# ~'Annulla_atto_contabile("SAL")
# ~'Exit sub
        # ~Barra_Apri_Chiudi_5("                         Sto elaborando Stato di Avanzamento Lavori...", 75)
    # ~If oRanges.hasByName("#SAL#" & nSal) Then
        # ~msgbox     "SAL n." & nSal & " già emesso.",48, "AVVISO!"
        # ~Exit Sub
    # ~End If 
# ######################################################################
# ######################################################################
# ######################################################################
# ####################### INIZIO COMPILAZIONE SAL ######################
# ######################################################################
# ######################################################################
# ######################################################################

    # ~Select Case nSal
        # ~Case = 1
        # ~If not thisComponent.Sheets.hasByName("SAL") Then ' se la sheet NON esiste
            # ~thisComponent.getSheets().insertNewByName("SAL",6,0) ' ricreala ALLA POSIZIONE 5
        # ~End If
            # ~oSheetSAL= thisComponent.Sheets.getByName("SAL") ' setta come corrente
            # ~oSheetSAL.PageStyle = "PageStyle_REGISTRO_A4" ' imposta stile pagina
            # ~ThisComponent.CurrentController.Select(oSheetSAL) ' seleziona la tag

            # ~ScriptPy("LeenoBasicBridge.py","setTabColor", 16769505)
            # ~ThisComponent.CurrentController.Select(oSheetSAL.getCellRangeByPosition(0,0,0,0)) ' focus sulla prima cella SAL
# ~'Annulla_atto_contabile("SAL")
# ----------------------------------------------------------------------
# idxCol e idxRow sono costanti settate nel modulo "_variabili"
# ----------------------------------------------------------------------
# riga di intestazione SAL
            # ~oSheetSAL.getCellRangeByPosition(idxCol+0,idxRow,7,idxRow).CellStyle="An.1v-Att Start" 'STILE
            # ~oSheetSAL.getCellByPosition(idxCol+0,idxRow).setstring("N. ord."+ chr(13) +"Articolo")
            # ~oSheetSAL.getCellByPosition(idxCol+1,idxRow).setstring("LAVORAZIONI"+ chr(13) + "E SOMMINISTRAZIONI")
            # ~oSheetSAL.getCellByPosition(idxCol+2,idxRow).setstring("U.M.")
            # ~oSheetSAL.getCellByPosition(idxCol+3,idxRow).setstring("Quantità")
            # ~oSheetSAL.getCellByPosition(idxCol+4,idxRow).setstring("Prezzo" + chr(13) + "unitario")
            # ~oSheetSAL.getCellByPosition(idxCol+5,idxRow).setstring("Importo")
            # ~oSheetSAL.getCellByPosition(idxCol+6,idxRow).setstring("Num." + chr(13) +"Pag.")
# larghezza colonne SAL
# ~'xray oSheetSAL
            # ~oSheetSAL.getCellByPosition(idxCol,idxRow).Columns.Width = 1600 'N. ord.
            # ~oSheetSAL.getCellByPosition(idxCol+1,idxRow).Columns.Width = 10100 'LAVORAZIONI
            # ~oSheetSAL.getCellByPosition(idxCol+2,idxRow).Columns.Width = 1500 'U.M.
            # ~oSheetSAL.getCellByPosition(idxCol+3,idxRow).Columns.Width = 1800 'Quantità
            # ~oSheetSAL.getCellByPosition(idxCol+4,idxRow).Columns.Width = 1400 'Prezzo
            # ~oSheetSAL.getCellByPosition(idxCol+5,idxRow).Columns.Width = 1900 'Importo
            # ~oSheetSAL.getCellByPosition(idxCol+6,idxRow).Columns.OptimalWidth = True 'n.pag
    
        # ~frow = getLastUsedRow(oSheetSAL)'+1 ' trovo il primo rigo vuoto
        # ~Case >1
            # ~oSheetSAL= thisComponent.Sheets.getByName("SAL") ' SETTA COME CORRENTE
            # ~oNamedRange=oRanges.getByName("#SAL#" & nSal-1).referredCells
            # ~With oNamedRange.RangeAddress
                # ~fRow = .StartRow
                # ~lRow = .EndRow
                # ~precRowSAL = .EndRow
            # ~End With
# ~'        InsRow = precRowSAL+1'+idxrow
# ----------------------------------------------------------------------
# ~'#########################################################################
# chiudo il vecchio SAL
            # ~oCell = oSheetSAL.getCellRangeByPosition(0,fRow,6,lRow)
            # ~oCell.Rows.IsVisible=false ' chiudi gruppo
    # ~End Select
# ~'    ThisComponent.CurrentController.ZoomValue = 100

# ~'fissa (0,fRow+1)
# ~fissa (0,idxrow+1)
# ~'vociSAL(), array (art, desc, um,  quant, prezzo, importo, sicurezza, mdo)
# ~'Print "qui"
# ~'xray vociSAL()
# ----------------------------------------------------------------------
# inserimento dati SAL
    # ~fRow=precRowSAL'+1
# ~'Print "vai" & frow
    # ~If fRow < 2 Then frow=2
    # ~num = 1
# ~'    Print ubound (vociSAL())
    # ~for each el in vociSAL()
        # ~oSheetSAL.getCellByPosition(idxCol, frow+num).String = num & chr(13) & el(0)' num, art
        # ~oSheetSAL.getCellByPosition(idxCol+1, frow+num).String = el(1)' descrizione
        # ~oSheetSAL.getCellByPosition(idxCol+2, frow+num).String = el(2)' um
        # ~oSheetSAL.getCellByPosition(idxCol+3, frow+num).value = el(3)' quant
        # ~oSheetSAL.getCellByPosition(idxCol+4, frow+num).value = el(4)' prezzo
        # ~oSheetSAL.getCellByPosition(idxCol+5, frow+num).value = el(5)' importo
        # ~num=1+num
    # ~Next
    # ~dariga=frow
    # ~ariga=frow+num-1
# ----------------------------------------------------------------------
# ~rem    gli do il colore SAL
        # ~oSheetSAL.getCellRangeByPosition(idxCol+0, daRiga+1, idxCol+1, ariga).cellstyle = "List-stringa-sin"'descr.
        # ~oSheetSAL.getCellRangeByPosition(idxCol+2, daRiga+1, idxCol+2, ariga).cellstyle = "List-num-centro"    'u. m.
        # ~oSheetSAL.getCellRangeByPosition(idxCol+3, daRiga+1, idxCol+3, ariga).cellstyle = "comp 1a"            'quant
        # ~oSheetSAL.getCellRangeByPosition(idxCol+4, daRiga+1, idxCol+5, ariga).cellstyle = "List-num-euro"    'soldi
# ----------------------------------------------------------------------
    # ~InsRow=ariga+2
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow).setstring("Parziale dei Lavori a Misura €")
    # ~ncol=ColumnNameOf(fcol+5)
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow).setformula("=SUBTOTAL(9;$"& ncol &"$"& precRowSAL+1 &":$"& ncol &"$" & InsRow & ")")
    # ~Row_Misura=InsRow ' posizione che serve per la pagina di riepilogo
    # .formula o .setformula() è uguale
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+2).formula = "=CONCATENATE(""Lavori a tutto il "";TEXT(NOW();""DD/MM/YYYY"");"" - T O T A L E   €"")"

    # ~ThisComponent.CurrentController.Select(oSheetSal.getCellByPosition(fcol+1, InsRow+2))
# ~copy_clip
# ~consolida_clip ' consolida la data
    # ~ncol=ColumnNameOf(fcol+5)
    # ~unSelect 'deseleziona
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+2).setformula("=SUBTOTAL(9;$"& ncol &"$"& precRowSAL+1 &":$"& ncol &"$" & InsRow+2 & ")")

# ~'    inizioFirme = InsRow+7
# ~'firme (inizioFirme)
    # ~fineFirme = getLastUsedRow(oSheetSal)+2
    # ~If precRowSAL<fRow Then precRowSAL =fRow

# ----------------------------------------------------------------------
# set area del SAL
# ~ncol=ColumnNameOf(fcol+5)
    # ~area="$A$" & precRowSAL+2 & ":$"& ncol &"$"&fineFirme+1
    # ~ScriptPy("LeenoBasicBridge.py","rifa_nomearea", ThisComponent, "SAL", area , "#SAL#" & nSal)
# set area di stampa SAL
        # ~oNamedRange=oRanges.getByName("#SAL#" & nSal).referredCells
        # ~With oNamedRange.RangeAddress
            # ~daRiga = .StartRow
            # ~aRiga = .EndRow
            # ~daColonna = .StartColumn
            # ~aColonna = .EndColumn
        # ~End With
    # ~ThisComponent.CurrentController.setFirstVisibleRow(daRiga)
# ----------------------------------------------------------------------
# area di stampa
    # ~Dim selAreaSAL(0) as new com.sun.star.table.CellRangeAddress
        # ~selAreaSAL(0).StartColumn = daColonna
        # ~selAreaSAL(0).StartRow = daRiga
        # ~selAreaSAL(0).EndColumn = aColonna
        # ~selAreaSAL(0).EndRow = aRiga
# ~'        xray selArea()
# ~'        xxx() = oNamedRange.RangeAddress()
# ~'        xray xxx()
# set intestazione area di stampa
        # ~oTitlesSAL = createUnoStruct("com.sun.star.table.CellRangeAddress")
        # ~oTitlesSAL.startRow = 2' riga dell'intestazione
        # ~oTitlesSAL.EndRow = 2' riga dell'intestazione
        # ~oTitlesSAL.startColumn = daColonna
        # ~oTitlesSAL.EndColumn = aColonna
        # ~oSheetSal.setTitleRows(oTitlesSAL)
        # ~oSheetSal.setPrintareas(selAreaSAL())
        # ~oSheetSal.setPrintTitleRows(True)
# ----------------------------------------------------------------------
# ~Visualizza_PageBreak
    # ~fineFirme = getLastUsedRow(oSheetSal)+2
# ~'    Barra_Apri_Chiudi_5("                                           Sto sistemando il SAL...",0)
# sistemo i totali SAL
# il settaggio degli stili, messo qui e ripetuto qualche riga sotto,
# serve a regolare bene l'altezza delle celle e calcolare correttamente il salto pagina
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow).CellStyle = "Ultimus_destra"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow).CellStyle = "Ultimus_destra_totali"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+2).CellStyle = "Ultimus_destra"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+2).CellStyle = "Ultimus_destra_totali"
    # ~i=1
    # ~Do While oSheetSal.getCellByPosition(fcol+0,fineFirme).rows.IsStartOfNewPage = False
# ~'        oSheetSal.getCellByPosition(fcol+1 , fineFirme).setstring("Sto sistemando il SAL...")
        # ~insRows (fineFirme,1) 'insertByIndex non funziona
            # ~oSheetSal.getCellByPosition(fcol+1, fineFirme).setstring("====================")
        # ~fineFirme = fineFirme+1
        # ~i=i+1
# ~'Barra_Apri_leggera
# ~'                Barra_Apri_Chiudi_5("                                           "& _
# ~'                 oSheetSal.getCellByPosition(0,fineFirme).rows.IsStartOfNewPage, 0)
    # ~Loop
# ~'Print InsRow
# ~'Print fineFirme
    # ~oSheetSal.getCellRangeByPosition (fcol+0,InsRow-1,fcol+5,fineFirme-1).CellStyle = "Ultimus_centro_bordi_lati"
    # ~oSheetSal.getCellRangeByPosition (fcol+0,fineFirme-1,fcol+5,fineFirme-1).CellStyle = "comp Descr"

# sistemo i totali SAL
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow).CellStyle = "Ultimus_destra"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow).CellStyle = "Ultimus_destra_totali"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+2).CellStyle = "Ultimus_destra"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+2).CellStyle = "Ultimus_destra_totali"
# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
# pagina RIEPILOGO SAL
# ----------------------------------------------------------------------
    # ~insRow=getLastUsedRow(oSheetSal)+1 ' SERVE PER PROSEGUIRE CON LA PAGINA DI RIEPILOGO

    # ~ThisComponent.CurrentController.Select(oSheetSal.getCellByPosition(fcol+1, fineFirme-1))
    # ~ThisComponent.CurrentController.setFirstVisibleRow (fineFirme)
# ~cancella_dati
    # ~unSelect 'deseleziona
# LA RIPETIZIONE DEL SETTAGGIO DEGLI STILI E' VOLUTA - VEDI # DI SOPRA
    # ~insRows (fineFirme,1) 'insertByIndex non funziona
    # ~oSheetSal.getCellRangeByPosition (fcol+0,fineFirme,fcol+5,fineFirme).CellStyle = "Ultimus_centro_bordi_lati"
    # ~fineFirme = fineFirme+1
# ~i=1
    # ~Do While oSheetSal.getCellByPosition(fcol+0,fineFirme).rows.IsStartOfNewPage = False
# ~'        oSheetSal.getCellByPosition(fcol+1 , fineFirme).setstring("Sto sistemando il SAL...")
        # ~insRows (fineFirme,1) 'insertByIndex non funziona
    # ~'        oSheetSal.getCellByPosition(fcol+1, fineFirme).setstring("====================")
        # ~fineFirme = fineFirme+1
        # ~i=i+1
    # ~Loop 
    # ~oSheetSal.getCellRangeByPosition (fcol+0,fineFirme,fcol+5,fineFirme).CellStyle = "comp Descr"
    # ~fineFirme=fineFirme-1
    # ~oSheetSal.rows.removeByIndex (fineFirme-1, 2)
# ----------------------------------------------------------------------
# Pagina di Riepilogo
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+1).CellStyle = "Ultimus_centro_Dsottolineato"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+1).setstring("R I E P I L O G O   S A L")
# ----------------------------------------------------------------------
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+3, fcol+1, InsRow+4).CellStyle = "Ultimus_sx_italic"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+3).setstring("Appalto: a misura")
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+4).setstring("Offerta: unico ribasso")
# ----------------------------------------------------------------------
#     IMPOSTA LA COLONNA DEI VALORI
    # ~oSheetSal.getCellRangeByPosition (fcol+5, InsRow+6,fcol+5, InsRow+15).CellStyle = "ULTIMUS"
# ----------------------------------------------------------------------
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+6).CellStyle = "Ultimus_sx_bold"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+6).setstring("Lavori a Misura €")
# ----------------------------------------------------------------------
    # ~ncol=ColumnNameOf(fcol+5)
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+7, fcol+1, InsRow+8).CellStyle = "Ultimus_sx"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+6).formula = "=$"& ncol &"$" & Row_Misura+1
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+7).setstring("Di cui importo per la Sicurezza")
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+7).value= SALsicurezza*-1

    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+8).setstring("Di cui importo per la Manodopera")
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+8).CellStyle = "Ultimus_"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+8).value= SALmdo*-1
# ----------------------------------------------------------------------
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+9, fcol+1, InsRow+10).CellStyle = "Ultimus_destra"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+9).string= "Importo dei Lavori a Misura su cui applicare il ribasso"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+9).formula= "=SUM(" & ncol & InsRow+7 & ":" & ncol & InsRow+9 &")"

    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+10).formula= _
    # ~"=CONCATENATE(""RIBASSO del "";TEXT($S2.$C$78*100;""#.##0,00"");""%"")"

    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+10).formula= "=-"& ncol & InsRow+10 &"*$S2.$C$78" ' RIBASSO
# ----------------------------------------------------------------------
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+11, fcol+1, InsRow+12).CellStyle = "Ultimus_sx"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+11).setstring("Importo per la Sicurezza")
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+11).value= SALsicurezza

    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+12).setstring("Importo per la Manodopera")
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+12).CellStyle = "Ultimus_"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+12).value= SALmdo
# ----------------------------------------------------------------------
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+13, fcol+1, InsRow+13).CellStyle = "Ultimus_destra_bold"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+13).string= "PER I LAVORI A MISURA €"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+13).formula= "=SUM(" & ncol & InsRow+10 & ":" & ncol & InsRow+13 &")"
# ----------------------------------------------------------------------
# IL TOTALE ANDRA' RISISTEMATO QUANDO SARANNO PRONTE LE ALTRE MODALITA' DI APPALTO: IN ECONOMIA E A CORPO
    # ~oSheetSal.getCellRangeByPosition(fcol+1, InsRow+15, fcol+1, InsRow+15).CellStyle = "Ultimus_destra_bold"
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+15).string= "T O T A L E  €"
    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+15).CellStyle = "Ultimus_destra_totali"

    # ~oSheetSal.getCellByPosition(fcol+5, InsRow+15).formula= "=SUM(" & ncol & InsRow+10 & ":" & ncol & InsRow+13 &")"
# ----------------------------------------------------------------------
# ~firme (InsRow+17)


    # ~seleziona_area("#SAL#" & nSal)
# ~'Raggruppa_righe
# ~'MOSTRA_righe ("off")
    # ~oSheetSal.getCellByPosition(fcol+1, InsRow+27).setstring("====================")
    # ~ThisComponent.CurrentController.Select(oSheetSal.getCellByPosition(fcol+1, InsRow+27))
# ~copy_clip
    # ~ThisComponent.CurrentController.Select(oSheetSal.GetCellRangeByPosition(fcol+1, InsRow+27, fcol+1, getLastUsedRow(oSheetSal)-1))
    # ~ScriptPy("LeenoBasicBridge.py","paste_clip", "0") 'sovrappone dati
    # ~oSheetSal.rows.removeByIndex (InsRow+27,1)
    # ~ThisComponent.currentController.removeRangeSelectionListener(oRangeSelectionListener) 'deseleziona
# ~Ripristina_statusLine 'Barra_chiudi_sempre_4
# ~'RiDefinisci_Area_Elenco_prezzi ' non capisco come mai l'area di elenco_prezzi viene sminchiata succede con LIBREOFFICE 4.3.*
# ~'Exit Sub
# ----------------------------------------------------------------------

# trovo l'ultimo effettivo numero di pagina del sal
    # ~inumPag =0
    # ~For i = precRowSAL+1 to getLastUsedRow(oSheetSal)
        # ~if oSheetSal.getCellByPosition(0,i).rows.IsStartOfNewPage = True then 
            # ~inumPag = inumPag+1
        # ~end If
    # ~Next
# ~'inumPag = inumPag-1 'ultimo numero pagina ESCLUSA la copertina"
# ----------------------------------------------------------------------

# inserisco la prima riga del documento
    # ~oNamedRange=oRanges.getByName("#SAL#" & nSal).referredCells
# ~'    oNamedRange=oRanges.getByName(nomearea).referredCells
    # ~ins = oNamedRange.RangeAddress.StartRow
    # ~insRows (ins, 1) 'insertByIndex non funziona
    # ~oSheetSal.getCellRangeByPosition (0,ins,10,ins).CellStyle = "uuuuu" '"Ultimus_Bordo_sotto"
# ~fissa (0, ins + 1)
# ~'print
# ----------------------------------------------------------------------
# ci metto un po' di informazioni
    # ~davoce=elVoci(0)(0) 'ultima voce
    # ~avoce=elVoci(ubound(elvoci()))(0) 'ultima voce
    # ~oSheetSal.getCellByPosition(1,ins).string = "segue Stato di Avanzamento Lavori n." & nSal & " - " & davoce & "÷" & avoce
    # ~oSheetSal.getCellByPosition(6,ins).value = inumPag 'ultimo numero pagina
    # ~oSheetSal.getCellByPosition(fcol+5, ins).setformula("=SUBTOTAL(9;$"& ncol &"$"& precRowSAL+1 &":$"& ncol &"$" & InsRow & ")")
    # ~oSheetSal.getCellByPosition(fcol+5, ins).cellstyle = "comp sotto Euro 3_R"
    # ~ThisComponent.CurrentController.Select(oSheetSal.getCellByPosition(fcol, ins))
# ----------------------------------------------------------------------
# ~Struttura_Contab ("#SAL#")
# ~end Sub 'genera_registro
    return


########################################################################
def GeneraAttiContabili():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
    if oSheet.Name != "CONTABILITA":
        return
    if Dialogs.YesNoDialog(Title='Avviso',
        Text= '''Prima di procedere è consigliabile salvare il lavoro.

Se decidi di continuare, devi attendere il messaggio di procedura completata.

Procedo senza salvare?''') == 0:
        return
    try:
        nSal, daVoce, aVoce, fpRow, lpRow = GeneraLibretto(oDoc)
    except:
        return
    GeneraRegistro(oDoc)
    DLG.chi((nSal, daVoce, aVoce, daRiga, aRiga))
    Dialogs.Info(Title = 'Voci registrate!',
                 Text='''La generazione degli allegati contabili è stata completata.
Grazie per l'attesa.''')


# CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA ## CONTABILITA #
########################################################################
########################################################################
g_exportedScripts = GeneraAttiContabili