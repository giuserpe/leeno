'''
Funzioni di utilità per la manipolazione dei fogli
relativamente alle funzionalità specifiche di LeenO
'''
import uno
from com.sun.star.sheet.CellFlags import HARDATTR, EDITATTR, FORMATTED

import pyleeno as PL
import LeenoUtils
import LeenoSheetUtils
import SheetUtils
import LeenoAnalysis
import LeenoComputo
import Dialogs


def ScriviNomeDocumentoPrincipaleInFoglio(oSheet):
    '''
    Indica qual è il Documento Principale
    nell'apposita area del foglio corrente
    '''
    # legge il percorso del documento principale
    sUltimus = LeenoUtils.getGlobalVar('sUltimus')

    # dal foglio risale al documento proprietario
    oDoc = SheetUtils.getDocumentFromSheet(oSheet)

    # se si sta lavorando sul Documento Principale, non fa nulla
    try:
        if sUltimus == uno.fileUrlToSystemPath(oDoc.getURL()):
            return
    except Exception:
        # file senza nome
        return

    d = {
        'COMPUTO': 'F1',
        'VARIANTE': 'F1',
        'Elenco Prezzi': 'A1',
        'CONTABILITA': 'F1',
        'Analisi di Prezzo': 'A1'
    }
    cell = d.get(oSheet.Name)
    if cell is None:
        return

    oSheet.getCellRangeByName(cell).String = 'DP: ' + sUltimus
    oSheet.getCellRangeByName("A1:AT1").clearContents(EDITATTR + FORMATTED + HARDATTR)

# ###############################################################

def SbiancaCellePrintArea():
    '''
    area 
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
    
    oPrintArea = oSheet.getPrintAreas()

    oSheet.getCellRangeByPosition(
        oPrintArea[0].StartColumn, oPrintArea[0].StartRow,
        oPrintArea[0].EndColumn, oPrintArea[0].EndRow
        ).CellBackColor = 16777215 #sbianca
    return

# ###############################################################

def setVisibilitaColonne(oSheet, sValori):
    '''
    sValori { string } : una stringa di configurazione della visibilità colonne
    permette di visualizzare/nascondere un set di colonne
    T = visualizza
    F = nasconde
    '''
    n = 0
    for el in sValori:
        oSheet.getCellByPosition(n, 2).Columns.IsVisible = True if el == 'T' else False
        n += 1

# ###############################################################

def setLarghezzaColonne(oSheet):
    '''
    regola la larghezza delle colonne a seconda della sheet
    '''
    if oSheet.Name == 'Analisi di Prezzo':
        for col, width in {'A':2100, 'B':12000, 'C':1600, 'D':2000, 'E':3400, 'F':3400,
                           'G':2700, 'H':2700, 'I':2000, 'J':2000, 'K':2000}.items():
            oSheet.Columns[col].Width = width
        SheetUtils.freezeRowCol(oSheet, 0, 2)

    elif oSheet.Name == 'CONTABILITA':
        setVisibilitaColonne(oSheet, 'TTTFFTTTTTFTFTFTFTFTTFTTFTFTTTTFFFFFF')
        # larghezza colonne importi
        oSheet.getCellRangeByPosition(13, 0, 1023, 0).Columns.Width = 1900
        # larghezza colonne importi
        oSheet.getCellRangeByPosition(19, 0, 23, 0).Columns.Width = 1000
        # nascondi colonne
        oSheet.getCellRangeByPosition(51, 0, 1023, 0).Columns.IsVisible = False

        for col, width in {'A':600, 'B':1500, 'C':6300, 'F':1300, 'G':1300,
                           'H':1300, 'I':1300, 'J':1700, 'L':1700, 'N':1900,
                           'P':1900, 'T':1000, 'U':1000, 'W':1000, 'X':1000,
                           'Z':1900, 'AC':1700, 'AD':1700, 'AE':1700,
                           'AX':1900, 'AY':1900}.items():
            oSheet.Columns[col].Width = width
        SheetUtils.freezeRowCol(oSheet, 0, 3)

    elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
        # mostra colonne
        oSheet.getCellRangeByPosition(5, 0, 8, 0).Columns.IsVisible = True

        oSheet.getColumns().getByName('A').Columns.Width = 600
        oSheet.getColumns().getByName('B').Columns.Width = 1500
        oSheet.getColumns().getByName('C').Columns.Width = 6300  # 7800
        oSheet.getColumns().getByName('F').Columns.Width = 1500
        oSheet.getColumns().getByName('G').Columns.Width = 1300
        oSheet.getColumns().getByName('H').Columns.Width = 1300
        oSheet.getColumns().getByName('I').Columns.Width = 1300
        oSheet.getColumns().getByName('J').Columns.Width = 1700
        oSheet.getColumns().getByName('L').Columns.Width = 1700
        oSheet.getColumns().getByName('S').Columns.Width = 1700
        oSheet.getColumns().getByName('AC').Columns.Width = 1700
        oSheet.getColumns().getByName('AD').Columns.Width = 1700
        oSheet.getColumns().getByName('AE').Columns.Width = 1700
        SheetUtils.freezeRowCol(oSheet, 0, 3)
        setVisibilitaColonne(oSheet, 'TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT')
    if oSheet.Name == 'Elenco Prezzi':
        oSheet.getColumns().getByName('A').Columns.Width = 1600
        oSheet.getColumns().getByName('B').Columns.Width = 10000
        oSheet.getColumns().getByName('C').Columns.Width = 1500
        oSheet.getColumns().getByName('D').Columns.Width = 1500
        oSheet.getColumns().getByName('E').Columns.Width = 1600
        oSheet.getColumns().getByName('F').Columns.Width = 1500
        oSheet.getColumns().getByName('G').Columns.Width = 1500
        oSheet.getColumns().getByName('H').Columns.Width = 1600
        oSheet.getColumns().getByName('I').Columns.Width = 1200
        oSheet.getColumns().getByName('J').Columns.Width = 1200
        oSheet.getColumns().getByName('K').Columns.Width = 100
        oSheet.getColumns().getByName('L').Columns.Width = 1600
        oSheet.getColumns().getByName('M').Columns.Width = 1600
        oSheet.getColumns().getByName('N').Columns.Width = 1600
        oSheet.getColumns().getByName('O').Columns.Width = 100
        oSheet.getColumns().getByName('P').Columns.Width = 1600
        oSheet.getColumns().getByName('Q').Columns.Width = 1600
        oSheet.getColumns().getByName('R').Columns.Width = 1600
        oSheet.getColumns().getByName('S').Columns.Width = 100
        oSheet.getColumns().getByName('T').Columns.Width = 1600
        oSheet.getColumns().getByName('U').Columns.Width = 1600
        oSheet.getColumns().getByName('V').Columns.Width = 1600
        oSheet.getColumns().getByName('W').Columns.Width = 100
        oSheet.getColumns().getByName('X').Columns.Width = 1600
        oSheet.getColumns().getByName('Y').Columns.Width = 1600
        oSheet.getColumns().getByName('Z').Columns.Width = 1600
        oSheet.getColumns().getByName('AA').Columns.Width = 1600
        SheetUtils.freezeRowCol(oSheet, 0, 3)
    adattaAltezzaRiga(oSheet)

# ###############################################################
def rRow(oSheet):
    '''
    Restituisce la posizione della riga rossa
    '''
    nRow = SheetUtils.getLastUsedRow(oSheet) +10
    for n in reversed(range(0, nRow)):
        if oSheet.getCellByPosition(
                0,
                n).CellStyle == 'Riga_rossa_Chiudi':
            return n

def cercaUltimaVoce(oSheet):
    nRow = SheetUtils.getLastUsedRow(oSheet) +1
    if nRow == 0:
        return 0
    for n in reversed(range(0, nRow)):
        # if oSheet.getCellByPosition(0, n).CellStyle in('Comp TOTALI'):
        if oSheet.getCellByPosition(
                0,
                n).CellStyle in ('EP-aS', 'EP-Cs', 'An-sfondo-basso Att End',
                                 'Comp End Attributo', 'Comp End Attributo_R',
                                 'comp Int_colonna',
                                 'comp Int_colonna_R_prima',
                                 'Livello-0-scritta', 'Livello-1-scritta',
                                 'livello2 valuta'):
            break
    return n


# ###############################################################


def cercaPartenza(oSheet, lrow):
    '''
    oSheet      foglio corrente
    lrow        riga corrente nel foglio
    Ritorna il nome del foglio [0] e l'id della riga di codice prezzo componente [1]
    il flag '#reg' solo per la contabilità.
    partenza = (nome_foglio, id_rcodice, flag_contabilità)
    '''
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')

    # COMPUTO, VARIANTE
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili_computo:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        partenza = (oSheet.Name, sStRange.RangeAddress.StartRow + 1)

    # CONTABILITA
    elif oSheet.getCellByPosition(0, lrow).CellStyle in stili_contab:
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)

        try:
            partenza = (oSheet.Name, sStRange.RangeAddress.StartRow + 1,
                        oSheet.getCellByPosition(22,
                        sStRange.RangeAddress.StartRow + 1).String)
        except:
            lrow = 3
            partenza = (oSheet.Name, lrow, '')

    # ANALISI o riga totale
    elif oSheet.getCellByPosition(0, lrow).CellStyle in ('An-lavoraz-Cod-sx', 'Comp TOTALI'):
        partenza = (oSheet.Name, lrow)

    # nulla di quanto sopra
    else:
        partenza = (oSheet.Name, lrow, '')

    return partenza


# ###############################################################


def selezionaVoce(oSheet, lrow):
    '''
    Restituisce inizio e fine riga di una voce in COMPUTO, VARIANTE,
    CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga all'interno della voce
    '''
    if oSheet.Name in ('Elenco Prezzi'):
        return lrow, lrow

    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    elif oSheet.Name == 'Analisi di Prezzo':
        sStRange = LeenoAnalysis.circoscriveAnalisi(oSheet, lrow)
    ###
    elif oSheet.Name == 'CONTABILITA':
        partenza = cercaPartenza(oSheet, lrow)
        if partenza[2] == '#reg':
            PL.sblocca_cont()
            if LeenoUtils.getGlobalVar('sblocca_computo') == 0:
                return
            pass
        else:
            pass
        sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    else:
        raise

    SR = sStRange.RangeAddress.StartRow
    ER = sStRange.RangeAddress.EndRow
    # ~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(0, SR, 250, ER))
    return SR, ER

# ###############################################################

def prossimaVoce(oSheet, lrow, n=1):
    '''
    oSheet { obect }
    lrow { double }   : riga di riferimento
    n    { integer }  : se 0 sposta prima della voce corrente
                        se 1 sposta dopo della voce corrente
    sposta il cursore prima o dopo la voce corrente restituendo un idrow
    '''
    stili_cat = LeenoUtils.getGlobalVar('stili_cat')
    stili_computo = LeenoUtils.getGlobalVar('stili_computo')
    stili_contab = LeenoUtils.getGlobalVar('stili_contab')
    noVoce = LeenoUtils.getGlobalVar('noVoce')
    stili = stili_computo + stili_contab

    # ~lrow = PL.LeggiPosizioneCorrente()[1]
    if lrow == 0:
        while oSheet.getCellByPosition(0, lrow).CellStyle not in stili:
            lrow += 1
        return lrow
    fine = cercaUltimaVoce(oSheet) + 1
    # la parte che segue sposta il focus alla voce successiva
    if lrow >= fine:
        return lrow
    if oSheet.getCellByPosition(0, lrow).CellStyle in stili:
        if n == 0:
            sopra = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.StartRow
            lrow = sopra
        elif n == 1:
            sotto = LeenoComputo.circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow
            lrow = sotto + 1
    while oSheet.getCellByPosition(0, lrow).CellStyle in stili_cat:
        lrow += 1
    while oSheet.getCellByPosition(0, lrow).CellStyle in ('uuuuu', 'Ultimus_centro_bordi_lati'):
        lrow += 1
    return lrow
# ###############################################################

def eliminaVoce(oSheet, lrow):
    '''
    usata in PL.MENU_elimina_voci_azzerate()
    
    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    '''
    voce = selezionaVoce(oSheet, lrow)
    SR = voce[0]
    ER = voce[1]

    oSheet.getRows().removeByIndex(SR, ER - SR + 1)
    
def elimina_voce(lrow=None, msg=1):
    '''
    @@@ MODIFICA IN CORSO CON 'LeenoSheetUtils.eliminaVoce'
    Elimina una voce in COMPUTO, VARIANTE, CONTABILITA o Analisi di Prezzo
    lrow { long }  : numero riga
    msg  { bit }   : 1 chiedi conferma con messaggio
                     0 esegui senza conferma
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)

    if oSheet.Name == 'Elenco Prezzi':
        Dialogs.Info(Title = 'Info', Text="""Per eliminare una o più voci dall'Elenco Prezzi
devi selezionarle ed utilizzare il comando 'Elimina righe' di Calc.""")
        return

    if oSheet.Name not in ('COMPUTO', 'CONTABILITA', 'VARIANTE', 'Analisi di Prezzo'):
        return

    try:
        SR = PL.seleziona_voce()[0]
    except:
        return
    ER = PL.seleziona_voce()[1]
    if msg == 1:
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
            0, SR, 250, ER))
        if '$C$' in oSheet.getCellByPosition(9, ER).queryDependents(False).AbsoluteName:
            undo = 1
            PL._gotoCella(9, ER)
            PL.comando ('ClearArrowDependents')
            PL.comando ('ShowDependents')
            oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
                0, SR, 250, ER))
            messaggio= """
Da questa voce dipende almeno un Vedi Voce.
VUOI PROCEDERE UGUALMENTE?"""
        else:
            messaggio = """OPERAZIONE NON ANNULLABILE!\n
Stai per eliminare la voce selezionata.
            Voi Procedere?\n"""
        # ~return
        if Dialogs.YesNoDialog(Title='*** A T T E N Z I O N E ! ***',
            Text= messaggio) == 1:
            try:
                undo
                comando ('Undo')
            except: 
                pass
            oSheet.getRows().removeByIndex(SR, ER - SR + 1)
            PL._gotoCella(0, SR+1)
        else:
            oDoc.CurrentController.select(oSheet.getCellRangeByPosition(
                0, SR, 250, ER))
            return
    elif msg == 0:
        oSheet.getRows().removeByIndex(SR, ER - SR + 1)
    if oSheet.Name != 'Analisi di Prezzo':
        PL.numera_voci(0)
    else:
        PL._gotoCella(0, SR+2)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))

# ###############################################################

def inserisciRigaRossa(oSheet):
    '''
    Inserisce la riga rossa di chiusura degli elaborati nel foglio specificato
    Questa riga è un riferimento per varie operazioni
    Errore se il foglio non è un foglio di LeenO
    '''
    lrow = 0
    nome = oSheet.Name
    if nome in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        lrow = cercaUltimaVoce(oSheet) + 2
        for n in range(lrow, SheetUtils.getLastUsedRow(oSheet) + 2):
            if oSheet.getCellByPosition(0, n).CellStyle == 'Riga_rossa_Chiudi':
                return
        oSheet.getRows().insertByIndex(lrow, 1)
        oSheet.getCellByPosition(0, lrow).String = 'Fine Computo'
        oSheet.getCellRangeByPosition(0, lrow, 34, lrow).CellStyle = 'Riga_rossa_Chiudi'
        oSheet.getCellByPosition(2, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'
    elif nome == 'Analisi di Prezzo':
        lrow = cercaUltimaVoce(oSheet) + 2
        oSheet.getCellByPosition(0, lrow).String = 'Fine ANALISI'
        oSheet.getCellRangeByPosition(0, lrow, 10, lrow).CellStyle = 'Riga_rossa_Chiudi'
        oSheet.getCellByPosition(1, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'
    elif nome == 'Elenco Prezzi':
        lrow = cercaUltimaVoce(oSheet) + 1
        if oSheet.getCellByPosition(0, lrow).CellStyle != 'Riga_rossa_Chiudi':
            lrow += 1
        oSheet.getCellByPosition(0, lrow).String = 'Fine elenco'
        oSheet.getCellRangeByPosition(0, lrow, 9, lrow).CellStyle = 'Riga_rossa_Chiudi'
        oSheet.getCellRangeByPosition(11, lrow, 21, lrow).CellStyle = 'EP statistiche_Contab'
        oSheet.getCellRangeByPosition(23, lrow, 25, lrow).CellStyle = 'EP statistiche'
        oSheet.getCellRangeByPosition(26, lrow, 26, lrow).CellStyle = 'EP-mezzo %'
        s = str(lrow + 1)
        oSheet.getCellByPosition(12, lrow).String = 'TOTALE'
        oSheet.getCellByPosition(13, lrow).Formula = '=SUBTOTAL(9;N3:N' + s + ')'
        oSheet.getCellByPosition(16, lrow).String = 'TOTALE'
        oSheet.getCellByPosition(17, lrow).Formula = '=SUBTOTAL(9;R3:R' + s + ')'
        oSheet.getCellByPosition(20, lrow).String = 'TOTALE'
        oSheet.getCellByPosition(21, lrow).Formula = '=SUBTOTAL(9;V3:V' + s + ')'
        oSheet.getCellByPosition(23, lrow).String = 'TOTALE'
        oSheet.getCellByPosition(24, lrow).Formula = '=SUBTOTAL(9;Y3:Y' + s + ')'
        oSheet.getCellByPosition(25, lrow).Formula = '=SUBTOTAL(9;Z3:Z' + s + ')'
        oSheet.getCellByPosition(26, lrow).Formula = '=IFERROR(IFS(AND(N' + s + '>R' + s + ';R' + s + '=0);-1;AND(N' + s + '<R' + s + ';N' + s + '=0);1;N' + s + '=R' + s + ';"--";N' + s + '>R' + s + ';-(N' + s + '-R' + s + ')/N' + s + ';N'+ s + '<R' + s + ';-(N' + s + '-R' + s + ')/N' + s + ');"--")'
        oSheet.getCellByPosition(1, lrow
        ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'

# ###############################################################


def adattaAltezzaRiga(oSheet):
    '''
    Adatta l'altezza delle righe al contenuto delle celle.
    imposta l'altezza ottimale delle celle
    usata in PL.Menu_adattaAltezzaRiga()
    '''
    oDoc = LeenoUtils.getDocument()
    # ~oDoc = SheetUtils.getDocumentFromSheet(oSheet)
    if not oDoc.getSheets().hasByName('S1'):
        return

    usedArea = SheetUtils.getUsedArea(oSheet)
    oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows.OptimalHeight = True

    # DALLA VERSIONE 6.4.2 IL PROBLEMA è RISOLTO
    # DALLA VERSIONE 7 IL PROBLEMA è PRESENTE
    if float(PL.loVersion()[:5].replace('.', '')) >= 642:
        return

    # se la versione di LibreOffice è maggiore della 5.2
    # esegue il comando agendo direttamente sullo stile
    lista_stili = ('comp 1-a', 'Comp-Bianche in mezzo Descr_R',
                   'Comp-Bianche in mezzo Descr', 'EP-a',
                   'Ultimus_centro_bordi_lati')
    # NELLE VERSIONI DA 5.4.2 A 6.4.1
    if(
       float(PL.loVersion()[:5].replace('.', '')) > 520 or
       float(PL.loVersion()[:5].replace('.', '')) < 642):
        for stile_cella in lista_stili:
            try:
                oDoc.StyleFamilies.getByName("CellStyles").getByName(stile_cella).IsTextWrapped = True
            except Exception:
                pass

        test = usedArea.EndRow + 1

        for y in range(0, test):
            if oSheet.getCellByPosition(2, y).CellStyle in lista_stili:
                oSheet.getCellRangeByPosition(0, y, usedArea.EndColumn, y).Rows.OptimalHeight = True

    if oSheet.Name in ('Elenco Prezzi', 'VARIANTE', 'COMPUTO', 'CONTABILITA'):
        oSheet.getCellByPosition(0, 2).Rows.Height = 800
    if oSheet.Name == 'Elenco Prezzi':
        test = usedArea.EndRow + 1
        for y in range(0, test):
            oSheet.getCellRangeByPosition(0, y, usedArea.EndColumn, y).Rows.OptimalHeight = True
    return


# ###############################################################


def inserSuperCapitolo(oSheet, lrow, sTesto='Super Categoria'):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della categoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1

    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'Livello-0-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-0-scritta mini'
    oSheet.getCellRangeByPosition( 18, lrow, 18, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-0-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-0-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-0-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1


# ###############################################################


def inserCapitolo(oSheet, lrow, sTesto='Categoria'):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della categoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'

    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1
    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41, lrow).CellStyle = 'Livello-1-scritta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'Livello-1-scritta mini'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'Livello-1-scritta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'Livello-1-scritta mini val'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1


# ###############################################################


def inserSottoCapitolo(oSheet, lrow, sTesto):
    '''
    lrow    { double } : id della riga di inserimento
    sTesto  { string } : titolo della sottocategoria
    '''
    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        return
    if not sTesto:
        sTesto ='senza_titolo'

    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Default':
        # se oltre la riga rossa
        lrow -= 2
    if oSheet.getCellByPosition(1, lrow).CellStyle == 'Riga_rossa_Chiudi':
        # se riga rossa
        lrow -= 1

    oSheet.getRows().insertByIndex(lrow, 1)
    oSheet.getCellByPosition(2, lrow).String = sTesto

    # inserisco i valori e le formule
    oSheet.getCellRangeByPosition(0, lrow, 41,lrow).CellStyle = 'livello2 valuta'
    oSheet.getCellRangeByPosition(2, lrow, 17, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(18, lrow, 18, lrow).CellStyle = 'livello2 scritta mini'
    oSheet.getCellRangeByPosition(24, lrow, 24, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(29, lrow, 29, lrow).CellStyle = 'livello2 valuta mini %'
    oSheet.getCellRangeByPosition(30, lrow, 30, lrow).CellStyle = 'livello2 valuta mini'
    oSheet.getCellRangeByPosition(31, lrow, 33, lrow).CellStyle = 'livello2_'
    oSheet.getCellRangeByPosition(2, lrow, 11, lrow).merge(True)

    # oSheet.getCellByPosition(1, lrow).Formula = '=AF' + str(lrow+1) + '''&"."&''' + 'AG' + str(lrow+1)
    # rinumero e ricalcolo
    # ocellBaseA = oSheet.getCellByPosition(1, lrow)
    # ocellBaseR = oSheet.getCellByPosition(31, lrow)

    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(32, lrowProvv).CellStyle != 'livello2 valuta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(
        32, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value + 1
    lrowProvv = lrow - 1
    while oSheet.getCellByPosition(31, lrowProvv).CellStyle != 'Livello-1-scritta':
        if lrowProvv > 4:
            lrowProvv -= 1
        else:
            break
    oSheet.getCellByPosition(31, lrow).Value = oSheet.getCellByPosition(1, lrowProvv).Value
    # SubSum_Cap(lrow)


# ###############################################################


def invertiUnSegno(oSheet, lrow):
    '''
    Inverte il segno delle formule di quantità nel rigo di misurazione lrow.
    lrow    { int }  : riga di riferimento
    usata con XPWE_it
    '''
    if oSheet.Name in ('COMPUTO', 'VARIANTE'):
        if 'comp 1-a' in oSheet.getCellByPosition(2, lrow).CellStyle:
            if 'ROSSO' in oSheet.getCellByPosition(2, lrow).CellStyle:
                oSheet.getCellByPosition(
                    9, lrow
                ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                    lrow + 1) + ')=0;"";PRODUCT(E' + str(
                        lrow + 1) + ':I' + str(lrow + 1) + '))'
                          # se VediVoce
                # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                # ~ oSheet.getCellByPosition(9, lrow).Formula='=IF(PRODUCT(E' +
                # str(lrow+1) + ':I' + str(lrow+1) + ')=0;"";PRODUCT(E' +
                # str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                # ~ else:
                # ~ oSheet.getCellByPosition(9, lrow).Formula=
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) +
                # ')=0;"";PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                for x in range(2, 10):
                    oSheet.getCellByPosition(
                        x, lrow).CellStyle = oSheet.getCellByPosition(
                            x, lrow).CellStyle.split(' ROSSO')[0]
            else:
                oSheet.getCellByPosition(
                    9, lrow
                ).Formula = '=IF(PRODUCT(E' + str(lrow + 1) + ':I' + str(
                    lrow + 1) + ')=0;"";-PRODUCT(E' + str(
                        lrow + 1) + ':I' + str(lrow + 1) + '))'  # se VediVoce
                # ~ if oSheet.getCellByPosition(4, lrow).Type.value != 'EMPTY':
                # ~ oSheet.getCellByPosition(9, lrow).Formula =
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;
                # "";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))' # se VediVoce
                # ~ else:
                # ~ oSheet.getCellByPosition(9, lrow).Formula =
                # '=IF(PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + ')=0;
                # "";-PRODUCT(E' + str(lrow+1) + ':I' + str(lrow+1) + '))'
                for x in range(2, 10):
                    oSheet.getCellByPosition(
                        x, lrow).CellStyle = oSheet.getCellByPosition(
                            x, lrow).CellStyle + ' ROSSO'
    if oSheet.Name in ('CONTABILITA'):
        formula1 = oSheet.getCellByPosition(9, lrow).Formula
        formula2 = oSheet.getCellByPosition(11, lrow).Formula
        oSheet.getCellByPosition(11, lrow).Formula = formula1
        oSheet.getCellByPosition(9, lrow).Formula = formula2
        if oSheet.getCellByPosition(11, lrow).String != '':
            for x in range(2, 12):
                oSheet.getCellByPosition(
                    x, lrow).CellStyle = oSheet.getCellByPosition(
                        x, lrow).CellStyle + ' ROSSO'
        else:
            for x in range(2, 12):
                oSheet.getCellByPosition(
                    x, lrow).CellStyle = oSheet.getCellByPosition(
                        x, lrow).CellStyle.split(' ROSSO')[0]


# ###############################################################

def numeraVoci(oSheet, lrow, all):
    '''
    all { boolean }  : True  rinumera tutto
                       False rinumera dalla voce corrente in giù
    '''
    lastRow = SheetUtils.getUsedArea(oSheet).EndRow + 1
    n = 1

    if not all:
        for x in reversed(range(0, lrow)):
            if(
               oSheet.getCellByPosition(1, x).CellStyle in ('comp Art-EP', 'comp Art-EP_R') and
               oSheet.getCellByPosition(1, x).CellBackColor != 15066597):
                n = oSheet.getCellByPosition(0, x).Value + 1
                break
        for row in range(lrow, lastRow):
            if oSheet.getCellByPosition(1, row).CellBackColor == 15066597:
                oSheet.getCellByPosition(0, row).String = ''
            elif oSheet.getCellByPosition(1,row).CellStyle in ('comp Art-EP', 'comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n += 1
    else:
        for row in range(0, lastRow):
            if oSheet.getCellByPosition(1, row).CellStyle in ('comp Art-EP','comp Art-EP_R'):
                oSheet.getCellByPosition(0, row).Value = n
                n = n + 1
