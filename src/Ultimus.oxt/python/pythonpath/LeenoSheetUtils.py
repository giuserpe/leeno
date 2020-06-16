'''
Funzioni di utilità per la manipolazione dei fogli
relativamente alle funzionalità specifiche di LeenO
'''
import uno
import pyleeno as PL
import LeenoUtils
import SheetUtils

# ###############################################################

def setLarghezzaColonne(oSheet):
    '''
    regola la larghezza delle colonne a seconda della sheet
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if oSheet.Name == 'Analisi di Prezzo':
        for col, width in {'A':2100, 'B':12000, 'C':1600, 'D':2000, 'E':3400, 'F':3400,
                           'G':2700, 'H':2700, 'I':2000, 'J':2000, 'K':2000}.items():
            oSheet.Columns[col].Width = width
        oDoc.CurrentController.freezeAtPosition(0, 2)

    elif oSheet.Name == 'CONTABILITA':
        viste_nuove('TTTFFTTTTTFTFTFTFTFTTFTTFTFTTTTFFFFFF')
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
        oDoc.CurrentController.freezeAtPosition(0, 3)

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
        oDoc.CurrentController.freezeAtPosition(0, 3)
        PL.viste_nuove('TTTFFTTTTTFTFFFFFFTFFFFFFFFFFFFFFFFFFFFFFFFFTT')
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
        oDoc.CurrentController.freezeAtPosition(0, 3)
    PL.adatta_altezza_riga(oSheet.Name)

# ###############################################################

def cercaUltimaVoce(oSheet):
    nRow = SheetUtils.getLastUsedRow(oSheet)
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
    elif nome == 'Analisi di Prezzo':
        lrow = cercaUltimaVoce(oSheet) + 2
        oSheet.getCellByPosition(0, lrow).String = 'Fine ANALISI'
        oSheet.getCellRangeByPosition(0, lrow, 10, lrow).CellStyle = 'Riga_rossa_Chiudi'
    elif nome == 'Elenco Prezzi':
        lrow = cercaUltimaVoce(oSheet) + 1
        oSheet.getCellByPosition(0, lrow).String = 'Fine elenco'
        oSheet.getCellRangeByPosition(0, lrow, 7, lrow).CellStyle = 'Riga_rossa_Chiudi'
    oSheet.getCellByPosition(2, lrow
    ).String = 'Questa riga NON deve essere cancellata, MAI!!!(ma può rimanere tranquillamente NASCOSTA!)'

