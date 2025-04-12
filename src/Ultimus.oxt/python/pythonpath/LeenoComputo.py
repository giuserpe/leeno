from com.sun.star.table import CellRangeAddress
import SheetUtils
import LeenoUtils
import LeenoSheetUtils
import LeenoConfig
import Dialogs

import LeenoDialogs as DLG

import pyleeno as PL

def datiVoceComputo(oSheet, lrow):
    '''
    Ricava i dati dalla voce di COMPUTO o CONTABILITA.

    Parametri:
    oSheet {Sheet} : Il foglio attivo da cui estrarre i dati.
    lrow {int} : La riga di riferimento da cui iniziare.

    Restituisce:
    - Se CONTABILITA: Una tupla (REG, SAL)
    - Se COMPUTO o VARIANTE: Una tupla (voce)
    '''

    # Circoscrive la voce di computo.
    sStRange = circoscriveVoceComputo(oSheet, lrow)
    if not sStRange:
        return None

    start_row = sStRange.RangeAddress.StartRow
    end_row = sStRange.RangeAddress.EndRow

    # Ricava i dati di base
    num = oSheet.getCellByPosition(0, start_row + 1).String
    art = oSheet.getCellByPosition(1, start_row + 1).String
    desc = oSheet.getCellByPosition(2, start_row + 1).String
    quantP = oSheet.getCellByPosition(9, end_row).Value
    mdo = oSheet.getCellByPosition(30, end_row).Value
    sic = oSheet.getCellByPosition(17, end_row).Value

    voce = []
    REG = []
    SAL = []

    if oSheet.Name == 'CONTABILITA':
        # Gestione quantità negativa.
        quantN = ''
        if quantP < 0:
            quantN = quantP
            quantP = ''

        data = oSheet.getCellByPosition(1, start_row + 2).String
        um = oSheet.getCellByPosition(8, end_row).String.split('[')[-1].split(']')[0]
        Nlib = int(oSheet.getCellByPosition(19, start_row + 1).Value)
        Plib = int(oSheet.getCellByPosition(20, start_row + 1).Value)
        flag = oSheet.getCellByPosition(22, start_row + 1).String
        nSal = int(oSheet.getCellByPosition(23, start_row + 1).Value)
        prezzo = oSheet.getCellByPosition(13, end_row).Value
        importo = oSheet.getCellByPosition(15, end_row).Value

        REG = (
            num + '\n' + art + '\n' + data, desc, Nlib, Plib, um, quantP, quantN,
            prezzo, importo
        )
        
        quant = quantP if quantP != '' else quantN

        SAL = (num, art, desc, um, quant, prezzo, importo, sic, mdo)
        return REG, SAL

    elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
        um = oSheet.getCellByPosition(8, end_row).String.split('[')[-1].split(']')[0]
        prezzo = oSheet.getCellByPosition(11, end_row).Value
        importo = oSheet.getCellByPosition(18, end_row).Value

        voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)
        return voce

    return None



def circoscriveVoceComputo(oSheet, lrow, misure = False):
    '''
    lrow { int } : Riga di riferimento per la selezione dell'intera voce.

    Circoscrive una voce di COMPUTO, VARIANTE o CONTABILITÀ
    partendo dalla posizione corrente del cursore.
    '''
    
    # Inizializza le righe di inizio e fine al valore di partenza.
    start_row = lrow
    end_row = lrow

    # Definisci stili per evitare ripetizioni.
    scritte_stili = {'Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'}
    comp_stili = {
        'comp progress', 'comp 10 s',
        'Comp Start Attributo', 'Comp End Attributo',
        'Comp Start Attributo_R', 'comp 10 s_R',
        'Comp End Attributo_R'
    }

    # Cerca la fine della voce.
    while oSheet.getCellByPosition(0, lrow).CellStyle in scritte_stili:
        lrow += 1
        if lrow >= oSheet.Rows.Count:  # Verifica se lrow supera il numero di righe del foglio.
            return

    if oSheet.getCellByPosition(0, lrow).CellStyle in comp_stili.union(scritte_stili):
        y = lrow
        # Cerca l'inizio e la fine della voce.
        while oSheet.getCellByPosition(0, y).CellStyle not in {'Comp End Attributo', 'Comp End Attributo_R'}:
            y += 1
            if y >= oSheet.Rows.Count:
                return
        end_row = y

        y = lrow
        while y > 0 and oSheet.getCellByPosition(0, y).CellStyle not in {'Comp Start Attributo', 'Comp Start Attributo_R'}:
            y -= 1
        start_row = y

    # Trova il range di firme in CONTABILITA.
    elif oSheet.getCellByPosition(0, lrow).CellStyle == 'Ultimus_centro_bordi_lati':
        for y in reversed(range(0, lrow)):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                start_row = y + 1
                break
        for y in range(lrow, SheetUtils.getLastUsedRow(oSheet)):
            if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                end_row = y - 1
                break

    # Caso specifico per ULTIMUS.
    elif 'ULTIMUS' in oSheet.getCellByPosition(0, lrow).CellStyle:
        start_row = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
        end_row = LeenoSheetUtils.rRow(oSheet) - 1

    else:
        return

    # Restituisce il range di celle individuato.
    if not misure:
        return oSheet.getCellRangeByPosition(0, start_row, 50, end_row)
    else:
        end_offset = -2 if oSheet.Name == 'CONTABILITA' else -1
        return oSheet.getCellRangeByPosition(2, start_row + 2, 8, end_row + end_offset)


    
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
    
    # se le colonne di misura sono nascoste, vengono viasulizzate
    if oSheet.getColumns().getByIndex(5).Columns.IsVisible == False:
        lrow = 4
        n = SheetUtils.getLastUsedRow(oSheet)
        for el in range(4, n):
            if oSheet.getCellByPosition(2, el).CellStyle == "comp sotto centro":
                oSheet.getCellByPosition(2, el).Formula = ''
        for el in range (5, 9):
            oSheet.getColumns().getByIndex(el).Columns.IsVisible = True
    
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

    # ~PL.numera_voci(0)
    LeenoSheetUtils.numeraVoci(oSheet, lrow + 1, False)
    if LeenoConfig.Config().read('Generale', 'pesca_auto') == '1':
        PL.pesca_cod()

def Menu_computoSenzaPrezzi():
    '''
    Duplica il COMPUTO/VARIANTE aggiungendo il suffissio '_copia'
    e cancella i prezzi unitari dal nuovo foglio
    '''
    PL.chiudi_dialoghi()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    nSheet = oSheet.Name
    if nSheet not in ('COMPUTO', 'VARIANTE'):
        return
    tag="_copia"
    idSheet = oSheet.RangeAddress.Sheet + 1
    if oDoc.getSheets().hasByName(nSheet + tag):
        Dialogs.Exclamation(Title = 'ATTENZIONE!',
        Text=f'La tabella di nome {nSheet}{tag} è già presente.')
        return
    else:
        oDoc.Sheets.copyByName(nSheet, nSheet + tag, idSheet)
    nSheet = nSheet + tag
    PL.GotoSheet(nSheet)
    # ~oSheet.protect('')  # 
    # ~Dialogs.Info(Title = 'Ricerca conclusa', Text=f'Il foglio {nSheet} è stato protetto, ma senza password.')

    oSheet = oDoc.getSheets().getByName(nSheet)
    PL.setTabColor(10079487)

    ultima_voce = LeenoSheetUtils.cercaUltimaVoce(oSheet)
    
    ultime_voci = set([circoscriveVoceComputo(oSheet, lrow).RangeAddress.EndRow for lrow in range(6, ultima_voce)])

    for lrow in ultime_voci:
        oSheet.getCellByPosition(11, lrow).String = ''
        oSheet.getCellByPosition(11, lrow).CellStyle = 'comp 1-a'
    LeenoUtils.DocumentRefresh(True)

###############################################################################

class DatiVoce:
    """
    Classe per l'estrazione e la gestione dei dati relativi a una voce di COMPUTO o CONTABILITÀ 
    da un foglio di calcolo LibreOffice Calc.

    Questa classe consente di accedere, in modo strutturato, ai dati associati a una voce, 
    partendo da una riga specificata nel foglio. A seconda del tipo di foglio (COMPUTO, VARIANTE, 
    CONTABILITA), estrae e organizza le informazioni in attributi distinti.

    Attributi:
        oSheet (object): Oggetto Sheet attivo, da cui leggere i dati.
        lrow (int): Riga iniziale di riferimento per individuare la voce.
        range_voce (object): Oggetto CellRange corrispondente alla voce (o None se non trovata).
        SR (int): Riga iniziale della voce.
        ER (int): Riga finale della voce.
        voce (tuple): Dati voce computo/variante.
        REG (tuple): Dati registro della contabilità.
        SAL (tuple): Dati SAL della contabilità.

    Utilizzo:
        dv = DatiVoce(oSheet, lrow)
        print(dv.SR)       # riga di inizio voce
        print(dv.voce)     # dati voce per COMPUTO
        print(dv.REG)      # dati registro per CONTABILITA
        print(dv.SAL)      # dati SAL per CONTABILITA
    """

    def __init__(self, oSheet, lrow):
        self.oSheet = oSheet
        self.lrow = lrow
        self._range = None
        self._SR = None
        self._ER = None
        self._REG = None
        self._SAL = None
        self._voce = None

    @property
    def range(self):
        if self._range is None:
            self._range = self._circoscrive_voce()
        return self._range

    @property
    def SR(self):
        if self._SR is None and self.range:
            self._SR = self.range.RangeAddress.StartRow
        return self._SR

    @property
    def ER(self):
        if self._ER is None and self.range:
            self._ER = self.range.RangeAddress.EndRow
        return self._ER

    @property
    def voce(self):
        if self._voce is None:
            self._estrai_dati()
        return self._voce

    @property
    def num(self):
        if self._voce is None:
            self._estrai_dati()
        return self._voce[0]

    @property
    def art(self):
        if self._voce is None:
            self._estrai_dati()
        return self._voce[1]

    @property
    def desc(self):
        if self._voce is None:
            self._estrai_dati()
        return self._voce[2]

    @property
    def REG(self):
        if self._REG is None:
            self._estrai_dati()
        return self._REG

    @property
    def SAL(self):
        if self._SAL is None:
            self._estrai_dati()
        return self._SAL

    def _circoscrive_voce(self):
        oSheet = self.oSheet
        lrow = self.lrow
        start_row = lrow
        end_row = lrow

        scritte_stili = {'Livello-0-scritta', 'Livello-1-scritta', 'livello2 valuta'}
        comp_stili = {
            'comp progress', 'comp 10 s',
            'Comp Start Attributo', 'Comp End Attributo',
            'Comp Start Attributo_R', 'comp 10 s_R',
            'Comp End Attributo_R'
        }

        while oSheet.getCellByPosition(0, lrow).CellStyle in scritte_stili:
            DLG.chi(oSheet.getCellByPosition(0, lrow).CellStyle)
            lrow += 1
            if lrow >= oSheet.Rows.Count:
                return None

        cs = oSheet.getCellByPosition(0, lrow).CellStyle

        if cs in comp_stili.union(scritte_stili):
            y = lrow
            while oSheet.getCellByPosition(0, y).CellStyle not in {'Comp End Attributo', 'Comp End Attributo_R'}:
                y += 1
                if y >= oSheet.Rows.Count:
                    return None
            end_row = y

            y = lrow
            while y > 0 and oSheet.getCellByPosition(0, y).CellStyle not in {'Comp Start Attributo', 'Comp Start Attributo_R'}:
                y -= 1
            start_row = y

        elif cs == 'Ultimus_centro_bordi_lati':
            for y in reversed(range(0, lrow)):
                if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                    start_row = y + 1
                    break
            for y in range(lrow, SheetUtils.getLastUsedRow(oSheet)):
                if oSheet.getCellByPosition(0, y).CellStyle != 'Ultimus_centro_bordi_lati':
                    end_row = y - 1
                    break

        elif 'ULTIMUS' in cs:
            start_row = LeenoSheetUtils.cercaUltimaVoce(oSheet) + 2
            end_row = LeenoSheetUtils.rRow(oSheet) - 1

        else:
            return cs

        return oSheet.getCellRangeByPosition(0, start_row, 50, end_row)

    def _estrai_dati(self):
        if not self.range:
            return

        sr = self.SR
        er = self.ER
        oSheet = self.oSheet

        num = oSheet.getCellByPosition(0, sr + 1).String
        art = oSheet.getCellByPosition(1, sr + 1).String
        desc = oSheet.getCellByPosition(2, sr + 1).String
        quantP = oSheet.getCellByPosition(9, er).Value
        mdo = oSheet.getCellByPosition(30, er).Value
        sic = oSheet.getCellByPosition(17, er).Value

        if oSheet.Name == 'CONTABILITA':
            quantN = ''
            if quantP < 0:
                quantN = quantP
                quantP = ''

            data = oSheet.getCellByPosition(1, sr + 2).String
            um = oSheet.getCellByPosition(8, er).String.split('[')[-1].split(']')[0]
            Nlib = int(oSheet.getCellByPosition(19, sr + 1).Value)
            Plib = int(oSheet.getCellByPosition(20, sr + 1).Value)
            flag = oSheet.getCellByPosition(22, sr + 1).String
            nSal = int(oSheet.getCellByPosition(23, sr + 1).Value)
            prezzo = oSheet.getCellByPosition(13, er).Value
            importo = oSheet.getCellByPosition(15, er).Value

            self._REG = (
                num + '\n' + art + '\n' + data, desc, Nlib, Plib, um,
                quantP, quantN, prezzo, importo
            )
            quant = quantP if quantP != '' else quantN
            self._SAL = (num, art, desc, um, quant, prezzo, importo, sic, mdo)

        elif oSheet.Name in ('COMPUTO', 'VARIANTE'):
            um = oSheet.getCellByPosition(8, er).String.split('[')[-1].split(']')[0]
            prezzo = oSheet.getCellByPosition(11, er).Value
            importo = oSheet.getCellByPosition(18, er).Value

            self._voce = (num, art, desc, um, quantP, prezzo, importo, sic, mdo)

