'''
Utilities to handle worksheets

Copyright 2020 by Massimo Del Fedele
'''
import random
# pyrefly: ignore [missing-import]
import uno
# pyrefly: ignore [missing-import]
from com.sun.star.xml import AttributeData
# pyrefly: ignore [missing-import]
from com.sun.star.beans import PropertyValue
# pyrefly: ignore [missing-import]
from com.sun.star.util import SortField
import LeenoUtils
import DocUtils
import LeenoDialogs as DLG

from datetime import date

'''
User defined attributes handling in worksheets
Not supporting (by now) namespaces, it allows to
insert strings attributes into spreadsheets
'''

def setSheetUserDefinedAttribute(oSheet, name, value):
    """Imposta o aggiorna un attributo personalizzato (UserDefinedAttribute) nel foglio specificato."""
    userAttributes = oSheet.UserDefinedAttributes
    attr = AttributeData()
    attr.Type = "CDATA"
    attr.Value = value
    if userAttributes.hasByName(name):
        userAttributes.replaceByName(name, attr)
    else:
        userAttributes.insertByName(name, attr)
    oSheet.UserDefinedAttributes = userAttributes

def getSheetUserDefinedAttribute(oSheet, name):
    """Restituisce il valore di un attributo personalizzato del foglio specificato, oppure None se assente."""
    userAttributes = oSheet.UserDefinedAttributes
    if userAttributes.hasByName(name):
        return userAttributes.getByName(name).Value
    return None

def hasSheetUserDefinedAttribute(oSheet, name):
    """Verifica se il foglio specificato possiede l'attributo personalizzato indicato."""
    userAttributes = oSheet.UserDefinedAttributes
    return userAttributes.hasByName(name)

def removeSheetUserDefinedAttribute(oSheet, name):
    """Rimuove un attributo personalizzato dal foglio specificato se presente."""
    userAttributes = oSheet.UserDefinedAttributes
    if userAttributes.hasByName(name):
        userAttributes.removeByName(name)
        oSheet.UserDefinedAttributes = userAttributes


# ###############################################################


def replaceText(sheet, replaceDict):
    '''
    in un foglio cerca tutte le occorrenze delle chiavi
    contenute in 'replaceDict' e le sostituisce con i rispettivi valori
    '''
    replace = sheet.createReplaceDescriptor()
    for key, val in replaceDict.items():
        replace.SearchString = key
        if type(val) == date:
            # for date use a nice format....
            val = LeenoUtils.date2String(val, 0)
        else:
            # be sure 'val' is a string...
            val = str(val)
        replace.ReplaceString = val
        sheet.replaceAll(replace)


# ###############################################################


def leggi_dato_etichettato(oSheet, label, occorrenza=None):
    '''
    Legge un dato "etichetta e valore" (come nel foglio S2): cerca `label`
    (senza distinguere maiuscole e minuscole, anche come parte del testo) nelle
    colonne B (1) e A (0) e ritorna il valore della colonna C (2) della stessa riga,
    senza spazi iniziali e finali.

    occorrenza None : ritorna il primo valore non vuoto trovato.
    occorrenza n >= 1: ritorna il valore della n-esima riga che contiene l'etichetta,
                       anche se vuoto; serve per le etichette ripetute, come
                       "Importo contrattuale" nei tre blocchi dei dati contratto.
    Se non trova nulla ritorna una stringa vuota.
    '''
    try:
        label_up = label.upper()
        # Ottieni l'area usata per limitare la ricerca
        oCell = oSheet.getCellByPosition(0, 0)
        oCursor = oSheet.createCursorByRange(oCell)
        oCursor.gotoEndOfUsedArea(True)
        max_row = oCursor.RangeAddress.EndRow

        for col in (1, 0):
            trovate = 0
            for r in range(max_row + 1):
                cell_txt = oSheet.getCellByPosition(col, r).String.upper()
                if label_up not in cell_txt:
                    continue
                val = oSheet.getCellByPosition(2, r).String
                if occorrenza is None:
                    if val:
                        return val.strip()
                else:
                    trovate += 1
                    if trovate == occorrenza:
                        return val.strip()
    except Exception as _exc:
        LeenoUtils.log_errore('leggi_dato_etichettato: ' + str(label), _exc)
    return ""


def getDocumentFromSheet(oSheet):
    '''
    given a sheet object returns containing document
    sadly there's no built-in interface for it
    and there's no simple way to do it...
    '''
    #  get all documents from desktop
    attrName = str(random.random())
    setSheetUserDefinedAttribute(oSheet, attrName, "JUSTSOMETHING")

    oEnum = LeenoUtils.getDesktop().Components.createEnumeration()
    while oEnum.hasMoreElements():
        oDoc = oEnum.nextElement()
        try:
            sheets = oDoc.Sheets
            idx = 0
            while idx < sheets.Count:
                sheet = sheets[idx]
                if getSheetUserDefinedAttribute(sheet, attrName) == "JUSTSOMETHING":
                    # found, return it
                    removeSheetUserDefinedAttribute(oSheet, attrName)
                    return oDoc
                idx += 1
        except Exception:
            pass
    # not found
    removeSheetUserDefinedAttribute(oSheet, attrName)
    return None


# ###############################################################

def isCurrentSheet(oSheet):
    '''
    check if sheet is the current one in its document
    '''
    oDoc = getDocumentFromSheet(oSheet)
    contr = oDoc.CurrentController
    return contr.ActiveSheet.Name == oSheet.Name


def setCurrentSheet(oSheet):
    '''
    set the active sheet
    '''
    oDoc = getDocumentFromSheet(oSheet)
    contr = oDoc.CurrentController
    contr.ActiveSheet = oSheet


def getCurrentSheet(oDoc):
    '''
    set the active sheet in given document
    '''
    contr = oDoc.CurrentController
    return contr.ActiveSheet


# ###############################################################


def getSheetNames(filePath):
    '''
    dato il file legge i nomi dei fogli contenuti
    '''
    oDoc = DocUtils.loadDocument(filePath)
    if oDoc is None:
        return tuple()
    sheets = oDoc.Sheets
    res = []
    for sheet in sheets:
        res.append(sheet.Name)
    oDoc.dispose()

    return tuple(res)


# ###############################################################


def tempCopySheet(oDoc, sourceName):
    '''
    crea una copia dello spreadsheet avente nome 'sourceName'
    del documento 'oDoc' e lo mette in coda
    '''
    sheets = oDoc.Sheets
    nSheets = sheets.Count
    if not sheets.hasByName(sourceName):
        return None

    while True:
        newName = sourceName + '_' + str(int(random.random() * 10000))
        if not sheets.hasByName(newName):
            break
    sheets.copyByName(sourceName, newName, nSheets)
    newSheet = sheets.getByName(newName)

    # copiamo anche le aree di stampa e le intestazioni...
    oldSheet = sheets.getByName(sourceName)
    newSheet.TitleRows = oldSheet.TitleRows
    newSheet.PrintTitleRows = oldSheet.PrintTitleRows
    newSheet.PrintAreas = oldSheet.PrintAreas

    return newSheet


# ###############################################################


def copyPageStyle(nDoc, style):
    '''
    copy a page style to a given document
    (or make properties identical if name already present)
    '''
    styleName = style.Name
    nPageStyles = nDoc.StyleFamilies.getByName('PageStyles')

    if nPageStyles.hasByName(styleName):
        # geyt page style
        nPageStyle = nPageStyles.getByName(styleName)
    else:
        # create the style inside new document
        nPageStyle = nDoc.createInstance('com.sun.star.style.PageStyle')

        # append to nDoc page styles
        nPageStyles.insertByName(styleName, nPageStyle)

    # copy all properties
    propSetInfo = style.PropertySetInfo
    props = propSetInfo.Properties
    for prop in props:
        name = prop.Name
        nPageStyle.setPropertyValue(name, style.getPropertyValue(name))

    # page scale is NOT correctly copied above... so do it again
    nPageStyle.setPropertyValue('PageScale', style.getPropertyValue('PageScale'))


def paginationFields(oDoc, oTxt):
    """Sostituisce i segnaposto [PAGINA] e [PAGINE] nell'oggetto di testo con i relativi campi numerici."""
    if oTxt.String.find('[PAGINA]') >= 0:
        oField = oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
        LeenoUtils.replacePatternWithField(oTxt, '[PAGINA]', oField)

    if oTxt.String.find('[PAGINE]') >= 0:
        oField = oDoc.createInstance("com.sun.star.text.TextField.PageCount")
        LeenoUtils.replacePatternWithField(oTxt, '[PAGINE]', oField)


def pdfExport(oDoc, sheets, destPath, preparaCopia=None, coverBuilder = None, printAreas=None):
    '''
    export a sequence of spreadsheets to a PDF file
    preparaCopia(nDoc, nomiStili) prepara il documento di stampa: stili di pagina dei fogli
    esportati (geometria, intestazione e piè di pagina, vedi LeenoPagina) e aspetto delle celle
    coverBuilder(oDoc, nDoc) takes current document as parameter, the
    print document and adds a cover to the latter at end
    if coverBuilder is None, no cover will be added
    printAreas e' un dizionario opzionale {nome foglio: CellRangeAddress}: l'area
    sostituisce quella del foglio sorgente e le sue righe vengono rese visibili
    nel solo documento temporaneo (es. un singolo SAL dentro un foglio condiviso)
    '''
    # create an empty document, senza mostrarlo all'utente
    nDoc = DocUtils.createSheetDocument(Hidden=True)

    # if there's a cover, copy it inside the new document
    # and fill it
    if coverBuilder is not None:
        coverBuilder(oDoc, nDoc)

    # we need to copy the page styles too... they don't get copied
    # along with sheet
    styleSet = set()
    for sheet in sheets:
        styleSet.add(sheet.PageStyle)

    pageStyles = oDoc.StyleFamilies.getByName('PageStyles')
    for styleName in styleSet:
        style = pageStyles.getByName(styleName)
        copyPageStyle(nDoc, style)

    # nDoc è un documento usa-e-getta: disattiviamo la stampa delle note
    # direttamente sui page style copiati, senza toccare il documento
    # originale. Esclude le note dall'export PDF (sia come testo a fine
    # pagina sia come annotazioni PDF cliccabili).
    nPageStyles = nDoc.StyleFamilies.getByName('PageStyles')
    for styleName in styleSet:
        nStyle = nPageStyles.getByName(styleName)
        nStyle.PrintAnnotations = False

    # copy required sheets on new document it
    # setting also the correct pagestyles...
    # and, if present, copy print area too
    for sheet in sheets:
        pos = nDoc.Sheets.Count
        nDoc.Sheets.importSheet(oDoc, sheet.Name, pos)
        nDoc.Sheets[pos].PageStyle = sheet.PageStyle
        if len(sheet.PrintAreas) > 0:
            nDoc.Sheets[pos].PrintAreas = sheet.PrintAreas
        if printAreas and sheet.Name in printAreas:
            area = printAreas[sheet.Name]
            nDoc.Sheets[pos].PrintAreas = (area,)
            nDoc.Sheets[pos].getCellRangeByPosition(
                area.StartColumn, area.StartRow, area.EndColumn, area.EndRow).Rows.IsVisible = True
    nDoc.Sheets.removeByName(nDoc.Sheets[0].Name)

    # PrintAnnotations sopra sopprime solo l'elenco testuale delle note a
    # fine pagina: l'indicatore visivo della nota sulla cella viene comunque
    # incluso da Calc nel rendering di stampa/export. nDoc è un documento
    # usa-e-getta importato da importSheet(), che copia anche le
    # annotazioni: le rimuoviamo esplicitamente per escluderle davvero
    # dal PDF.
    for nSheet in nDoc.Sheets:
        annots = nSheet.Annotations
        for i in range(annots.getCount() - 1, -1, -1):
            annots.removeByIndex(i)

    # do NOT specity first page number
    # otherwise numbering in header/footer will be wrong
    nPageStyles = nDoc.StyleFamilies.getByName('PageStyles')
    for nSheet in nDoc.Sheets:
        nPageStyles.getByName(nSheet.PageStyle).FirstPageNumber = 0

    # geometria, intestazione e piè di pagina: solo sugli stili dei fogli esportati,
    # la copertina ha il suo stile e non si tocca
    if preparaCopia:
        preparaCopia(nDoc, sorted(styleSet))

    # SelectPdfVersion=2 forza l'export in PDF/A-2b: formato idoneo alla
    # conservazione a norma degli atti contabili (registro, computo, ecc.),
    # a differenza del PDF generico prodotto dal filtro senza FilterData.
    # NOTA: FilterData deve essere tipizzato esplicitamente come
    # sequence<PropertyValue> (uno.Any), altrimenti PyUNO lo marshalla come
    # sequence<any> generico e il filtro calc_pdf_Export lo ignora in
    # silenzio, senza sollevare errore.
    storeArgs = {
        'FilterName': 'calc_pdf_Export',
        'FilterData': LeenoUtils.dictToProperties({'SelectPdfVersion': 2}, unoAny=True),
    }

    destUrl = uno.systemPathToFileUrl(destPath)
    nDoc.storeToURL(destUrl, LeenoUtils.dictToProperties(storeArgs))
    nDoc.close(False)
    del nDoc


# ###############################################################

def freezeRowCol(oSheet, row, col):
    '''
    freeze row and column up to row and col
    on sheet oSheet
    Sadly it must use the controller, so it can't be
    done headless... but we try to preserve all we can
    '''
    # get document from sheet
    oDoc = getDocumentFromSheet(oSheet)

    # get controller from document
    controller = oDoc.CurrentController

    # if current sheet is not the one we want to setup, we
    # must change it
    if isCurrentSheet(oSheet):
        curSheet = None
    else:
        curSheet = getCurrentSheet(oDoc)
        setCurrentSheet(oSheet)

    # now change freeze point
    controller.freezeAtPosition(row, col)

    # if current sheet was another one, restore it
    if curSheet is not None:
        setCurrentSheet(curSheet)


# ###############################################################


def setTabColor(oSheet, color):
    '''
    colore   { integer } : id colore
    attribuisce al tab del foglio oSheet un colore a scelta
    '''
    oSheet.TabColor = color

# ###############################################################


def getUsedArea(oSheet):
    '''
    Restituisce l'indirizzo dell' area usata nello spreadsheet
    in forma di oggetto CellRangeAddress
    I membri sono:
        Sheet   numero intero indice dello Sheet contenente l'area
                occhio che è un indice, NON un oggetto spreadsheet
        StartColumn
        StartRow
        EndColumn
        EndRow
    '''
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    return aAddress  # .EndColumn, aAddress.EndRow)


def getRigaSentinella(oSheet, sentinella):
    '''
    oSheet     { object } : foglio in cui cercare
    sentinella { string } : testo della riga sentinella (es. 'Fine elenco')

    Restituisce l'indice di riga (0-based) della prima riga con la
    sentinella in colonna A. Se non esiste, l'ultima riga usata.
    '''
    row = uFindStringCol(sentinella, 0, oSheet)
    if row is None:
        row = getUsedArea(oSheet).EndRow
    return row


def colonnaInLettera(n):
    '''
    n { int } : indice di colonna 0-based

    Restituisce il nome della colonna (0 -> A, 25 -> Z, 26 -> AA).
    '''
    result = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def getLastUsedRow(oSheet):
    ''' l'ultima riga usata nello spreadsheet '''
    return getUsedArea(oSheet).EndRow


def getLastUsedColumn(oSheet):
    ''' l'ultima colonna usata nello spreadsheet '''
    return getUsedArea(oSheet).EndColumn

# ###############################################################

def uFindStringCol(sString, nCol, oSheet, start=2, equal=0, up=False):
    '''
    sString { string }  : stringa da cercare
    nCol    { integer } : indice di colonna
    oSheet  { object }  :
    start   { integer } : riga di partenza
    equal   { integer } : se equal = 1 fa una ricerca per cella intera

    Trova la prima ricorrenza di una stringa(sString) nella
    colonna nCol di un foglio di calcolo(oSheet) e restituisce
    in numero di riga
    '''
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress

    righe = range(start, aAddress.EndRow + 1)
    if up==True:
        righe = reversed (righe)

    # progress = Dialogs.Progress(Title='Preparazione in corso...', Text='Rimane il')
    # progress.setLimits(0, aAddress.EndRow)
    # progress.setValue(0)
    # progress.show()

    for nRow in righe:
        # progress.setValue(nRow)  # Aggiorna progresso
        cell_value = oSheet.getCellByPosition(nCol, nRow).String
        if (equal == 1 and cell_value == sString) or (equal == 0 and sString in cell_value):
            # progress.hide()  # Chiude la barra di progresso
            return nRow
    # progress.hide()  # Chiude la barra di progresso
    return None

        

def sStrColtoList(sString, nCol, oSheet, start=2, equal=0):
    '''
    sString { string }  : stringa da cercare
    nCol    { integer } : indice di colonna
    oSheet  { object }  :
    start   { integer } : riga di partenza

    Trova tutte le ricorrenze di una stringa (sString) nella
    colonna nCol di un foglio di calcolo (oSheet) e restituisce
    la lista delle righe
    '''
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    ricorrenze = []
    for nRow in range(start, aAddress.EndRow + 1):
        if sString.upper() in oSheet.getCellByPosition(nCol, nRow).String.upper():
            ricorrenze.append(nRow)
    return ricorrenze

def uFindString(sString, oSheet, up=False, equal=1, start_col=0, start_row=2):
    '''
    sString { string }  : stringa da cercare
    oSheet  { object }  :

    Trova la prima ricorrenza di una stringa(sString) riga
    per riga in un foglio di calcolo(oSheet) e restituisce
    una tupla(IDcolonna, IDriga)
    '''
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    aAddress = oCursor.RangeAddress
    righe = range (start_row, aAddress.EndRow + 1)
    if up==True:
        righe = reversed (righe)

    for nRow in righe:
        for nCol in range(start_col, aAddress.EndColumn + 1):
            # ritocco di +Daniele Zambelli:
            if (equal == 1 and sString.lower() == oSheet.getCellByPosition(nCol, nRow).String.lower()) or (equal == 0 and sString.lower() in oSheet.getCellByPosition(nCol, nRow).String.lower()):
                return (nCol, nRow)

# ###############################################################

def createSortField(column, sortAscending):
    '''
    create a sort field to be used in sortColumns()
    column is the column to sort for (integer)
    sortAscending is a boolean
    '''
    oSortField = SortField()
    oSortField.Field = column
    oSortField.SortAscending = sortAscending
    return oSortField


def sortColumns(oRange, sortFields):
    '''
    sort a range of cells based on given sortFields
    sortfields are given by a tuple
    so you can order by more criterions
    '''
    oSortDesc = [PropertyValue()]
    oSortDesc[0].Name = "SortFields"
    oSortDesc[0].Value = uno.Any("[]com.sun.star.util.SortField", sortFields)
    oRange.sort(oSortDesc)


def simpleSortColumn(oRange, column, sortAscending):
    '''
    simple sort of a range by a column
    sort direction given by 'sortAscending'
    '''
    sortField = createSortField(column, sortAscending)
    sortColumns(oRange, (sortField,))


# ###############################################################


def NominaArea(oDoc, sSheet, sRange, sName):
    '''
    Definisce o ridefinisce un'area di dati a cui far riferimento
    sSheet = nome del foglio, es.: 'S5'
    sRange = individuazione del range di celle, es.:'$B$89:$L$89'
    sName = nome da attribuire all'area scelta, es.: "manodopera"
    '''
    sPath = "$'" + sSheet + "'." + sRange
    oRanges = oDoc.NamedRanges
    oCellAddress = oDoc.Sheets.getByName(sSheet).getCellRangeByName('A1').getCellAddress()
    if oRanges.hasByName(sName):
        oRanges.removeByName(sName)
    oRanges.addNewByName(sName, sPath, oCellAddress, 0)


# ###############################################################


def FixNamedArea():
    '''
    Corregge i nomi di range dati che contengono '#' sostituendo con '_'
    '''
    oDoc = LeenoUtils.getDocument()
    oRange = oDoc.NamedRanges
    tNamedArea = oRange.ElementNames
    for el in tNamedArea:
        if '#' in el:
            aName = oRange.getByName(el).ReferredCells.AbsoluteName
            sSheet = aName.split('.')[0][1:]
            sRange = aName.split('.')[1]
            sName = el.replace('#', '_')
            NominaArea(oDoc, sSheet, sRange, sName)
            oDoc.NamedRanges.removeByName(el)
    return


# ###############################################################

def remove_bad_ranges():
    '''Recupera tutti gli intervalli nominati definiti nel documento e
    rimuove gli intervalli corrotti o non validi.'''

    oDoc = LeenoUtils.getDocument()
    oNamedRanges = oDoc.NamedRanges
    tNamedArea = oNamedRanges.getElementNames()
    
    # DLG.chi(f"Numero di intervalli: {len(tNamedArea)}")
    
    for i, name in enumerate(tNamedArea, 1):
        oNamedRange = oNamedRanges.getByName(name)
        ref_cells = oNamedRange.ReferredCells
        try:
            ref_cells.RangeAddress
            # addr = ref_cells.RangeAddress
            # DLG.chi(f'''{name}:
            #     - Riga iniziale: {addr.StartRow + 1}
            #     - Riga finale: {addr.EndRow + 1}
            #     - Colonna iniziale: {addr.StartColumn + 1}
            #     - Colonna finale: {addr.EndColumn + 1}
            #     - Foglio: {addr.Sheet}''')
        except Exception:
            oDoc.NamedRanges.removeByName(name)
            # DLG.chi(f"Area {name} rimossa")
    return
    
# ###############################################################


def visualizza_PageBreak(arg=True):
    '''
    Mostra i salti di pagina dell'area di stampa definita.
    arg       { boolean }
    '''
    # oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.CurrentController.ActiveSheet
    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    oProp = PropertyValue()
    oProp.Name = 'PagebreakMode'
    oProp.Value = arg
    properties = (oProp, )

    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
    dispatchHelper.executeDispatch(oFrame, ".uno:PagebreakMode", "", 0, properties)


# ###############################################################

def MENU_unisci_fogli():
    '''
    unisci fogli
    serve per unire tanti fogli in un unico foglio
    '''
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames
    if not oDoc.getSheets().hasByName('unione_fogli'):
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        unione = oDoc.Sheets.insertByName('unione_fogli', sheet)
        unione = oDoc.getSheets().getByName('unione_fogli')
        test = True
        for el in lista_fogli:
            if test == False:
                start = 0
            else:
                start = 1
            oSheet = oDoc.getSheets().getByName(el)
            oRangeAddress = oSheet.getCellRangeByPosition(
                0, start, (getUsedArea(oSheet).EndColumn),
                (getUsedArea(oSheet).EndRow)).getRangeAddress()
            oCellAddress = unione.getCellByPosition(
                0,
                getUsedArea(unione).EndRow + 1).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
            test = True
        DLG.MsgBox('Unione dei fogli eseguita.', 'Avviso')
    else:
        unione = oDoc.getSheets().getByName('unione_fogli')
        DLG.MsgBox('Il foglio "unione_fogli" è già esistente, quindi non procedo.', 'Avviso!')
    oDoc.CurrentController.setActiveSheet(unione)


# ###############################################################

def column_name_of(nColumnNumber):
    """
    Converte un numero di colonna in un nome alfanumerico (es. 0 -> "A", 1 -> "B", 26 -> "AA").
    
    Args:
        nColumnNumber (int): Numero della colonna (0-based).
    
    Returns:
        str: Nome della colonna in formato alfanumerico.
    """
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByIndex(0)  # Deve esserci sempre un foglio con indice 0
    lcolumns = oSheet.Columns
    lcolumn = lcolumns.getByIndex(nColumnNumber)
    cColumnName = lcolumn.Name
    return cColumnName
