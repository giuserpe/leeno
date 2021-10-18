'''
Utilities to handle worksheets

Copyright 2020 by Massimo Del Fedele
'''
import random
import uno
from com.sun.star.xml import AttributeData
from com.sun.star.beans import PropertyValue
from com.sun.star.util import SortField
import LeenoUtils
import LeenoSettings
import DocUtils
import LeenoDialogs as DLG

from datetime import date

'''
User defined attributes handling in worksheets
Not supporting (by now) namespaces, it allows to
insert strings attributes into spreadsheets
'''

def setSheetUserDefinedAttribute(oSheet, name, value):
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
    userAttributes = oSheet.UserDefinedAttributes
    if userAttributes.hasByName(name):
        return userAttributes.getByName(name).Value
    return None

def hasSheetUserDefinedAttribute(oSheet, name):
    userAttributes = oSheet.UserDefinedAttributes
    return userAttributes.hasByName(name)

def removeSheetUserDefinedAttribute(oSheet, name):
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
    if oTxt.String.find('[PAGINA]') >= 0:
        oField = oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
        LeenoUtils.replacePatternWithField(oTxt, '[PAGINA]', oField)

    if oTxt.String.find('[PAGINE]') >= 0:
        oField = oDoc.createInstance("com.sun.star.text.TextField.PageCount")
        LeenoUtils.replacePatternWithField(oTxt, '[PAGINE]', oField)


def pdfExport(oDoc, sheets, destPath, HeaderFooter=None, coverBuilder = None):
    '''
    export a sequence of spreadsheets to a PDF file
    coverBuilder(oDoc, nDoc) takes current document as parameter, the
    print document and adds a cover to the latter at end
    if coverBuilder is None, no cover will be added
    '''
    # create an empty document
    nDoc = DocUtils.createSheetDocument(Hidden=False)

    # if there's a cover, copy it inside the new document
    # and fill it
    if coverBuilder is not None:
        hasCover = coverBuilder(oDoc, nDoc)
    else:
        hasCover = False

    # we need to copy the page styles too... they don't get copied
    # along with sheet
    styleSet = set()
    for sheet in sheets:
        styleSet.add(sheet.PageStyle)

    pageStyles = oDoc.StyleFamilies.getByName('PageStyles')
    for styleName in styleSet:
        style = pageStyles.getByName(styleName)
        copyPageStyle(nDoc, style)

    # copy required sheets on new document it
    # setting also the correct pagestyles...
    # and, if present, copy print area too
    for sheet in sheets:
        pos = nDoc.Sheets.Count
        nDoc.Sheets.importSheet(oDoc, sheet.Name, pos)
        nDoc.Sheets[pos].PageStyle = sheet.PageStyle
        if len(sheet.PrintAreas) > 0:
            nDoc.Sheets[pos].PrintAreas = sheet.PrintAreas
    nDoc.Sheets.removeByName(nDoc.Sheets[0].Name)

    # finally we must apply header/footers to page styles
    # we do it ONLY on page styles which already have
    # header or footer enabled. Other page styles (like cover ones)
    # are left alone
    if HeaderFooter:
        pageStyles = nDoc.StyleFamilies.getByName('PageStyles')
        styleSet = set()
        for sheet in nDoc.Sheets:
            styleSet.add(sheet.PageStyle)
        for styleName in styleSet:
            pageStyle = pageStyles.getByName(styleName)

            # do NOT specity first page number
            # otherwise numbering in header/footer will be wrong
            pageStyle.FirstPageNumber = 0

            if pageStyle.HeaderIsOn:
                print("HEADER")
                left = HeaderFooter.get('intSx', '')
                center = HeaderFooter.get('intCenter', '')
                right = HeaderFooter.get('intDx', '')
                print("  Left  :", left)
                print("  Center:", center)
                print("  Right :", right)
                content = pageStyle.LeftPageHeaderContent
                content.LeftText.String = left
                content.CenterText.String = center
                content.RightText.String = right

                # do PAGINA and PAGINE fields management
                paginationFields(nDoc, content.LeftText)
                paginationFields(nDoc, content.CenterText)
                paginationFields(nDoc, content.RightText)

                pageStyle.RightPageHeaderContent = content

            if pageStyle.FooterIsOn:
                print("FOOTER")
                left = HeaderFooter.get('ppSx', '')
                center = HeaderFooter.get('ppCenter', '')
                right = HeaderFooter.get('ppDx', '')
                print("  Left  :", left)
                print("  Center:", center)
                print("  Right :", right)
                content = pageStyle.LeftPageFooterContent
                content.LeftText.String = left
                content.CenterText.String = center
                content.RightText.String = right

                # do PAGINA and PAGINE fields management
                paginationFields(nDoc, content.LeftText)
                paginationFields(nDoc, content.CenterText)
                paginationFields(nDoc, content.RightText)

                pageStyle.RightPageFooterContent = content

    storeArgs = {
        'FilterName': 'calc_pdf_Export',
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
        Sheet 	numero intero indice dello Sheet contenente l'area
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


def getLastUsedRow(oSheet):
    ''' l'ultima riga usata nello spreadsheet '''
    return getUsedArea(oSheet).EndRow


def getLastUsedColumn(oSheet):
    ''' l'ultima colonna usata nello spreadsheet '''
    return getUsedArea(oSheet).EndColumn

# ###############################################################

def uFindStringCol(sString, nCol, oSheet, start=2, equal=0):
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
    for nRow in range(start, aAddress.EndRow + 1):
        if sString in oSheet.getCellByPosition(nCol, nRow).String:
            return nRow

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
    ricorrenze = list()
    for nRow in range(start, aAddress.EndRow + 1):
        if sString.upper() in oSheet.getCellByPosition(nCol, nRow).String.upper():
            ricorrenze.append(nRow)
    return ricorrenze

def uFindString(sString, oSheet):
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
    for nRow in range(0, aAddress.EndRow + 1):
        for nCol in range(0, aAddress.EndColumn + 1):
            # ritocco di +Daniele Zambelli:
            if sString in oSheet.getCellByPosition(nCol, nRow).String:
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


def visualizza_PageBreak(arg=True):
    '''
    Mostra i salti di pagina dell'area di stampa definita.
    arg       { boolean }
    '''
    # oDoc = LeenoUtils.getDocument()
    #  oSheet = oDoc.getSheets().getByName(oDoc.CurrentController.ActiveSheet.Name)
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
        for el in lista_fogli:
            oSheet = oDoc.getSheets().getByName(el)
            oRangeAddress = oSheet.getCellRangeByPosition(
                0, 0, (getUsedArea(oSheet).EndColumn),
                (getUsedArea(oSheet).EndRow)).getRangeAddress()
            oCellAddress = unione.getCellByPosition(
                0,
                getUsedArea(unione).EndRow + 1).getCellAddress()
            oSheet.copyRange(oCellAddress, oRangeAddress)
        DLG.MsgBox('Unione dei fogli eseguita.', 'Avviso')
    else:
        unione = oDoc.getSheets().getByName('unione_fogli')
        DLG.MsgBox('Il foglio "unione_fogli" è già esistente, quindi non procedo.', 'Avviso!')
    oDoc.CurrentController.setActiveSheet(unione)


# ###############################################################
