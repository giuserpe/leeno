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


def pdfExport0(oDoc, sheets, destPath):
    '''
    export a sequence of spreadsheets to a PDF file
    '''
    # due giorni per trovare 'sta caxxxxxxata....

    # questo crea un oggetto UNO dello stesso tipo di oDoc.CurrentSelection
    # che FUNZIONA per stampare correttamente sheets multipli
    # ora bisogna solo guardare cosa ci sta dentro...
    rgs = oDoc.createInstance("com.sun.star.sheet.SheetCellRanges")

    #for sheet in sheets:
    #    areas = sheet.PrintAreas
    #    if len(areas) > 0:
    #        rgs.addRangeAddresses(areas, False)
        #print(sheet.getCellRangeByName('Q55'))
        #rgs.addRangeAddress(sheet.getCellRangeByName('A1').RangeAddress, False)
        #rgs.addRangeAddress(sheet.RangeAddress, False)

    #print("rgs-STR:", rgs.RangeAddressesAsString)
    #print("rgs:", rgs)
    #print("curr-STR:", oDoc.CurrentSelection.RangeAddressesAsString)
    #print("curr:", oDoc.CurrentSelection)

    curr = oDoc.CurrentSelection
    #for i in range(0, curr.Count):
    #    rgs.addRangeAddress(curr[i].RangeAddress, False)
    #rgs.addRangeAddresses(curr, False)

    filterProps = {
        'Selection': rgs,
        #'Selection': oDoc.CurrentSelection
    }

    storeArgs = {
        'FilterName': 'calc_pdf_Export',
        'FilterData': LeenoUtils.dictToProperties(filterProps, True),
    }

    destUrl = uno.systemPathToFileUrl(destPath)
    print(destUrl)
    oDoc.storeToURL(destUrl, LeenoUtils.dictToProperties(storeArgs))


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


def pdfExport(oDoc, sheets, destPath):
    '''
    export a sequence of spreadsheets to a PDF file
    '''
    # create an empty document
    desktop = LeenoUtils.getDesktop()
    pth = 'private:factory/scalc'
    nDoc = desktop.loadComponentFromURL(pth, '_default', 0, ())

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

    storeArgs = {
        'FilterName': 'calc_pdf_Export',
    }

    destUrl = uno.systemPathToFileUrl(destPath)
    nDoc.storeToURL(destUrl, LeenoUtils.dictToProperties(storeArgs))
    #nDoc.close(False)
    #del nDoc


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
