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
User defined attributes handlint in worksheets
Not supporting (by now) namespaces, it allows to
insert strings attributes into spreadsheets
'''

def setUserDefinedAttribute(oSheet, name, value):
    userAttributes = oSheet.UserDefinedAttributes
    attr = AttributeData()
    attr.Type = "CDATA"
    attr.Value = value
    if userAttributes.hasByName(name):
        userAttributes.replaceByName(name, attr)
    else:
        userAttributes.insertByName(name, attr)
    oSheet.UserDefinedAttributes = userAttributes

def getUserDefinedAttribute(oSheet, name):
    userAttributes = oSheet.UserDefinedAttributes
    if userAttributes.hasByName(name):
        return userAttributes.getByName(name).Value
    return None

def hasUserDefinedAttribute(oSheet, name):
    userAttributes = oSheet.UserDefinedAttributes
    return userAttributes.hasByName(name)

def removeUserDefinedAttribute(oSheet, name):
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
    setUserDefinedAttribute(oSheet, attrName, "JUSTSOMETHING")

    oEnum = LeenoUtils.getDesktop().Components.createEnumeration()
    while oEnum.hasMoreElements():
        oDoc = oEnum.nextElement()
        try:
            sheets = oDoc.Sheets
            idx = 0
            while idx < sheets.Count:
                sheet = sheets[idx]
                if getUserDefinedAttribute(sheet, attrName) == "JUSTSOMETHING":
                    # found, return it
                    removeUserDefinedAttribute(oSheet, attrName)
                    return oDoc
                idx += 1
        except Exception:
            pass
    # not found
    removeUserDefinedAttribute(oSheet, attrName)
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

