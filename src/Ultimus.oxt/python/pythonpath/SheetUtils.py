'''
Utilities to handle worksheets

Copyright 2020 by Massimo Del Fedele
'''
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.util import SortField

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

