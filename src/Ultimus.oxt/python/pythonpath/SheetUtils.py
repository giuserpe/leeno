'''
Utilities to handle worksheets
'''
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.util import SortField

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


