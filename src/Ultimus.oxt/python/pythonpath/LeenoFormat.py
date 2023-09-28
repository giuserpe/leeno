import uno
import LeenoUtils


def getNumFormat(FormatString):
    '''
    Restituisce il numero identificativo del formato sulla base di una
    stringa di riferimento.
    FormatString { string } : codifica letterale del numero; es.: "#.##0,00"
    '''
    oDoc = LeenoUtils.getDocument()

    LocalSettings = uno.createUnoStruct("com.sun.star.lang.Locale")
    LocalSettings.Language = "it"
    LocalSettings.Country = "IT"
    NumberFormats = oDoc.NumberFormats
    #  FormatString # = "#.##0,00"
    NumberFormatId = NumberFormats.queryKey(FormatString, LocalSettings, True)

    if NumberFormatId == -1:
        NumberFormatId = NumberFormats.addNew(FormatString, LocalSettings)
    return NumberFormatId


def getFormatString(stile_cella):
    '''
    Recupera la stringa di riferimento dal nome dello stile di cella.
    stile_cella { string } : nome dello stile di cella
    '''
    oDoc = LeenoUtils.getDocument()
    num = oDoc.StyleFamilies.getByName("CellStyles").getByName(stile_cella).NumberFormat
    return oDoc.getNumberFormats().getByKey(num).FormatString


def setCellStyleDecimalPlaces(nome_stile, n):
    '''
    Cambia il numero dei decimali dello stile di cella.
    stile_cella { string } : nome stile di cella
    n { int } : nuovo numero decimali
    '''
    oDoc = LeenoUtils.getDocument()
    stringa = getFormatString(nome_stile).split(';')
    new = []
    for el in stringa:
        new.append(el.split(',')[0] + ',' + '0' * n)
    oDoc.StyleFamilies.getByName('CellStyles').getByName(nome_stile).NumberFormat = getNumFormat(';'.join(new))
