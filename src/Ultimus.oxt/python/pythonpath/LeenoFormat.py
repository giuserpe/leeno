import uno
import LeenoUtils


def getNumFormat(FormatString):
    '''
    Restituisce il numero identificativo del formato sulla base di una
    stringa di riferimento.
    FormatString { string } : codifica letterale del numero; es.: "#.##0,00"
    '''
    oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet

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
    # oSheet = oDoc.CurrentController.ActiveSheet
    num = oDoc.StyleFamilies.getByName("CellStyles").getByName(stile_cella).NumberFormat
    return oDoc.getNumberFormats().getByKey(num).FormatString


def setCellStyleDecimalPlaces(nome_stile, n):
    '''
    Cambia il numero dei decimali dello stile di cella.
    stile_cella { string } : nome stile di cella
    n { int } : nuovo numero decimali
    '''
    oDoc = LeenoUtils.getDocument()
    # oSheet = oDoc.CurrentController.ActiveSheet
    stringa = getFormatString(nome_stile).split(';')
    new = list()
    for el in stringa:
        new.append(el.split(',')[0] + ',' + '0' * n)
    #  oDoc.StyleFamilies.getByName('CellStyles').getByName(nome_stile).NumberFormat = getNumFormat(strall("#.##0,", 6+int(PartiUguali), 1))
    oDoc.StyleFamilies.getByName('CellStyles').getByName(nome_stile).NumberFormat = getNumFormat(';'.join(new))
