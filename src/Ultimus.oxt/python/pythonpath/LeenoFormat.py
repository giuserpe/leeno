# pyrefly: ignore [missing-import]
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

def getPercentFormat():
    '''
    Restituisce il formato per percentuale con 3 decimali,
    rosso quando negativo: es. 26.354% oppure -26.354% in rosso
    '''
    # Sezione 1: positivi/zero (colore default)
    # Sezione 2: negativi (rosso)
    FormatString = '#.##0,000%;[ROSSO]-#.##0,000%'
    return getNumFormat(FormatString)

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
    # ~ import LeenoDialogs as DLG
    try:
        for el in stringa:
            new.append(el.split(',')[0] + ',' + '0' * n)
        oDoc.StyleFamilies.getByName('CellStyles').getByName(nome_stile).NumberFormat = getNumFormat(';'.join(new))
    except Exception as e:
        # ~ DLG.chi(f"Errore durante l'elaborazione: {e}")
        pass


def col_to_index(col):
    '''
    Converte un identificativo di colonna (stringa come "A", "C", "AA" oppure intero)
    in un indice di colonna 0-based.
    '''
    if isinstance(col, int):
        return col
    if isinstance(col, str):
        col_str = col.strip().upper()
        if col_str.isdigit():
            return int(col_str)
        index = 0
        for char in col_str:
            if 'A' <= char <= 'Z':
                index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1 if index > 0 else 0
    return 0


def sostituisci_stile_colonna(nome_foglio, colonna, stile_origine, stile_destinazione, oDoc=None):
    '''
    Nell'ambito di uno sheet - solo per una specifica colonna,
    sostituisce un dato stile di cella con uno diverso.
    Esempio: foglio "CONTABILITA", colonna "C", stile cella "Comp-Bianche sopra_R" > "Comp-Bianche sopraS"

    :param nome_foglio: str, object o None - Nome del foglio (es. "CONTABILITA"), oggetto foglio UNO,
                        oppure None/stringa vuota per usare il foglio attivo.
    :param colonna: str o int - Colonna (es. "C", "AA" oppure indice 0-based/1-based).
    :param stile_origine: str - Nome dello stile di cella da sostituire.
    :param stile_destinazione: str - Nome del nuovo stile di cella da applicare.
    :param oDoc: oggetto documento Calc (opzionale).
    :return: int - Numero di celle/intervalli modificati.
    '''
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()

    if not nome_foglio:
        oSheet = oDoc.CurrentController.ActiveSheet
    elif isinstance(nome_foglio, str):
        oSheet = oDoc.getSheets().getByName(nome_foglio)
    else:
        oSheet = nome_foglio

    col_idx = col_to_index(colonna)

    try:
        import SheetUtils
        max_row = SheetUtils.getLastUsedRow(oSheet)
    except Exception:
        max_row = 1048575

    if max_row < 0:
        max_row = 1048575

    oColRange = oSheet.getCellRangeByPosition(col_idx, 0, col_idx, max_row)

    if hasattr(oColRange, "createReplaceDescriptor"):
        replace = oColRange.createReplaceDescriptor()
        replace.SearchStyles = True
        replace.SearchString = stile_origine
        replace.ReplaceString = stile_destinazione
        return oColRange.replaceAll(replace)

    search = oColRange.createSearchDescriptor()
    search.SearchStyles = True
    search.SearchString = stile_origine

    found = oColRange.findAll(search)
    updated_count = 0

    if found is not None:
        if hasattr(found, "Count"):
            for i in range(found.Count):
                item = found.getByIndex(i)
                item.CellStyle = stile_destinazione
                if hasattr(item, "getRangeAddress"):
                    addr = item.getRangeAddress()
                    updated_count += (addr.EndRow - addr.StartRow + 1) * (addr.EndColumn - addr.StartColumn + 1)
                else:
                    updated_count += 1
        elif hasattr(found, "CellStyle"):
            found.CellStyle = stile_destinazione
            if hasattr(found, "getRangeAddress"):
                addr = found.getRangeAddress()
                updated_count = (addr.EndRow - addr.StartRow + 1) * (addr.EndColumn - addr.StartColumn + 1)
            else:
                updated_count = 1

    return updated_count
