'''
Often used utility functions

Copyright 2020 by Massimo Del Fedele
'''
import sys
import uno
import unohelper

from com.sun.star.beans import PropertyValue
from datetime import date
from contextlib import contextmanager

import calendar
import PyPDF2

# ============================================================================
# CORE INFRASTRUCTURE (Must be at the top to avoid circular import issues)
# ============================================================================

def getComponentContext():
    '''
    Get current application's component context
    '''
    try:
        if "__global_context__" in globals() and __global_context__ is not None:
            return __global_context__
        return uno.getComponentContext()
    except Exception:
        return uno.getComponentContext()


def getDesktop():
    '''
    Get current application's LibreOffice desktop
    '''
    ctx = getComponentContext()
    return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def getDocument():
    try:
        ctx = getComponentContext()
        if ctx is None:
            return None

        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def is_valid_calc(comp):
            if comp is None: return False
            try:
                # Verifica che abbia i fogli e non sia stato eliminato
                return hasattr(comp, "getSheets") and not getattr(comp, "isDisposed", False)
            except:
                return False

        # Tenta il recupero diretto
        oDoc = desktop.getCurrentComponent()
        if is_valid_calc(oDoc):
            return oDoc

        # Fallback: Scansione dei componenti attivi
        components = desktop.getComponents().createEnumeration()
        while components.hasMoreElements():
            comp = components.nextElement()
            if is_valid_calc(comp):
                return comp

        return None

    except Exception:
        return None


def getServiceManager():
    '''
    Gets the service manager
    '''
    return getComponentContext().ServiceManager

# ============================================================================
# PROJECT IMPORTS
# ============================================================================

import LeenoDialogs as DLG
import Dialogs
from LeenoConfig import COLORE_ROSSO_AVVISO
import LeenoGlobals
import pyleeno as PL

def getDispatcher():
    '''
    Restituisce un DispatchHelper per l'invio di comandi .uno:
    '''
    ctx = getComponentContext()
    return ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)


def isPasswordProtected(oDoc=None):
    '''
    Verifica se il documento ha protezioni attive (fogli o struttura).
    Ignora la password di apertura del file perché gestita da LibreOffice.
    '''
    if oDoc is None:
        oDoc = getDocument()

    if oDoc is None:
        return False

    try:
        # 1. Verifica la protezione della struttura del documento
        if hasattr(oDoc, "isProtected") and oDoc.isProtected():
            return True

        # 2. Verifica se almeno un foglio è protetto
        if hasattr(oDoc, "getSheets"):
            sheets = oDoc.getSheets()
            for i in range(sheets.getCount()):
                if sheets.getByIndex(i).isProtected():
                    return True

        # 3. Verifica se il documento è in sola lettura (spesso dovuto a opzioni di condivisione)
        args = oDoc.getArgs()
        for arg in args:
            if arg.Name == "ReadOnly" and arg.Value:
                return True

    except Exception:
        pass

    return False






def createUnoService(serv):
    '''
    create an UNO service
    '''
    return getComponentContext().getServiceManager().createInstance(serv)


def isLeenoDocument():
    '''
    check if current document is a LeenO document
    '''
    try:
        return getDocument().getSheets().hasByName('S2')
    except Exception:
        return False

###############################################################################
def findOpenDocument(filepath):
    '''
    Check if a document is already open and return it.
    '''
    desktop = getDesktop()
    file_url = unohelper.systemPathToFileUrl(filepath)

    for component in desktop.getComponents():
        if hasattr(component, "getURL"):
            if component.getURL() == file_url:
                return component
    return None

def openAndSetActiveDocument(filepath):
    '''
    Open a document if not already open, and set it as active.
    '''
    desktop = getDesktop()
    file_url = unohelper.systemPathToFileUrl(filepath)

    # Controlla se il documento è già aperto
    document = findOpenDocument(filepath)

    if not document:
        properties = (PropertyValue("Hidden", 0, False, 0),)
        document = desktop.loadComponentFromURL(file_url, "_blank", 0, properties)

    if document:
        frame = desktop.getCurrentFrame()
        frame.activate()
        return document

    return None

def getCursorPosition(document):
    '''
    Get the current cursor position in the active sheet of a Calc document.
    Returns (row, column) or None if the document is not a Calc spreadsheet.
    '''
    # if not document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
    #     print("Il documento non è un foglio di calcolo.")
    #     return None

    controller = document.getCurrentController()
    # active_sheet = controller.getActiveSheet()
    selection = controller.getSelection()

    if selection is not None and hasattr(selection, "getRangeAddress"):
        address = selection.getRangeAddress()
        return address.StartRow, address.StartColumn

    return None
###############################################################################


def DocumentRefresh(boo):
    '''
    Abilita / disabilita il refresh per accelerare le procedure
    '''
    oDoc = getDocument()
    if oDoc is None:
        return  # Esci silenziosamente se non c'è un documento attivo
    # L'ordine che segue non va cambiato!!!
    if boo:
        oDoc.IsAdjustHeightEnabled = True
        oDoc.enableAutomaticCalculation(True)
        oDoc.removeActionLock()
        oDoc.resetActionLocks()
        oDoc.unlockControllers()
        oDoc.calculateAll()
    else:
        oDoc.lockControllers()
        oDoc.addActionLock()
        oDoc.IsAdjustHeightEnabled = False
        oDoc.enableAutomaticCalculation(False)


# import Dialogs as DLG
@contextmanager
def DocumentRefreshContext(enable_refresh: bool):
    """
    Context manager per gestire lo stato di refresh del documento,
    evitando incompatibilità con UNO in caso di eccezione.
    """
    original_state = not enable_refresh
    DocumentRefresh(enable_refresh)
    try:
        yield
    except Exception as e:
        pass

        # DLG.errore(str(e))

        # Evita crash UNO: rimuove il traceback
#         IconType = "error"
#         Title = 'ATTENZIONE!'
#         Text='''
# [1] Prima di procedere è meglio dare un nome al file.

# Lavorando su un file senza nome
# potresti avere dei malfunzionamenti.
# '''
#         import Dialogs
        # Dialogs.NotifyDialog(IconType = IconType, Title = Title, Text = Text)
        # raise Exception(str(e)) from None
    finally:
        DocumentRefresh(original_state)

###############################################################################
###############################################################################
"""
Decorator e Context Manager per disabilitare il refresh in LeenO.
"""

import functools
import logging
from contextlib import contextmanager

logger = logging.getLogger(__name__)


# ============================================================================
# CONTEXT MANAGER (per uso con 'with')
# ============================================================================

@contextmanager
def no_refresh_context():
    """
    Context manager per disabilitare temporaneamente il refresh.

    Uso:
        with no_refresh_context():
            # Il refresh è disabilitato qui
            pass
        # Il refresh è riattivato qui
    """
    # Setup: disabilita refresh
    DocumentRefresh(False)

    try:
        # Yield control al blocco with
        yield
    except Exception as e:
        # Se l'eccezione è un oggetto UNO (ha il metodo getTypes), 
        # lo trasformiamo in una Exception Python standard per evitare 
        # che contextlib crashi cercando di impostare __traceback__
        if hasattr(e, "getTypes"):
            raise Exception(f"Errore UNO: {str(e)}") from None
        raise
    finally:
        # Cleanup: riattiva sempre il refresh
        DocumentRefresh(True)


# ============================================================================
# DECORATOR (per decorare funzioni)
# ============================================================================

def no_refresh(func):
    """
    Decorator che disabilita il refresh durante l'esecuzione della funzione.

    Uso:
        @no_refresh
        def mia_funzione():
            # Il refresh è disabilitato qui
            pass
        # Il refresh è riattivato qui
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        # Usa il context manager
        with no_refresh_context():
            return func(*args, **kwargs)

    return wrapper

# ============================================================================
# RIEPILOGO UTILIZZO
# ============================================================================

"""
SCELTA TRA DECORATOR E CONTEXT MANAGER:

1. Usa il DECORATOR quando:
   - Vuoi disabilitare il refresh per tutta la funzione
   - La funzione è ben definita e riutilizzabile
   - Vuoi codice pulito e dichiarativo

   @no_refresh
   def mia_funzione():
       pass

2. Usa il CONTEXT MANAGER quando:
   - Vuoi controllo granulare (solo parte della funzione)
   - Hai logica condizionale
   - Vuoi gestire manualmente gli scope

   with no_refresh_context():
       # solo questa parte
       pass

ENTRAMBI garantiscono che il refresh venga riattivato, anche in caso di errori!
"""
###############################################################################
###############################################################################

# def getGlobalVar(name):
#     if type(__builtins__) == type(sys):
#         bDict = __builtins__.__dict__
#     else:
#         bDict = __builtins__
#     return bDict.get('LEENO_GLOBAL_' + name)


# def setGlobalVar(name, value):
#     if type(__builtins__) == type(sys):
#         bDict = __builtins__.__dict__
#     else:
#         bDict = __builtins__
#     bDict['LEENO_GLOBAL_' + name] = value


# def initGlobalVars(dict):
#     if type(__builtins__) == type(sys):
#         bDict = __builtins__.__dict__
#     else:
#         bDict = __builtins__
#     for key, value in dict.items():
#         bDict['LEENO_GLOBAL_' + key] = value


def dictToProperties(values, unoAny=False):
    '''
    convert a dictionary in a tuple of UNO properties
    if unoAny is True, return the result in an UNO Any variable
    otherwise use a python tuple
    '''
    ps = tuple([PropertyValue(Name=n, Value=v) for n, v in values.items()])
    if unoAny:
        ps = uno.Any('[]com.sun.star.beans.PropertyValue', ps)
    return ps

import calendar
def daysInMonth(dat):
    '''
    Restituisce il numero di giorni nel mese della data passata.
    Gestisce correttamente anni bisestili e mesi di diversa durata.
    '''
    # calendar.monthrange restituisce una tupla (giorno_settimana_inizio, numero_giorni)
    # Prendiamo l'indice [1] per avere il numero totale di giorni.
    return calendar.monthrange(dat.year, dat.month)[1]

def firstWeekDay(dat):
    '''
    returns first week day in month from dat
    monday is 0
    '''
    return calendar.weekday(dat.year, dat.month, 1)


DAYNAMES = ['Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab', 'Dom']
MONTHNAMES = [
    'Gennaio', 'Febbraio', 'Marzo', 'Aprile',
    'Maggio', 'Giugno', 'Luglio', 'Agosto',
    'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
]

def date2String(dat, fmt = 0):
    '''
    conversione data in stringa
    fmt = 0     25 Febbraio 2020
    fmt = 1     25/2/2020
    fmt = 2     25-02-2020
    fmt = 3     25.02.2020
    '''
    d = dat.day
    m = dat.month
    if m < 10:
        ms = '0' + str(m)
    else:
        ms = str(m)
    y = dat.year
    if fmt == 1:
        return str(d) + '/' + ms + '/' + str(y)
    elif fmt == 2:
        return str(d) + '-' + ms + '-' + str(y)
    elif fmt == 3:
        return str(d) + '.' + ms + '.' + str(y)
    else:
        return str(d) + ' ' + MONTHNAMES[m - 1] + ' ' + str(y)

def string2Date(s):
    if '.' in s:
        sp = s.split('.')
    elif '/' in s:
        sp = s.split('/')
    elif '-' in s:
        sp = s.split('-')
    else:
        return date.today()
    if len(sp) != 3:
        raise Exception
    day = int(sp[0])
    month = int(sp[1])
    year = int(sp[2])
    return date(day=day, month=month, year=year)

def countPdfPages(path):
    '''
    Returns the number of pages in a PDF document
    using external PyPDF2 module
    '''
    with open(path, 'rb') as f:
        pdf = PyPDF2.PdfFileReader(f)
        return pdf.getNumPages()


def replacePatternWithField(oTxt, pattern, oField):
    '''
    Replaces a string pattern in a Text object
    (for example '[PATTERN]') with the given field
    '''
    # pattern may be there many times...
    repl = False
    pos = oTxt.String.find(pattern)
    while pos >= 0:
        #create a cursor
        cursor = oTxt.createTextCursor()

        # use it to select the pattern
        cursor.collapseToStart()
        cursor.goRight(pos, False)
        cursor.goRight(len(pattern), True)

        # remove the pattern from text
        cursor.String = ''

        # insert the field at cursor's position
        cursor.collapseToStart()
        oTxt.insertTextContent(cursor, oField, False)

        # next occurrence of pattern
        pos = oTxt.String.find(pattern)

        repl = True
    return repl

########################################################################

def indirizzo_in_forma_leggibile():
    """
    Restituisce l'indirizzo leggibile della cella attualmente selezionata.

    - Usa `CellAddressConversion` per convertire l'indirizzo della cella selezionata.
    - Restituisce la rappresentazione leggibile per l'interfaccia utente.

    Returns:
        str: L'indirizzo della cella in formato leggibile.
    """
    oDoc = getDocument()

    # Controlla che il documento sia un foglio di calcolo
    if not hasattr(oDoc, "Sheets"):
        print("Il documento corrente non è un foglio di calcolo.")
        return None

    # Ottieni la cella attiva
    active_cell = oDoc.CurrentSelection
    cell_address = active_cell.CellAddress

    # Converte l'indirizzo in una rappresentazione leggibile
    converter = oDoc.createInstance("com.sun.star.table.CellAddressConversion")
    converter.Address = cell_address

    user_representation = converter.UserInterfaceRepresentation
    persistent_representation = converter.PersistentRepresentation

    # Stampa le rappresentazioni (opzionale)

    return user_representation

########################################################################

def reset_properties(o_range, cell_formatting=False, character_formatting=False,
                     paragraph_formatting=False, border_and_table_formatting=False,
                     number_formatting=False, alignment_and_justification=False,
                     validation=False, shadow_and_effects=False):
    """
    Ripristina le proprietà di una selezione di celle (o range) ai valori predefiniti.

    Args:
        o_range (object): L'oggetto range che rappresenta una selezione di celle.
        cell_formatting (bool): Se True, ripristina le proprietà relative alla formattazione delle celle.
        character_formatting (bool): Se True, ripristina le proprietà relative alla formattazione del carattere.
        paragraph_formatting (bool): Se True, ripristina le proprietà relative alla formattazione del paragrafo.
        border_and_table_formatting (bool): Se True, ripristina le proprietà relative ai bordi e alla formattazione della tabella.
        number_formatting (bool): Se True, ripristina le proprietà relative alla formattazione numerica.
        alignment_and_justification (bool): Se True, ripristina le proprietà relative all'allineamento e giustificazione.
        validation (bool): Se True, ripristina le proprietà relative alla validazione.
        shadow_and_effects (bool): Se True, ripristina le proprietà relative agli effetti di ombreggiatura.

    ### ESEMPIO D'USO:
        oDoc = getDocument()
        o_range = oDoc.CurrentSelection
        reset_properties(o_range, cell_formatting=True, character_formatting=True)
    """
    # Proprietà ordinate per categoria
    cell_formatting_properties = [
        'CellBackColor', 'CellBackgroundComplexColor', 'CellProtection', 'CellStyle', 'IsCellBackgroundTransparent'
    ]

    character_formatting_properties = [
        'CharColor', 'CharComplexColor', 'CharContoured', 'CharCrossedOut', 'CharEmphasis', 'CharFont',
        'CharFontCharSet', 'CharFontCharSetAsian', 'CharFontCharSetComplex', 'CharFontFamily', 'CharFontFamilyAsian',
        'CharFontFamilyComplex', 'CharFontName', 'CharFontNameAsian', 'CharFontNameComplex', 'CharFontPitch',
        'CharFontPitchAsian', 'CharFontPitchComplex', 'CharFontStyleName', 'CharFontStyleNameAsian',
        'CharFontStyleNameComplex', 'CharHeight', 'CharHeightAsian', 'CharHeightComplex', 'CharLocale',
        'CharLocaleAsian', 'CharLocaleComplex', 'CharOverline', 'CharOverlineColor', 'CharOverlineHasColor',
        'CharPosture', 'CharPostureAsian', 'CharPostureComplex', 'CharRelief', 'CharShadowed', 'CharStrikeout',
        'CharUnderline', 'CharUnderlineColor', 'CharUnderlineHasColor', 'CharWeight', 'CharWeightAsian',
        'CharWeightComplex', 'CharWordMode'
    ]

    paragraph_formatting_properties = [
        'ParaAdjust', 'ParaBottomMargin', 'ParaIndent', 'ParaIsCharacterDistance', 'ParaIsForbiddenRules',
        'ParaIsHangingPunctuation', 'ParaIsHyphenation', 'ParaLastLineAdjust', 'ParaLeftMargin', 'ParaRightMargin',
        'ParaTopMargin'
    ]

    border_and_table_formatting_properties = [
        'BottomBorder', 'BottomBorder2', 'BottomBorderComplexColor', 'LeftBorder', 'LeftBorder2',
        'LeftBorderComplexColor', 'RightBorder', 'RightBorder2', 'RightBorderComplexColor', 'TopBorder',
        'TopBorder2', 'TopBorderComplexColor', 'TableBorder', 'TableBorder2'
    ]

    number_formatting_properties = [
        'NumberFormat', 'NumberingRules'
    ]

    alignment_and_justification_properties = [
        'HoriJustify', 'HoriJustifyMethod', 'VertJustify', 'VertJustifyMethod', 'IsTextWrapped'
    ]

    validation_properties = [
        'Validation', 'ValidationLocal', 'ValidationXML'
    ]

    shadow_and_effects_properties = [
        'ShadowFormat', 'ShrinkToFit', 'CharPosture', 'CharStrikeout', 'CharShadowed'
    ]

    # Creiamo una lista di tutte le proprietà da ripristinare
    properties_to_reset = []

    if cell_formatting:
        properties_to_reset.extend(cell_formatting_properties)
    if character_formatting:
        properties_to_reset.extend(character_formatting_properties)
    if paragraph_formatting:
        properties_to_reset.extend(paragraph_formatting_properties)
    if border_and_table_formatting:
        properties_to_reset.extend(border_and_table_formatting_properties)
    if number_formatting:
        properties_to_reset.extend(number_formatting_properties)
    if alignment_and_justification:
        properties_to_reset.extend(alignment_and_justification_properties)
    if validation:
        properties_to_reset.extend(validation_properties)
    if shadow_and_effects:
        properties_to_reset.extend(shadow_and_effects_properties)

    # Ripristina tutte le proprietà ai valori predefiniti

    # for prop in properties_to_reset:
    #     try:
    #         o_range.setPropertyToDefault(prop)
    #     except Exception as e:
    #         pass


########################################################################

def imposta_schermo_intero(stato):
    """
    Attiva o disattiva la modalità schermo intero.

    :param stato: Booleano, True per abilitare, False per disabilitare lo schermo intero.
    """
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    document = desktop.getCurrentComponent()

    if document is not None:
        controller = document.getCurrentController()
        frame = controller.getFrame()
        dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)

        # Configura i parametri per il comando FullScreen
        args = [uno.createUnoStruct("com.sun.star.beans.PropertyValue")]
        args[0].Name = "FullScreen"
        args[0].Value = stato

        # Esegue il comando FullScreen
        dispatcher.executeDispatch(frame, ".uno:FullScreen", "", 0, tuple(args))

def massimizza():
    """Abilita la modalità schermo intero."""
    imposta_schermo_intero(True)

def torna_a_schermo_normale():
    """Disabilita la modalità schermo intero."""
    imposta_schermo_intero(False)

########################################################################
########################################################################
########################################################################
def int_to_italian(n: int) -> str:
    if n == 0:
        return "zero"
    units = ["","uno","due","tre","quattro","cinque","sei","sette","otto","nove",
             "dieci","undici","dodici","tredici","quattordici","quindici",
             "sedici","diciassette","diciotto","diciannove"]
    tens_names = {20:"venti",30:"trenta",40:"quaranta",50:"cinquanta",
                  60:"sessanta",70:"settanta",80:"ottanta",90:"novanta"}

    def under_thousand(num: int) -> str:
        res = ""
        if num >= 100:
            hundreds = num // 100
            rest = num % 100
            if rest >= 80 and rest < 90:
                if hundreds == 1:
                    res += "cent"
                else:
                    res += units[hundreds] + "cent"
            else:
                if hundreds == 1:
                    res += "cento"
                else:
                    res += units[hundreds] + "cento"
            num = rest
        if num >= 20:
            tens = (num // 10) * 10
            unit = num % 10
            tens_word = tens_names[tens]
            if unit == 1 or unit == 8:
                tens_word = tens_word[:-1]
            res += tens_word
            if unit:
                res += units[unit]
        elif num > 0:
            res += units[num]
        return res

    parts = []
    billions = n // 1_000_000_000
    if billions:
        if billions == 1:
            parts.append("unmiliardo")
        else:
            parts.append(int_to_italian(billions) + "miliardi")
        n %= 1_000_000_000
    millions = n // 1_000_000
    if millions:
        if millions == 1:
            parts.append("unmilione")
        else:
            parts.append(int_to_italian(millions) + "milioni")
        n %= 1_000_000
    thousands = n // 1000
    if thousands:
        if thousands == 1:
            parts.append("mille")
        else:
            parts.append(int_to_italian(thousands) + "mila")
        n %= 1000
    if n:
        parts.append(under_thousand(n))
    return "".join(parts)

def convert_number_string(s: str) -> str:
    s = s.strip()
    if not s:
        raise ValueError("Stringa vuota.")
    negative = s.startswith("-")
    if negative:
        s = s[1:].strip()
    if ',' in s:
        integer_part, frac_part = s.split(',', 1)
    elif '.' in s:
        integer_part, frac_part = s.split('.', 1)
    else:
        integer_part, frac_part = s, ""
    if integer_part == "":
        integer_value = 0
    else:
        if not integer_part.isdigit():
            raise ValueError("Parte intera non valida: deve contenere solo cifre.")
        integer_value = int(integer_part)
    words = int_to_italian(integer_value)
    if negative:
        words = "meno" + words
    if frac_part != "":
        frac_clean = "".join(ch for ch in frac_part if ch.isdigit())
        return f"{words}/{frac_clean}"
    else:
        return words

import textwrap

def wrap_text(text: str, width=72) -> str:
    # return "\n".join(textwrap.wrap(text, width=50))
    lines = text.splitlines()  # mantiene il testo così com'è diviso
    wrapped_lines = [ "\n".join(textwrap.wrap(line, width)) if line else "" for line in lines ]
    return "\n".join(wrapped_lines)

def wrap_path(path, max_len=60):
    """Versione ultra-compatta per wrapping percorsi"""
    parts = path.split('\\')
    result = parts[0]
    current = parts[0]

    for p in parts[1:]:
        if len(current + '\\' + p) > max_len:
            result += '\\\n' + p
            current = '\\' + p
        else:
            result += '\\' + p
            current += '\\' + p

    return result

##########################################################################
import functools



def memorizza_posizione(step=0):
    """Memorizza la posizione corrente del cursore, con incremento opzionale della riga"""
    ctx = getComponentContext()
    doc = getDocument()
    controller = doc.getCurrentController()

    # Ottieni la selezione corrente
    selection = controller.getSelection()

    # Gestione per diversi tipi di selezione
    if selection.supportsService("com.sun.star.sheet.SheetCell"):
        # Singola cella
        cell_addr = selection.getCellAddress()
        pos_data = {
            'type': 'cell',
            'sheet': cell_addr.Sheet,
            'col': cell_addr.Column,
            'row': cell_addr.Row + step  # incremento opzionale
        }
    elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):
        # Range di celle
        range_addr = selection.getRangeAddress()
        pos_data = {
            'type': 'range',
            'sheet': range_addr.Sheet,
            'col': range_addr.StartColumn,
            'row': range_addr.StartRow + step,      # incremento opzionale
            'end_col': range_addr.EndColumn,
            'end_row': range_addr.EndRow + step     # incremento opzionale
        }
    else:
        # DLG.chi("Tipo di selezione non supportato")
        return

    # Memorizza i dati
    LeenoGlobals.setGlobalVar('ultima_posizione', pos_data)

    # DLG.chi(f"Posizione salvata: Foglio {pos_data['sheet']}, Riga {pos_data['row']}, Col {pos_data['col']}")

def ripristina_posizione():
    """Ripristina la posizione memorizzata"""
    pos_data = LeenoGlobals.getGlobalVar('ultima_posizione')
    if not pos_data:
        DLG.chi("Nessuna posizione memorizzata trovata")
        return

    doc = getDocument()
    controller = doc.getCurrentController()
    sheets = doc.getSheets()

    try:
        sheet = sheets.getByIndex(pos_data['sheet'])

        if pos_data['type'] == 'cell':
            # Ripristina singola cella
            cell = sheet.getCellByPosition(pos_data['col'], pos_data['row'])
            controller.select(cell)
        else:
            # Ripristina range di celle
            cell_range = sheet.getCellRangeByPosition(
                pos_data['col'], pos_data['row'],
                pos_data['end_col'], pos_data['end_row']
            )
            controller.select(cell_range)

    except Exception as e:
        DLG.chi(f"Errore nel ripristino: {str(e)}")
    doc.CurrentController.select(doc.createInstance("com.sun.star.sheet.SheetCellRanges"))

#############################################################################

def preserva_posizione(step=0):
    """
    Decorator che memorizza la posizione del cursore prima della funzione
    e la ripristina alla fine.
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Memorizza
            memorizza_posizione(step=0) # Salviamo la posizione iniziale "vera"
            try:
                result = func(*args, **kwargs)
                return result
            finally:
                # Ripristina (anche se la funzione va in errore)
                # Se passi un valore a 'step', lo usiamo per spostarci DOPO l'operazione
                if step != 0:
                    pos = LeenoGlobals.getGlobalVar('ultima_posizione')
                    if pos:
                        pos['row'] += step
                        if 'end_row' in pos: pos['end_row'] += step
                        LeenoGlobals.setGlobalVar('ultima_posizione', pos)
                ripristina_posizione()
        return wrapper
    return decorator

from contextlib import contextmanager

@contextmanager
def CursorContext(step=0):
    memorizza_posizione(step)
    try:
        yield
    finally:
        ripristina_posizione()

#############################################################################

@contextmanager
def ProtezioneFoglioContext(sheet_or_name, password="", oDoc=None):
    """
    Context manager che sblocca un foglio e lo riprotegge alla fine.

    :param sheet_or_name: Oggetto foglio o stringa col nome del foglio
    :param password: Password del foglio (default vuota)
    :param oDoc: Documento (opzionale, se sheet_or_name è una stringa)
    """
    # Se passiamo un nome, recuperiamo l'oggetto foglio
    if isinstance(sheet_or_name, str):
        if oDoc is None:
            oDoc = getDocument()
        oSheet = oDoc.getSheets().getByName(sheet_or_name)
    else:
        oSheet = sheet_or_name

    # Verifichiamo se il foglio è effettivamente protetto
    was_protected = oSheet.isProtected()

    if was_protected:
        oSheet.unprotect(password)

    try:
        # Qui il codice all'interno del blocco 'with' viene eseguito
        yield oSheet
    finally:
        # Questo viene eseguito SEMPRE, anche in caso di errore
        if was_protected:
            oSheet.protect(password)

########################################################################


def _fingerprint_voce(oSheet, SR, ER):
    '''
    Calcola una chiave univoca (hashable) per una voce di computo/contabilità.
    La chiave è basata su: codice articolo, righe di misura (colonne C-I), totale quantità.

    Parametri:
    oSheet  {Sheet} : foglio di lavoro
    SR      {int}   : StartRow della voce (Comp Start Attributo)
    ER      {int}   : EndRow della voce (Comp End Attributo)

    Ritorna una tupla hashable o None se la voce non è valida.
    '''
    try:
        art = oSheet.getCellByPosition(1, SR + 1).String.strip()
        if not art:
            return None
        # righe di misura: da SR+2 a ER-1 (escluso), colonne C-I (indici 2-8)
        misure = []
        for r in range(SR + 2, ER):
            riga = tuple(
                oSheet.getCellByPosition(c, r).getString().strip()
                for c in range(2, 9)
            )
            misure.append(riga)
        quant = round(oSheet.getCellByPosition(9, ER).Value, 6)
        return (art, tuple(misure), quant)
    except Exception:
        return None


########################################################################


@no_refresh
def MENU_trova_duplicati():
    '''
    Scansiona il foglio attivo (COMPUTO, VARIANTE o CONTABILITA) e individua
    le voci duplicate: stesso codice articolo, stesse righe di misura, stesso totale.

    - Le voci NON duplicate vengono raggruppate e chiuse (outline).
    - Le voci duplicate vengono evidenziate con COLORE_ROSSO_AVVISO sulla cella A del SR.
    - Viene mostrato un dialogo riepilogativo.
    '''
    PL.struttura_off()
    oDoc = getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    if oSheet.Name not in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
        Dialogs.Exclamation(
            Title='ATTENZIONE!',
            Text='Funzione disponibile solo nei fogli COMPUTO, VARIANTE e CONTABILITA.')
        return

    stili_computo = set(LeenoGlobals.getGlobalVar('stili_computo'))
    stili_contab  = set(LeenoGlobals.getGlobalVar('stili_contab'))
    stili_validi  = stili_computo | stili_contab
    stili_cat     = set(LeenoGlobals.getGlobalVar('stili_cat'))
    stili_skip    = stili_cat | {'uuuuu', 'Ultimus_centro_bordi_lati',
                                  'comp Int_colonna', 'ULTIMUS',
                                  'ULTIMUS_1', 'ULTIMUS_2', 'ULTIMUS_3', ''}

    last_row = SheetUtils.getLastUsedRow(oSheet)
    iSheet   = oSheet.RangeAddress.Sheet

    # --- Passata 1: raccolta fingerprint ---
    # fingerprint -> lista di (SR, ER)
    fp_map = {}
    lrow = 3  # le prime righe sono intestazione
    while lrow <= last_row:
        stile = oSheet.getCellByPosition(0, lrow).CellStyle
        if stile in stili_skip:
            lrow += 1
            continue
        if stile not in stili_validi:
            lrow += 1
            continue

        vrange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        if not vrange:
            lrow += 1
            continue

        SR = vrange.RangeAddress.StartRow
        ER = vrange.RangeAddress.EndRow

        fp = _fingerprint_voce(oSheet, SR, ER)
        if fp is not None:
            fp_map.setdefault(fp, []).append((SR, ER))

        lrow = ER + 1  # salta alla voce successiva

    # --- Separazione duplicati / non duplicati ---
    duplicati   = {fp: lst for fp, lst in fp_map.items() if len(lst) > 1}
    singole     = {fp: lst for fp, lst in fp_map.items() if len(lst) == 1}

    sr_duplicati = set()
    for lst in duplicati.values():
        for SR, ER in lst:
            sr_duplicati.add(SR)

    # --- Passata 2: evidenziazione e raggruppamento ---
    def _make_range_addr(start, end):
        addr = uno.createUnoStruct('com.sun.star.table.CellRangeAddress')
        addr.Sheet       = iSheet
        addr.StartColumn = 0
        addr.EndColumn   = 0
        addr.StartRow    = start
        addr.EndRow      = end
        return addr

    for lst in duplicati.values():
        for SR, ER in lst:
            oSheet.getCellByPosition(0, SR).CellBackColor = COLORE_ROSSO_AVVISO

    for lst in singole.values():
        SR, ER = lst[0]
        if SR <= ER and SR >= 3:
            try:
                oSheet.group(_make_range_addr(SR, ER), 1)
                oSheet.getCellRangeByPosition(0, SR, 0, ER).Rows.IsVisible = False
            except Exception:
                pass

    # Chiude tutti i gruppi (livello 1) tramite dispatch
    if singole:
        try:
            dispatcher = getDispatcher()
            frame = oDoc.getCurrentController().Frame
            dispatcher.executeDispatch(frame, '.uno:HideDetail', '', 0, ())
        except Exception:
            pass

    # --- Dialogo riepilogativo ---
    if duplicati:
        righe = []
        for (art, misure, quant), lst in sorted(duplicati.items(), key=lambda x: x[0][0]):
            righe.append(f'  {art}  ({len(lst)} volte, q.tà = {quant})')
        testo = 'Voci duplicate trovate:\n\n' + '\n'.join(righe)
    else:
        struttura_off()
        testo = 'Nessuna voce duplicata trovata.'

    Dialogs.notizia(Title='Duplicati', Text=testo)
