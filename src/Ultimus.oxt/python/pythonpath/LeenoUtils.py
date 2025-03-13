'''
Often used utility functions

Copyright 2020 by Massimo Del Fedele
'''
import sys
import uno
import unohelper

from com.sun.star.beans import PropertyValue
from datetime import date

import calendar

import PyPDF2
'''
ALCUNE COSE UTILI

La finestra che contiene il documento (o componente) corrente:
    desktop.CurrentFrame.ContainerWindow
Non cambia nulla se è aperto un dialogo non modale,
ritorna SEMPRE il frame del documento.

    desktop.ContainerWindow ritorna un None -- non so a che serva

Per ottenere le top windows, c'è il toolkit...
    tk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    tk.getTopWindowCount()      ritorna il numero delle topwindow
    tk.getTopWIndow(i)          ritorna una topwindow dell'elenco
    tk.getActiveTopWindow ()    ritorna la topwindow attiva
La topwindow attiva, per essere attiva deve, appunto, essere attiva, indi avere il focus
Se si fa il debug, ad esempio, è probabile che la finestra attiva sia None

Resta quindi SEMPRE il problema di capire come fare a centrare un dialogo sul componente corrente.
Se non ci sono dialoghi in esecuzione, il dialogo creato prende come parent la ContainerWindow(si suppone...)
e quindi viene posizionato in base a quella
Se c'è un dialogo aperto e nell'event handler se ne apre un altro, l'ultimo prende come parent il precedente,
e viene quindi posizionato in base a quello e non alla schermata principale.
Serve quindi un metodo per trovare le dimensioni DELLA FINESTRA PARENT di un dialogo, per posizionarlo.

L'oggetto UnoControlDialog permette di risalire al XWindowPeer (che non serve ad una cippa), alla XView
(che mi fornisce la dimensione del dialogo ma NON la parent...), al UnoControlDialogModel, che fornisce
la proprietà 'DesktopAsParent' che mi dice SOLO se il dialogo è modale (False) o non modale (True)

L'unica soluzione che mi viene in mente è tentare con tk.ActiveTopWindow e, se None, prendere quella del desktop

'''

def getComponentContext():
    '''
    Get current application's component context
    '''
    try:
        if __global_context__ is not None:
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
    '''
    Get active document
    '''
    desktop = getDesktop()

    # try to activate current frame
    # needed sometimes because UNO doesnt' find the correct window
    # when debugging.
    try:
        desktop.getCurrentFrame().activate()
    except Exception:
        pass

    return desktop.getCurrentComponent()


def getServiceManager():
    '''
    Gets the service manager
    '''
    return getComponentContext().ServiceManager


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
    if not document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        print("Il documento non è un foglio di calcolo.")
        return None

    controller = document.getCurrentController()
    active_sheet = controller.getActiveSheet()
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
    # L'ordine che segue non va cambiato!!!
    if boo:
        oDoc.IsAdjustHeightEnabled = True
        oDoc.enableAutomaticCalculation(True)
        oDoc.removeActionLock()
        oDoc.resetActionLocks()
        oDoc.unlockControllers()
        oDoc.calculateAll()
    else:
        oDoc.IsAdjustHeightEnabled = False
        oDoc.enableAutomaticCalculation(False)
        oDoc.lockControllers()
        oDoc.addActionLock()


def getGlobalVar(name):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    return bDict.get('LEENO_GLOBAL_' + name)


def setGlobalVar(name, value):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    bDict['LEENO_GLOBAL_' + name] = value


def initGlobalVars(dict):
    if type(__builtins__) == type(sys):
        bDict = __builtins__.__dict__
    else:
        bDict = __builtins__
    for key, value in dict.items():
        bDict['LEENO_GLOBAL_' + key] = value


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


def daysInMonth(dat):
    '''
    returns days in month of date dat
    '''
    month = dat.month + 1
    year = dat.year
    if month > 12:
        month = 1
        year += 1
    dat2 = date(year=year, month=month, day=dat.day)
    t = dat2 - dat
    return t.days


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

def elimina_nomi_area_errati():
    '''
    Rimuove i nomi di area con riferimenti non validi (#REF! o #rif!).
    '''
    oDoc = getDocument()

    sheets = oDoc.Sheets
    named_ranges = oDoc.NamedRanges
    nomi_area = named_ranges.ElementNames
    n = len(nomi_area)

    # Crea o ottieni il foglio "duplicati"
    if sheets.hasByName("duplicati"):
        sheet = sheets.getByName("duplicati")
    else:
        sheet = oDoc.createInstance("com.sun.star.sheet.Spreadsheet")
        sheets.insertByName("duplicati", sheet)

    # Analizza i nomi di area
    for i, nome_area in enumerate(nomi_area):
        cella_nome = sheet.getCellByPosition(0, i)  # Prima colonna
        cella_nome.String = nome_area

        cella_contenuto = sheet.getCellByPosition(1, i)  # Seconda colonna
        contenuto = named_ranges.getByName(nome_area).Content
        cella_contenuto.String = contenuto

        # Rimuovi i nomi di area con riferimenti non validi
        if "#REF!" in contenuto or "#rif!" in contenuto:
            named_ranges.removeByName(nome_area)

    # Rimuovi il foglio temporaneo
    sheets.removeByName("duplicati")

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
        oDoc = LeenoUtils.getDocument()
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