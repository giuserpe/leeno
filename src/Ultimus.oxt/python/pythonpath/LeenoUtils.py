'''
Often used utility functions

Copyright 2020 by Massimo Del Fedele
'''
import sys
import uno
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



def DocumentRefresh(boo):
    oDoc = getDocument()
    if boo == True:
        oDoc.enableAutomaticCalculation(True)
        oDoc.CurrentController.ZoomValue = 100
        oDoc.calculateAll()

        oDoc.removeActionLock()
        oDoc.unlockControllers()


    elif boo == False:
        oDoc.enableAutomaticCalculation(False)
        oDoc.CurrentController.ZoomValue = 400

        oDoc.addActionLock()
        oDoc.lockControllers()


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

