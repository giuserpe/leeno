'''
Often used utility functions

Copyright 2020 by Massimo Del Fedele
'''
import uno

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

def DisableDocumentRefresh(oDoc):
    '''
    Disabilita il refresh per accelerare le procedure
    '''
    oDoc.lockControllers()
    oDoc.addActionLock()


def EnableDocumentRefresh(oDoc):
    '''
    Riabilita il refresh
    '''
    oDoc.removeActionLock()
    oDoc.unlockControllers()
