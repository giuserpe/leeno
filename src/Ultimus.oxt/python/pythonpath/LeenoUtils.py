'''
    Often used utility functions
'''
import uno


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
