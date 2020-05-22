'''
    Often used utility functions
'''
import uno


def getComponentContext():
    '''
    Get current application's component context
    '''
    return uno.getComponentContext()


def getDesktop():
    '''
    Get current application's LibreOffice desktop
    '''
    ctx = uno.getComponentContext()
    return ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)


def getDocument():
    '''
    Get active document
    '''
    return getDesktop().getCurrentComponent()
