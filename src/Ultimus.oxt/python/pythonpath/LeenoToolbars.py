'''
LeenoToolbars
Gestione delle toolbars di LeenO
'''
from com.sun.star.awt import Point

import LeenoUtils
from LeenoConfig import Config

# i nome delle toolbars di LeenO
_TOOLBAR_NAMES = (
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG',
    'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA',
)

import pyleeno as PL

def Vedi():
    '''
    accende tutte le toolbars (se non sono richieste quelle contestuali)
    oppure solo quelle relative alla pagina visualizzata, se richieste le contestuali
    '''
    oDoc = LeenoUtils.getDocument()
    try:
        oLayout = oDoc.CurrentController.getFrame().LayoutManager

        if Config().read('Generale', 'toolbar_contestuali') == '0':
            # toolbar sempre visibili
            AllOn()
        else:
            # toolbar contestualizzate
            AllOff()
        #  oLayout.hideElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV")
        Ordina()
        oLayout.showElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar")
        nSheet = oDoc.CurrentController.ActiveSheet.Name

        if nSheet == 'Elenco Prezzi':
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO', 1)
        elif nSheet == 'Analisi di Prezzo':
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI', 1)
        elif nSheet in ('COMPUTO', 'VARIANTE'):
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO', 1)
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG', 1)
        elif nSheet == 'CONTABILITA':
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO', 1)
            # ~ On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG', 1)
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA', 1)
    except Exception:
        pass
    PL.fissa()


def On(toolbarURL, flag):
    '''
    toolbarURL  { string } : indirizzo toolbar
    flag { integer } : 1 = acceso; 0 = spento
    Visualizza o nascondi una toolbar
    '''
    oDoc = LeenoUtils.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    if flag:
        oLayout.showElement(toolbarURL)
    else:
        oLayout.hideElement(toolbarURL)


def Ordina():
    '''
    @@ DA DOCUMENTARE
    '''
    #  https://www.openoffice.org/api/docs/common/ref/com/sun/star/ui/DockingArea.html
    oDoc = LeenoUtils.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    i = 0
    for aBar in _TOOLBAR_NAMES:
        oLayout.dockWindow(aBar, 'DOCKINGAREA_TOP', Point(i, 4))
        i += 1
    oLayout.dockWindow(
        'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV',
        'DOCKINGAREA_RIGHT', Point(0, 0))


def AllOn(flag=True):
    '''
    Accende o spegne tutte le toolbar di LeenO
    '''
    for aBar in _TOOLBAR_NAMES:
        On(aBar, flag)


def AllOff():
    '''
    Spegne tutte le toolbar di LeenO
    '''
    AllOn(False)


def Switch(arg):
    '''
    Nasconde o mostra le toolbar di Libreoffice.
    '''
    oDoc = LeenoUtils.getDocument()
    oLayout = oDoc.CurrentController.getFrame().LayoutManager
    for el in oLayout.Elements:
        if el.ResourceURL not in _TOOLBAR_NAMES + (
                'private:resource/menubar/menubar',
                'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV',
                'private:resource/toolbar/findbar',
                'private:resource/statusbar/statusbar',
        ):
            #  if oLayout.isElementVisible(el.ResourceURL):
            if arg:
                oLayout.showElement(el.ResourceURL)
            else:
                oLayout.hideElement(el.ResourceURL)
