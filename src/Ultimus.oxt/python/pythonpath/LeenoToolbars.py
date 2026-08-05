'''
LeenoToolbars
Gestione delle toolbars di LeenO
'''
# pyrefly: ignore [missing-import]
from com.sun.star.awt import Point

import os
import sys
import LeenoUtils
import pyleeno as PL
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


# def Vedi(arg=None):
#     '''
#     accende tutte le toolbars (se non sono richieste quelle contestuali)
#     oppure solo quelle relative alla pagina visualizzata, se richieste le contestuali
#     '''
#     oDoc = LeenoUtils.getDocument()

#     if sys.platform == 'linux' or sys.platform == 'darwin':
#         var = 'HOME'
#     else:
#         var = 'HOMEPATH'
#     try:
#         if 'giuserpe' in os.getlogin():
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 1)
#         else:
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV', 0)
#     except:
#         pass
#     try:
#         oLayout = oDoc.CurrentController.getFrame().LayoutManager

#         if Config().read('Generale', 'toolbar_contestuali') == '0':
#             # toolbar sempre visibili
#             AllOn()
#         else:
#             # toolbar contestualizzate
#             AllOff()
#         Ordina()
#         oLayout.showElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar")
#         nSheet = oDoc.CurrentController.ActiveSheet.Name

#         if nSheet == 'Elenco Prezzi':
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ELENCO', 1)
#         elif nSheet == 'Analisi di Prezzo':
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_ANALISI', 1)
#         elif nSheet in ('COMPUTO', 'VARIANTE'):
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO', 1)
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG', 1)
#         elif nSheet == 'CONTABILITA':
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_COMPUTO', 1)
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA', 1)
#             On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG', 1)

#     except Exception:
#         pass
#     PL.dp()
#     # PL.fissa()

import time
import LeenoDialogs as DLG
@LeenoUtils.release_ram
def Vedi(arg=None):
    '''
    accende tutte le toolbars (se non sono richieste quelle contestuali)
    oppure solo quelle relative alla pagina visualizzata, se richieste le contestuali
    '''
    import time
    t0 = time.time()

    oDoc = LeenoUtils.getDocument()

    try:
        user = os.environ.get('USERNAME') or os.environ.get('USER') or ''
        On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV',
           1 if 'giuserpe' in user else 0)
    except:
        pass
    # DLG.chi(f"checkpoint 1 (user/toolbar DEV): {time.time()-t0:.3f}s")

    try:
        oLayout = oDoc.CurrentController.getFrame().LayoutManager
        # DLG.chi(f"checkpoint 2 (LayoutManager): {time.time()-t0:.3f}s")

        if Config().read('Generale', 'toolbar_contestuali') == '0':
            AllOn()
        else:
            AllOff()
        # DLG.chi(f"checkpoint 3 (dopo AllOn/AllOff): {time.time()-t0:.3f}s")

        Ordina()
        # DLG.chi(f"checkpoint 4 (dopo Ordina): {time.time()-t0:.3f}s")

        oLayout.showElement("private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar")
        # DLG.chi(f"checkpoint 5 (dopo showElement): {time.time()-t0:.3f}s")

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
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CONTABILITA', 1)
            On('private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_CATEG', 1)
        # DLG.chi(f"checkpoint 6 (dopo toolbar sheet): {time.time()-t0:.3f}s")

    except Exception as e:
        # DLG.chi(f"Eccezione nel blocco toolbar: {e} — a {time.time()-t0:.3f}s")
        pass

    # DLG.chi(f"checkpoint 7 (prima di PL.dp): {time.time()-t0:.3f}s")
    PL.dp()
    # DLG.chi(f"checkpoint 8 (fine Vedi): {time.time()-t0:.3f}s")



def _get_layout_manager(oDoc, caller_name):
    '''
    Risolve in modo difensivo oDoc.CurrentController.getFrame().LayoutManager,
    delegando la validazione/fallback del documento a
    LeenoUtils.resolve_document() (unica fonte di verità, usata anche da
    altri moduli come pyleeno.descrizione_in_una_colonna).

    Ritorna (oLayout, oDoc_effettivo) oppure (None, None) se nessun
    tentativo va a buon fine; il chiamante deve loggare e uscire.
    '''
    oDoc = LeenoUtils.resolve_document(oDoc)
    if oDoc is None:
        DLG.chi(f"Toolbars.{caller_name}: nessun documento utilizzabile, richiesta ignorata")
        return None, None
    try:
        return oDoc.CurrentController.getFrame().LayoutManager, oDoc
    except Exception as e:
        DLG.chi(f"Toolbars.{caller_name}: documento risolto ma non utilizzabile ({e})")
        return None, None


def On(toolbarURL, flag, oDoc=None):
    '''
    toolbarURL  { string } : indirizzo toolbar
    flag { integer } : 1 = acceso; 0 = spento
    oDoc { document, opzionale } : documento già risolto dal chiamante.
        Se non fornito, o se non più utilizzabile, viene ri-risolto con
        LeenoUtils.getDocument().
    Visualizza o nascondi una toolbar
    '''
    oLayout, oDoc = _get_layout_manager(oDoc, "On")
    if oLayout is None:
        return
    if flag:
        oLayout.showElement(toolbarURL)
    else:
        oLayout.hideElement(toolbarURL)


def Ordina(oDoc=None):
    '''
    @@ DA DOCUMENTARE
    oDoc { document, opzionale } : documento già risolto dal chiamante.
    '''
    #  https://www.openoffice.org/api/docs/common/ref/com/sun/star/ui/DockingArea.html
    oLayout, oDoc = _get_layout_manager(oDoc, "Ordina")
    if oLayout is None:
        return
    i = 0
    for aBar in _TOOLBAR_NAMES:
        oLayout.dockWindow(aBar, 'DOCKINGAREA_TOP', Point(i, 4))
        i += 1
    oLayout.dockWindow(
        'private:resource/toolbar/addon_ULTIMUS_3.OfficeToolBar_DEV',
        'DOCKINGAREA_RIGHT', Point(0, 0))


def AllOn(flag=True, oDoc=None):
    '''
    Accende o spegne tutte le toolbar di LeenO
    oDoc { document, opzionale } : documento già risolto dal chiamante.
    '''
    for aBar in _TOOLBAR_NAMES:
        On(aBar, flag, oDoc=oDoc)


def AllOff(oDoc=None):
    '''
    Spegne tutte le toolbar di LeenO
    oDoc { document, opzionale } : documento già risolto dal chiamante.
    '''
    AllOn(False, oDoc=oDoc)


def Switch(arg, oDoc=None):
    '''
    Nasconde o mostra le toolbar di Libreoffice.
    oDoc { document, opzionale } : documento già risolto dal chiamante.
        Va passato esplicitamente da chi lo ha già ottenuto prima
        dell'esecuzione di un dialogo modale. Anche così non è garanzia
        assoluta: se tra la risoluzione di oDoc e questa chiamata è girata
        una funzione che tocca la configurazione dell'estensione (es.
        pyleeno.nuove_icone() -> Debug.aggiorna_configurazione_leeno(),
        che può arrivare a invocare desktop.terminate()), il frame/
        controller del documento può risultare smontato: in quel caso
        oDoc.CurrentController solleva UnknownPropertyException invece di
        essere semplicemente None. Per questo la risoluzione del
        LayoutManager passa sempre da _get_layout_manager(), che tenta
        anche una ri-risoluzione fresca del documento prima di rinunciare.
    '''
    oLayout, oDoc = _get_layout_manager(oDoc, "Switch")
    if oLayout is None:
        return
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
