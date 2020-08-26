"""
    LeenO - modulo di gestione eventi del documento e dei fogli
"""
import uno
from com.sun.star.beans import PropertyValue
import LeenoUtils

def macro_SHEET(nSheet, nEvento, miamacro):
    '''
    Attribuisce specifica macro ad evento di un foglio
    '''
 # ~nEvento:
 # ~"OnFocus"       entrando nel foglio
 # ~"OnUnfocus"     uscendo dal foglio
 # ~"OnSelect"      selezionando
 # ~"OnDoubleClick" doppio click
 # ~"OnRightClick"  click destro
 # ~"OnChange"      modificando il contenuto
 # ~"OnCalculate"   mboh...
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'EventType'
    oProp0.Value = 'Script'
    oProp1 = PropertyValue()
    oProp1.Name = 'Script'
    oProp1.Value = miamacro
    
    oProp.append(oProp0)
    oProp.append(oProp1)

    properties = tuple(oProp)
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName(nSheet)
    uno.invoke(
        oSheet.Events, 'replaceByName',
        (nEvento, uno.Any('[]com.sun.star.beans.PropertyValue', properties))
    )
    return

########################################################################

def macro_DOC(nEvento, miamacro):
    '''
    Attribuisce specifica macro ad evento del documento
    '''
# ~http://bit.ly/1EgROQt
# ~esempio: macro_DOC("OnUnfocus", "vnd.sun.star.script:UltimusFree2.Menu.SeeComponentsElements?language=Basic&location=application")

    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'EventType'
    oProp0.Value = 'Script'
    oProp1 = PropertyValue()
    oProp1.Name = 'Script'
    oProp1.Value = miamacro # persorso macro da assegnare
    
    oProp.append(oProp0)
    oProp.append(oProp1)

    properties = tuple(oProp)
    oDoc = LeenoUtils.getDocument()
   
    uno.invoke(
        oDoc.Events, 'replaceByName',
        (nEvento, uno.Any('[]com.sun.star.beans.PropertyValue', properties))
    )
    return

########################################################################

def assegna():
    # ~OnFocus
    # ~OnUnfocus
    # ~OnSelect
    # ~OnDoubleClick
    # ~OnRightClick
    # ~OnChange
    # ~OnCalculate
    oDoc = LeenoUtils.getDocument()

    macro_SHEET ("Elenco Prezzi", "OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.elenco_OnFocus?language=Basic&location=application")
    macro_SHEET ("Elenco Prezzi", "OnSelect", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnSelect?language=Basic&location=application")

    if oDoc.getSheets().hasByName('Analisi di Prezzo'):
        macro_SHEET ("Analisi di Prezzo", "OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.analisi_OnFocus?language=Basic&location=application")
        macro_SHEET ("Analisi di Prezzo", "OnSelect", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnSelect?language=Basic&location=application")

    macro_SHEET ("COMPUTO", "OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnFocus?language=Basic&location=application")
    macro_SHEET ("COMPUTO", "OnSelect", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnSelect?language=Basic&location=application")

    if oDoc.getSheets().hasByName("VARIANTE"):
        macro_SHEET ("VARIANTE", "OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnFocus?language=Basic&location=application")
        macro_SHEET ("VARIANTE", "OnSelect", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnSelect?language=Basic&location=application")

    if oDoc.getSheets().hasByName("CONTABILITA"):
        macro_SHEET ("CONTABILITA", "OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnFocus?language=Basic&location=application")
        macro_SHEET ("CONTABILITA", "OnSelect", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnSelect?language=Basic&location=application")
    macro_SHEET ("S2", "OnUnfocus", "vnd.sun.star.script:UltimusFree2.Header_Footer.set_header_auto?language=Basic&location=application")
    macro_SHEET ("S1", "OnUnfocus", "vnd.sun.star.script:UltimusFree2.Menu.LeenoToolbars_Vedi?language=Basic&location=application")
    # ~macro_SHEET ("S1", "OnUnfocus","service:org.giuseppe-vizziello.leeno.dispatcher?LeenoToolbars.Vedi")
    # ~OnStartApp
    # ~OnCloseApp
    # ~macro_DOC ("OnCreate", "vnd.sun.star.script:Standard.Controllo.Controlla_Esistenza_LibUltimus?language=Basic&location=document")
    macro_DOC ("OnNew", "vnd.sun.star.script:Standard.Controllo.Controlla_Esistenza_LibUltimus?language=Basic&location=document")
    # ~OnLoadFinished
    macro_DOC ("OnLoad", "vnd.sun.star.script:Standard.Controllo.Controlla_Esistenza_LibUltimus?language=Basic&location=document")
    macro_DOC ("OnPrepareUnload", "vnd.sun.star.script:UltimusFree2._variabili.autoexec_off?language=Basic&location=application")
    macro_DOC ("OnUnload", "vnd.sun.star.script:UltimusFree2.Lupo_0.Svuota_Globale?language=Basic&location=application")
    macro_DOC ("OnSave", "vnd.sun.star.script:UltimusFree2.Menu.pyleeno_bak0?language=Basic&location=application")
    # ~macro_DOC ("OnSave", "service:org.giuseppe-vizziello.leeno.dispatcher?LeenoBasicBridge.bak0")
    # ~OnSaveDone
    # ~OnSaveFailed
    macro_DOC ("OnSaveAs", "vnd.sun.star.script:UltimusFree2.Lupo_0.Svuota_Globale?language=Basic&location=application")
    # ~OnSaveAsDone
    macro_DOC ("OnSaveAsFailed", "vnd.sun.star.script:UltimusFree2._variabili.autoexec?language=Basic&location=application")
    # ~macro_DOC ("OnCopyTo", "vnd.sun.star.script:UltimusFree2.Lupo_0.Svuota_Globale?language=Basic&location=application")
    # ~OnCopyToDone
    # ~OnCopyToFailed
    macro_DOC ("OnFocus", "vnd.sun.star.script:UltimusFree2.Menu.computo_OnFocus?language=Basic&location=application")
    # ~macro_DOC ("OnUnfocus", "vnd.sun.star.script:UltimusFree2.PY_bridge.ScriviNomeDocumentoPrincipale?language=Basic&location=application")
    # ~OnPrint
    # ~OnViewCreated
    # ~OnPrepareViewClosing
    # ~OnViewClosed
    # ~OnModifyChanged
    # ~OnTitleChanged
    # ~OnVisAreaChanged
    # ~OnModeChanged
    # ~OnStorageChanged

########################################################################


def pulisci():
    oDoc = LeenoUtils.getDocument()
    lista_fogli = oDoc.Sheets.ElementNames

    eventi = oDoc.CurrentController.ActiveSheet.Events.ElementNames
    eventi_doc = oDoc.Events.ElementNames
    for nome in lista_fogli:
        for ev in eventi:
            
            macro_SHEET(nome, ev, '')
    for ev in eventi_doc:
        macro_DOC(ev, '')
    macro_DOC ("OnNew", "vnd.sun.star.script:Standard.Controllo.Controlla_Esistenza_LibUltimus?language=Basic&location=document")
    macro_DOC ("OnLoad", "vnd.sun.star.script:Standard.Controllo.Controlla_Esistenza_LibUltimus?language=Basic&location=document")

########################################################################
