########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import unohelper
from com.sun.star.awt import XActionListener

import LeenoUtils
import DocUtils
import SheetUtils
import Dialogs
import LeenoSettings
import LeenoConfig
import LeenoDialogs as DLG


_EXPORTSETTINGSITEMS = (
    'npElencoPrezzi',
    'npComputoMetrico',
    'npCostiManodopera',
    'npQuadroEconomico',
    'cbElencoPrezzi',
    'cbComputoMetrico',
    'cbCostiManodopera',
    'cbQuadroEconomico',
)

def loadExportSettings(oDoc):
    """Carica le impostazioni di esportazione PDF dal documento o dalla configurazione."""
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'ImpostazioniExport')
    if data is None or len(data) == 0:
        data = cfg.readBlock('ImpostazioniExport', True)
    return data

def storeExportSettings(oDoc, es):
    """Salva le impostazioni di esportazione PDF nel documento e nella configurazione."""
    cfg = LeenoConfig.Config()

    DocUtils.storeDataBlock(oDoc, 'ImpostazioniExport', es)
    cfg.writeBlock('ImpostazioniExport', es, True)


def prepareCover(oDoc, nDoc, docSubst):
    '''
    prepare cover page, if there's one
    copy to nDoc document and fill it's data
    return true if we've got a cover, false otherwise
    docSubst is a dictionary with additional variable replacements
    mostly used for [PAGINE], [OGGETTO] and [NUMERO_DOCUMENTO]
    which are document dependent data
    '''
    # load print settings and look for cover
    data, covers = LeenoSettings.loadPrintSettings(oDoc)
    fileCopertine = data.get('fileCopertine')
    copertina = data.get('copertina')
    if fileCopertine is None or copertina is None:
        return False
    if fileCopertine == '' or copertina == '':
        return False

    cDoc = DocUtils.loadDocument(fileCopertine)
    if cDoc is None:
        return False
    if not copertina in cDoc.Sheets:
        cDoc.close(False)
        del cDoc
        return False

    # we need to copy page style too
    sheet = cDoc.Sheets[copertina]
    pageStyle = sheet.PageStyle
    if pageStyle is not None and pageStyle != '':
        print("PAGE HAS STYLE")
        pageStyles = cDoc.StyleFamilies.getByName('PageStyles')
        style = pageStyles.getByName(pageStyle)
        SheetUtils.copyPageStyle(nDoc, style)

    # cover is OK, copy to new document
    pos = nDoc.Sheets.Count
    nDoc.Sheets.importSheet(cDoc, copertina, pos)

    # if page has a print area, copy it too...
    nDoc.Sheets[pos].PageStyle = sheet.PageStyle
    if len(sheet.PrintAreas) > 0:
        print("PAGE HAS PRINT AREA")
        nDoc.Sheets[pos].PrintAreas = sheet.PrintAreas

    # replaces all placeholders with settings ones
    settings = LeenoSettings.loadPageReplacements(oDoc)
    for key, val in docSubst.items():
        settings[key] = val
    SheetUtils.replaceText(nDoc.Sheets[pos], settings)

    # close cover document and return
    cDoc.close(False)
    del cDoc
    return True

def prepareHeaderFooter(oDoc, docSubst):
    """Prepara il dizionario di intestazione e piè di pagina sostituendo i segnaposto."""

    res = {}

    # load print settings, we need header and footer data
    printSettings, dummy = LeenoSettings.loadPrintSettings(oDoc)

    # load replacement templates
    replDict = LeenoSettings.loadPageReplacements(oDoc)
    for key, val in docSubst.items():
        replDict[key] = val

    # replace placeholders
    for psKey in ('intSx', 'intCenter', 'intDx', 'ppSx', 'ppCenter', 'ppDx'):
        if psKey in printSettings:
            psVal = printSettings[psKey]
            for replKey, replVal in replDict.items():

                # pagination needs some extra steps
                if replKey in ('[PAGINA]', '[PAGINE]'):
                    continue

                while replKey in psVal:
                    psVal = psVal.replace(replKey, replVal)
            res[psKey] = psVal

    return res

_DLG_ESPORTAZIONE_PDF = (
    'vnd.sun.star.script:UltimusFree2.DlgEsportazionePdf'
    '?language=Basic&location=application')

_EXPORTSETTINGSCHECKS = tuple(k for k in _EXPORTSETTINGSITEMS if k.startswith('cb'))


class _SfogliaCartella(unohelper.Base, XActionListener):
    """Pulsante di selezione della cartella di destinazione."""
    def __init__(self, oDlg):
        self._oDlg = oDlg

    def actionPerformed(self, ev):
        try:
            folder = Dialogs.FolderSelect()
            if folder is not None:
                oEdit = self._oDlg.getModel().getByName('pathEdit')
                oEdit.Text = os.path.join(folder, '')
                oEdit.HelpText = oEdit.Text
        except Exception as _exc:
            LeenoUtils.log_errore('PdfDialog: selezione cartella', _exc)

    def disposing(self, ev):
        return None


def PdfDialog(es):
    """
    Crea la finestra di dialogo per l'esportazione dei documenti in PDF
    (UltimusFree2.DlgEsportazionePdf) precompilata con le impostazioni es.
    Restituisce (oDlg, listener): il listener va tenuto vivo fino alla chiusura.
    """
    ctx = LeenoUtils.getComponentContext()
    dp = ctx.ServiceManager.createInstance('com.sun.star.awt.DialogProvider')
    oDlg = dp.createDialog(_DLG_ESPORTAZIONE_PDF)
    oModel = oDlg.getModel()

    base = LeenoUtils.LeenO_path() + '/python/pythonpath/'
    oModel.getByName('imgPdf').ImageURL = base + 'Icons-Big/pdf.png'
    for nome, icona in (('btnOk', 'Icons-24x24/ok.png'),
                        ('btnAnnulla', 'Icons-24x24/cancel.png'),
                        ('btnCartella', 'Icons-16x16/folder.png')):
        oModel.getByName(nome).ImageURL = base + icona
        oModel.getByName(nome).ImageAlign = 0

    for key in _EXPORTSETTINGSITEMS:
        if key in _EXPORTSETTINGSCHECKS:
            oModel.getByName(key).State = 1 if es.get(key, False) else 0
        else:
            oModel.getByName(key).Text = str(es.get(key, ''))

    percorso = Dialogs.getDefaultPath()
    oModel.getByName('pathEdit').Text = percorso
    oModel.getByName('pathEdit').HelpText = percorso

    azione = _SfogliaCartella(oDlg)
    oDlg.getControl('btnCartella').addActionListener(azione)
    return oDlg, azione


def _leggiImpostazioniExport(oDlg):
    """Legge dal dialogo i valori elencati in _EXPORTSETTINGSITEMS."""
    oModel = oDlg.getModel()
    res = {}
    for key in _EXPORTSETTINGSITEMS:
        if key in _EXPORTSETTINGSCHECKS:
            res[key] = bool(oModel.getByName(key).State)
        else:
            res[key] = oModel.getByName(key).Text
    return res


def _esportaPdfFoglio(destFolder, nTavola, nomeFoglio, nomeFile, oggetto):
    """Esporta in formato PDF il foglio nomeFoglio nella cartella destFolder."""

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName(nomeFoglio)

    # lancia l'export
    nDoc = str(nTavola)
    baseName = nDoc + '-' if nDoc else ''
    destPath = os.path.join(destFolder, baseName + nomeFile)
    selection = [oSheet, ]
    docSubst = {
        '[OGGETTO]': oggetto,
        '[NUMERO_DOCUMENTO]': str(nTavola),
    }
    headerFooter = prepareHeaderFooter(oDoc, docSubst)
    docSubst['[PAGINE]'] = len(oSheet.RowPageBreaks) - 1
    SheetUtils.pdfExport(oDoc, selection, destPath, headerFooter, lambda oDoc, nDoc: prepareCover(oDoc, nDoc, docSubst))


def PdfElencoPrezzi(destFolder, nTavola):
    """Esporta in formato PDF il foglio Elenco Prezzi nella cartella specificata."""
    _esportaPdfFoglio(destFolder, nTavola, 'Elenco Prezzi', 'ElencoPrezzi.pdf', 'Elenco Prezzi')


def PdfComputoMetrico(destFolder, nTavola):
    """Esporta in formato PDF il foglio COMPUTO nella cartella specificata."""
    _esportaPdfFoglio(destFolder, nTavola, 'COMPUTO', 'ComputoMetrico.pdf', 'Computo Metrico')


def MENU_Pdf():
    """Mostra la finestra di dialogo ed esegue l'esportazione PDF dei documenti selezionati."""

    oDoc = LeenoUtils.getDocument()
    es = loadExportSettings(oDoc)

    oDlg, azione = PdfDialog(es)
    try:
        esito = oDlg.execute()
        if esito == 1:
            es = _leggiImpostazioniExport(oDlg)
            destFolder = oDlg.getModel().getByName('pathEdit').Text
    finally:
        oDlg.dispose()

    # se premuto "annulla" non fa nulla
    if esito != 1:
        return

    storeExportSettings(oDoc, es)

    # controlla se selezionato elenco prezzi
    if es['cbElencoPrezzi']:
        PdfElencoPrezzi(destFolder, es['npElencoPrezzi'])

    # controlla se selezionato computo metrico
    if es['cbComputoMetrico']:
        PdfComputoMetrico(destFolder, es['npComputoMetrico'])
