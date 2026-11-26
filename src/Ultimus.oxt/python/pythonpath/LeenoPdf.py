########################################################################
# LeenO - Computo Metrico
# Copyright (C) Giuseppe Vizziello - supporto@leeno.org
# Licenza LGPL http://www.gnu.org/licenses/lgpl.html
########################################################################

import os
import unohelper
from com.sun.star.awt import XActionListener

import LeenoUtils
import LeenoContab
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
    'npLibretto',
    'npRegistro',
    'npSAL',
    'npCdP',
    'cbElencoPrezzi',
    'cbComputoMetrico',
    'cbCostiManodopera',
    'cbQuadroEconomico',
    'cbLibretto',
    'cbRegistro',
    'cbSAL',
    'cbCdP',
)

# numero di tavola proposto se il valore manca (documenti o config precedenti)
_EXPORTNPDEFAULT = {'npLibretto': '5', 'npRegistro': '6', 'npSAL': '7', 'npCdP': '8'}

# atti contabili esportabili per singolo SAL: (nome file, chiave del documento in LeenoSettings)
_DOCUMENTI_SAL = {
    'Libretto': ('LibrettoMisure', 'LIBRETTO'),
    'Registro': ('RegistroContabilita', 'REGISTRO'),
    'SAL': ('SAL', 'SAL'),
    'CdP': ('CertificatoPagamento', 'CDP'),
}

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
        pageStyles = cDoc.StyleFamilies.getByName('PageStyles')
        style = pageStyles.getByName(pageStyle)
        SheetUtils.copyPageStyle(nDoc, style)

    # cover is OK, copy to new document
    pos = nDoc.Sheets.Count
    nDoc.Sheets.importSheet(cDoc, copertina, pos)

    # if page has a print area, copy it too...
    nDoc.Sheets[pos].PageStyle = sheet.PageStyle
    if len(sheet.PrintAreas) > 0:
        nDoc.Sheets[pos].PrintAreas = sheet.PrintAreas

    # replaces all placeholders: dati di S2 (con ripiego sulle impostazioni di lavoro) e docSubst
    settings = LeenoSettings.valori_segnaposto(
        oDoc, nomefile=LeenoSettings.nome_file_documento(oDoc), extra=docSubst)
    SheetUtils.replaceText(nDoc.Sheets[pos], settings)

    # close cover document and return
    cDoc.close(False)
    del cDoc
    return True

def prepareHeaderFooter(oDoc, docSubst, chiave):
    """
    Prepara il dizionario di intestazione e piè di pagina del documento `chiave`
    (vedi LeenoSettings._DOCUMENTI_STAMPA) sostituendo i segnaposto.
    docSubst: segnaposto legati al documento prodotto ([OGGETTO], [NUMERO_DOCUMENTO], [SAL]).
    [PAGINA] e [PAGINE] restano: diventano campi numerici in SheetUtils.pdfExport.
    """
    printSettings, dummy = LeenoSettings.loadPrintSettings(oDoc)
    testi = LeenoSettings.impostazioni_documento(printSettings, chiave)
    valori = LeenoSettings.valori_segnaposto(
        oDoc, nomefile=LeenoSettings.nome_file_documento(oDoc), extra=docSubst)
    # la paginazione richiede passaggi in piu': restano segnaposto fino a SheetUtils.pdfExport
    for segnaposto in ('[PAGINA]', '[PAGINE]'):
        valori.pop(segnaposto, None)
    testata = {psKey: LeenoSettings.sostituisci_segnaposto(psVal, valori) for psKey, psVal in testi.items()}
    # [FIRME] nel piè di pagina: la riga delle firme va da sola in fondo, senza sovrapporsi al resto
    testata.update(LeenoSettings.piede_con_firme(testata))
    return testata

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


def PdfDialog(es, listaSal=()):
    """
    Crea la finestra di dialogo per l'esportazione dei documenti in PDF
    (UltimusFree2.DlgEsportazionePdf) precompilata con le impostazioni es.
    listaSal: numeri di SAL registrati (stringhe); se vuota gli atti contabili
    vengono disabilitati. Restituisce (oDlg, listener): il listener va tenuto
    vivo fino alla chiusura.
    """
    ctx = LeenoUtils.getComponentContext()
    dp = ctx.ServiceManager.createInstance('com.sun.star.awt.DialogProvider')
    oDlg = dp.createDialog(_DLG_ESPORTAZIONE_PDF)
    oModel = oDlg.getModel()

    base = LeenoUtils.LeenO_path() + '/python/pythonpath/'
    oModel.getByName('imgPdf').ImageURL = base + 'Icons-Big/pdf.png'
    # for nome, icona in (('btnOk', 'Icons-24x24/ok.png'),
    #                     ('btnAnnulla', 'Icons-24x24/cancel.png'),
    #                     ('btnCartella', 'Icons-16x16/folder.png')):
    #     oModel.getByName(nome).ImageURL = base + icona
    #     oModel.getByName(nome).ImageAlign = 0

    for key in _EXPORTSETTINGSITEMS:
        if key in _EXPORTSETTINGSCHECKS:
            oModel.getByName(key).State = 1 if es.get(key, False) else 0
        else:
            oModel.getByName(key).Text = str(es.get(key, _EXPORTNPDEFAULT.get(key, '')))

    # atti contabili: disponibili solo se esiste almeno un SAL
    oList = oDlg.getControl('listSal')
    if len(listaSal) > 0:
        oList.addItems(tuple(listaSal), 0)
        oList.selectItem(listaSal[-1], True)
    else:
        for tipo in _DOCUMENTI_SAL:
            oModel.getByName('np' + tipo).Enabled = False
            oModel.getByName('cb' + tipo).Enabled = False
            oModel.getByName('cb' + tipo).State = 0
        oModel.getByName('listSal').Enabled = False

    percorso = Dialogs.getDefaultPath()
    oModel.getByName('pathEdit').Text = percorso
    oModel.getByName('pathEdit').HelpText = percorso

    azione = _SfogliaCartella(oDlg)
    oDlg.getControl('btnCartella').addActionListener(azione)
    return oDlg, azione


def _leggiImpostazioniExport(oDlg, es):
    """
    Legge dal dialogo i valori elencati in _EXPORTSETTINGSITEMS.
    I controlli disabilitati conservano il valore precedente di es.
    """
    oModel = oDlg.getModel()
    res = {}
    for key in _EXPORTSETTINGSITEMS:
        oItem = oModel.getByName(key)
        if not oItem.Enabled:
            res[key] = es.get(key, _EXPORTNPDEFAULT.get(key, False if key in _EXPORTSETTINGSCHECKS else ''))
        elif key in _EXPORTSETTINGSCHECKS:
            res[key] = bool(oItem.State)
        else:
            res[key] = oItem.Text
    return res


def _pagineFoglio(oSheet, area=None):
    """Numero di pagine di stampa del foglio o, se indicata, dell'area."""
    breaks = oSheet.RowPageBreaks
    if area is None:
        return len(breaks) - 1
    # interruzioni interne all'area + 1
    return len([b for b in breaks if area.StartRow < b.Position <= area.EndRow]) + 1


def _esportaPdfFoglio(destFolder, nTavola, nomeFoglio, nomeFile, chiave, area=None, oDoc=None, nSal=None):
    """
    Esporta in formato PDF il foglio nomeFoglio nella cartella destFolder.
    chiave: documento (intestazione e piè di pagina da LeenoSettings); nSal: numero del SAL
    per gli atti contabili. Se area (CellRangeAddress) e' indicata esporta solo quell'area.
    """

    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.Sheets.getByName(nomeFoglio)

    # lancia l'export
    nDoc = str(nTavola)
    baseName = nDoc + '-' if nDoc else ''
    destPath = os.path.join(destFolder, baseName + nomeFile)
    selection = [oSheet, ]
    docSubst = {
        '[OGGETTO]': LeenoSettings.titolo_documento(chiave),
        '[NUMERO_DOCUMENTO]': str(nTavola),
        '[SAL]': '' if nSal is None else str(nSal),
    }
    headerFooter = prepareHeaderFooter(oDoc, docSubst, chiave)
    docSubst['[PAGINE]'] = _pagineFoglio(oSheet, area)
    printAreas = {nomeFoglio: area} if area is not None else None
    SheetUtils.pdfExport(oDoc, selection, destPath, headerFooter,
                         lambda oDoc, nDoc: prepareCover(oDoc, nDoc, docSubst), printAreas)


def PdfElencoPrezzi(destFolder, nTavola):
    """Esporta in formato PDF il foglio Elenco Prezzi nella cartella specificata."""
    _esportaPdfFoglio(destFolder, nTavola, 'Elenco Prezzi', 'ElencoPrezzi.pdf', 'ELENCO')


def PdfComputoMetrico(destFolder, nTavola):
    """Esporta in formato PDF il foglio COMPUTO nella cartella specificata."""
    _esportaPdfFoglio(destFolder, nTavola, 'COMPUTO', 'ComputoMetrico.pdf', 'COMPUTO')


def PdfAttoSal(destFolder, nTavola, tipo, nSal, oDoc=None):
    """
    Esporta in formato PDF un atto contabile (tipo: Libretto, Registro, SAL, CdP)
    del SAL nSal. Ritorna False se l'area del SAL non esiste.
    """
    if oDoc is None:
        oDoc = LeenoUtils.getDocument()
    nomeFoglio, prefisso, colFine = LeenoContab.AREE_SAL[tipo]
    nomeArea = prefisso + str(nSal)
    if not oDoc.NamedRanges.hasByName(nomeArea):
        return False
    area = oDoc.NamedRanges.getByName(nomeArea).ReferredCells.RangeAddress
    area.EndColumn = colFine
    nomeFile, chiave = _DOCUMENTI_SAL[tipo]
    _esportaPdfFoglio(destFolder, nTavola, nomeFoglio,
                      f'{nomeFile}_SAL{nSal}.pdf', chiave, area, oDoc, nSal)
    return True


def MENU_Pdf():
    """Mostra la finestra di dialogo ed esegue l'esportazione PDF dei documenti selezionati."""

    oDoc = LeenoUtils.getDocument()
    es = loadExportSettings(oDoc)

    oDlg, azione = PdfDialog(es, LeenoContab.ultimo_sal())
    try:
        esito = oDlg.execute()
        if esito == 1:
            es = _leggiImpostazioniExport(oDlg, es)
            destFolder = oDlg.getModel().getByName('pathEdit').Text
            nSal = oDlg.getControl('listSal').getSelectedItem()
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

    # atti contabili del SAL scelto
    mancanti = []
    if nSal:
        for tipo in _DOCUMENTI_SAL:
            if es['cb' + tipo]:
                try:
                    if not PdfAttoSal(destFolder, es['np' + tipo], tipo, nSal, oDoc):
                        mancanti.append(LeenoSettings.titolo_documento(_DOCUMENTI_SAL[tipo][1]))
                except Exception as _exc:
                    LeenoUtils.log_errore('MENU_Pdf: esportazione ' + tipo, _exc)
                    mancanti.append(LeenoSettings.titolo_documento(_DOCUMENTI_SAL[tipo][1]))
    if mancanti:
        Dialogs.Exclamation(Title='Esportazione PDF',
                            Text='Per il SAL n. ' + str(nSal) + ' non sono stati esportati: '
                            + ', '.join(mancanti) + '.')
