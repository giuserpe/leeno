import os
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
    cfg = LeenoConfig.Config()
    data = DocUtils.loadDataBlock(oDoc, 'ImpostazioniExport')
    if data is None or len(data) == 0:
        data = cfg.readBlock('ImpostazioniExport', True)
    return data

def storeExportSettings(oDoc, es):
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

def PdfDialog():
    # dimensione verticale dei checkbox == dimensione bottoni
    #dummy, hItems = Dialogs.getButtonSize('', Icon="Icons-24x24/settings.png")
    nWidth, hItems = Dialogs.getEditBox('aa')

    # dimensione dell'icona col PDF
    imgW = Dialogs.getBigIconSize()[0] * 2

    return Dialogs.Dialog(Title='Esportazione documenti PDF',  Horz=False, CanClose=True,  Items=[
        Dialogs.HSizer(Items=[
            Dialogs.VSizer(Items=[
                Dialogs.Spacer(),
                Dialogs.ImageControl(Image='Icons-Big/pdf.png', MinWidth=imgW),
                Dialogs.Spacer(),
            ]),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Tavola'),
                Dialogs.Spacer(),
                Dialogs.Edit(Id='npElencoPrezzi', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                Dialogs.Spacer(),
                Dialogs.Edit(Id='npComputoMetrico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                Dialogs.Spacer(),
                Dialogs.Edit(Id='npCostiManodopera', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
                Dialogs.Spacer(),
                Dialogs.Edit(Id='npQuadroEconomico', Align=1, FixedHeight=hItems, FixedWidth=nWidth),
            ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.FixedText(Text='Oggetto'),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbElencoPrezzi", Label="Elenco prezzi", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbComputoMetrico", Label="Computo metrico", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbCostiManodopera", Label="Costi manodopera", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbQuadroEconomico", Label="Quadro economico", FixedHeight=hItems),
            ]),
            Dialogs.Spacer(),
        ]),
        Dialogs.Spacer(),
        Dialogs.Spacer(),
        Dialogs.FixedText(Text='Cartella di destinazione:'),
        Dialogs.Spacer(),
        Dialogs.PathControl(Id="pathEdit"),
        Dialogs.Spacer(),
        Dialogs.HSizer(Items=[
            Dialogs.Spacer(),
            Dialogs.Button(Label='Ok', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/ok.png',  RetVal=1),
            Dialogs.Spacer(),
            Dialogs.Button(Label='Annulla', MinWidth=Dialogs.MINBTNWIDTH, Icon='Icons-24x24/cancel.png',  RetVal=-1),
            Dialogs.Spacer()
        ])
    ])


def PdfElencoPrezzi(destFolder, nTavola):

    oDoc = LeenoUtils.getDocument()
    ep = oDoc.Sheets.getByName('Elenco Prezzi')

    # lancia l'export
    nDoc = str(nTavola)
    baseName = ''
    if nDoc != '' and nDoc is not None:
        baseName = nDoc + '-'
    destPath = os.path.join(destFolder, baseName + 'ElencoPrezzi.pdf')
    print(f"Export to '{destPath}' file")
    selection = [ep, ]
    docSubst = {
        '[OGGETTO]':'Elenco Prezzi',
        '[NUMERO_DOCUMENTO]': str(nTavola),
    }
    headerFooter = prepareHeaderFooter(oDoc, docSubst)
    nPages = len(ep.RowPageBreaks) - 1

    # ~ nPages = LeenoUtils.countPdfPages(destPath)
    docSubst['[PAGINE]'] = nPages
    SheetUtils.pdfExport(oDoc, selection, destPath, headerFooter, lambda oDoc, nDoc: prepareCover(oDoc, nDoc, docSubst))


def PdfComputoMetrico(destFolder, nTavola):

    oDoc = LeenoUtils.getDocument()
    ep = oDoc.Sheets.getByName('COMPUTO')

    # lancia l'export
    nDoc = str(nTavola)
    baseName = ''
    if nDoc != '' and nDoc is not None:
        baseName = nDoc + '-'
    destPath = os.path.join(destFolder, baseName + 'ComputoMetrico.pdf')
    # ~ print(f"Export to '{destPath}' file")
    selection = [ep, ]
    docSubst = {
        '[OGGETTO]':'Computo Metrico',
        '[NUMERO_DOCUMENTO]': str(nTavola),
    }
    headerFooter = prepareHeaderFooter(oDoc, docSubst)

    nPages = len(ep.RowPageBreaks) - 1
    # ~ nPages = LeenoUtils.countPdfPages(destPath)
    docSubst['[PAGINE]'] = nPages
    SheetUtils.pdfExport(oDoc, selection, destPath, headerFooter, lambda oDoc, nDoc: prepareCover(oDoc, nDoc, docSubst))


def MENU_Pdf():

    oDoc = LeenoUtils.getDocument()
    es = loadExportSettings(oDoc)

    dlg = PdfDialog()
    dlg.setData(es)

    # se premuto "annulla" non fa nulla
    if dlg.run() < 0:
        return

    es = dlg.getData(_EXPORTSETTINGSITEMS)
    storeExportSettings(oDoc, es)

    # estrae la path
    destFolder = dlg['pathEdit'].getPath()
    # ~ destFolder = 'W:\\_dwg\\ULTIMUSFREE\\_SRC'
    
    # ~ import LeenoDialogs as DLG
    # ~ DLG.chi(destFolder)
    # ~ return

    # controlla se selezionato elenco prezzi
    if dlg['cbElencoPrezzi'].getState():
        PdfElencoPrezzi(destFolder, es['npElencoPrezzi'])

    # controlla se selezionato computo metrico
    if dlg['cbComputoMetrico'].getState():
        PdfComputoMetrico(destFolder, es['npComputoMetrico'])

