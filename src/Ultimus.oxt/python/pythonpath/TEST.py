########
def debug_link():
    '''
    @@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()
    window = oDoc.getCurrentController().getFrame().getContainerWindow()
    ctx = LeenoUtils.getComponentContext()

    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)

    toolkit = create("com.sun.star.awt.Toolkit")
    msgbox = toolkit.createMessageBox(window, 0, 1, "Message", 'foo')
    link = create("com.sun.star.awt.UnoControlFixedHyperlink")
    link_model = create("com.sun.star.awt.UnoControlFixedHyperlinkModel")
    link.setModel(link_model)
    link.createPeer(toolkit, msgbox)
    link.setPosSize(35, 8, 100, 15, 15)
    link.setText("Canale Telegram")
    link.setURL("https://t.me/leeno_computometrico")
    link.setVisible(True)
    msgbox.execute()
    msgbox.dispose()
########################################################################
def debug_errore():
    '''
    @@ DA DOCUMENTARE
    '''
    #  sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
    #  return

    try:
        # sStRange = LeenoComputo.circoscriveVoceComputo(oSheet, lrow)
        LeenoComputo.circoscriveVoceComputo(oSheet, lrow)

    except Exception as e:
        #  MsgBox ("CSV Import failure exception " + str(type(e)) +
        #  ". Messaggio: " + str(e) + " args " + str(e.args) +
        #  traceback.format_exc());
        DLG.MsgBox("Eccezione " + str(type(e)) + "\nMessaggio: " + str(e.args) + '\n' + traceback.format_exc())
########################################################################
# def debug_():
#     '''
#     @@ DA DOCUMENTARE
#     '''
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     if oSheet.Name in ('COMPUTO', 'VARIANTE', 'CONTABILITA'):
#         try:
#             sRow = oDoc.getCurrentSelection().getRangeAddresses()[0].StartRow
#             eRow = oDoc.getCurrentSelection().getRangeAddresses()[0].EndRow
#         except Exception:
#             sRow = oDoc.getCurrentSelection().getRangeAddress().StartRow
#             eRow = oDoc.getCurrentSelection().getRangeAddress().EndRow
#     DLG.chi((sRow, eRow))


# def debug_():
#     oDoc = LeenoUtils.getDocument()
#     oSheet = oDoc.CurrentController.ActiveSheet
#     try:
#         oRangeAddress = oDoc.getCurrentSelection().getRangeAddresses()
#     except AttributeError:
#         oRangeAddress = oDoc.getCurrentSelection().getRangeAddress()
#     el_y = []
#     try:
#         len(oRangeAddress)
#         for el in oRangeAddress:
#             el_y.append((el.StartRow, el.EndRow))
#     except TypeError:
#         el_y.append((oRangeAddress.StartRow, oRangeAddress.EndRow))
#     lista = []
#     for y in el_y:
#         for el in range(y[0], y[1] + 1):
#             lista.append(el)
#     for el in lista:
#         oSheet.getCellByPosition(
#             7, el).Formula = '=' + oSheet.getCellByPosition(
#                 6, el).Formula + '*' + oSheet.getCellByPosition(7, el).Formula
#         oSheet.getCellByPosition(6, el).String = ''



########################################################################
def debug_syspath():
    '''
    @@ DA DOCUMENTARE
    '''
    # ~pydevd.settrace()
    # pathsstring = "paths \n"
    somestring = ''
    for i in sys.path:
        somestring = somestring + i + "\n"
    DLG.chi(somestring)
########################################################################
# def debug_():
    # '''cambio data contabilità'''
    # oDoc = LeenoUtils.getDocument()
    # #  DLG.mri(oDoc)
    # oSheet = oDoc.CurrentController.ActiveSheet
    # DLG.chi(oDoc.getCurrentSelection().CellBackColor)
    # # ~return
    # fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
    # for i in range(0, fine):
    #     if oSheet.getCellByPosition(1, i).String == 'Data_bianca':
    #         oSheet.getCellByPosition(1, i).Value = 43861.0
########################################################################
import itertools
import operator
import functools
import LeenoImport as LI

def MENU_debug():
    # ~import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()
    # ~sistema_cose()
    # ~MENU_nasconde_voci_azzerate()
    # ~oDoc = LeenoUtils.getDocument()
    # ~oSheet = oDoc.CurrentController.ActiveSheet
    # ~lrow = LeggiPosizioneCorrente()[1]

    # ~raggruppa_righe_voce(lrow, 1)
    return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    for el in reversed(range (1, lr)):
        if oSheet.getCellByPosition(2, el).CellStyle == 'comp 1-a' and \
            "'" in oSheet.getCellByPosition(2, el).Formula:
            ff = oSheet.getCellByPosition(2, el).Formula.split("'")
            oSheet.getCellByPosition(2, el).Formula = ff[0] + ff[-1][1:]

    return
    # ~LeenoSheetUtils.setAdatta()
    # ~sistema_cose()
    # ~return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    usedArea = SheetUtils.getUsedArea(oSheet)
    # ~oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows.OptimalHeight = False
    oSheet.getCellRangeByPosition(0, 0, 1023, 1048575).Rows.OptimalHeight = False
    oSheet.getCellRangeByPosition(0, 0, 1023, 1048575).Rows.Height = 1576
    DLG.mri(oSheet.getCellRangeByPosition(0, 0, usedArea.EndColumn, usedArea.EndRow).Rows)
    return
    lr = SheetUtils.getLastUsedRow(oSheet) + 1
    for el in reversed(range (1, lr)):
        if oSheet.getCellByPosition(2, el).CellStyle == 'comp 1-a' and \
            oSheet.getCellByPosition(2, el).String == '' and \
            oSheet.getCellByPosition(9, el).String == '':
            oSheet.getRows().removeByIndex(el, 1)
        elif oSheet.getCellByPosition(2, el).Type.value == 'TEXT':
            oSheet.getCellByPosition(2, el).String = '- ' + oSheet.getCellByPosition(2, el).String
    return

def MENU_debug():

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheets = oDoc.Sheets.ElementNames
    set_area_stampa()
    orig = oDoc.getURL()

    dest = '.'.join(os.path.basename(orig).split('.')[0:-1]) + '.pdf'
    orig = uno.fileUrlToSystemPath(orig)
    dir_bak = os.path.dirname(oDoc.getURL())
    # ~DelPrintArea()
    oDoc.storeToURL(dir_bak + '/' + dest, [])

    # ~DLG.chi(dir_bak)
    return


def MENU_debug():
    '''

    '''
    DelPrintArea()
    oDoc = LeenoUtils.getDocument()
    # ~oProp = []
    # ~oProp0 = PropertyValue()
    # ~oProp0.Name = 'Overwrite'
    # ~oProp0.Value = True
    # ~oProp1 = PropertyValue()
    # ~oProp1.Name = 'FilterName'
    # ~oProp1.Value = 'calc_pdf_Export'
    # ~oProp.append(oProp0)
    # ~oProp.append(oProp1)
    # ~properties = tuple(oProp)
    # ~sUrl = "file:///W:/test.pdf"
    # ~oDoc.storeToURL(sUrl, properties)

    # ~'crea proprietà e valori in filterData, che verranno passati a filterProps
    filterData = []
    filterData0 = PropertyValue()
    filterData0.Name = "Selection"
    filterData0.Value = oDoc.CurrentController.ActiveSheet
    filterData1 = PropertyValue()
    filterData1.Name = "IsAddStream"
    filterData1.Value = True
    filterData.append(filterData0)
    filterData.append(filterData1)

    # ~'crea proprietà e valori in filterProps, che verranno passati alla funzione di esportazione storeToURL
    filterProps = []
    filterProps0 = PropertyValue()
    filterProps0.Name = "FilterName"
    filterProps0.Value = "calc_pdf_Export"
    filterProps1 = PropertyValue()
    filterProps1.Name = "FilterData"
    filterProps1.Value = tuple(filterData)
    filterProps.append(filterProps0)
    filterProps.append(filterProps1)
    
    properties = tuple(filterProps)

    sUrl = "file:///W:/test.pdf"
    oDoc.storeToURL(sUrl, properties)
def MENU_debug():
    
    # ~DlgPDF()
    # ~return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    DLG.chi(len(oSheet.RowPageBreaks))
    return
    # ~testo = oSheet.getCellByPosition(0, 0).String
    # ~txt = " ".join(testo.split())
    # ~oSheet.getCellByPosition(0, 1).String = txt
    # ~DLG.chi(txt)
    import LeenoSettings
    LeenoSettings.MENU_PrintSettings()
    # ~LeenoSettings.MENU_JobSettings()
    return
    import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()
    # ~return
    
    oDoc = LeenoUtils.getDocument()
    es = LeenoPdf.loadExportSettings(oDoc)
    
    # ~DLG.chi(es)
    # ~return

    # ~dlg = PdfDlg()
    dlg = LeenoPdf.PdfDialog()
    dlg.setData(es)

    # se premuto "annulla" non fa nulla
    if dlg.run() < 0:
        return

    es = dlg.getData(_EXPORTSETTINGSITEMS)
    storeExportSettings(oDoc, es)

    # estrae la path
    # ~destFolder = dlg['pathEdit'].getPath()
    destFolder = 'W:\\_dwg\\ULTIMUSFREE\\_SRC'
    
    # ~import LeenoDialogs as DLG
    # ~DLG.chi(destFolder)
    # ~return

    # controlla se selezionato elenco prezzi
    if dlg['cbElencoPrezzi'].getState():
        PdfElencoPrezzi(destFolder, es['npElencoPrezzi'])

    # controlla se selezionato computo metrico
    if dlg['cbComputoMetrico'].getState():
        PdfComputoMetrico(destFolder, es['npComputoMetrico'])
    return

    oDoc = LeenoUtils.getDocument()

    oSheets = list(oDoc.getSheets().getElementNames())
    # ~DLG.chi(oSheets)
    # ~DLG.chi(oSheets)
    # ~nWidth, hItems = Dialogs.getEditBox('g')

    # ~Dialogs.FolderSelect()
    # ~Dialogs.ListBox(Id=None, List=oSheets, Current=None)
    # ~nWidth, hItems = Dialogs.getEditBox('aa')
    Dialogs.ListBox.setList(self, oSheets)
    return
    
    return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = LeggiPosizioneCorrente()[1]
    rigenera_voce(lrow)
    lrow = LeenoSheetUtils.prossimaVoce(oSheet, lrow, 1)

    # ~dispatchHelper.executeDispatch(oFrame, '.uno:DataSort', '', 0, properties)

    # ~For Each oSh In oSheets
        # ~If oSh.Name <> "cP_Cop" and oSh.Name <> oActiveSheet Then ' and oSh.Name <> "copyright_LeenO" Then
        # ~p = 0

        # ~'    ThisComponent.CurrentController.Select(ThisComponent.Sheets.GetByName(oSh.Name).getCellByPosition(0,0))
        # ~'    oSh.IsVisible = False
        # ~Else

            # ~Set_Area_Stampa_N("NO_messaggio")
            # ~If     oSh.Name = oActiveSheet Then
                # ~ThisComponent.CurrentController.Select(oSh.getCellRangeByposition(0,0,getLastUsedCol(oSh),getLastUsedRow(oSh)))
                # ~if msgbox (CHR$(10) &"Preferisci nascondere i colori?",36, "") = 6 Then ScriptPy("LeenoSheetUtils.py","SbiancaCellePrintArea")
                # ~unSelect 'unselect ranges 
            # ~Else
            # ~End If
        # ~End If
    # ~Next
# ~'parametri di esportazione
    # ~dim dispatcher as Object
    # ~dim document as Object
    # ~document   = ThisComponent.CurrentController.Frame
    # ~dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

    
    # ~rem ----------------------------------------------------------------------
import LeenoUtils
import LeenoEvents
import LeenoContab

import LeenoImport
from com.sun.star.sheet.GeneralFunction import MAX


def MENU_debug():
    testo = ''
    suffisso = InputBox(
        testo, t='Inserisci il suffisso per il Codice Articolo (es: "BAS22/1_").')
    if suffisso in (None, '', ' '):
        return
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = SheetUtils.getLastUsedRow(oSheet)

    # attiva la progressbar
    progress = Dialogs.Progress(Title='Operazione in corso...', Text="Progressione")
    n = 0
    progress.setLimits(n, lrow)
    progress.show()
    progress.setValue(0)
    for y in range(0, lrow):
        if oSheet.getCellByPosition(0, y).CellStyle == "EP-aS" and \
        oSheet.getCellByPosition(0, y).String != "000":
            oSheet.getCellByPosition(0, y).String = suffisso + oSheet.getCellByPosition(0, y).String
        progress.setValue(y)
    progress.hide()
    # ~DLG.chi(oSheet.getCellByPosition(0, 3).CellBackColor)
    return
    oDoc = LeenoUtils.getDocument()
    oStyleFam = oDoc.StyleFamilies
    oTablePageStyles = oStyleFam.getByName("PageStyles")
    oCpyStyle = oDoc.createInstance("com.sun.star.style.PageStyle")
    # ~oTablePageStyles.insertByName('PageStyle_REGISTRO_A4', oCpyStyle)
    stili = ("VARIANTE", "COMPUTO", "COMPUTO_print", 'Elenco Prezzi', 'CONTABILITA', 'Registro', 'SAL')
    for el in stili:
        try:
            oTablePageStyles.insertByName(el, oCpyStyle)
        except:
            pass
    return
    oDoc = LeenoUtils.getDocument()
    DLG.mri(oDoc.StyleFamilies.getByName('PageStyles')[1])
    return
    stili = oDoc.StyleFamilies.getByName('PageStyles').getElementNames()
    oDoc.getStyleFamilies().loadStylesFromURL(filename, [])

    DLG.chi(stili)

    return

    
    oColumn = oSheet.getColumns().getByIndex(23)
    DLG.chi(int(oColumn.computeFunction(MAX)))
        # ~i= LeenoSheetUtils.prossimaVoce(oSheet, i, saltaCat=True)

    return
    lrow = LeggiPosizioneCorrente()[1]

    DLG.chi(oSheet.getCellRangeByName("A1").CellBackColor) 
    return
    sistema_cose()
    return

    return
    oDoc = LeenoUtils.getDocument()
    oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
    SR = oRange.StartRow + 1
    ER = oRange.EndRow
    oSheet = oDoc.CurrentController.ActiveSheet

    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, SR, 1, ER -1))
    return
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    lrow = SheetUtils.getUsedArea(oSheet).EndRow + 1
    lCol = SheetUtils.getUsedArea(oSheet).EndColumn 
    for y in reversed(range(1, lrow)):
        if oSheet.getCellByPosition(1, y).String ==  "CAM":
            oSheet.getCellByPosition(2, y).String = "CAM - " + oSheet.getCellByPosition(2, y).String
            # ~oSheet.getRows().removeByIndex(y, 1)

    LeenoUtils.DocumentRefresh(True)


    # ~ LeenoSheetUtils.elimina_righe_vuote()
    # ~SheetUtils.MENU_unisci_fogli()
    # ~DLG.chi(loVersion())
    # ~LeenoEvents.assegna()

    return
    # ~vista_terra_terra()
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet

    # ~DLG.mri(oSheet.getColumns().getCount())
    DLG.mri(oSheet)
    return
    import LeenoPdf
    # ~LeenoPdf.MENU_Pdf()

    dlg = LeenoPdf.PdfDialog()

    return
    # ~LeenoSheetUtils.elimina_righe_vuote()
    # ~sistema_cose()
    # ~LI.MENU_emilia_romagna()
    # ~return

# ~'ACCODA PIù FILE DI CALC IN UNO SOLO
	# ~Dim DocName as object, DocUlr as string, dummy()
	# ~Doc = ThisComponent
	# ~Sheet = Doc.Sheets(0) 
	# ~sPath ="W:/_dwg/ULTIMUSFREE/elenchi/Piemonte/2022_luglio/"  ' cartella con i documenti da copiare (non ci deve essere il file destinazione con la macro
	# ~sFileName = Dir(sPath & "*.ods", 0)
# ~'	Barra_Apri_Chiudi_5(".......................Sto lavorando su "& sFileName, 0)
	# ~Do While (sFileName <> "")
		# ~c = Sheet.createCursor
		# ~c.gotoEndOfUsedArea(false)
		# ~LastRow = c.RangeAddress.EndRow + 1
		# ~DocUrl = ConvertToURL(sPath & sFileName)
# ~'on error goto errore
		# ~DocName = StarDesktop.loadComponentFromURL (DocUrl, "_blank",0, Dummy() )
		# ~Sheet1 = DocName.Sheets(0) ' questo indica l'index del foglio da copiare
		# ~c = Sheet1.createCursor
		# ~c.gotoEndOfUsedArea(false)
		# ~LastRow1 = c.RangeAddress.EndRow
	# ~'	oStart=uFindString("ATTENZIONE!", Sheet1)
	# ~'	Srow=oStart.RangeAddress.EndRow+1
	# ~Srow = 2
		# ~Range = Sheet1.getCellRangeByPosition(0, Srow,  12, LastRow1).getDataArray '(1^ colonna, 1^ riga, 10^ colonna, ultima riga)
		# ~DocName.dispose
		# ~dRange  = Sheet.getCellRangeByPosition(0, LastRow, 12, LastRow1 + LastRow-Srow)
		# ~dRange.setDataArray(Range)
		# ~sFileName = Dir()
	# ~Loop
	# ~print "fatto!"
	# ~errore:
# ~End Sub
