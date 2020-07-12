import LeenoUtils
import SheetUtils
import Dialogs
import os

def PdfDialog():
    # dimensione verticale dei checkbox == dimensione bottoni
    dummy, hItems = Dialogs.getButtonSize('', Icon="Icons-24x24/settings.png")

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
                Dialogs.CheckBox(Id="cbElencoPrezzi", Label="Elenco prezzi", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbComputoMetrico", Label="Computo metrico", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbCostiManodopera", Label="Costi manodopera", FixedHeight=hItems),
                Dialogs.Spacer(),
                Dialogs.CheckBox(Id="cbQuadroEconomici", Label="Quadro economico", FixedHeight=hItems),
            ]),
            Dialogs.Spacer(),
            Dialogs.VSizer(Items=[
                Dialogs.Button(Id="btnElencoPrezzi", Icon="Icons-24x24/settings2.png"),
                Dialogs.Spacer(),
                Dialogs.Button(Id="btnComputoMetrico", Icon="Icons-24x24/settings2.png"),
                Dialogs.Spacer(),
                Dialogs.Button(Id="btnCostiManodopera", Icon="Icons-24x24/settings2.png"),
                Dialogs.Spacer(),
                Dialogs.Button(Id="btnQuadroEconomici", Icon="Icons-24x24/settings2.png"),
            ]),
        ]),
        Dialogs.Spacer(),
        Dialogs.FixedText(Text='Cartella di destinazione:'),
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


def PdfElencoPrezzi(destFolder):

    oDoc = LeenoUtils.getDocument()


    # se la copertina è presente, la copia in coda
    #cover = SheetUtils.tempCopySheet(oDoc, 'Copertina')
    #ep = SheetUtils.tempCopySheet(oDoc, 'Elenco Prezzi')

    cover = oDoc.Sheets.getByName('Copertina')
    ep = oDoc.Sheets.getByName('Elenco Prezzi')

    # copia l'elenco prezzi a fine fogli

    # seleziona entrambi

    # lancia l'export
    destPath = os.path.join(destFolder, 'test.pdf')
    print(f"Export to '{destPath}' file")
    selection = [cover, ep, ]
    SheetUtils.pdfExport(oDoc, selection, destPath)

    #oDoc.Sheets.removeByName(cover.Name)
    #oDoc.Sheets.removeByName(ep.Name)


def MENU_Pdf():
    # crea ed esegue il dialogo
    dlg = PdfDialog()

    # se premuto "annulla" non fa nulla
    if dlg.run() < 0:
        return

    # estrae la path
    destFolder = dlg['pathEdit'].getPath()

    # controlla se selezionato elenco prezzi
    if dlg['cbElencoPrezzi'].getState():
        PdfElencoPrezzi(destFolder)

