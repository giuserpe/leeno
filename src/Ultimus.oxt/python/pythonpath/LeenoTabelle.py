'''
Gestione tabelle Pesi e Misure
*** DA REVISIONARE ***

con questa versione LibreOffice va in freeze e il file tabelle.ods rimane aperto più volte

'''
# ~import uno
# ~import pyleeno as PL

# ~import LeenoUtils
# ~import LeenoDialogs as DLG
# ~import DocUtils


# ~class tabelle_dati:
    
    # ~def __init__ (self):
        # ~""" Class initialiser """
        # ~pass
    # ~try:
        # ~oDoc.dispose()
    # ~except:
        # ~pass

    # ~filename = uno.fileUrlToSystemPath(PL.LeenO_path()) + '/data/tabelle.ods'
    # ~oDoc = DocUtils.loadDocument(filename, Hidden=True)

    # ~tabelle = {'Tondo per c.a.': 'tondo_ca',
        # ~'Reti elettrosaldate' : 'reti_els',
        # ~'Pesi specifici' : 'pesi_specifici',
        # ~'Categorie' : 'categorie'
        # ~}

    # ~fogli = oDoc.getSheets().getElementNames()
    
    # ~psm = LeenoUtils.getComponentContext().ServiceManager
    # ~dp = psm.createInstance('com.sun.star.awt.DialogProvider')

    # ~oDlg = dp.createDialog(
        # ~"vnd.sun.star.script:UltimusFree2.Dialog_tabelle?language=Basic&location=application"
    # ~)
    # ~combobox = oDlg.getControl('ComboBox1')
    # ~sString = combobox
    # ~sString.Text = fogli[0]
    # ~listbox = oDlg.getControl('ListBox1')

    # ~def init():
        # ~combobox = tabelle_dati.combobox
        # ~listbox = tabelle_dati.listbox

        # ~combobox.addItems(tabelle_dati.fogli, 0)
        # ~nome = combobox.Text

        # ~oSheet = tabelle_dati.oDoc.Sheets.getByName(nome)
        # ~Dati = oSheet.getCellRangeByName(tabelle_dati.tabelle[nome]).DataArray
        # ~lista = []
        # ~for el in Dati:
            # ~lista.append(list(el)[0])
        # ~listbox.addItems(lista, 0)
        # ~# tabelle_dati.oDlg.execute()
        # ~# try:
            # ~# tabelle_dati.oDlg.endExecute()
        # ~# except:
            # ~# pass
        # ~# DLG.chi(combobox.Text)
        # ~return

    # ~def compila():
        # ~# return
        # ~combobox = tabelle_dati.combobox
        # ~listbox = tabelle_dati.listbox
        # ~listbox.removeItems(0, len(listbox.Items))
        
        # ~nome = combobox.Text
        # ~# DLG.mri(combobox)
        # ~# return
        
        # ~oSheet = tabelle_dati.oDoc.Sheets.getByName(nome)
        # ~Dati = oSheet.getCellRangeByName(tabelle_dati.tabelle[nome]).DataArray

        # ~lista = []
        # ~for el in Dati:
            # ~lista.append(list(el)[0])

        # ~listbox.addItems(lista, 0)
        # ~# tabelle_dati.oDlg.endExecute()
        # ~tabelle_dati.oDlg.execute()
        # ~return
    # ~def ok():
        # ~tabelle_dati.oDlg.endExecute()
        # ~tabelle_dati.oDlg.dispose()
        # ~tabelle_dati.oDoc.dispose()
        # ~return
    # ~pass

# ~def tabella_compila():
    # ~tabelle_dati.compila()
# ~def tabella_ok():
    # ~tabella_dati.ok()
