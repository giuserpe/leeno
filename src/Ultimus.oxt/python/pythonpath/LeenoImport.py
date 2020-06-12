"""
    LeenO - modulo di importazione prezzari
"""
import logging
import threading

from xml.etree.ElementTree import ElementTree

import uno

from com.sun.star.sheet.CellFlags import (VALUE, DATETIME, STRING,
                                          ANNOTATION, FORMULA,
                                          OBJECTS, EDITATTR)

import LeenoUtils
import pyleeno as PL
import LeenoDialogs as DLG
import LeenoToolbars as Toolbars
from LeenoConfig import Config

import Dialogs


def ImportErrorDlg(msg):
    """ Generico dialogo di errore di importazione con messaggio
        DA FARE
    """
    print("Import error:", msg)


########################################################################
def MENU_importa_listino_leeno():
    '''
    @@ DA DOCUMENTARE
    '''
    importa_listino_leeno_th().start()


class importa_listino_leeno_th(threading.Thread):
    '''
    @@ DA DOCUMENTARE
    '''
    def __init__(self):
        threading.Thread.__init__(self)

    def run(self):
        importa_listino_leeno_run()


###
def importa_listino_leeno_run():
    '''
    Esegue la conversione di un listino (formato LeenO) in template Computo
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    #  giallo(16777072,16777120,16777168)
    #  verde(9502608,13696976,15794160)
    #  viola(12632319,13684991,15790335)
    lista_articoli = list()
    nome = oSheet.getCellByPosition(2, 0).String
    test = PL.uFindStringCol('ATTENZIONE!', 5, oSheet) + 1
    assembla = DLG.DlgSiNo(
        '''Il riconoscimento di descrizioni e sottodescrizioni
dipende dalla colorazione di sfondo delle righe.

Nel caso in cui questa fosse alterata, il risultato finale
della conversione potrebbe essere inatteso.

Considera anche la possibilità di recuperare il formato XML(SIX)
di questo prezzario dal sito ufficiale dell'ente che lo rilascia.

Vuoi assemblare descrizioni e sottodescrizioni?''', 'Richiesta')
    oDialogo_attesa = DLG.dlg_attesa()
    DLG.attesa().start()  # mostra il dialogo

    if assembla == 2:
        PL.colora_vecchio_elenco()
    orig = oDoc.getURL()
    dest0 = orig[0:-4] + '_new.ods'

    orig = uno.fileUrlToSystemPath(PL.LeenO_path() + '/template/leeno/Computo_LeenO.ots')
    dest = uno.fileUrlToSystemPath(dest0)

    PL.shutil.copyfile(orig, dest)
    madre = ''
    for el in range(test, PL.getLastUsedCell(oSheet).EndRow + 1):
        tariffa = oSheet.getCellByPosition(2, el).String
        descrizione = oSheet.getCellByPosition(4, el).String
        um = oSheet.getCellByPosition(6, el).String
        sic = oSheet.getCellByPosition(11, el).String
        prezzo = oSheet.getCellByPosition(7, el).String
        mdo_p = oSheet.getCellByPosition(8, el).String
        mdo = oSheet.getCellByPosition(9, el).String
        if oSheet.getCellByPosition(2,
                                    el).CellBackColor in (16777072, 16777120,
                                                          9502608, 13696976,
                                                          12632319, 13684991):
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        elif oSheet.getCellByPosition(2,
                                      el).CellBackColor in (16777168, 15794160,
                                                            15790335):
            if assembla == 2:
                madre = descrizione
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        else:
            if madre == '':
                descrizione = oSheet.getCellByPosition(4, el).String
            else:
                descrizione = madre + ' \n- ' + oSheet.getCellByPosition(
                    4, el).String
            articolo = (
                tariffa,
                descrizione,
                um,
                sic,
                prezzo,
                mdo_p,
                mdo,
            )
        lista_articoli.append(articolo)
    oDialogo_attesa.endExecute()
    PL._gotoDoc(dest)  # vado sul nuovo file
    # compilo la tabella ###################################################
    oDoc = LeenoUtils.getDocument()
    oDialogo_attesa = DLG.dlg_attesa()
    DLG.attesa().start()  # mostra il dialogo

    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = nome
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 1).String = nome

    oSheet.getRows().insertByIndex(4, len(lista_articoli))
    lista_come_array = tuple(lista_articoli)
    # Parametrizzo il range di celle a seconda della dimensione della lista
    colonne_lista = len(lista_come_array[1]
                        )  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(
        lista_come_array)  # numero di righe necessarie per ospitare i dati
    oRange = oSheet.getCellRangeByPosition(
        0,
        4,
        colonne_lista - 1,  # l'indice parte da 0
        righe_lista + 4 - 1)
    oRange.setDataArray(lista_come_array)
    oSheet.getRows().removeByIndex(3, 1)
    oDoc.CurrentController.setActiveSheet(oSheet)
    oDialogo_attesa.endExecute()
    procedo = DLG.DlgSiNo(
        '''Vuoi mettere in ordine la visualizzazione del prezzario?

Le righe senza prezzo avranno una tonalità di sfondo
diversa dalle altre e potranno essere facilmente nascoste.

Questa operazione potrebbe richiedere del tempo.''', 'Richiesta...')
    if procedo == 2:
        DLG.attesa().start()  # mostra il dialogo
        #  struttura_Elenco()
        oDialogo_attesa.endExecute()
    DLG.MsgBox('Conversione eseguita con successo!', '')
    PL.autoexec()


def MENU_XML_toscana_import():
    '''
    Importazione di un prezzario XML della regione Toscana
    in tabella Elenco Prezzi del template COMPUTO.
    '''
    oDoc = LeenoUtils.getDocument()

    DLG.MsgBox('Questa operazione potrebbe richiedere del tempo.', 'Avviso')
    PL.New_file.computo(0)

    try:
        filename = Dialogs.FileSelect('Scegli il file XML Toscana da importare', '*.xml')
        oDialogo_attesa = DLG.dlg_attesa()

        # mostra il dialogo
        DLG.attesa().start()
        if filename is None:
            return
    except Exception:
        ImportErrorDlg("Errore di importazione")
        return

    if not oDoc.getSheets().hasByName('COMPUTO'):
        if (len(oDoc.getURL()) == 0 and
                PL.getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndColumn == 0 and
                PL.getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndRow == 0):
            oDoc.close(True)

    # effettua il parsing del file XML
    tree = ElementTree()

    try:
        tree.parse(filename)
    except Exception:
        PL.ns_ins(filename)
        tree.parse(filename)
    # ~except Exception as e:
        # ~MsgBox ("Eccezione " + str(type(e)) +
        # ~"\nMessaggio: " + str(e.args) + '\n' +
        # ~traceback.format_exc());
        # ~return

    root = tree.getroot()
    iterator = tree.getiterator()

    PRT = '{' + str(iterator[0].getchildren()[0]).split('}')[0].split('{')[-1] + '}'  # xmlns
    # nome del prezzario
    intestazione = root.find(PRT + 'intestazione')
    titolo = ('Prezzario ' +
              intestazione.get('autore') +
              ' - ' +
              intestazione[0].get('area') +
              ' ' +
              intestazione[0].get('anno'))

    licenza = (intestazione[1].get('descrizione').split(':')[0] +
               ' ' +
               intestazione[1].get('tipo'))

    titolo = (titolo +
              '\nCopyright: ' +
              licenza +
              '\n\nhttp://prezzariollpp.regione.toscana.it')

    # Contenuto = root.find(PRT+'Contenuto')

    voci = root.getchildren()[1]

    tipo_lista = list()
    cap_lista = list()
    lista_articoli = list()
    lista_cap = list()
    lista_subcap = list()
    for el in voci:
        if el.tag == PRT + 'Articolo':
            codice = el.get('codice')
            codicesp = codice.split('.')

        voce = el.getchildren()[2].text
        articolo = el.getchildren()[3].text

        if articolo is None:
            desc_voce = voce
        else:
            desc_voce = voce + ' ' + articolo
        udm = el.getchildren()[4].text

        try:
            sic = float(el.getchildren()[-1][-4].get('valore'))
        except IndexError:
            sic = ''

        try:
            prezzo = float(el.getchildren()[5].text)
        except Exception:
            prezzo = float(el.getchildren()[5].text.split('.')[0] +
                           el.getchildren()[5].text.split('.')[1] +
                           '.' +
                           el.getchildren()[5].text.split('.')[2])

        try:
            mdo = float(el.getchildren()[-1][-1].get('percentuale')) / 100
            mdoE = mdo * prezzo
        except IndexError:
            mdo = ''
            mdoE = ''

        if codicesp[0] not in tipo_lista:
            tipo_lista.append(codicesp[0])
            cap = (codicesp[0], el.getchildren()[0].text, '', '', '', '', '')
            lista_cap.append(cap)
        if codicesp[0] + '.' + codicesp[1] not in cap_lista:
            cap_lista.append(codicesp[0] + '.' + codicesp[1])
            cap = (codicesp[0] +
                   '.' +
                   codicesp[1], el.getchildren()[1].text, '', '', '', '', '', '')

            lista_subcap.append(cap)
        voceel = (codice, desc_voce, udm, sic, prezzo, mdo, mdoE)
        lista_articoli.append(voceel)

    # compilo ##########################################################
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = titolo
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    flags = (VALUE + DATETIME + STRING + ANNOTATION +
             FORMULA + OBJECTS + EDITATTR)  # FORMATTED + HARDATTR
    oSheet.getCellRangeByName('D1:V1').clearContents(flags)
    oDoc.getSheets().getByName('COMPUTO').IsVisible = False
    oSheet.getCellByPosition(1, 0).String = titolo
    oSheet.getCellByPosition(2, 0).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili \
sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''

    oSheet.getCellByPosition(1, 0).CellStyle = 'EP-mezzo'
    n = 0

    for el in (lista_articoli, lista_cap, lista_subcap):
        oSheet.getRows().insertByIndex(4, len(el))
        lista_come_array = tuple(el)
        # Parametrizzo il range di celle a seconda della dimensione della lista
        # scarto_colonne = 0  # numero colonne da saltare a partire da sinistra
        # scarto_righe = 4  # numero righe da saltare a partire dall'alto
        colonne_lista = len(lista_come_array[1])  # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array)  # numero di righe necessarie per ospitare i dati
        oRange = oSheet.getCellRangeByPosition(0, 4, colonne_lista + 0 - 1, righe_lista + 4 - 1)
        oRange.setDataArray(lista_come_array)
        # ~ oSheet.getRows().removeByIndex(3, 1)
        oDoc.CurrentController.setActiveSheet(oSheet)

        oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
        oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
        oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
        oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
        oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"
        oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
        oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
        oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche'
        if n == 1:
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777120
        elif n == 2:
            oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellBackColor = 16777168
        n += 1
    # ~ set_larghezza_colonne()
    Toolbars.Vedi()
    # ~ adatta_altezza_riga('Elenco Prezzi')
    # ~ riordina_ElencoPrezzi()
    oDialogo_attesa.endExecute()
    PL.struttura_Elenco()
    oSheet.getCellRangeByName('F2').String = 'prezzi'
    oSheet.getCellRangeByName('E2').Formula = ('=COUNT(E3:E' + str(PL.getLastUsedCell(oSheet).EndRow + 1) +
                                               ')')
    dest = filename[0:-4] + '.ods'
    PL.salva_come(dest)
    DLG.MsgBox('''
Importazione eseguita con successo!

ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale prima di accedere al Prezzario.

    ''', 'ATTENZIONE!')
########################################################################


def MENU_sardegna_2019():
    '''
    @@@ DA DOCUMENTARE
    '''
    oDoc = LeenoUtils.getDocument()

    try:
        oDoc.getSheets().insertNewByName('nuova_tabella', 2)
    except Exception:
        pass

    oSheet0 = oDoc.getSheets().getByName('Worksheet')
    oSheet1 = oDoc.getSheets().getByName('nuova_tabella')
    # fine = PL.getLastUsedCell(oSheet0).EndRow + 1
    n = 1
    test1 = test2 = test3 = test4 = 1
    for i in range(2, 50):
        cod = oSheet0.getCellByPosition(0, i).String
        cods = cod.split('.')
        # ~ chi(cod)
        cod0 = cods[0]
        if test1 == 1:
            cod1 = cods[1]
            # ~ test1 =1
        if test2 == 1:
            cod2 = cods[2]
            # ~ test2 =1
        # if test3 == 1:
        #    cod3 = cods[3]
        # ~ test3 =1
        cap1 = oSheet0.getCellByPosition(1, i).String
        cap2 = oSheet0.getCellByPosition(2, i).String
        cap3 = oSheet0.getCellByPosition(3, i).String
        des = oSheet0.getCellByPosition(4, i).String
        um = oSheet0.getCellByPosition(5, i).String
        sic = oSheet0.getCellByPosition(10, i).Value
        prz = oSheet0.getCellByPosition(7, i).Value
        mdo = oSheet0.getCellByPosition(13, i).Value

        if test1 == 1:
            oSheet1.getCellByPosition(0, n).String = cod0
            oSheet1.getCellByPosition(1, n).String = cap1
            test1 = 0
        elif test2 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod0 + '.' + cod1
            oSheet1.getCellByPosition(1, n).String = cap2
            test2 = 0
        elif test3 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod0 + '.' + cod1 + '.' + cod2
            oSheet1.getCellByPosition(1, n).String = cap3
            test3 = 0
        elif test4 == 1:
            n += 1
            oSheet1.getCellByPosition(0, n).String = cod
            oSheet1.getCellByPosition(1, n).String = des
            oSheet1.getCellByPosition(2, n).String = um
            oSheet1.getCellByPosition(3, n).String = sic
            oSheet1.getCellByPosition(4, n).String = prz
            oSheet1.getCellByPosition(5, n).String = mdo
            # ~ n += 1

########################################################################


def MENU_basilicata_2020():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Basilicata
    partendo dalle colonne: CODICE	DESCRIZIONE	U. MISURA	PREZZO	MANODOPERA
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    oDoc = LeenoUtils.getDocument()
    for el in ('CAPITOLI', 'CATEGORIE', 'VOCI'):
        oSheet = oDoc.getSheets().getByName(el)
        oSheet.getRows().removeByIndex(0, 1)
    oSheet = oDoc.getSheets().getByName('CATEGORIE')
    PL._gotoSheet('CATEGORIE')
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    for i in range(0, fine):
        oSheet.getCellByPosition(1, i).String = (
            oSheet.getCellByPosition(0, i).String +
            "." +
            oSheet.getCellByPosition(1, i).String)

    oSheet.getColumns().removeByIndex(0, 1)
    oSheet = oDoc.getSheets().getByName('VOCI')
    PL._gotoSheet('VOCI')
    oSheet.getColumns().removeByIndex(0, 3)
    oSheet = oDoc.getSheets().getByName('SOTTOVOCI')
    PL._gotoSheet('SOTTOVOCI')
    oSheet.getColumns().removeByIndex(0, 4)
    PL.join_sheets()
    oSheet = oDoc.getSheets().getByName('unione_fogli')
    PL._gotoSheet('unione_fogli')
    oSheet.getRows().removeByIndex(0, 1)
    PL.ordina_col(1)
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3:
            madre = oSheet.getCellByPosition(1, i).String
        elif len(oSheet.getCellByPosition(0, i).String.split('.')) == 4:
            if oSheet.getCellByPosition(1, i).String != '':
                oSheet.getCellByPosition(1, i).String = (
                    madre +
                    "\n- " +
                    oSheet.getCellByPosition(1, i).String)
            else:
                oSheet.getCellByPosition(1, i).String = madre
            oSheet.getCellByPosition(4, i).Value = oSheet.getCellByPosition(4, i).Value / 100
    for i in reversed(range(0, fine)):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3:
            oSheet.getRows().removeByIndex(i, 1)
    oSheet.getRows().removeByIndex(0, 1)
    oSheet.getColumns().insertByIndex(3, 1)

########################################################################


def MENU_Piemonte_2019():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Piemonte
    partendo dalle colonne: Sez.	Codice	Descrizione	U.M.	Euro	Manod. lorda	% Manod.	Note
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = PL.getLastUsedCell(oSheet).EndRow + 1
    elenco = list()
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(1, i).String.split('.')) <= 2:
            cod = oSheet.getCellByPosition(1, i).String
            des = oSheet.getCellByPosition(2, i).String.replace('\n\n', '\n')
            um = ''
            eur = ''
            mdol = ''
            mdo = ''
            if oSheet.getCellByPosition(7, i).String != '':
                des = des + '\n(' + oSheet.getCellByPosition(7, i).String + ')'
            elenco.append((cod, des, um, '', eur, mdo, mdol))

        if len(oSheet.getCellByPosition(1, i).String.split('.')) == 3:
            cod = oSheet.getCellByPosition(1, i).String
            des = oSheet.getCellByPosition(2, i).String.replace(' \n\n', '')
            madre = des
            um = ''
            eur = ''
            mdol = ''
            mdo = ''
            if oSheet.getCellByPosition(7, i).String != '':
                des = des + '\n(' + oSheet.getCellByPosition(7, i).String + ')'
            # ~elenco.append ((cod, des, um, '', eur, mdo, mdol))
        if len(oSheet.getCellByPosition(1, i).String.split('.')) == 4:
            cod = oSheet.getCellByPosition(1, i).String
            des = madre
            if oSheet.getCellByPosition(2, i).String != '...':
                des = madre + '\n- ' + oSheet.getCellByPosition(2, i).String.replace('\n\n', '')
            um = oSheet.getCellByPosition(3, i).String
            eur = ''
            if oSheet.getCellByPosition(4, i).Value != 0:
                eur = oSheet.getCellByPosition(4, i).Value
            mdol = ''
            if oSheet.getCellByPosition(5, i).Value != 0:
                mdol = oSheet.getCellByPosition(5, i).Value
            mdo = ''
            if oSheet.getCellByPosition(6, i).Value != 0:
                mdo = oSheet.getCellByPosition(6, i).Value
            # ~note= oSheet.getCellByPosition(7, i).String
            elenco.append((cod, des, um, '', eur, mdo, mdol))

    try:
        oDoc.getSheets().insertNewByName('nuova_tabella', 2)
    except Exception:
        pass

    PL._gotoSheet('nuova_tabella')
    oSheet = oDoc.getSheets().getByName('nuova_tabella')
    elenco = tuple(elenco)
    oRange = oSheet.getCellRangeByPosition(0,
                                           0,
                                           # l'indice parte da 0
                                           len(elenco[0]) - 1,
                                           len(elenco) - 1)
    oRange.setDataArray(elenco)


def MENU_fuf():
    '''
    Traduce un particolare formato DAT usato in falegnameria - non c'entra un tubo con LeenO.
    E' solo una cortesia per un amico.
    '''
    filename = Dialogs.FileSelect('Scegli il file DAT da importare', '*.dat')
    riga = list()
    try:
        f = open(filename, 'r')
    except TypeError:
        return
    ordini = list()
    riga = ('Codice', 'Descrizione articolo', 'Quantità', 'Data consegna',
            'Conto lavoro', 'Prezzo(€)')
    ordini.append(riga)

    for row in f:
        art = row[:15]
        if art[0:4] not in ('HEAD', 'FEET'):
            art = art[4:]
            des = row[22:62]
            num = 1  # row[72:78].replace(' ','')
            # car = row[78:87]
            dataC = row[96:104]
            dataC = '=DATE(' + dataC[:4] + ';' + dataC[4:6] + ';' + dataC[
                6:] + ')'
            clav = row[120:130]
            prz = row[142:-1]
            riga = (art, des, num, dataC, clav, float(prz.strip()))
            ordini.append(riga)

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    lista_come_array = tuple(ordini)
    colonne_lista = len(lista_come_array[0]
                        )  # numero di colonne necessarie per ospitare i dati
    righe_lista = len(
        lista_come_array)  # numero di righe necessarie per ospitare i dati

    oRange = oSheet.getCellRangeByPosition(
        0,
        0,
        colonne_lista - 1,  # l'indice parte da 0
        righe_lista - 1)
    oRange.setFormulaArray(lista_come_array)

    oSheet.getCellRangeByPosition(
        0, 0,
        getLastUsedCell(oSheet).EndColumn,
        getLastUsedCell(oSheet).EndRow).Columns.OptimalWidth = True

    return
    copy_clip()

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext(
        'com.sun.star.frame.DispatchHelper', ctx)
    oProp = []
    oProp0 = PropertyValue()
    oProp0.Name = 'Flags'
    oProp0.Value = 'D'
    oProp1 = PropertyValue()
    oProp1.Name = 'FormulaCommand'
    oProp1.Value = 0
    oProp2 = PropertyValue()
    oProp2.Name = 'SkipEmptyCells'
    oProp2.Value = False
    oProp3 = PropertyValue()
    oProp3.Name = 'Transpose'
    oProp3.Value = False
    oProp4 = PropertyValue()
    oProp4.Name = 'AsLink'
    oProp4.Value = False
    oProp5 = PropertyValue()
    oProp5.Name = 'MoveMode'
    oProp5.Value = 4
    oProp.append(oProp0)
    oProp.append(oProp1)
    oProp.append(oProp2)
    oProp.append(oProp3)
    oProp.append(oProp4)
    oProp.append(oProp5)
    properties = tuple(oProp)
    #  _gotoCella(6,1)

    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0,
                                   properties)
    oDoc.CurrentController.select(
        oSheet.getCellRangeByPosition(0, 1, 5,
                                      getLastUsedCell(oSheet).EndRow + 1))

    ordina_col(3)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect

