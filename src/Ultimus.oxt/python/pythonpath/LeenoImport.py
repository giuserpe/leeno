"""
    LeenO - modulo di importazione prezzari
"""
import os
import threading
import uno
from com.sun.star.beans import PropertyValue

from io import StringIO
import xml.etree.ElementTree as ET

import LeenoUtils
import pyleeno as PL
import LeenoDialogs as DLG

import SheetUtils
import LeenoSheetUtils

import Dialogs
import DocUtils

import LeenoImport_XmlSix
import LeenoImport_XmlToscana
import LeenoImport_XmlSardegna
import LeenoImport_XmlLiguria
import LeenoImport_XmlVeneto
import LeenoImport_XmlBasilicata
import LeenoImport_XmlLombardia


def fixParagraphSize(txt):
    '''
    corregge il paragrafo della descrizione negli elenchi prezzi
    in modo che LibreOffice calcoli correttamente l'altezza della cella
    '''
    minLen = 130
    splitted = txt.split('\n')
    if len(splitted) > 1:
        spl0 = splitted[0]
        spl1 = splitted[1]
        if len(spl0) + len(spl1) < minLen:
            dl = minLen - len(spl0) - len(spl1)
            spl0 = spl0 + dl * " "
            txt = spl0 + '\n' + spl1
            for t in splitted[2:]:
                txt += '\n' + t
    return txt


def stripXMLNamespaces(data):
    '''
    prende una stringa contenente un file XML
    elimina i namespaces dai dati
    e ritorna il root dell' XML
    '''
    it = ET.iterparse(StringIO(data))
    for _, el in it:
        # strip namespaces
        _, _, el.tag = el.tag.rpartition('}')
    return it.root


def findXmlParser(xmlText):
    '''
    fa un pre-esame del contenuto xml della stringa fornita
    per determinare se si tratta di un tipo noto
    (nel qual caso fornisce un parser adatto) oppure no
    (nel qual caso avvisa di inviare il file allo staff)
    '''

    parsers = {
        'xmlns="six.xsd"': LeenoImport_XmlSix.parseXML,
        'autore="Regione Toscana"': LeenoImport_XmlToscana.parseXML,
        'autore="Regione Calabria"': LeenoImport_XmlToscana.parseXML,
        'autore="Regione Campania"': LeenoImport_XmlToscana.parseXML,
        'autore="Regione Sardegna"': LeenoImport_XmlSardegna.parseXML,
        'autore="Regione Liguria"': LeenoImport_XmlLiguria.parseXML,
        'rks=': LeenoImport_XmlVeneto.parseXML,
        '<pdf>Prezzario_Regione_Basilicata': LeenoImport_XmlBasilicata.parseXML,
        '<autore>Regione Lombardia': LeenoImport_XmlLombardia.parseXML,
        '<autore>LOM': LeenoImport_XmlLombardia.parseXML,
        'xsi:noNamespaceSchemaLocation="Parte': LeenoImport_XmlLombardia.parseXML1,
    }

    # controlla se il file è di tipo conosciuto...
    for pattern, xmlParser in parsers.items():
        if pattern in xmlText:
            # si, ritorna il parser corrispondente
            return xmlParser

    # non trovato... ritorna None
    return None

def compilaElencoPrezzi(oDoc, dati, progress):
    '''
    Scrive la pagina dell' Elenco Prezzi di un documento LeenO
    Il documento deve essere vuoto (appena creato)
    I dati DEVONO essere nel formato seguente :

        articolo = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': mdo,
            'sicurezza': oneriSic
        }
        artList = { codice : articolo, ... }

        superCatList = { codice : descrizione, ... }
        catList = { codice : descrizione, ... }

        dati = {
            'titolo': titolo,
            'superCategorie': superCatList,
            'categorie': catList,
            'articoli' : artList
        }

        progress è una progressbar già visualizzata

    '''

    # inserisce supercategorie e categorie nella lista
    # articoli, creando quindi un blocco unico
    LeenoUtils.DocumentRefresh(False)
    artList = dati['articoli']
    superCatList = dati['superCategorie']
    catList = dati['categorie']
    for codice, desc in superCatList.items():
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': '',
            'prezzo': '',
            'mdo': '',
            'sicurezza': ''
        }
    for codice, desc in catList.items():
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': '',
            'prezzo': '',
            'mdo': '',
            'sicurezza': ''
        }

    # ordina l'elenco per codice articolo
    sortedArtList = sorted(artList.items())

    # ora, per velocità di compilazione, deve creare un array
    # contenente tante tuples quanti articoli
    # ognuna con la sequenza corretta per l'inserimento nel foglio
    # (codice, desc, um, sicurezza, prezzo, mdo)
    artArray = []
    for item in sortedArtList:
        itemData = item[1]
        prezzo = itemData['prezzo']
        mdo = itemData['mdo']
        if mdo == 0:
            mdo = ''
        # ~if isinstance(prezzo, str) or isinstance(mdo, str):
            # ~mdoVal = ''
        # ~else:
            # ~mdoVal = prezzo * mdo

        mdoVal = ''

        artArray.append((
            itemData['codice'],
            itemData['desc'],
            itemData['um'],
            itemData['sicurezza'],
            prezzo,
            mdo,
            mdoVal
        ))

    numItems = len(artArray)
    numColumns = len(artArray[0])

    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = dati['titolo']
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oSheet.getCellByPosition(1, 0).String = dati['titolo']
    oSheet.getRows().insertByIndex(4, numItems)

    # riga e colonna di partenza del blocco da riempire
    startRow = 4
    startCol = 0

    # fissa i limiti della progress
    progress.setLimits(0, numItems)
    progress.setValue(0)

    item = 0
    step = 100
    while item < numItems:
        progress.setValue(item)
        sliced = artArray[item:item + step]
        num = len(sliced)
        oRange = oSheet.getCellRangeByPosition(
            startCol,
            startRow + item,
            # l'indice parte da 0
            startCol + numColumns - 1,
            startRow + item + num - 1)
        oRange.setDataArray(sliced)

        item += step

    return True


def MENU_ImportElencoPrezziXML():
    '''
    Routine di importazione di un prezzario XML in tabella Elenco Prezzi
    '''
    LeenoUtils.DocumentRefresh(False)

    filename = Dialogs.FileSelect('Scegli il file XML da importare', '*.xml')
    if filename is None:
        return

    Dialogs.Info(
        Title ='Informazione',
        Text = '''L'operazione potrebbe richiedere del tempo e
LibreOffice potrebbe sembrare bloccato.''')

    # se il file non contiene un titolo, utilizziamo il nome del file
    # come titolo di default
    defaultTitle = os.path.split(filename)[1]

    # legge il file XML in una stringa
    with open(filename, 'r', errors='ignore', encoding="utf8") as file:
      data = file.read()

    # cerca un parser adatto
    xmlParser = findXmlParser(data)

    # se non trovato, il file è di tipo sconosciuto
    if xmlParser is None:
        Dialogs.Exclamation(
            Title = "File sconosciuto",
            Text = """Il file fornito sembra essere di un tipo sconosciuto.
Verrà tentata un'importazione utilizzando il formato XPWE."""
        )
        import LeenoImport_XPWE as LXPWE
        LXPWE.MENU_XPWE_import(filename)

    else:

        # ~ try:
        dati = xmlParser(data, defaultTitle)
        # ~ except Exception:
            # ~ Dialogs.Exclamation(
               # ~ Title="Errore nel file XML",
               # ~ Text=f"Riscontrato errore nel file XML\n'{filename}'\nControllarlo e riprovare")
            # ~ return

        # il parser può gestirsi l'errore direttamente, nel qual caso
        # ritorna None ed occorre uscire
        if dati is None:
            return

        # creo nuovo file di computo
        oDoc = PL.creaComputo(0)

        PL.GotoSheet("Elenco Prezzi")
        LeenoUtils.DocumentRefresh(False)

        # visualizza la progressbar
        progress = Dialogs.Progress(
            Title="Importazione prezzario",
            Text="Compilazione prezzario in corso")
        progress.show()

        # compila l'elenco prezzi
        compilaElencoPrezzi(oDoc, dati, progress)

        # si posiziona sul foglio di computo appena caricato
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oDoc.CurrentController.setActiveSheet(oSheet)

        dest = os.path.split(filename)[0] + '\\' + '_'.join(oSheet.getCellRangeByName('B1').String.split('\n')[0].split(' ')) + '.ods'
        dest = ''.join(dest.split('-_'))
        # ~dest = filename[0:-4]+ '.ods'
        # salva il file col nome del titolo
        PL.salva_come(dest)

        # nasconde la progressbar
        progress.hide()

    try:
        oSheet
    except:
        oDoc = LeenoUtils.getDocument()
        oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
        oSheet.getRows().insertByIndex(3, 1)

    # aggiunge informazioni nel foglio
    oSheet.getCellByPosition(11, 3).String = ''
    oSheet.getCellByPosition(12, 3).String = ''
    oSheet.getCellByPosition(13, 3).String = ''
    oSheet.getCellByPosition(0, 3).String = '000'
    oSheet.getCellByPosition(1, 3).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.
N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''

    # ~ if Dialogs.YesNoDialog(Title='AVVISO!',
    # ~ Text='''Vuoi ripulire le descrizioni dagli spazi e dai salti riga in eccesso?

# ~ L'operazione potrebbe richiedere del tempo e
# ~ LibreOffice potrebbe sembrare bloccato!

# ~ Vuoi procedere comunque?''') == 0:
        # ~pass
    # ~ else:
        # ~ oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        # ~ SR = oRange.StartRow + 1
        # ~ ER = oRange.EndRow
        # ~ oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, SR, 1, ER -1))
        # ~ PL.sistema_cose()
        # ~ oDoc.CurrentController.select(
            # ~ oDoc.createInstance(
                # ~ "com.sun.star.sheet.SheetCellRanges"))  # unselect

    # evidenzia e struttura i capitoli
    # ~ PL.struttura_Elenco()
    oSheet.getCellRangeByName('E2').Formula = '=COUNT(E:E) & " prezzi"'

    PL._gotoCella(0, 3)
    LeenoUtils.DocumentRefresh(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    try:
        dati
    except:
        Dialogs.Exclamation(
            Title = "IMPORTAZIONE ANNULLATA!",
            Text = """IMPORTAZIONE ANNULLATA!

Formato del file non riconosciuto.

Ti invitiamo a inviare una copia del file
al team di LeenO, affinché il formato possa essere
supportato nella prossima versione del programma."""
        )
        return
    Dialogs.Info(
        Title =f'Importate {len(dati["articoli"])} voci di Elenco Prezzi',
        Text = '''
ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili
        sul sito istituzionale ufficiale prima di accedere al Prezzario.'''
        )
    PL.dlg_donazioni()

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
    lista_articoli = []
    nome = oSheet.getCellByPosition(2, 0).String
    try:
        test = SheetUtils.uFindStringCol('ATTENZIONE!', 5, oSheet) + 1
    except:
        test = 5
    fine = SheetUtils.getUsedArea(oSheet).EndRow + 1
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
    for el in range(test, SheetUtils.getLastUsedRow(oSheet) + 1):
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

    PL.struttura_Elenco()
    DLG.MsgBox('Conversione eseguita con successo!', '')
    PL.autoexec()

########################################################################

def MENU_emilia_romagna():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Emilia Romagna
    
    *** IMPRATICABILE: IL FILE DI ORIGINE È PARECCHIO DISORDINATO ***
    
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = SheetUtils.getLastUsedRow(oSheet) + 1
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3 and \
                oSheet.getCellByPosition(3, i).Type.value != 'EMPTY':
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
    LeenoUtils.DocumentRefresh(True)

########################################################################
def MENU_umbria():
    '''
    Adatta la struttura del prezzario rilasciato dalla regione Umbria
    
    Il risultato ottenuto va inserito in Elenco Prezzi.
    '''
    LeenoUtils.DocumentRefresh(False)
    SheetUtils.MENU_unisci_fogli()
    PL.GotoSheet("unione_fogli")
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.Columns.insertByIndex(4, 1)
    oSheet.getCellByPosition(4, 0).String = 'Incidenza MdO\n%'
    fine = SheetUtils.getLastUsedRow(oSheet) + 1
    for i in range(1, fine):
        oSheet.getCellByPosition(0, i).String = oSheet.getCellByPosition(0, i).String
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 1 and \
        len(oSheet.getCellByPosition(0, i).String.split('.')) == 2:
            pass
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 3 and \
        oSheet.getCellByPosition(3, i).Type.value != 'EMPTY':
            mdo = oSheet.getCellByPosition(5, i).Value
            prz = oSheet.getCellByPosition(3, i).Value
            oSheet.getCellByPosition(4, i).Value = mdo / prz
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 4 and \
        oSheet.getCellByPosition(3, i).Type.value == 'EMPTY':
            madre = oSheet.getCellByPosition(1, i).String
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 4 and \
        oSheet.getCellByPosition(3, i).Type.value != 'EMPTY':
            oSheet.getCellByPosition(1, i).String = madre +"\n- " + oSheet.getCellByPosition(1, i).String
            mdo = oSheet.getCellByPosition(5, i).Value
            prz = oSheet.getCellByPosition(3, i).Value
            oSheet.getCellByPosition(4, i).Value = mdo / prz
#> codice di Lorenzo Vagnarelli
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 5 and \
        oSheet.getCellByPosition(3, i).Type.value == 'EMPTY':
            madre = oSheet.getCellByPosition(1, i).String
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 5 and \
        oSheet.getCellByPosition(3, i).Type.value != 'EMPTY':
            oSheet.getCellByPosition(1, i).String = madre +"\n- " + oSheet.getCellByPosition(1, i).String
            mdo = oSheet.getCellByPosition(5, i).Value
            prz = oSheet.getCellByPosition(3, i).Value
            oSheet.getCellByPosition(4, i).Value = mdo / prz
#< codice di Lorenzo Vagnarelli
    oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, 0, 1, fine))
    PL.sistema_cose()
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect
    LeenoUtils.DocumentRefresh(True)

########################################################################


def MENU_ValdAosta():
    '''Non va: spesso il file di origine non è ordinato'''
    oDoc = LeenoUtils.getDocument()
    LeenoUtils.DocumentRefresh(False)
    oSheet = oDoc.CurrentController.ActiveSheet
    oSheet.Columns.insertByIndex(3, 1)
    primo_nome = oSheet.Name
    oSheet.Name = 'COMPUTO'
    LeenoSheetUtils.MENU_elimina_righe_vuote()
    oSheet.Name = oSheet.Name
    fine = SheetUtils.getLastUsedRow(oSheet) + 1
    for i in range(0, fine):
        if len(oSheet.getCellByPosition(0, i).String.split('.')) == 2:
            madre = oSheet.getCellByPosition(1, i).String
        elif len(oSheet.getCellByPosition(0, i).String.split('.')) == 3:
            oSheet.getCellByPosition(1, i).String = madre + '\n- ' + oSheet.getCellByPosition(1, i).String

########################################################################

def MENU_Piemonte():

    Dialogs.Info(
        Title = "Importazione Prezzario Regione Piemonte",
        Text = '''
NOTA: Questo processo di importazione richiede la selezione di un file
        del set, in formato XLS, così come scaricato dal sito istituzionale
        della Regione Piemonte e posto in un'unica cartella del disco.

'''
# ~https://www.regione.piemonte.it/web/temi/protezione-civile-difesa-suolo-opere-pubbliche/opere-pubbliche

        )
    import glob
    import DocUtils

    artList = {}
    superCatList = {}
    catList = {}

    filename = Dialogs.FileSelect('Scegli un file del set da importare...', '*.xls')
    if filename in ('Cancel', '', None):
        return
    dest = os.path.dirname(filename).split('.')[0].split('/')
    directory = os.path.dirname(filename).split('.')[0].replace('\\', '/') + '/'
    files = glob.glob(directory + '*.xls')

    progress = Dialogs.Progress(Title='Caricamento dei dati in corso...', Text="Progressione")
    n = 0
    progress.setLimits(n, len(files))
    progress.show()
    progress.setValue(0)

    for el in files:
        oDoc = DocUtils.loadDocument(el, Hidden=True)
        oSheet = oDoc.getSheets().getByIndex(0)
        lrow = SheetUtils.getLastUsedRow(oSheet) + 1

        for i in range(2, lrow):
            codice = oSheet.getCellByPosition(1, i).String
            desc = oSheet.getCellByPosition(2, i).String
            um = oSheet.getCellByPosition(3, i).String
            prezzo = oSheet.getCellByPosition(4, i).Value
            if prezzo == 0:
                prezzo = ''
            mdo = oSheet.getCellByPosition(6, i).Value
            if mdo == 0:
                mdo = ''

            if len(codice.split('.')) == 3 and prezzo == '':
                madre = desc
            elif len(codice.split('.')) == 4:
                desc = madre + '\n- ' + desc

            artList[codice] = {
                'codice': codice,
                'desc': desc,
                'um': um,
                'prezzo': prezzo,
                'mdo': mdo,
                'sicurezza': ''
            }
        n += 1
        progress.setValue(n)
    progress.hide()

    titolo = 'Piemonte ' + oSheet.getCellRangeByName('A1').String.split('\n')[0]
    dati = {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }

    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.CurrentController.ActiveSheet
    if(len(oDoc.getURL()) == 0 and
       SheetUtils.getUsedArea(oSheet).EndColumn == 0 and
       SheetUtils.getUsedArea(oSheet).EndRow == 0):
        oDoc.close(True)

    # creo nuovo file di computo
    oDoc = PL.creaComputo(0)
    PL.GotoSheet("Elenco Prezzi")
    LeenoUtils.DocumentRefresh(False)

    # visualizza la progressbar
    progress = Dialogs.Progress(
        Title="Importazione prezzario",
        Text="Compilazione prezzario in corso")
    progress.show()

    # compila l'elenco prezzi
    compilaElencoPrezzi(oDoc, dati, progress)
    # si posiziona sul foglio di computo appena caricato
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oDoc.CurrentController.setActiveSheet(oSheet)

    # messaggio di ok
    Dialogs.Ok(Text=f'Importate {len(dati["articoli"])} voci\ndi elenco prezzi')

    # nasconde la progressbar
    progress.hide()
    # aggiunge informazioni nel foglio
    oSheet.getCellByPosition(11, 3).String = ''
    oSheet.getCellByPosition(12, 3).String = ''
    oSheet.getCellByPosition(13, 3).String = ''
    oSheet.getCellByPosition(0, 3).String = '000'
    oSheet.getCellByPosition(1, 3).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.
N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''

    if Dialogs.YesNoDialog(Title='AVVISO!',
    Text='''Vuoi ripulire le descrizioni dagli spazi e dai salti riga in eccesso?

L'operazione potrebbe richiedere del tempo e
LibreOffice potrebbe sembrare bloccato!

Vuoi procedere comunque?''') == 0:
        pass
    else:
        oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        SR = oRange.StartRow + 1
        ER = oRange.EndRow
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, SR, 1, ER -1))
        PL.sistema_cose()
        oDoc.CurrentController.select(
            oDoc.createInstance(
                "com.sun.star.sheet.SheetCellRanges"))  # unselect

    oSheet.getCellRangeByName('E2').Formula = '=COUNT(E:E) & " prezzi"'
    PL._gotoCella(0, 3)

    # salva il file
    PL.salva_come(directory + titolo.replace(' ', '_') + '.ods')
    # evidenzia e struttura i capitoli
    PL.struttura_Elenco()

    LeenoUtils.DocumentRefresh(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    Dialogs.Info(
        Title = "Importazione eseguita con successo!",
        Text = '''
ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili
        sul sito istituzionale ufficiale prima di accedere al Prezzario.'''
        )
    return

########################################################################

def MENU_fuf():
    '''
    Traduce un particolare formato DAT usato in falegnameria - non c'entra un tubo con LeenO.
    E' solo una cortesia per un amico.
    '''
    filename = Dialogs.FileSelect('Scegli il file DAT da importare', '*.dat')
    riga = []
    try:
        f = open(filename, 'r')
    except TypeError:
        return
    ordini = []
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
        SheetUtils.getLastUsedColumn(oSheet),
        SheetUtils.getLastUsedRow(oSheet)).Columns.OptimalWidth = True
    oDoc.CurrentController.freezeAtPosition(0, 1)
    PL._gotoCella(0, 1)
    oDoc.CurrentController.ShowGrid = True
    oSheet.getCellRangeByName('A1:F1').CellStyle = 'Accent 3'
    return
    PL.comando('Copy')

    ctx = LeenoUtils.getComponentContext()
    desktop = LeenoUtils.getDesktop()
    oFrame = desktop.getCurrentFrame()
    dispatchHelper = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.DispatchHelper', ctx)
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
    dispatchHelper.executeDispatch(oFrame, '.uno:InsertContents', '', 0, properties)

    oDoc.CurrentController.select( oSheet.getCellRangeByPosition(0, 1, 5, SheetUtils.getLastUsedRow(oSheet) + 1))
    PL.ordina_col(3)
    oDoc.CurrentController.select(
        oDoc.createInstance("com.sun.star.sheet.SheetCellRanges"))  # unselect

########################################################################

def MENU_FVG():
    '''
    Importa Prezzario Friuli Venezia Giulia secondo schema tabellare
    rilasciato a gennaio 2023
    '''
    filename = Dialogs.FileSelect('Scegli il Prezzario Friuli Venezia Giulia da importare...', '*.xlsx')
    if filename in ('Cancel', '', None):
        return


    progress = Dialogs.Progress(
        Title="Importazione prezzario",
        Text="Compilazione prezzario in corso")

    oDoc = DocUtils.loadDocument(filename, Hidden=True)
    progress.show()
    # ~oDoc = LeenoUtils.getDocument()
    # ~filename = uno.fileUrlToSystemPath(oDoc.getURL())
    oSheet = oDoc.CurrentController.ActiveSheet
    fine = SheetUtils.getLastUsedRow(oSheet) +1
    
    oDoc.CurrentController.select(
        oSheet.getCellRangeByPosition(2, 3, 2, fine))
    PL.sistema_cose()
    artList = {}
        # ~artList[codice] = {
            # ~'codice': codice,
            # ~'desc': desc,
            # ~'um': um,
            # ~'prezzo': prezzo,
            # ~'mdo': mdo,          #%
            # ~'sicurezza': oneriSic
        # ~}
    superCatList = {}
    catList = {}
    titolo = oSheet.getCellRangeByName('A1').String
    for i in range(3, fine):
        codice = oSheet.getCellByPosition(0, i).String
        desc = oSheet.getCellByPosition(1, i).String
        if oSheet.getCellByPosition(2, i).Type.value == 'TEXT':
            madre = oSheet.getCellByPosition(2, i).String
        um = ''
        prezzo = ''
        mdo = ''
        oneriSic = ''
        if oSheet.getCellByPosition(2, i).Type.value == 'EMPTY' and \
        oSheet.getCellByPosition(6, i).Type.value == 'VALUE':
            desc = madre + '\n- ' + oSheet.getCellByPosition(1, i).String
        if oSheet.getCellByPosition(2, i).Type.value == 'TEXT' and \
        oSheet.getCellByPosition(6, i).Type.value == 'VALUE':
            madre = oSheet.getCellByPosition(2, i).String
            desc = oSheet.getCellByPosition(1, i).String + '\n' + madre

        if oSheet.getCellByPosition(6, i).Type.value == 'VALUE':
            um = oSheet.getCellByPosition(5, i).String
            prezzo = oSheet.getCellByPosition(6, i).Value
            mdo = oSheet.getCellByPosition(8, i).Value / 100
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': mdo,
            'sicurezza': oneriSic
        }
    dati = {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
        }
    # creo nuovo file di computo
    # ~oDoc = PL.creaComputo(0)
    orig = oDoc.getURL()
    dest0 = orig[0:-5] + '.ods'

    orig = uno.fileUrlToSystemPath(PL.LeenO_path() + '/template/leeno/Computo_LeenO.ots')
    dest = uno.fileUrlToSystemPath(dest0)

    PL.shutil.copyfile(orig, dest)
    PL._gotoDoc(dest)  # vado sul nuovo file

    progress.hide()
    oDoc = LeenoUtils.getDocument()
    progress.show()
    
    LeenoUtils.DocumentRefresh(False)

    # compila l'elenco prezzi
    compilaElencoPrezzi(oDoc, dati, progress)

    # si posiziona sul foglio di computo appena caricato
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oDoc.CurrentController.setActiveSheet(oSheet)

    progress.hide()

    # messaggio di ok
    Dialogs.Ok(Text=f'Importate {len(dati["articoli"])} voci\ndi elenco prezzi')

    # aggiunge informazioni nel foglio
    # ~oSheet.getRows().insertByIndex(3, 1)
    oSheet.getCellByPosition(11, 3).String = ''
    oSheet.getCellByPosition(12, 3).String = ''
    oSheet.getCellByPosition(13, 3).String = ''
    oSheet.getCellByPosition(0, 3).String = '000'
    oSheet.getCellByPosition(1, 3).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.
N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''

    if Dialogs.YesNoDialog(Title='AVVISO!',
    Text='''Vuoi ripulire le descrizioni dagli spazi e dai salti riga in eccesso?

L'operazione potrebbe richiedere del tempo e
LibreOffice potrebbe sembrare bloccato!

Vuoi procedere comunque?''') == 0:
        pass
    else:
        oRange = oDoc.NamedRanges.elenco_prezzi.ReferredCells.RangeAddress
        SR = oRange.StartRow + 1
        ER = oRange.EndRow
        oDoc.CurrentController.select(oSheet.getCellRangeByPosition(1, SR, 1, ER -1))
        PL.sistema_cose()
        oDoc.CurrentController.select(
            oDoc.createInstance(
                "com.sun.star.sheet.SheetCellRanges"))  # unselect
    # evidenzia e struttura i capitoli
    PL.struttura_Elenco()
    oSheet.getCellRangeByName('E2').Formula = '=COUNT(E:E) & " prezzi"'
    PL._gotoCella(0, 3)
    LeenoUtils.DocumentRefresh(True)
    LeenoSheetUtils.adattaAltezzaRiga(oSheet)
    Dialogs.Info(
        Title = "Importazione eseguita con successo!",
        Text = '''
ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili
        sul sito istituzionale ufficiale prima di accedere al Prezzario.'''
        )
    return
