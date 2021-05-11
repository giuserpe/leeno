import Dialogs
import pyleeno as PL

from io import StringIO
import xml.etree.ElementTree as ET

import codecs
import shutil
import LeenoImport
import LeenoUtils
import LeenoDialogs as DLG
import SheetUtils

from com.sun.star.sheet.CellFlags import \
    VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

def parseXML(data, defaultTitle=None):
    '''
    estrae dal file XML i dati dell'elenco prezzi
    I dati estratti avranno il formato seguente:

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
    '''
    # alcuni files sono degli XML-SIX con un bug
    # consistente nella mancata dichiarazione del namespace
    # quindi lo aggiungiamo a manina nei dati
    if data.find("xmlns:PRT=") < 0:
        pattern = "<PRT:Prezzario>"
        pos = data.find(pattern) + len(pattern) - 1
        data = data[:pos] + ' xmlns:PRT="mynamespace"' + data[pos:]
        print(data[:1000])

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    intestazione = root.find('intestazione')
    autore = intestazione.attrib['autore']
    # versione = intestazione.attrib['versione']

    dettaglio = intestazione.find('dettaglio')
    anno = dettaglio.attrib['anno']
    area = dettaglio.attrib['area']

    # copyright = intestazione.find('copyright')
    # ccType = copyright.attrib['tipo']
    # ccDesc = copyright.attrib['descrizione']

    # crea il titolo dell' EP
    titolo = "Elenco prezzi - " + autore + " - " + area + " - anno " + anno

    contenuto = root.find('Contenuto')
    articoli = contenuto.findall('Articolo')

    artList = {}
    superCatList = {}
    catList = {}

    for articolo in articoli:
        # rimuovo il 'TOS20_' dal codice
        # ~codice = articolo.attrib['codice'].split('_')[1]
        codice = articolo.attrib['codice']

        # divide il codice per ottenere i codici di supercategoria e categoria
        codiceSplit = codice.split('.')
        codiceSuperCat = codiceSplit[0]
        codiceCat = codiceSuperCat + '.' + codiceSplit[1]

        # estrae supercategoria e categoria
        superCat = articolo.find('tipo').text
        cat = articolo.find('capitolo').text

        # li inserisce se necessario nelle liste
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat
        if not codiceCat in catList:
            catList[codiceCat] = cat

        voce = articolo.find('voce').text
        if voce is None:
            voce = ''
        art = articolo.find('articolo').text
        if art is None:
            art = ''
        desc = voce + '\n' + art

        # giochino per garantire che la prima stringa abbia una lunghezza minima
        # in modo che LO formatti correttamente la cella
        # ~desc = LeenoImport.fixParagraphSize(desc)

        # un po' di pulizia nel testo
        desc = desc.replace(
            '\t', ' ').replace('\n', ' ').replace('\n\n', '\n').replace('Ã¨', 'è').replace(
                'Â°', '°').replace('Ã', 'à').replace(' $', '')
        while '  ' in desc:
            desc = desc.replace('  ', ' ')
        while '\n\n' in desc:
            desc = desc.replace('\n\n', '\n')

        um = articolo.find('um').text
        prezzo = articolo.find('prezzo').text

        # in 'sto benedetto prezzario ci sono numeri (grandi) con un punto
        # per separare le migliaia OLTRE al punto per separare i decimali
        # quindi... se trovo più di un punto decimale, devo eliminare i primi
        if prezzo is not None:
            prSplit = prezzo.split('.')
            prezzo = ''
            for p in prSplit[0:-1]:
                prezzo += p
            prezzo += '.' + prSplit[-1]
            prezzo = float(prezzo)

        analisi = articolo.find('Analisi')
        if analisi is not None:
            # se c'è l'analisi, estrae incidenza MDO e costi sicurezza da quella
            try:
                oneriSic = float(analisi.find('onerisicurezza').attrib['valore'])
            except Exception:
                oneriSic = ''

            try:
                mdo = float(analisi.find('incidenzamanodopera').attrib['percentuale']) / 100
            except Exception:
                mdo = ''
        else:
            # niente analisi, la voce non dispone di incidenza MDO e costi sicurezza
            oneriSic = ''
            mdo = ''

        # compone l'articolo e lo mette in lista
        artList[codice] = {
            'codice': codice,
            'desc': desc,
            'um': um,
            'prezzo': prezzo,
            'mdo': mdo,
            'sicurezza': oneriSic
        }

    # ritorna un dizionario contenente tutto il necessario
    # per costruire l'elenco prezzi
    return {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }


def MENU_XML_toscana_import():
    '''
    Routine di importazione di un prezzario XML-SIX in tabella Elenco Prezzi
    del template COMPUTO.
    '''
    filename = Dialogs.FileSelect('Scegli il file XML da importare', '*.xml')
    if filename is None:
        return

    # legge il file XML in una stringa
    with open(filename, 'r') as file:
      data = file.read()

    # lo analizza eliminando i namespaces
    # (che qui rompono solo le scatole...)
    it = ET.iterparse(StringIO(data))
    for _, el in it:
        # strip namespaces
        _, _, el.tag = el.tag.rpartition('}')
    root = it.root

    try:
        dati = parseXML(root)

    except Exception:
        Dialogs.Exclamation(
           Title="Errore nel file XML",
           Text=f"Riscontrato errore nel file XML\n'{filename}'\nControllarlo e riprovare")
        return

    # il parser può gestirsi l'errore direttamente, nel qual caso
    # ritorna None ed occorre uscire
    if dati is None:
        return

    # creo nuovo file di computo
    oDoc = PL.creaComputo(0)

    # visualizza la progressbar
    progress = Dialogs.Progress(
        Title="Importazione prezzario",
        Text="Compilazione prezzario in corso")
    progress.show()

    # compila l'elenco prezzi
    LeenoImport.compilaElencoPrezzi(oDoc, dati, progress)

    # si posiziona sul foglio di computo appena caricato
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    oDoc.CurrentController.setActiveSheet(oSheet)

    # messaggio di ok
    Dialogs.Ok(Text=f'Importate {len(dati["articoli"])} voci\ndi elenco prezzi')

    # nasconde la progressbar
    progress.hide()

########################################################################

def ns_ins(filename):
    '''
    Se assente, inserisce il namespace nel file XML.
    '''
    f = codecs.open(filename, 'r', 'utf-8')
    out_file = '.'.join(filename.split('.')[:-1]) + '.bak'
    of = codecs.open(out_file, 'w', 'utf-8')

    for row in f:
        nrow = row.replace(
            '<PRT:Prezzario>',
            '<PRT:Prezzario xmlns="http://www.regione.toscana.it/Prezzario" xmlns:PRT="http://www.regione.toscana.it/Prezzario/Prezzario.xsd">'
        )
        of.write(nrow)
    f.close()
    of.close()
    shutil.move(out_file, filename)

########################################################################

def MENU_XML_toscana_import_old(arg=None):
    '''
    Importazione di un prezzario XML della regione Toscana
    in tabella Elenco Prezzi del template COMPUTO.
    '''
    oDoc = PL.creaComputo(0)

    try:
        filename = Dialogs.FileSelect('Scegli il file XML da importare', '*.xml')

        if filename == None: return
    except:
        return
    if oDoc.getSheets().hasByName('COMPUTO') == False:
        if len(oDoc.getURL())==0 and \
        getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndColumn ==0 and \
        getLastUsedCell(oDoc.CurrentController.ActiveSheet).EndRow ==0:
            oDoc.close(True)

    # effettua il parsing del file XML
    tree = ET.ElementTree()
    try:
        tree.parse(filename)
    except:
        ns_ins(filename)
        tree.parse(filename)
    # ~except Exception as e:
        # ~MsgBox ("Eccezione " + str(type(e)) +
                # ~"\nMessaggio: " + str(e.args) + '\n' +
                # ~traceback.format_exc());
        # ~return
    root = tree.getroot()
    iter = tree.getiterator()

    PRT = '{' + str(iter[0].getchildren()[0]).split('}')[0].split('{')[-1] + '}' # xmlns
    # nome del prezzario
    intestazione = root.find(PRT+'intestazione')
    titolo = 'Prezzario '+ intestazione.get('autore') + ' - ' + intestazione[0].get('area') +' '+ intestazione[0].get('anno')
    licenza = intestazione[1].get('descrizione').split(':')[0] +' '+ intestazione[1].get('tipo')
    titolo = titolo + '\nCopyright: ' + licenza  + '\n\nhttp://prezzariollpp.regione.toscana.it'

    Contenuto = root.find(PRT+'Contenuto')

    voci = root.getchildren()[1]

    tipo_lista = list()
    cap_lista = list()
    lista_articoli = list()
    lista_cap = list()
    lista_subcap = list()
    # attiva la progressbar
    n = 0
    progress = Dialogs.Progress(Title='Importazione XML in corso...', Text="Lettura dati")
    progress.setLimits(0, len(voci))
    progress.setValue(n)
    progress.show()
    for el in voci:
        n += 1
        progress.setValue(n)
        if el.tag == PRT+'Articolo':
            codice = el.get('codice').replace('TOS21_', '')
            codicesp = codice.split('.')

        voce = el.getchildren()[2].text
        articolo = el.getchildren()[3].text
        if articolo == None:
            desc_voce = voce
        else:
            desc_voce = voce + ' ' + articolo
        udm = el.getchildren()[4].text

        try:
            sic = float(el.getchildren()[-1][-4].get('valore'))
        except IndexError:
            sic =''
        try:
            prezzo = float(el.getchildren()[5].text)
        except:
            prezzo = float(el.getchildren()[5].text.split('.')[0]+el.getchildren()[5].text.split('.')[1]+'.'+el.getchildren()[5].text.split('.')[2])
        try:
            mdo = float(el.getchildren()[-1][-1].get('percentuale'))/100
            mdoE = mdo * prezzo
        except IndexError:
            mdo =''
            mdoE = ''
        if codicesp[0] not in tipo_lista:
            tipo_lista.append(codicesp[0])
            cap =(codicesp[0], el.getchildren()[0].text, '', '', '', '', '')
            lista_cap.append(cap)
        if codicesp[0]+'.'+codicesp[1] not in cap_lista:
            cap_lista.append(codicesp[0]+'.'+codicesp[1])
            cap =(codicesp[0]+'.'+codicesp[1], el.getchildren()[1].text, '', '', '', '', '', '')
            lista_subcap.append(cap)
        voceel =(codice, desc_voce, udm, sic, prezzo, mdo, mdoE)
        lista_articoli.append(voceel)
    progress.hide()
# compilo ##############################################################
    oDoc = LeenoUtils.getDocument()
    oSheet = oDoc.getSheets().getByName('S2')
    oSheet.getCellByPosition(2, 2).String = titolo
    oSheet = oDoc.getSheets().getByName('Elenco Prezzi')
    flags = VALUE + DATETIME + STRING + ANNOTATION + FORMULA + OBJECTS + EDITATTR # FORMATTED + HARDATTR
    oSheet.getCellRangeByName('D1:V1').clearContents(flags)
    oDoc.getSheets().getByName('COMPUTO').IsVisible = False
    oSheet.getCellByPosition(1, 0).String = titolo
    oSheet.getCellByPosition(0, 3).String = '000'
    oSheet.getCellByPosition(1, 3).String = '''ATTENZIONE!
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale di riferimento prima di accedere al prezzario.'''
    oSheet.getCellByPosition(1, 0).CellStyle = 'EP-mezzo'
    # attiva la progressbar
    x = 0
    progress = Dialogs.Progress(Title='Importazione XML in corso...', Text="Lettura dati")
    progress.setLimits(0, 3)
    progress.setValue(x)
    progress.show()

    for el in (lista_articoli, lista_subcap, lista_cap):
        x += 1
        progress.setValue(x)
        oSheet.getRows().insertByIndex(4, len(el))
        lista_come_array = tuple(el)
        # Parametrizzo il range di celle a seconda della dimensione della lista
        colonne_lista = len(lista_come_array[1]) # numero di colonne necessarie per ospitare i dati
        righe_lista = len(lista_come_array) # numero di righe necessarie per ospitare i dati
        oRange = oSheet.getCellRangeByPosition( 0, 4, colonne_lista + 0 - 1, righe_lista + 4 - 1)
        oRange.setDataArray(lista_come_array)
        oDoc.CurrentController.setActiveSheet(oSheet)

    oSheet.getCellRangeByPosition(0, 3, 0, righe_lista + 3 - 1).CellStyle = "EP-aS"
    oSheet.getCellRangeByPosition(1, 3, 1, righe_lista + 3 - 1).CellStyle = "EP-a"
    oSheet.getCellRangeByPosition(2, 3, 7, righe_lista + 3 - 1).CellStyle = "EP-mezzo"
    oSheet.getCellRangeByPosition(5, 3, 5, righe_lista + 3 - 1).CellStyle = "EP-mezzo %"
    oSheet.getCellRangeByPosition(8, 3, 9, righe_lista + 3 - 1).CellStyle = "EP-sfondo"
    oSheet.getCellRangeByPosition(11, 3, 11, righe_lista + 3 - 1).CellStyle = 'EP-mezzo %'
    oSheet.getCellRangeByPosition(12, 3, 12, righe_lista + 3 - 1).CellStyle = 'EP statistiche_q'
    oSheet.getCellRangeByPosition(13, 3, 13, righe_lista + 3 - 1).CellStyle = 'EP statistiche'

    PL.riordina_ElencoPrezzi(oDoc)
    progress.hide()
    PL.struttura_Elenco()
    oSheet.getCellRangeByName('F2').String = 'prezzi'
    oSheet.getCellRangeByName('E2').Formula = '=COUNT(E3:E' + str(SheetUtils.getLastUsedRow(oSheet)+1) +')'
    dest = filename[0:-4]+ '.ods'
    PL.salva_come(dest)
    PL._gotoCella(0, 3)
    DLG.MsgBox('''
Importazione eseguita con successo!

ATTENZIONE:
1. Lo staff di LeenO non si assume alcuna responsabilità riguardo al contenuto del prezzario.
2. L’utente finale è tenuto a verificare il contenuto dei prezzari sulla base di documenti ufficiali.
3. L’utente finale è il solo responsabile degli elaborati ottenuti con l'uso di questo prezzario.

N.B.: Si rimanda ad una attenta lettura delle note informative disponibili sul sito istituzionale ufficiale prima di accedere al Prezzario.

    ''','ATTENZIONE!')
#~ ########################################################################
