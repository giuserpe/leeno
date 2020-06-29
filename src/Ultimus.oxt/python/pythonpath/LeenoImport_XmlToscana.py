import Dialogs
import pyleeno as PL

from io import StringIO
import xml.etree.ElementTree as ET

import LeenoImport


def parseXML(root, defaultTitle):
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
        codice = articolo.attrib['codice'].split('_')[1]

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
        desc = LeenoImport.fixParagraphSize(desc)

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

