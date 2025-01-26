import re
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
    #ripulisce il testo da caratteri non stampabili
    # data = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', data)
    data = PL.clean_text(data)
    
    # alcuni files sono degli XML-SIX con un bug
    # consistente nella mancata dichiarazione del namespace
    # quindi lo aggiungiamo a manina nei dati
    # ~def PRT():
        # ~return "PRT"
    def trovaTipo(xmlText):
        Dati = {
            'PRT="https://prezzariollpp.regione.toscana.it/PrezzarioRT.xsd"': 'PRT',
            'EASY="https://prezzariollpp.regione.toscana.it/prezzario.xsd"': 'EASY',
            'EASY="https://prezzario.regione.campania.it/prezzario.xsd"': 'EASY',
            'EASY="https://prezzario.calabriallpp.it/prezzario.xsd"': 'EASY',
        }
        # controlla se il file è di tipo conosciuto
        # la Regione Toscana ha l'abitudine di cambiare i tags dei formati XML
        for pattern, tipo in Dati.items():
            if pattern in xmlText:
                return tipo
        return None
 
    if trovaTipo(data) == 'EASY':
        if data.find("xmlns:EASY=") < 0:
            pattern = "<EASY:Prezzario>"
            pos = data.find(pattern) + len(pattern) - 1
            data = data[:pos] + ' xmlns:EASY="mynamespace"' + data[pos:]

    if trovaTipo(data) == 'PRT' or trovaTipo(data) == None:
        if data.find("xmlns:PRT=") < 0:
            pattern = "<PRT:Prezzario>"
            pos = data.find(pattern) + len(pattern) - 1
            data = data[:pos] + ' xmlns:PRT="mynamespace"' + data[pos:]

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    
    try:
        root = LeenoImport.stripXMLNamespaces(data)
    except Exception as e:
        Dialogs.Exclamation(
        Title="ERRORE xmlns:",
        Text=f"""{e}

Il namespace dichiarato nel file fornito
non è incluso nel set di importazione.

Ti invitiamo ad inviare una copia del file XML
al team di LeenO, affinché il formato possa essere
supportato nella prossima versione del programma""")
# ~ in caso di errore, aggiungere il namespace richiesto in Dati{}
        return

    intestazione = root.find('intestazione')
    autore = intestazione.attrib['autore']
    # versione = intestazione.attrib['versione']

    dettaglio = intestazione.find('dettaglio')
    anno = dettaglio.attrib['anno']
    area = dettaglio.attrib['area']

    copyright = intestazione.find('copyright')
    ccType = copyright.attrib['tipo']
    ccDesc = copyright.attrib['descrizione']

    # crea il titolo dell' EP
    # ~titolo = "Elenco prezzi - " + autore + " - " + area + " - anno " + anno
    titolo = "Elenco prezzi - " + area + " - anno " + anno + "\n"\
    + "Copyright: " + ccType + " - " + ccDesc

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
        try:
            superCat = articolo.find('tipo').text
        except AttributeError:
            superCat = articolo.find('livello1').text
        try:
            cat = articolo.find('capitolo').text
        except AttributeError:
            cat = articolo.find('livello2').text

        # li inserisce se necessario nelle liste
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat
        if not codiceCat in catList:
            catList[codiceCat] = cat

        try:
            voce = articolo.find('voce').text
        except AttributeError:
            voce = articolo.find('livello3').text
        if voce is None:
            voce = ''
        
        try:
            art = articolo.find('articolo').text
        except AttributeError:
            art = articolo.find('livello4').text
        if art is None:
            art = ''
        desc = voce + '\n' + art

        # giochino per garantire che la prima stringa abbia una lunghezza minima
        # in modo che LO formatti correttamente la cella
        # ~desc = LeenoImport.fixParagraphSize(desc)

        # un po' di pulizia nel testo
        # desc = PL.clean_text (desc)

        um = articolo.find('um').text
        prezzo = articolo.find('prezzo').text

        # in 'sto benedetto prezzario ci sono numeri (grandi) con un punto
        # per separare le migliaia OLTRE al punto per separare i decimali
        # quindi... se trovo più di un punto decimale, devo eliminare i primi
        if prezzo is not None:
            if '.' not in prezzo:
                prezzo = prezzo + '.0'
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
