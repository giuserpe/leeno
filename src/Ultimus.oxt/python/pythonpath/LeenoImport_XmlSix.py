"""
    LeenO - modulo parser XML per il formato XML-SIX
"""
from io import StringIO
import xml.etree.ElementTree as ET

import pyleeno as PL

import Dialogs
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

    prezzario = root.find('prezzario')
    descrizioni = prezzario.findall('przDescrizione')
    lingue = {}
    lingua = None
    lingueEstese = {'it': 'Italiano', 'de': 'Deutsch', 'en': 'English', 'fr': 'Français', 'es': 'Español'}
    try:
        for desc in descrizioni:
            lingua = desc.attrib['lingua']
            lExt = lingueEstese.get(lingua, lingua)
            lingue[lExt] = lingua
    except KeyError:
        pass

    if len(lingue) > 1:
        lingue['Tutte'] = 'tutte'
        lingue['Annulla'] = 'annulla'
        lingua = Dialogs.MultiButton(
            Icon="Icons-Big/question.png",
            Title="Scelta lingue",
            Text="Il file fornito è un prezzario multilinguale\n\nSelezionare la lingua da importare\noppure 'Tutte' per ottenere un prezzario multilinguale",
            Buttons=lingue)
        # se si chiude la finestra il dialogo ritorna 'None'
        # lo consideriamo come un 'Annulla'
        if lingua is None:
            lingua = 'annulla'
        if lingua == 'tutte':
            lingua = None
    else:
        lingua = None

    if lingua == 'annulla':
        return None

    # da qui, se lingua == None importa tutte le lingue presenti
    # altrimenti solo quella specificata

    # estrae il nome
    # se richiesta un lingua specifica, estrae quella
    # altrimenti le estrea tutte e le concatena una dopo l'altra
    nome = ""
    if lingua is None:
        nome = descrizioni[0].attrib['breve']
        for desc in range(1, len(descrizioni)):
            nome = nome + '\n' + descrizioni[desc].attrib['breve']
    else:
        for desc in descrizioni:
            if desc.attrib['lingua'] == lingua:
                nome = desc.attrib['breve']
                break

    # legge le unità di misura
    # siccome ci interessano soli i simboli e non il resto
    # non serve il processing per le lingue
    units = {}
    umList = prezzario.findall('unitaDiMisura')
    for um in umList:
        attr = um.attrib
        try:
            if 'simbolo' in attr:
                sym = attr['simbolo']
            else:
                sym = attr['udmId']
            umId = attr['unitaDiMisuraId']
            units[umId] = sym
        except KeyError:
            pass

    # se ci sono le categorie SOA, estrae prima quelle
    # in versione a una o più lingue a seconda del file
    # e di come viene richiesta la cosa
    # attualmente non servono, ma non si sa mai...
    categorieSOA = {}
    catList = root.findall('categoriaSOA')
    for cat in catList:
        attr = cat.attrib
        try:
            soaId = attr['soaId']
            soaCategoria = attr['soaCategoria']
            descs = cat.findall('soaDescrizione')
            text = ""
            for desc in descs:
                descAttr = desc.attrib
                try:
                    descLingua = descAttr['lingua']
                except KeyError:
                    descLingua = None
                if lingua is None or descLingua is None or lingua == descLingua:
                    text = text + descAttr['breve'] + '\n'
            if text != "":
                text = text[: -len('\n')]

            categorieSOA[soaCategoria] = {'soaId': soaId, 'descrizione': text}
        except KeyError:
            pass

    # infine tiriamo fuori il prezzario
    # utilizziamo le voci 'true' come base per le descrizioni
    # aggiungendo quelle delle voci specializzate '.a, .b...'
    baseCodice = ''
    baseTextBreve = ''
    baseTextEstesa = ''
    artList = {}

    productList = prezzario.findall('prodotto')
    for product in productList:
        attr = product.attrib
        try:
            # il codice del prodotto
            codice = attr['prdId']

            # se c'è, estrae l'unità di misura
            if 'unitaDiMisuraId' in attr:
                um = attr['unitaDiMisuraId']
                # converte l'unità dal codice al simbolo
                um = units.get(um, "*SCONOSCIUTA*")
            else:
                # unità non trovata - la lascia in bianco
                um = ""

            # il prezzo
            # alcune voci non hanno il campo del prezzo essendo
            # voci principali composte da sottovoci
            # le importo comunque, lasciando il valore nullo
            try:
                prezzo = float(product.find('prdQuotazione').attrib['valore'])
            except Exception:
                prezzo = ""
            if prezzo == 0:
                prezzo = ""

            # percentuale manodopera
            mdo = ""
            try:
                mdo = float(product.find('incidenzaManodopera').text) / 100
            except Exception:
                mdo = ""
            if mdo == 0:
                mdo = ""

            # oneri sicurezza
            try:
                oneriSic = float(attr['onereSicurezza'])
            except Exception:
                oneriSic = ""
            if oneriSic == 0:
                oneriSic = ""

            # per le descrizioni, come sempre... processing a seconda
            # della lingua disponibile / richiesta
            descs = product.findall('prdDescrizione')
            textBreve = ""
            textEstesa = ""
            for desc in descs:
                descAttr = desc.attrib
                try:
                    descLingua = descAttr['lingua']
                except KeyError:
                    descLingua = None
                if lingua is None or descLingua is None or lingua == descLingua:
                    if 'breve' in descAttr:
                        textBreve = textBreve + descAttr['breve'] + '\n'
                    if 'estesa' in descAttr:
                        textEstesa = textEstesa + descAttr['estesa'] + '\n'
            if textBreve != "":
                textBreve = textBreve[: -len('\n')]
            if textEstesa != "":
                textEstesa = textEstesa[: -len('\n')]

            # controlla se la voce è una voce 'base' o una specializzazione
            # della voce base. Il campo 'voce' è totalmente inaffidabile, quindi
            # consideriamo come 'base' delle voci a valore nullo e le azzeriamo
            # ad ogni nuova base e/o cambio di numerazione
            base = (prezzo == "")
            if not base and not codice.startswith(baseCodice):
                baseTextBreve = ""
                baseTextEstesa = ""

            # se voce base, tiene buona la descrizione e la salva anche come base
            # per le prossime voci (estese). Per funzionare, questo presuppone che
            # nell' XML le voci estese seguano quella base in ordine, e che tutte le voci
            # siano correttamente etichettate come voce = 'true' se voci base
            # probabilmente si può fare di meglio...
            if base:
                baseTextBreve = textBreve + '\n'
                baseTextEstesa = textEstesa + '\n'
                baseCodice = codice

            if not base and codice.startswith(baseCodice):
                textBreve = baseTextBreve + textBreve
                textEstesa = baseTextEstesa + textEstesa

            # utilizza solo la descrizione lunga per LeenO
            if len(textBreve) > len(textEstesa):
                desc = textBreve
            else:
                desc = textEstesa

            # giochino per garantire che la prima stringa abbia una lunghezza minima
            # in modo che LO formatti correttamente la cella
            desc = LeenoImport.fixParagraphSize(desc)

            # compone l'articolo e lo mette in lista
            artList[codice] = {
                'codice': codice,
                'desc': desc,
                'um': um,
                'prezzo': prezzo,
                'mdo': mdo,
                'sicurezza': oneriSic
            }
        except KeyError:
            pass

    superCatList = {}
    catList = {}

    return {
        'titolo': defaultTitle,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }
