import re
import pyleeno as PL
import LeenoImport
import xml.etree.ElementTree as ET
# ~import LeenoDialogs as DLG

# ~from com.sun.star.sheet.CellFlags import \
    # ~VALUE, DATETIME, STRING, ANNOTATION, FORMULA, HARDATTR, OBJECTS, EDITATTR, FORMATTED

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
    data = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', data)

    # alcuni files sono degli XML-SIX con un bug
    # consistente nella mancata dichiarazione del namespace
    # quindi lo aggiungiamo a manina nei dati

    # ~ if data.find("xmlns=") < 0:
        # ~ pattern = "<PRT:Prezzario>"
        # ~ pos = data.find(pattern) + len(pattern) - 1
        # ~ data = data[:pos] + ' xmlns:PRT="mynamespace"' + data[pos:]
        # ~ print(data[:1000])

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    titolo = root.find('pdf').text
    if '.pdf' in titolo:
        titolo = titolo[: -4]
    
    titolo = ' '.join(titolo.split('_'))

    artList = {}
    superCatList = {}
    catList = {}

    capitoli = root.find('capitoli') # è una sola ricorrenza

    for Capitolo in capitoli:

        # estrae supercategoria e categoria
        codiceSuperCat = Capitolo.find('codice').text
        superCat = Capitolo.find('descrizione').text.strip()

        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        categorie = Capitolo.find('categorie')
        for Categoria in categorie:
            codiceCat = codiceSuperCat + '.' + Categoria.find('codice').text
            Cat = Categoria.find('descrizione').text
            if not codiceCat in catList:
                catList[codiceCat] = Cat

            # estrae voci e sottovoci
            voci = Categoria.find('voci')
            for Voce in voci:
                voce = Voce.find('descrizione').text
                # ~hashcode = Voce.find('hashcode').text # il dato c'è, ma per ora non serve
                Scodice = codiceCat + '.' + Voce.find('codice').text
                sottovoci = Voce.find('sottovoci')
                for Sottovoce in sottovoci:
                    codice = Scodice + '.' + Sottovoce.find('codice').text
                    try:
                        desc = voce + '\n- ' + Sottovoce.find('descrizione').text
                    except:
                        desc = voce
                    um = Sottovoce.find('unitaMisura').find('codice').text.strip()
                    prezzo = float(Sottovoce.find('prezzo').text)
                    mdo = float(Sottovoce.find('manodopera').text) / 100
                    if mdo ==0:
                        mdo = ''

                    # un po' di pulizia nel testo
                    desc = PL.clean_text (desc)

                    # compone l'articolo e lo mette in lista
                    artList[codice] = {
                        'codice': codice,
                        'desc': desc,
                        'um': um,
                        'prezzo': prezzo,
                        'mdo': mdo,
                        'sicurezza': ''
                    }

    # ritorna un dizionario contenente tutto il necessario
    # per costruire l'elenco prezzi
    return {
        'titolo': titolo,
        'superCategorie': superCatList,
        'categorie': catList,
        'articoli' : artList
    }
