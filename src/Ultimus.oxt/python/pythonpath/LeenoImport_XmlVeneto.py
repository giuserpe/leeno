import LeenoImport
import xml.etree.ElementTree as ET
import LeenoDialogs as DLG

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

    # ~desc = root.items()[1][1]
    prezzario = root.get('desc')
    prezzario = root.get('ver')
    # ~articoli = contenuto.findall('Articolo')

    # ~desc = settore.attrib['desc']
    artList = {}
    superCatList = {}
    catList = {}

    settori = root.findall('settore')
    for settore in settori:
        # estrae supercategoria e categoria
        codiceSuperCat = settore.attrib['cod']
        superCat = settore.attrib['desc']

        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat
        
        capitoli = settore.findall('capitolo')
        for capitolo in capitoli:
            codiceCat = capitolo.attrib['cod'] 
            Cat = capitolo.attrib['desc'] 
            
            if not codiceCat in catList:
                catList[codiceCat] = Cat

            paragrafi = capitolo.findall('paragrafo')
            for paragrafo in paragrafi:
                codiceCat = paragrafo.attrib['cod'] 
                # ~Cat = paragrafo.attrib['desc'] 
                
                if not codiceCat in catList:
                    catList[codiceCat] = paragrafo.find('sint').text
                    # ~DLG.chi(catList)
                    # ~return


            # ~voce = articolo.find('voce').text
            # ~if voce is None:
                # ~voce = ''
            # ~art = articolo.find('articolo').text
            # ~if art is None:
                # ~art = ''
            # ~desc = voce + '\n- ' + art

    # ~DLG.chi(superCatList)
    return

    for articolo in articoli:
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
        desc = voce + '\n- ' + art

        # un po' di pulizia nel testo
        desc = desc.replace('\t', ' ').replace('Ã¨', 'è').replace(
        'Â°', '°').replace('Ã', 'à').replace(' $', '')
        while '  ' in desc:
            desc = desc.replace('  ', ' ')
        while '\n\n' in desc:
            desc = desc.replace('\n\n', '\n')

        um = articolo.find('um').text.split('(')[-1][: -1]
        prezzo = articolo.find('prezzo').attrib['valore']

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
        mdo = float(articolo.find('mo').text) / 100
        oneriSic = float(articolo.find('sicurezza').text)

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
