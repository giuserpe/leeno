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
    titolo = root.get('desc') #+ ' - ' + root.get('ver')
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
                    voce = paragrafo.find('estesa').text
                    try:
                        paragrafo.find('tipologia').text
                        if paragrafo.find('tipologia').text == 'Manodopera':
                            mdo = 1
                        else:
                            mdo = ''
                    except:
                        mdo = ''
                    prezzi = paragrafo.findall('prezzi')
                    for el in prezzi[0]:
                        art = el.text
                        if voce == None: voce = ''
                        if voce in art:
                            desc = art
                        else:
                            desc = voce + '\n- ' + art

                        # un po' di pulizia nel testo
                        desc = desc.replace('\t', ' ').replace('Ã¨', 'è'
                        ).replace('Â°', '°').replace('Ã', 'à').replace(
                        ' $', '').replace('#13;', ' ').replace('\n \n', '\n')
                        while '  ' in desc:
                            desc = desc.replace('  ', ' ')
                        while '\n\n' in desc:
                            desc = desc.replace('\n\n', '\n')
                        
                        
                        codice = el.attrib['cod']
                        um = el.attrib['umi']
                        prezzo = float(el.attrib['val'])
                        if mdo == '':
                            try:
                                mdo = float(el.attrib['man']) / 100
                            except:
                                pass
                        # ~oneriSic = float(el.attrib['sicurezza'])

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
