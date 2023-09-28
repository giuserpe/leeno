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

    # elimina i namespaces dai dati ed ottiene
    # elemento radice dell' albero XML
    root = LeenoImport.stripXMLNamespaces(data)

    voci = root.find('voci')
    voci = voci.find('voci')
    rifvoce = voci.find('riferimenti_voce')
    titolo = rifvoce.find('autore').text + ' ' + rifvoce.find('invigore').text + ' ' + rifvoce.find('anno').text

    artList = {}
    superCatList = {}
    catList = {}
    
    # prendo il suffisso da aggiungere al codice categorie
    suff = root[0][0].getchildren()[1].attrib['CMPcodifica_voce'][:8]

    voci = root[0]
    for voce in voci:
        dettaglio_voce = voce.getchildren()[1]
        # estrae supercategoria e categoria
        codiceSuperCat = suff + dettaglio_voce.attrib['codifica_I_livello_voce']
        superCat = dettaglio_voce.attrib['declaratoria_I_livello_voce']
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        codiceSuperCat = suff + dettaglio_voce.attrib['codifica_II_livello_voce']
        superCat = dettaglio_voce.attrib['declaratoria_II_livello_voce']
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        codiceSuperCat = suff + dettaglio_voce.attrib['codifica_III_livello_voce']
        superCat = dettaglio_voce.attrib['declaratoria_III_livello_voce']
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        codiceSuperCat = suff + dettaglio_voce.attrib['codifica_IV_livello_voce']
        superCat = dettaglio_voce.attrib['declaratoria_IV_livello_voce']
        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        try:
            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_V_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_V_livello_voce']
        except:
            pass

        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        try:
            codiceSuperCat = suff + dettaglio_voce.attrib['codifica_VI_livello_voce']
            superCat = dettaglio_voce.attrib['declaratoria_VI_livello_voce']
        except:
            pass

        if not codiceSuperCat in superCatList:
            superCatList[codiceSuperCat] = superCat

        codice = dettaglio_voce.attrib['CMPcodifica_voce']
        desc = dettaglio_voce.find('declaratoria_voce').text
        um = dettaglio_voce.attrib['udm_voce']
        prezzo = float(dettaglio_voce.attrib['prezzo_voce'])
        risorse = dettaglio_voce.find('risorse')

        for el in risorse:
            if el.attrib['tipologia_risorsa'] == "MANODOPERA":
                mdo = float (el.attrib['perc_importo_tipo_risorsa'])
                break
            else:
                mdo = ''
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
